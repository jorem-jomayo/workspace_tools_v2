importScripts("vendor/xlsx.full.min.js", "vendor/jszip.min.js");

const PREVIEW_ROW_LIMIT = 100;
const OUTPUT_COLUMNS = [41, 2, 3, 4, 27, 32, 36, 26];

function formatExcelDate(value) {
  if (value == null || value === "") return value;
  if (typeof value === "number") {
    const date = XLSX.SSF.parse_date_code(value);
    if (date) return new Date(Date.UTC(date.y, date.m - 1, date.d));
  }
  const text = String(value).trim();
  if (!/^\d{4}-\d{2}-\d{2}(?: \d{2}:\d{2}:\d{2}(?:\.\d+)?)?$/.test(text)
    && !/^\d{1,2}\/\d{1,2}\/\d{4}(?: \d{2}:\d{2}:\d{2}(?:\.\d+)?)?$/.test(text)) return value;
  const parsed = new Date(text);
  return Number.isFinite(parsed.getTime()) ? parsed : value;
}

function dateToUsString(value) {
  if (!(value instanceof Date)) return value == null ? "" : value;
  return `${String(value.getUTCMonth() + 1).padStart(2, "0")}/${String(value.getUTCDate()).padStart(2, "0")}/${value.getUTCFullYear()}`;
}

function cellValue(sheet, row, column) {
  const cell = sheet[XLSX.utils.encode_cell({ r: row, c: column })];
  if (!cell) return "";
  return cell.v ?? "";
}

function csvEscape(value) {
  const text = value == null ? "" : String(value);
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

async function processWorkbook(file) {
  self.postMessage({ type: "phase", phase: "Reading workbook" });
  const fileData = await file.arrayBuffer();
  const workbook = XLSX.read(fileData, { type: "array", cellDates: true });
  const sheetName = workbook.SheetNames?.find((name) =>
    name.trim().toUpperCase() === "RESULT"
  );
  if (!sheetName) throw new Error("A sheet named RESULT was not found in this workbook.");
  const sheetEntry = Object.entries(workbook.Sheets || {}).find(([name]) =>
    name.trim().toUpperCase() === "RESULT"
  );
  const sheet = sheetEntry?.[1];
  if (!sheet) {
    const loadedSheets = Object.keys(workbook.Sheets || {});
    const details = loadedSheets.length ? loadedSheets.join(", ") : "none";
    throw new Error(
      `The workbook lists "${sheetName}", but its cells could not be loaded (loaded sheets: ${details}). ` +
      "Open the workbook in Excel and save a fresh .xlsx copy. If it is very large, try splitting it into smaller workbooks."
    );
  }
  if (!sheet["!ref"]) throw new Error("The RESULT sheet is empty or contains no readable cells.");
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const totalRows = range.e.r - range.s.r + 1;
  const outputSheet = { "!ref": `A1:H${totalRows}` };
  const csvParts = ['\uFEFF"Ref Code","Field Name"\r\n'];
  const columnWidths = Array(8).fill(10);
  const previewRows = [];
  let csvLineCount = 0;

  self.postMessage({ type: "phase", phase: "Extracting columns", totalRows });
  for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
    const row = OUTPUT_COLUMNS.map((column) =>
      dateToUsString(formatExcelDate(cellValue(sheet, rowIndex, column)))
    );
    const outputRowIndex = rowIndex - range.s.r;
    if (outputRowIndex < PREVIEW_ROW_LIMIT) previewRows.push(row);

    for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
      columnWidths[columnIndex] = Math.max(
        columnWidths[columnIndex],
        String(row[columnIndex] ?? "").length + 3
      );
      const value = row[columnIndex];
      if (value === "" || value == null) continue;
      const address = XLSX.utils.encode_cell({ r: outputRowIndex, c: columnIndex });
      outputSheet[address] = {
        t: typeof value === "number" ? "n" : typeof value === "boolean" ? "b" : "s",
        v: value,
        s: { font: { name: "Calibri", sz: 9 } }
      };
    }

    if (rowIndex > range.s.r) {
      csvParts.push([
        csvEscape(cellValue(sheet, rowIndex, 41)),
        csvEscape(cellValue(sheet, rowIndex, 26))
      ].join(",") + "\r\n");
      csvLineCount += 1;
      if (csvLineCount % 10000 === 0) {
        const lastPart = csvParts.pop();
        csvParts.push(new Blob(csvParts, { type: "text/csv;charset=utf-8" }));
        csvParts.push(lastPart);
      }
    }
    if ((rowIndex - range.s.r + 1) % 25000 === 0) {
      self.postMessage({
        type: "progress",
        processedRows: rowIndex - range.s.r + 1,
        totalRows
      });
    }
  }

  self.postMessage({ type: "phase", phase: "Building the output workbook", totalRows });
  outputSheet["!cols"] = columnWidths.map((width) => ({ wch: width }));
  const outputBook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(outputBook, outputSheet, "RESULT");
  let outputArray = XLSX.write(outputBook, {
    bookType: "xlsx",
    type: "array",
    cellStyles: true
  });

  self.postMessage({ type: "phase", phase: "Applying workbook formatting", totalRows });
  const styledWorkbook = await JSZip.loadAsync(outputArray);
  const stylesFile = styledWorkbook.file("xl/styles.xml");
  if (!stylesFile) throw new Error("Could not apply the Field Result workbook formatting.");
  let stylesXml = await stylesFile.async("string");
  const fontSize = stylesXml.match(/<sz val="[^"]+"\s*\/>/);
  if (!fontSize) throw new Error("Could not apply the Field Result workbook font size.");
  stylesXml = stylesXml.replace(fontSize[0], '<sz val="9"/>');
  styledWorkbook.file("xl/styles.xml", stylesXml);
  outputArray = await styledWorkbook.generateAsync({ type: "uint8array" });

  self.postMessage({
    type: "complete",
    outputArray,
    previewRows,
    totalRows,
    extraCsvBlob: new Blob(csvParts, { type: "text/csv;charset=utf-8" })
  }, [outputArray.buffer]);
}

self.onmessage = async (event) => {
  try {
    await processWorkbook(event.data.file);
  } catch (error) {
    self.postMessage({
      type: "error",
      message: error instanceof Error ? error.message : String(error)
    });
  }
};
