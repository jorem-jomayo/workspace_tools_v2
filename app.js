"use strict";

const toolDetails = {
  vrp: {
    title: "VRP Mapper",
    subtitle: "Map VRP account files to the demand letter template and prepare release data."
  },
  field: {
    title: "Field Result",
    subtitle: "Extract and format selected columns from a Field Result workbook."
  },
  esign: {
    title: "E-SIGN FIXER",
    subtitle: "Prepare individual transparent signature images from a scanned page."
  }
};

const TEMPLATE_FILE = "demand_letter_template.csv";
const CMS_FILE = "cmd_id.xlsx";
const AMOUNTS_FILE = "amounts.csv";
const AREA_MAPPING = {
  ABRA: "NORTH LUZON", AURORA: "NORTH LUZON", BATAAN: "NORTH LUZON",
  BENGUET: "NORTH LUZON", BULACAN: "NORTH LUZON", CAGAYAN: "NORTH LUZON",
  "ILOCOS NORTE": "NORTH LUZON", "ILOCOS SUR": "NORTH LUZON", ISABELA: "NORTH LUZON",
  "LA UNION": "NORTH LUZON", "NUEVA ECIJA": "NORTH LUZON", PAMPANGA: "NORTH LUZON",
  PANGASINAN: "NORTH LUZON", ZAMBALES: "NORTH LUZON",
  NCR: "NCR", "METRO MANILA": "NCR",
  BATANGAS: "SOUTH LUZON", "CAMARINES SUR": "SOUTH LUZON", CAVITE: "SOUTH LUZON",
  LAGUNA: "SOUTH LUZON", "OCCIDENTAL MINDORO": "SOUTH LUZON",
  "ORIENTAL MINDORO": "SOUTH LUZON", "QUEZON PROVINCE": "SOUTH LUZON", RIZAL: "SOUTH LUZON",
  AKLAN: "VISAYAS", ANTIQUE: "VISAYAS", BOHOL: "VISAYAS", CAPIZ: "VISAYAS",
  CEBU: "VISAYAS", GUIMARAS: "VISAYAS", ILOILO: "VISAYAS", LEYTE: "VISAYAS",
  "NEGROS OCCIDENTAL": "VISAYAS", "NEGROS ORIENTAL": "VISAYAS",
  "AGUSAN DEL SUR": "MINDANAO", BUKIDNON: "MINDANAO", COTABATO: "MINDANAO",
  "DAVAO DE ORO": "MINDANAO", "DAVAO DEL NORTE": "MINDANAO", "DAVAO DEL SUR": "MINDANAO",
  "DAVAO ORIENTAL": "MINDANAO", "LANAO DEL NORTE": "MINDANAO", "LANAO DEL SUR": "MINDANAO",
  MAGUINDANAO: "MINDANAO", "MISAMIS OCCIDENTAL": "MINDANAO", "MISAMIS ORIENTAL": "MINDANAO",
  SARANGANI: "MINDANAO", "SOUTH COTABATO": "MINDANAO", "SULTAN KUDARAT": "MINDANAO",
  SULU: "MINDANAO", "SURIGAO DEL SUR": "MINDANAO", TARLAC: "MINDANAO", "TAWI-TAWI": "MINDANAO",
  "ZAMBOANGA DEL NORTE": "MINDANAO", "ZAMBOANGA DEL SUR": "MINDANAO",
  "ZAMBOANGA SIBUGAY": "MINDANAO"
};

const COLUMN_ALIASES = {
  account_no: ["ACCOUNT NUMBER", "ACCOUNT_NO", "ACCOUNTNUMBER"],
  bank: ["BANK"],
  placement: ["PLACEMENT"],
  ch_code: ["CH CODE", "CH_CODE", "CHCODE"],
  ch_name: ["CH NAME", "CH_NAME", "CHNAME"],
  address_type: ["ADD TYPE", "ADD_TYPE", "ADDTYPE"],
  address: ["ADDRESS"],
  final_area: ["FINAL AREA", "FINAL_AREA", "FINALAREA"],
  area: ["AREA"],
  municipality: ["MUNICIPALITY"],
  dl_type: ["DL TYPE", "DL_TYPE", "DLTYPE"],
  ref_code: ["REFERENCE CODE", "REF CODE", "REF_CODE", "REFCODE"],
  autofield_date: ["AUTOFIELD DATE", "AUTOFIELD_DATE", "AUTOFIELDDATE"],
  outstanding_balance: ["OB/PRINCIPAL", "OB_PRINCIPAL", "OBPRINCIPAL", "OB"],
  endorsement_date: ["ENDO DATE", "ENDO_DATE", "ENDODATE", "ENDORSEMENT DATE"],
  pullout_date: ["POUT DATE", "POUT_DATE", "POUTDATE", "PULLOUT DATE"],
  type_of_account: ["TYPE OF ACCOUNT", "TYPE_OF_ACCOUNT", "TYPEOFACCOUNT", "ACCOUNT TYPE", "ACCOUNT_TYPE"],
  assigned_printer_user: ["ASSIGNED PRINTER USER", "ASSIGNED_PRINTER_USER", "PRINTER USER"],
  shared_or_exclusive: ["SHARED OR EXCLUSIVE", "SHARED_OR_EXCLUSIVE", "SHARED/EXCLUSIVE"],
  due_date: ["DUE DATE", "DUE_DATE"],
  cycle: ["CYCLE"],
  amount_due: ["AMOUNT DUE", "AMOUNT_DUE", "PRINCIPAL AMOUNT DUE"],
  visit_type: ["VISIT TYPE", "VISIT_TYPE", "FIELD REMARS", "FIELD REMARKS", "FIELD REMARK"],
  unit_details: ["UNIT DETAILS", "UNIT_DETAILS"]
};

function normalizeHeader(value) {
  return String(value ?? "").trim().toUpperCase().replace(/[\s_\/]/g, "");
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let value = "";
  let quoted = false;
  const input = text.replace(/^\uFEFF/, "");
  for (let index = 0; index < input.length; index += 1) {
    const character = input[index];
    if (quoted) {
      if (character === '"' && input[index + 1] === '"') {
        value += '"';
        index += 1;
      } else if (character === '"') {
        quoted = false;
      } else {
        value += character;
      }
    } else if (character === '"' && value === "") {
      quoted = true;
    } else if (character === ",") {
      row.push(value);
      value = "";
    } else if (character === "\n" || character === "\r") {
      row.push(value);
      if (row.some((cell) => cell !== "")) rows.push(row);
      row = [];
      value = "";
      if (character === "\r" && input[index + 1] === "\n") index += 1;
    } else {
      value += character;
    }
  }
  row.push(value);
  if (row.some((cell) => cell !== "")) rows.push(row);
  return rows;
}

function recordsFromRows(rows) {
  if (!rows.length) return [];
  const headers = rows[0].map((header) => String(header).trim());
  return rows.slice(1).map((cells) => Object.fromEntries(
    headers.map((header, index) => [header, String(cells[index] ?? "").trim()])
  ));
}

function csvEscape(value) {
  const text = value == null ? "" : String(value);
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

function toCsv(headers, rows) {
  return [headers, ...rows.map((row) => headers.map((header) => row[header] ?? ""))]
    .map((row) => row.map(csvEscape).join(","))
    .join("\r\n");
}

function findSourceValue(record, aliases) {
  const desired = new Set(aliases.map(normalizeHeader));
  const key = Object.keys(record).find((candidate) => desired.has(normalizeHeader(candidate)));
  return key === undefined ? "" : String(record[key] ?? "").trim();
}

function findSourceValueOrColumn(record, aliases, columnIndex) {
  const matched = findSourceValue(record, aliases);
  if (matched) return matched;
  return String(Object.values(record)[columnIndex] ?? "").trim();
}

function formatGeneralNumber(value) {
  const text = String(value ?? "").trim().replace(/,/g, "");
  if (!text || ["nan", "none", "0", "0.0"].includes(text.toLowerCase())) return "0";
  const number = Number(text);
  if (!Number.isFinite(number)) return text;
  return Number.isInteger(number) ? String(number) : String(Number(number.toFixed(2)));
}

function formatDate(value, outputPattern) {
  const text = String(value ?? "").trim();
  if (!text) return "";
  let date;
  const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  const us = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (iso) date = new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  else if (us) date = new Date(Number(us[3]), Number(us[1]) - 1, Number(us[2]));
  else {
    const parsed = new Date(text);
    if (!Number.isFinite(parsed.getTime())) return "";
    date = new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
  }
  if (!Number.isFinite(date.getTime())) return "";
  const pad = (part) => String(part).padStart(2, "0");
  if (outputPattern === "DD-MM-YYYY") return `${pad(date.getDate())}-${pad(date.getMonth() + 1)}-${date.getFullYear()}`;
  return `${pad(date.getMonth() + 1)}/${pad(date.getDate())}/${date.getFullYear()}`;
}

function normalizeAccountType(value) {
  const upper = value.toUpperCase();
  if (upper.includes("TRANS") || upper.includes("DETAIL")) return "Trans/Details";
  if (upper === "DL" || upper.includes("DL")) return "DL";
  return value;
}

function normalizeVisitType(value) {
  const normalized = String(value ?? "").trim().toUpperCase().replace(/[^A-Z0-9]/g, "");
  if (normalized.includes("CARAVAN")) return "CARAVAN";
  if (normalized.includes("NEWENDO") || normalized.includes("REVISIT")) return "REGULAR";
  if (normalized === "OTS") return "OTS";
  if (normalized === "REGULAR") return "REGULAR";
  return "";
}

function cleanAddress(value) {
  return String(value ?? "")
    .replace(/\bBlk\b/gi, "Block")
    .replace(/\bBRGY\b/gi, "BARANGAY")
    .replace(/\bSTA\.?/gi, "SANTA")
    .replace(/\bSTO\.?/gi, "SANTO")
    .replace(/C\s*\/\s*O\.?\s*[:\-]?\s*/gi, "")
    .replace(/\s+/g, " ")
    .replace(/,\s*,/g, ",")
    .replace(/^[,\s]+|[,\s]+$/g, "")
    .toUpperCase();
}

function mapFileType(filename) {
  const name = filename.replace(/\.[^.]+$/, "").trim().toUpperCase();
  if (name.includes("REVISIT") && name.includes("NO DL")) return "Trans/Details";
  if (name.includes("WITH DL") || (name.includes("MC2") && name.includes("OTHERS"))) return "DL";
  return "DL";
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function downloadText(text, filename, mime = "text/csv;charset=utf-8") {
  downloadBlob(new Blob(["\uFEFF", text], { type: mime }), filename);
}

function setNotice(id, message, kind = "") {
  const notice = document.getElementById(id);
  notice.textContent = message;
  notice.className = `notice${kind ? ` ${kind}` : ""}`;
  notice.hidden = false;
}

function clearNotice(id) {
  const notice = document.getElementById(id);
  notice.hidden = true;
  notice.textContent = "";
  notice.className = "notice";
}

function renderTable(target, title, headers, rows) {
  const safe = (value) => String(value ?? "").replace(/[&<>"']/g, (character) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  })[character]);
  const section = document.createElement("section");
  section.innerHTML = `
    <div class="result-section">
      <h3>${safe(title)}</h3>
      <p class="result-summary">${rows.length.toLocaleString()} row(s)</p>
      <div class="table-wrap"><table class="result-table">
        <thead><tr>${headers.map((header) => `<th>${safe(header)}</th>`).join("")}</tr></thead>
        <tbody>${rows.map((row) => `<tr>${headers.map((header) => `<td>${safe(row[header])}</td>`).join("")}</tr>`).join("")}</tbody>
      </table></div>
    </div>`;
  target.append(section.firstElementChild);
}

function createIcon(symbol, className = "icon") {
  const icon = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  icon.classList.add(...className.split(" "));
  icon.setAttribute("aria-hidden", "true");
  const use = document.createElementNS("http://www.w3.org/2000/svg", "use");
  use.setAttribute("href", `#icon-${symbol}`);
  icon.append(use);
  return icon;
}

function renderDownloads(target, items) {
  const wrapper = document.createElement("div");
  wrapper.className = "actions result-section";
  for (const item of items) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "button button-primary";
    button.textContent = item.label;
    button.prepend(createIcon(item.icon || "download"));
    button.addEventListener("click", item.action);
    wrapper.append(button);
  }
  target.append(wrapper);
}

async function loadReferenceData() {
  const [templateResponse, cmsResponse, amountsResponse] = await Promise.all([
    fetch(TEMPLATE_FILE),
    fetch(CMS_FILE),
    fetch(AMOUNTS_FILE)
  ]);
  for (const [response, filename] of [
    [templateResponse, TEMPLATE_FILE], [cmsResponse, CMS_FILE], [amountsResponse, AMOUNTS_FILE]
  ]) {
    if (!response.ok) throw new Error(`Could not load ${filename} (HTTP ${response.status}).`);
  }
  const templateHeaders = parseCsv(await templateResponse.text())[0];
  if (!templateHeaders?.length) throw new Error(`${TEMPLATE_FILE} is empty or invalid.`);
  const cmsWorkbook = XLSX.read(await cmsResponse.arrayBuffer(), { type: "array", raw: false });
  const cmsSheet = cmsWorkbook.Sheets[cmsWorkbook.SheetNames[0]];
  const cmsRows = XLSX.utils.sheet_to_json(cmsSheet, { defval: "", raw: false });
  const cmsMap = new Map();
  for (const row of cmsRows) {
    const code = findSourceValue(row, ["CH CODE", "CH_CODE", "CHCODE"]).toUpperCase();
    const id = findSourceValue(row, ["CMS ID", "CMS_ID"]);
    if (code && id) cmsMap.set(code, id);
  }
  const amountRows = recordsFromRows(parseCsv(await amountsResponse.text()));
  const amountOb = new Map();
  const amountDue = new Map();
  for (const row of amountRows) {
    const code = findSourceValue(row, ["CH CODE", "CH_CODE", "CHCODE"]).toUpperCase();
    if (!code) continue;
    const ob = findSourceValue(row, ["AMOUNT OB"]);
    const due = findSourceValue(row, ["PRINCIPAL AMOUNT DUE"]);
    if (ob) amountOb.set(code, ob.replace(/,/g, ""));
    if (due) amountDue.set(code, due.replace(/,/g, ""));
  }
  return { templateHeaders, cmsMap, amountOb, amountDue };
}

function sourceRowsFromSheet(sheet) {
  return XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false })
    .map((row) => Object.fromEntries(Object.entries(row).map(([key, value]) => [String(key).trim(), String(value ?? "").trim()])));
}

async function processVrp() {
  const output = document.getElementById("vrp-output");
  output.replaceChildren();
  const files = Array.from(document.getElementById("vrp-files").files || []);
  const version = document.getElementById("vrp-version").value;
  if (!files.length) throw new Error(version === "MC2" ? "Please select at least one CSV file." : "Please select an Excel workbook.");
  const refs = document.getElementById("reference-codes").value
    .split(/[,\r\n]+/).map((code) => code.trim()).filter(Boolean);
  const refSet = new Set(refs);
  const sources = [];
  if (version === "MC2") {
    for (const file of files) {
      if (!file.name.toLowerCase().endsWith(".csv")) throw new Error(`${file.name}: MC2 accepts CSV files only.`);
      const rows = recordsFromRows(parseCsv(await file.text()));
      for (const row of rows) sources.push({ row, fileType: mapFileType(file.name), sourceName: file.name });
    }
  } else {
    const file = files[0];
    if (!file.name.toLowerCase().endsWith(".xlsx")) throw new Error("FCL accepts an .xlsx workbook.");
    const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", raw: false });
    const sheetName = workbook.SheetNames.find((name) => name.trim().toUpperCase() === "SUMMARY");
    if (!sheetName) throw new Error("The Excel workbook does not contain a sheet named SUMMARY.");
    for (const row of sourceRowsFromSheet(workbook.Sheets[sheetName])) sources.push({ row, fileType: "OTS", sourceName: file.name });
  }
  if (!sources.length) throw new Error("No account rows were found in the selected file(s).");
  const { templateHeaders, cmsMap, amountOb, amountDue } = await loadReferenceData();
  const refKey = Object.keys(sources[0].row).find((key) => ["REFCODE", "REFERENCECODE"].includes(normalizeHeader(key)));
  if (refs.length && !refKey) throw new Error("Reference code column was not found in the source data.");
  const filtered = refs.length ? sources.filter(({ row }) => refSet.has(String(row[refKey] ?? "").trim())) : sources;
  const outputRows = filtered.map(({ row, fileType, sourceName }) => {
    const result = Object.fromEntries(templateHeaders.map((header) => [header, ""]));
    for (const [target, aliases] of Object.entries(COLUMN_ALIASES)) {
      const header = templateHeaders.find((candidate) => normalizeHeader(candidate) === normalizeHeader(target));
      if (!header) continue;
      const value = findSourceValue(row, aliases);
      if (value) result[header] = value;
    }
    const getOutput = (key) => {
      const header = templateHeaders.find((candidate) => normalizeHeader(candidate) === normalizeHeader(key));
      return header ? result[header] : "";
    };
    const setOutput = (key, value) => {
      const header = templateHeaders.find((candidate) => normalizeHeader(candidate) === normalizeHeader(key));
      if (header) result[header] = value;
    };
    const chCode = getOutput("ch_code").toUpperCase();
    if (!getOutput("account_no")) setOutput("account_no", "0");
    if (getOutput("address")) setOutput("address", cleanAddress(getOutput("address")));
    setOutput("outstanding_balance", formatGeneralNumber(getOutput("outstanding_balance") || amountOb.get(chCode) || "0"));
    if (!getOutput("amount_due") && amountDue.has(chCode)) setOutput("amount_due", formatGeneralNumber(amountDue.get(chCode)));
    if (!getOutput("cms_id") && cmsMap.has(chCode)) setOutput("cms_id", cmsMap.get(chCode));
    if (!getOutput("final_area")) setOutput("final_area", getOutput("area"));
    const area = getOutput("area").trim().toUpperCase();
    setOutput("area_cluster", AREA_MAPPING[area] || "");
    setOutput("month", new Intl.DateTimeFormat("en", { month: "long" }).format(new Date()).toUpperCase());
    setOutput("account_type", "HOUSING");
    setOutput("form_code", "vid04qNT");
    if (!getOutput("type_of_account")) setOutput("type_of_account", fileType === "OTS" ? "DL" : fileType);
    else setOutput("type_of_account", normalizeAccountType(getOutput("type_of_account")));
    setOutput("autofield_date", formatDate(getOutput("autofield_date"), "MM/DD/YYYY"));
    setOutput("pullout_date", formatDate(getOutput("pullout_date"), "MM/DD/YYYY"));
    setOutput("endorsement_date", formatDate(getOutput("endorsement_date"), "DD-MM-YYYY"));
    if (version === "FCL") {
      setOutput("visit_type", "OTS");
    } else {
      const currentVisitType = findSourceValueOrColumn(row, COLUMN_ALIASES.visit_type, 0);
      const normalizedCurrentType = normalizeVisitType(currentVisitType);
      if (normalizedCurrentType) {
        setOutput("visit_type", normalizedCurrentType);
      } else {
        const visitTypeSource = [
          findSourceValue(row, ["REMARK", "REMARKS", "REMARK TYPE", "REMARKS TYPE", "REMARK DESCRIPTION", "REMARKS DESCRIPTION"]),
          findSourceValue(row, ["ENDO TYPE", "ENDO_TYPE", "ENDORSEMENT TYPE", "DL TYPE", "DL_TYPE"]),
          findSourceValue(row, ["VISIT TYPE", "VISIT_TYPE"])
        ].find((value) => normalizeVisitType(value));
        const visitType = normalizeVisitType(visitTypeSource || sourceName);
        if (visitType) setOutput("visit_type", visitType);
      }
    }
    return result;
  });
  const releaseRows = filtered.map(({ row }) => ({
    "REF CODE": findSourceValue(row, ["REF CODE", "REFERENCE CODE", "REF_CODE", "REFCODE"]),
    "RELEASED TO": findSourceValueOrColumn(row, ["TAG FM", "TAGFM", "RELEASED TO", "TAG"], 19)
  }));
  const bank = findSourceValue(filtered[0]?.row || {}, ["BANK"]).replace(/[^\w-]+/g, "_") || "BANK";
  const mainName = `${bank}_${outputRows.length}.csv`;
  const releaseName = `Released_to_${bank}_${outputRows.length}.csv`;
  const mainCsv = toCsv(templateHeaders, outputRows);
  const releaseCsv = toCsv(["REF CODE", "RELEASED TO"], releaseRows);
  if (version === "MC2") {
    const cmsHeader = templateHeaders.find((header) => normalizeHeader(header) === "CMSID");
    if (!cmsHeader) throw new Error(`${TEMPLATE_FILE} is missing the cms_id output column.`);
    const missingCmsRows = outputRows.filter((row) => {
      const cmsId = String(row[cmsHeader] ?? "").trim();
      return !cmsId || ["NAN", "NONE", "NULL"].includes(cmsId.toUpperCase());
    });
    if (missingCmsRows.length) {
      const missingCodes = [...new Set(missingCmsRows
        .map((row) => findSourceValue(row, ["CH CODE", "CH_CODE", "CHCODE"]))
        .filter(Boolean))];
      const codeDetails = missingCodes.length ? ` CH CODE(s): ${missingCodes.join(", ")}.` : "";
      throw new Error(`MC2 processing stopped: CMS ID not found for ${missingCmsRows.length} account(s). Check the CH CODE mapping in ${CMS_FILE}.${codeDetails}`);
    }
  }

  renderDownloads(output, [
    { label: `Download Main CSV (${mainName})`, icon: "table", action: () => downloadText(mainCsv, mainName) },
    { label: `Download Released-to CSV (${releaseName})`, icon: "file-text", action: () => downloadText(releaseCsv, releaseName) }
  ]);
  renderTable(output, "Main Data Preview", templateHeaders, outputRows);
  renderTable(output, "Release File Preview", ["REF CODE", "RELEASED TO"], releaseRows);
  return `${outputRows.length.toLocaleString()} account(s) processed.`;
}

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

async function processFieldResult() {
  const file = document.getElementById("field-file").files[0];
  if (!file) throw new Error("Please select an Excel workbook.");
  if (!file.name.toLowerCase().endsWith(".xlsx")) throw new Error("Please select an .xlsx workbook.");
  const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: true });
  const sheetName = workbook.SheetNames.find((name) => name.trim().toUpperCase() === "RESULT");
  if (!sheetName) throw new Error("A sheet named RESULT was not found in this workbook.");
  const sheet = workbook.Sheets[sheetName];
  const source = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: true });
  const sourceRow = (row, column) => row[column] ?? "";
  const targetRows = source.map((row) => [
    sourceRow(row, 41), sourceRow(row, 2), sourceRow(row, 3), sourceRow(row, 4),
    sourceRow(row, 27), sourceRow(row, 32), sourceRow(row, 36), sourceRow(row, 26)
  ].map(formatExcelDate));
  const extraRows = source.slice(1).map((row) => ({
    "Ref Code": sourceRow(row, 41),
    "Field Name": sourceRow(row, 26)
  }));
  const outputBook = XLSX.utils.book_new();
  const outputSheet = XLSX.utils.aoa_to_sheet(targetRows.map((row) => row.map(dateToUsString)));
  outputSheet["!cols"] = Array.from({ length: 8 }, (_, col) => {
    const width = targetRows.reduce((max, row) => Math.max(max, String(dateToUsString(row[col]) ?? "").length), 0);
    return { wch: Math.max(width + 3, 10) };
  });
  for (const address of Object.keys(outputSheet)) {
    if (address.startsWith("!")) continue;
    outputSheet[address].s = { font: { name: "Calibri", sz: 9 } };
  }
  XLSX.utils.book_append_sheet(outputBook, outputSheet, "RESULT");
  let outputArray = XLSX.write(outputBook, { bookType: "xlsx", type: "array", cellStyles: true });
  const styledWorkbook = await JSZip.loadAsync(outputArray);
  const stylesFile = styledWorkbook.file("xl/styles.xml");
  if (!stylesFile) throw new Error("Could not apply the Field Result workbook formatting.");
  let stylesXml = await stylesFile.async("string");
  const fontSize = stylesXml.match(/<sz val="[^"]+"\s*\/>/);
  if (!fontSize) throw new Error("Could not apply the Field Result workbook font size.");
  stylesXml = stylesXml.replace(fontSize[0], '<sz val="9"/>');
  styledWorkbook.file("xl/styles.xml", stylesXml);
  outputArray = await styledWorkbook.generateAsync({ type: "uint8array" });
  const date = new Date();
  const dateStamp = `${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}-${date.getFullYear()}`;
  const workbookName = `FIELD RESULT ${dateStamp}.xlsx`;
  const extraCsv = toCsv(["Ref Code", "Field Name"], extraRows);
  const output = document.getElementById("field-output");
  output.replaceChildren();
  renderDownloads(output, [
    { label: `Download Extracted File (${workbookName})`, icon: "sheet", action: () => downloadBlob(new Blob([outputArray], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), workbookName) },
    { label: "Download Ref Code & Field Name CSV (fieldman.csv)", icon: "file-text", action: () => downloadText(extraCsv, "fieldman.csv") }
  ]);
  renderTable(output, "Extracted Data Preview", ["A", "B", "C", "D", "E", "F", "G", "H"], targetRows.map((row) => Object.fromEntries(row.map((value, index) => [String.fromCharCode(65 + index), dateToUsString(value)]))));
  return `${targetRows.length.toLocaleString()} row(s) extracted.`;
}

function grayscaleImageData(imageData) {
  const { data, width, height } = imageData;
  const gray = new Uint8Array(width * height);
  for (let i = 0, pixel = 0; i < data.length; i += 4, pixel += 1) {
    gray[pixel] = Math.round(0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2]);
  }
  return gray;
}

function dilate(mask, width, height) {
  const result = new Uint8Array(mask.length);
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      let found = false;
      for (let dy = -1; dy <= 1 && !found; dy += 1) {
        const yy = y + dy;
        if (yy < 0 || yy >= height) continue;
        for (let dx = -1; dx <= 1; dx += 1) {
          const xx = x + dx;
          if (xx >= 0 && xx < width && mask[yy * width + xx]) {
            found = true;
            break;
          }
        }
      }
      result[y * width + x] = found ? 1 : 0;
    }
  }
  return result;
}

function findSignatureBounds(mask, width, height) {
  const visited = new Uint8Array(mask.length);
  const queue = new Int32Array(mask.length);
  const components = [];
  for (let start = 0; start < mask.length; start += 1) {
    if (!mask[start] || visited[start]) continue;
    let head = 0;
    let tail = 0;
    queue[tail++] = start;
    visited[start] = 1;
    let minX = width, minY = height, maxX = -1, maxY = -1, area = 0;
    while (head < tail) {
      const point = queue[head++];
      const x = point % width;
      const y = Math.floor(point / width);
      area += 1;
      minX = Math.min(minX, x); minY = Math.min(minY, y);
      maxX = Math.max(maxX, x); maxY = Math.max(maxY, y);
      const yStart = Math.max(0, y - 1), yEnd = Math.min(height - 1, y + 1);
      const xStart = Math.max(0, x - 1), xEnd = Math.min(width - 1, x + 1);
      for (let yy = yStart; yy <= yEnd; yy += 1) {
        const offset = yy * width;
        for (let xx = xStart; xx <= xEnd; xx += 1) {
          const next = offset + xx;
          if (mask[next] && !visited[next]) {
            visited[next] = 1;
            queue[tail++] = next;
          }
        }
      }
    }
    if (area >= 1000) components.push({ x: minX, y: minY, width: maxX - minX + 1, height: maxY - minY + 1 });
  }
  return components.sort((left, right) => left.y - right.y || left.x - right.x);
}

function makeSignaturePng(gray, width, height, bounds) {
  const padding = 15;
  const x1 = Math.max(0, bounds.x - padding), y1 = Math.max(0, bounds.y - padding);
  const x2 = Math.min(width, bounds.x + bounds.width + padding), y2 = Math.min(height, bounds.y + bounds.height + padding);
  let minX = x2, minY = y2, maxX = -1, maxY = -1;
  for (let y = y1; y < y2; y += 1) {
    for (let x = x1; x < x2; x += 1) {
      const offset = y * width + x;
      if (gray[offset] < 240) {
        minX = Math.min(minX, x); minY = Math.min(minY, y);
        maxX = Math.max(maxX, x); maxY = Math.max(maxY, y);
      }
    }
  }
  if (maxX < minX || maxY < minY) return null;
  const cropWidth = maxX - minX + 1, cropHeight = maxY - minY + 1;
  const squareSize = Math.max(cropWidth, cropHeight);
  const square = document.createElement("canvas");
  square.width = squareSize;
  square.height = squareSize;
  const context = square.getContext("2d");
  const rgba = context.createImageData(squareSize, squareSize);
  const xOffset = Math.floor((squareSize - cropWidth) / 2);
  const yOffset = Math.floor((squareSize - cropHeight) / 2);
  for (let y = 0; y < cropHeight; y += 1) {
    for (let x = 0; x < cropWidth; x += 1) {
      const value = gray[(minY + y) * width + minX + x];
      const offset = ((yOffset + y) * squareSize + xOffset + x) * 4;
      rgba.data[offset] = value;
      rgba.data[offset + 1] = value;
      rgba.data[offset + 2] = value;
      rgba.data[offset + 3] = value < 240 ? 255 : 0;
    }
  }
  context.putImageData(rgba, 0, 0);
  const finalCanvas = document.createElement("canvas");
  finalCanvas.width = 250;
  finalCanvas.height = 250;
  const finalContext = finalCanvas.getContext("2d");
  finalContext.imageSmoothingEnabled = true;
  finalContext.drawImage(square, 0, 0, 250, 250);
  return finalCanvas;
}

async function makeSignaturesDocument(signatures) {
  const zip = new JSZip();
  const relationships = signatures.map((_, index) =>
    `<Relationship Id="rId${index + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/signature_${index + 1}.png"/>`
  ).join("");
  const paragraphs = signatures.map((_, index) => `
    <w:p><w:r><w:t>Signature ${index + 1}:</w:t></w:r></w:p>
    <w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">
      <wp:extent cx="2286000" cy="2286000"/><wp:docPr id="${index + 1}" name="Signature ${index + 1}"/>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
        <pic:pic><pic:nvPicPr><pic:cNvPr id="${index + 1}" name="signature_${index + 1}.png"/><pic:cNvPicPr/></pic:nvPicPr>
        <pic:blipFill><a:blip r:embed="rId${index + 2}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
        <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2286000" cy="2286000"/></a:xfrm>
        <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>
      </a:graphicData></a:graphic>
    </wp:inline></w:drawing></w:r></w:p>`).join("");
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
      <w:body><w:p><w:r><w:t>Extracted E-Signatures</w:t></w:r></w:p>${paragraphs}<w:sectPr/></w:body>
    </w:document>`;
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
      <Default Extension="xml" ContentType="application/xml"/>
      <Default Extension="png" ContentType="image/png"/>
      <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    </Types>`);
  zip.folder("_rels").file(".rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    </Relationships>`);
  zip.folder("word").file("document.xml", documentXml);
  zip.folder("word").folder("_rels").file("document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}</Relationships>`);
  const media = zip.folder("word").folder("media");
  for (let index = 0; index < signatures.length; index += 1) {
    const image = await new Promise((resolve, reject) => signatures[index].toBlob((blob) => blob ? resolve(blob) : reject(new Error("Could not encode signature image.")), "image/png"));
    media.file(`signature_${index + 1}.png`, image);
  }
  return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

async function processSignatures() {
  const file = document.getElementById("esign-file").files[0];
  if (!file) throw new Error("Please select a PNG or JPG image.");
  const bitmap = await createImageBitmap(file);
  const canvas = document.createElement("canvas");
  canvas.width = bitmap.width;
  canvas.height = bitmap.height;
  const context = canvas.getContext("2d", { willReadFrequently: true });
  context.fillStyle = "#fff";
  context.fillRect(0, 0, canvas.width, canvas.height);
  context.drawImage(bitmap, 0, 0);
  bitmap.close();
  const pixels = context.getImageData(0, 0, canvas.width, canvas.height);
  const gray = grayscaleImageData(pixels);
  const binary = new Uint8Array(gray.length);
  for (let index = 0; index < gray.length; index += 1) binary[index] = gray[index] < 220 ? 1 : 0;
  const expanded = dilate(binary, canvas.width, canvas.height);
  const bounds = findSignatureBounds(expanded, canvas.width, canvas.height);
  const signatures = bounds.map((box) => makeSignaturePng(gray, canvas.width, canvas.height, box)).filter(Boolean);
  if (!signatures.length) throw new Error("No signatures were detected. Try a clearer image with darker signatures.");
  const zip = new JSZip();
  for (let index = 0; index < signatures.length; index += 1) {
    const png = await new Promise((resolve, reject) => signatures[index].toBlob((blob) => blob ? resolve(blob) : reject(new Error("Could not encode signature image.")), "image/png"));
    zip.file(`signature_${index + 1}.PNG`, png);
  }
  const [zipBlob, docxBlob] = await Promise.all([
    zip.generateAsync({ type: "blob", mimeType: "application/zip" }),
    makeSignaturesDocument(signatures)
  ]);
  const output = document.getElementById("esign-output");
  output.replaceChildren();
  renderDownloads(output, [
    { label: "Download Signature Images (ZIP)", icon: "archive", action: () => downloadBlob(zipBlob, "signatures.zip") },
    { label: "Download Word Document", icon: "file-text", action: () => downloadBlob(docxBlob, "signatures.docx") }
  ]);
  const heading = document.createElement("h3");
  heading.className = "result-section";
  heading.textContent = `Extracted ${signatures.length} signature(s)`;
  output.append(heading);
  const gallery = document.createElement("div");
  gallery.className = "output-hint";
  signatures.forEach((signature, index) => {
    const card = document.createElement("div");
    card.className = "output-card";
    const caption = document.createElement("strong");
    caption.textContent = `Signature ${index + 1}`;
    const image = document.createElement("img");
    image.className = "preview-image";
    image.alt = `Extracted signature ${index + 1}`;
    image.src = signature.toDataURL("image/png");
    card.append(caption, image);
    gallery.append(card);
  });
  output.append(gallery);
  return `${signatures.length} signature(s) extracted.`;
}

async function runProcess(form, noticeId, processor) {
  const button = form.querySelector('button[type="submit"]');
  const originalContents = Array.from(button.childNodes, (node) => node.cloneNode(true));
  button.disabled = true;
  button.textContent = "Processing...";
  clearNotice(noticeId);
  try {
    const result = await processor();
    setNotice(noticeId, result, "success");
  } catch (error) {
    setNotice(noticeId, error instanceof Error ? error.message : String(error), "error");
  } finally {
    button.disabled = false;
    button.replaceChildren(...originalContents);
  }
}

const navButtons = document.querySelectorAll(".nav-button");
const content = document.querySelector(".content");
let toolTransitionTimer;
let activeTool = "vrp";

navButtons.forEach((button) => {
  button.addEventListener("click", () => {
    const selectedTool = button.dataset.tool;
    if (selectedTool === activeTool && !toolTransitionTimer) return;

    window.clearTimeout(toolTransitionTimer);
    navButtons.forEach((item) => item.setAttribute("aria-current", item === button ? "page" : "false"));
    content.classList.add("is-transitioning");
    toolTransitionTimer = window.setTimeout(() => {
      document.querySelectorAll(".tool-view").forEach((view) => {
        view.hidden = view.id !== `tool-${selectedTool}`;
      });
      document.getElementById("current-tool").textContent = toolDetails[selectedTool].title;
      document.getElementById("page-title").textContent = toolDetails[selectedTool].title;
      document.getElementById("page-subtitle").textContent = toolDetails[selectedTool].subtitle;
      activeTool = selectedTool;
      toolTransitionTimer = undefined;
      window.requestAnimationFrame(() => content.classList.remove("is-transitioning"));
    }, 100);
  });
});

const versionSelect = document.getElementById("vrp-version");
const vrpFiles = document.getElementById("vrp-files");
const vrpFileList = document.getElementById("vrp-file-list");

function updateVrpUploadMode(clearSelection = true) {
  const isMc2 = versionSelect.value === "MC2";
  vrpFiles.accept = isMc2
    ? ".csv,text/csv"
    : ".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  vrpFiles.multiple = isMc2;
  document.getElementById("vrp-upload-label").textContent = isMc2
    ? "Upload VRP ACCOUNTS CSV(s)"
    : "Upload VRP ACCOUNTS Excel file";
  document.getElementById("vrp-upload-help").textContent = isMc2
    ? "CSV files · multiple files allowed"
    : "Excel workbook · must contain a SUMMARY sheet";
  if (clearSelection) {
    vrpFiles.value = "";
    vrpFileList.textContent = "";
  }
}

function showSelectedFiles(input, output) {
  const names = Array.from(input.files || []).map((file) => file.name);
  output.textContent = names.length ? `Selected: ${names.join(", ")}` : "";
}

versionSelect.addEventListener("change", () => updateVrpUploadMode());
vrpFiles.addEventListener("change", () => showSelectedFiles(vrpFiles, vrpFileList));
document.getElementById("field-file").addEventListener("change", (event) => {
  showSelectedFiles(event.currentTarget, document.getElementById("field-file-list"));
});
document.getElementById("esign-file").addEventListener("change", (event) => {
  showSelectedFiles(event.currentTarget, document.getElementById("esign-file-list"));
});

document.getElementById("vrp-form").addEventListener("submit", (event) => {
  event.preventDefault();
  runProcess(event.currentTarget, "vrp-notice", processVrp);
});
document.getElementById("field-form").addEventListener("submit", (event) => {
  event.preventDefault();
  runProcess(event.currentTarget, "field-notice", processFieldResult);
});
document.getElementById("esign-form").addEventListener("submit", (event) => {
  event.preventDefault();
  runProcess(event.currentTarget, "esign-notice", processSignatures);
});
document.getElementById("vrp-reset").addEventListener("click", () => {
  window.setTimeout(() => {
    updateVrpUploadMode(false);
    clearNotice("vrp-notice");
    document.getElementById("vrp-output").replaceChildren();
  }, 0);
});
document.getElementById("field-form").addEventListener("reset", () => {
  window.setTimeout(() => {
    document.getElementById("field-file-list").textContent = "";
    clearNotice("field-notice");
    document.getElementById("field-output").replaceChildren();
  }, 0);
});
document.getElementById("esign-form").addEventListener("reset", () => {
  window.setTimeout(() => {
    document.getElementById("esign-file-list").textContent = "";
    clearNotice("esign-notice");
    document.getElementById("esign-output").replaceChildren();
  }, 0);
});
