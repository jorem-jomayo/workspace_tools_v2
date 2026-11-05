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

function referenceCodesFromColumnM(rows) {
  return [...new Set(rows.slice(1)
    .map((row) => String(row[12] ?? "").trim())
    .filter(Boolean))];
}

async function extractVrpReferenceCodes() {
  const files = Array.from(vrpFiles.files || []);
  if (!files.length) throw new Error("Please select a source file first.");

  const rows = [];
  if (versionSelect.value === "MC2") {
    for (const file of files) {
      if (!file.name.toLowerCase().endsWith(".csv")) {
        throw new Error(`${file.name}: MC2 accepts CSV files only.`);
      }
      rows.push(...referenceCodesFromColumnM(parseCsv(await file.text())));
    }
  } else {
    const file = files[0];
    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      throw new Error("FCL accepts an .xlsx workbook.");
    }
    const workbook = XLSX.read(await file.arrayBuffer(), { type: "array", raw: false });
    const sheetName = workbook.SheetNames.find((name) => name.trim().toUpperCase() === "SUMMARY");
    if (!sheetName) throw new Error("The Excel workbook does not contain a sheet named SUMMARY.");
    const sheetRows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
      header: 1,
      defval: "",
      raw: false
    });
    rows.push(...referenceCodesFromColumnM(sheetRows));
  }

  const codes = [...new Set(rows)];
  if (!codes.length) throw new Error("No REF CODE values were found in column M.");
  return codes;
}

async function copyToClipboard(text) {
  if (navigator.clipboard?.writeText) {
    try {
      await navigator.clipboard.writeText(text);
      return;
    } catch {
      // Try the legacy copy method for browsers that reject Clipboard API access.
    }
  }

  const temporaryInput = document.createElement("textarea");
  temporaryInput.value = text;
  temporaryInput.setAttribute("readonly", "");
  temporaryInput.style.position = "fixed";
  temporaryInput.style.left = "-9999px";
  const previousFocus = document.activeElement;
  document.body.append(temporaryInput);
  let copied = false;
  try {
    temporaryInput.focus();
    temporaryInput.select();
    copied = document.execCommand("copy");
  } finally {
    temporaryInput.remove();
    previousFocus?.focus();
  }
  if (!copied) throw new Error("Clipboard access was blocked. Allow clipboard access and try again.");
}

async function processFieldResult() {
  const file = document.getElementById("field-file").files[0];
  if (!file) throw new Error("Please select an Excel workbook.");
  if (!file.name.toLowerCase().endsWith(".xlsx")) throw new Error("Please select an .xlsx workbook.");
  setNotice("field-notice", "Uploading workbook for memory-efficient processing…", "");
  const response = await fetch("/api/field-result", {
    method: "POST",
    body: file,
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    },
    cache: "no-store",
    credentials: "same-origin"
  });
  let result;
  try {
    result = await response.json();
  } catch {
    const isStaticHosting = location.protocol === "file:" ||
      location.hostname.endsWith(".github.io");
    if (response.status === 404 && isStaticHosting) {
      return processFieldResultInBrowser(file);
    }
    if (response.status === 404) {
      throw new Error("The streaming Field Result server is unavailable. Restart the app with npm.cmd start; the Python static server cannot process large workbooks.");
    }
    throw new Error(`The Field Result server returned an invalid response (HTTP ${response.status}).`);
  }
  if (!response.ok) {
    throw new Error(result.error || `Field Result processing failed (HTTP ${response.status}).`);
  }

  const output = document.getElementById("field-output");
  output.replaceChildren();
  renderDownloads(output, [
    {
      label: `Download Extracted File (${result.workbookName})`,
      icon: "sheet",
      action: () => downloadUrl(`/api/field-result/${result.id}/workbook`, result.workbookName)
    },
    {
      label: "Download Ref Code & Field Name CSV (fieldman.csv)",
      icon: "file-text",
      action: () => downloadUrl(`/api/field-result/${result.id}/csv`, "fieldman.csv")
    }
  ]);
  const previewObjects = result.previewRows.map((row) =>
    Object.fromEntries(row.map((value, index) => [String.fromCharCode(65 + index), value]))
  );
  renderTable(
    output,
    `Extracted Data Preview (first ${previewObjects.length.toLocaleString()} of ${result.rows.toLocaleString()} rows)`,
    ["A", "B", "C", "D", "E", "F", "G", "H"],
    previewObjects
  );
  return `${result.rows.toLocaleString()} row(s) extracted. Preview shows up to 100 rows.`;
}

function downloadUrl(url, filename) {
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
}

async function processFieldResultInBrowser(file) {
  const { outputArray, previewRows, totalRows, extraCsvBlob } = await new Promise((resolve, reject) => {
    const fieldWorker = new Worker("field-result-worker.js?v=4");
    fieldWorker.onmessage = ({ data }) => {
      if (data.type === "error") {
        fieldWorker.terminate();
        reject(new Error(data.message));
      } else if (data.type === "phase") {
        const total = data.totalRows ? ` (${data.totalRows.toLocaleString()} rows)` : "";
        setNotice("field-notice", `${data.phase}${total}…`, "");
      } else if (data.type === "progress") {
        setNotice(
          "field-notice",
          `Extracting columns: ${data.processedRows.toLocaleString()} of ${data.totalRows.toLocaleString()} rows…`,
          ""
        );
      } else if (data.type === "complete") {
        fieldWorker.terminate();
        resolve(data);
      }
    };
    fieldWorker.onerror = () => {
      fieldWorker.terminate();
      reject(new Error("Field Result background processing failed. The workbook may be too large for the available browser memory; try splitting it into smaller workbooks."));
    };
    fieldWorker.postMessage({ file });
  });
  const date = new Date();
  const dateStamp = `${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}-${date.getFullYear()}`;
  const workbookName = `FIELD RESULT ${dateStamp}.xlsx`;
  const output = document.getElementById("field-output");
  output.replaceChildren();
  renderDownloads(output, [
    { label: `Download Extracted File (${workbookName})`, icon: "sheet", action: () => downloadBlob(new Blob([outputArray], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), workbookName) },
    { label: "Download Ref Code & Field Name CSV (fieldman.csv)", icon: "file-text", action: () => downloadBlob(extraCsvBlob, "fieldman.csv") }
  ]);
  const previewObjects = previewRows.map((row) =>
    Object.fromEntries(row.map((value, index) => [String.fromCharCode(65 + index), value]))
  );
  renderTable(
    output,
    `Extracted Data Preview (first ${previewRows.length.toLocaleString()} of ${totalRows.toLocaleString()} rows)`,
    ["A", "B", "C", "D", "E", "F", "G", "H"],
    previewObjects
  );
  return `${totalRows.toLocaleString()} row(s) extracted. Preview shows up to 100 rows.`;
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

const fieldResultTimings = [];
const INITIAL_FIELD_RESULT_MS_PER_MIB = 5000;

function estimateFieldResultDuration(fileSize) {
  const mebibytes = Math.max(fileSize / (1024 * 1024), 0.1);
  const measuredRates = fieldResultTimings.map(({ duration, size }) => duration / Math.max(size / (1024 * 1024), 0.1));
  const msPerMib = measuredRates.length
    ? measuredRates.reduce((total, rate) => total + rate, 0) / measuredRates.length
    : INITIAL_FIELD_RESULT_MS_PER_MIB;
  return Math.max(15000, 10000 + mebibytes * msPerMib);
}

function formatRemainingTime(milliseconds) {
  const totalSeconds = Math.ceil(milliseconds / 1000);
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;
  if (!minutes) return `${seconds} sec`;
  return seconds ? `${minutes} min ${seconds} sec` : `${minutes} min`;
}

async function runProcess(form, noticeId, processor, estimate) {
  const button = form.querySelector('button[type="submit"]');
  const originalContents = Array.from(button.childNodes, (node) => node.cloneNode(true));
  button.disabled = true;
  button.textContent = "Processing...";
  clearNotice(noticeId);
  const startedAt = performance.now();
  let estimateTimer;
  if (estimate) {
    const updateEstimate = () => {
      const remaining = estimate.totalMs - (performance.now() - startedAt);
      setNotice(
        noticeId,
        remaining > 0
          ? `Uploading and processing: about ${formatRemainingTime(remaining)} remaining.`
          : "Uploading or processing is taking longer than the initial estimate.",
        ""
      );
    };
    updateEstimate();
    estimateTimer = window.setInterval(updateEstimate, 1000);
  }
  try {
    if (estimate) await new Promise((resolve) => window.requestAnimationFrame(resolve));
    const result = await processor();
    if (estimate?.onSuccess) estimate.onSuccess(performance.now() - startedAt);
    setNotice(noticeId, result, "success");
  } catch (error) {
    setNotice(noticeId, error instanceof Error ? error.message : String(error), "error");
  } finally {
    window.clearInterval(estimateTimer);
    button.disabled = false;
    button.replaceChildren(...originalContents);
  }
}

const navButtons = document.querySelectorAll(".nav-button");
const content = document.querySelector(".content");
const vrpNotifications = document.getElementById("vrp-notifications");
const notificationPopover = document.getElementById("notification-popover");
const notificationStatus = document.getElementById("notification-status");
const notificationCount = document.getElementById("notification-count");
const notificationBadge = document.getElementById("notification-badge");
const notificationExport = document.getElementById("notification-export");
let hasLoadedNotificationCount = false;
let toolTransitionTimer;
let activeTool = "vrp";

function showNotificationStatus(message, state = "") {
  notificationStatus.textContent = message;
  notificationStatus.dataset.state = state;
}

function closeNotificationPopover() {
  notificationPopover.hidden = true;
  vrpNotifications.setAttribute("aria-expanded", "false");
}

function formatRecordCount(count) {
  return `${count.toLocaleString()} ${count === 1 ? "record" : "records"} in the selected Lark view`;
}

async function getNotificationError(response) {
  if (response.status === 404) {
    return "The Lark API is not available on this server. Start the protected app with npm.cmd start.";
  }
  try {
    const body = await response.json();
    return body.error || `Request failed (${response.status}).`;
  } catch {
    return `Request failed (${response.status}).`;
  }
}

async function requestLocalNotificationApi(path) {
  let response;
  try {
    response = await fetch(path, {
      cache: "no-store",
      credentials: "same-origin"
    });
  } catch {
    throw new Error("Cannot reach the protected app server. Start it with npm.cmd start after updating .dev.vars.");
  }
  if (!response.ok) throw new Error(await getNotificationError(response));
  return response;
}

async function refreshLarkNotifications() {
  const refreshButton = document.getElementById("notification-refresh");
  refreshButton.disabled = true;
  notificationCount.textContent = "Checking…";
  notificationBadge.hidden = true;
  notificationExport.disabled = true;
  showNotificationStatus("Checking the Lark table…");
  try {
    const response = await requestLocalNotificationApi("/api/notifications");
    const result = await response.json();
    if (!Number.isSafeInteger(result.count) || result.count < 0) {
      throw new Error("The Lark API returned an invalid record count.");
    }
    notificationCount.textContent = formatRecordCount(result.count);
    notificationBadge.textContent = result.count > 999 ? "999+" : String(result.count);
    notificationBadge.hidden = result.count === 0;
    vrpNotifications.setAttribute("aria-label", `Notifications, ${result.count} records`);
    hasLoadedNotificationCount = true;
    notificationExport.disabled = false;
    const checkedAt = new Date(result.checkedAt);
    const timeLabel = Number.isNaN(checkedAt.getTime())
      ? ""
      : ` Last checked ${checkedAt.toLocaleTimeString()}.`;
    showNotificationStatus(`Connected to the Lark table.${timeLabel}`, "success");
  } catch (error) {
    hasLoadedNotificationCount = false;
    notificationCount.textContent = "Unavailable";
    notificationExport.disabled = true;
    vrpNotifications.setAttribute("aria-label", "Notifications");
    showNotificationStatus(error.message || "Could not check the Lark table.", "error");
  } finally {
    refreshButton.disabled = false;
  }
}

async function exportLarkRecords() {
  notificationExport.disabled = true;
  showNotificationStatus("Preparing the CSV export…");
  try {
    const response = await requestLocalNotificationApi("/api/export");
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "NEW VRP ACCOUNTS FOR TAG_PIF WITH DL_JOREM FOR UPLOAD.csv";
    document.body.append(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1000);
    showNotificationStatus("CSV export downloaded.", "success");
  } catch (error) {
    showNotificationStatus(error.message || "Could not export the Lark table.", "error");
  } finally {
    notificationExport.disabled = !hasLoadedNotificationCount;
  }
}

vrpNotifications.addEventListener("click", () => {
  const open = notificationPopover.hidden;
  notificationPopover.hidden = !open;
  vrpNotifications.setAttribute("aria-expanded", String(open));
  if (open) refreshLarkNotifications();
});

document.getElementById("notification-close").addEventListener("click", closeNotificationPopover);
document.getElementById("notification-refresh").addEventListener("click", refreshLarkNotifications);
notificationExport.addEventListener("click", exportLarkRecords);

document.addEventListener("click", (event) => {
  if (!event.target.closest("#notification-control")) closeNotificationPopover();
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") closeNotificationPopover();
});

window.setInterval(() => {
  if (activeTool === "vrp" && document.visibilityState === "visible") {
    refreshLarkNotifications();
  }
}, 5 * 60 * 1000);

refreshLarkNotifications();

navButtons.forEach((button) => {
  button.addEventListener("click", () => {
    const selectedTool = button.dataset.tool;
    if (selectedTool === activeTool && !toolTransitionTimer) return;

    if (selectedTool !== "vrp") closeNotificationPopover();
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
      vrpNotifications.hidden = selectedTool !== "vrp";
      activeTool = selectedTool;
      toolTransitionTimer = undefined;
      window.requestAnimationFrame(() => content.classList.remove("is-transitioning"));
    }, 100);
  });
});

const versionSelect = document.getElementById("vrp-version");
const vrpFiles = document.getElementById("vrp-files");
const vrpFileList = document.getElementById("vrp-file-list");
const extractRefCodesButton = document.getElementById("vrp-extract-ref-codes");
const extractRefCodesStatus = document.getElementById("vrp-extract-ref-codes-status");

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
    extractRefCodesButton.disabled = true;
    extractRefCodesStatus.textContent = "";
    delete extractRefCodesStatus.dataset.state;
  }
}

function showSelectedFiles(input, output) {
  const names = Array.from(input.files || []).map((file) => file.name);
  output.textContent = names.length ? `Selected: ${names.join(", ")}` : "";
}

versionSelect.addEventListener("change", () => updateVrpUploadMode());
vrpFiles.addEventListener("change", () => {
  showSelectedFiles(vrpFiles, vrpFileList);
  extractRefCodesButton.disabled = !vrpFiles.files.length;
  extractRefCodesStatus.textContent = "";
  delete extractRefCodesStatus.dataset.state;
});
extractRefCodesButton.addEventListener("click", async () => {
  extractRefCodesButton.disabled = true;
  extractRefCodesButton.textContent = "Extracting...";
  extractRefCodesStatus.textContent = "";
  delete extractRefCodesStatus.dataset.state;
  try {
    const codes = await extractVrpReferenceCodes();
    await copyToClipboard(codes.join("\n"));
    extractRefCodesStatus.textContent = `Copied ${codes.length.toLocaleString()} unique REF CODE(s).`;
    extractRefCodesStatus.dataset.state = "success";
  } catch (error) {
    extractRefCodesStatus.textContent = error instanceof Error ? error.message : String(error);
    extractRefCodesStatus.dataset.state = "error";
  } finally {
    extractRefCodesButton.textContent = "Get REF CODE";
    extractRefCodesButton.disabled = !vrpFiles.files.length;
  }
});
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
  const file = document.getElementById("field-file").files[0];
  const estimate = file ? {
    totalMs: estimateFieldResultDuration(file.size),
    onSuccess: (duration) => {
      fieldResultTimings.push({ duration, size: file.size });
      if (fieldResultTimings.length > 5) fieldResultTimings.shift();
    }
  } : undefined;
  runProcess(event.currentTarget, "field-notice", processFieldResult, estimate);
});
document.getElementById("esign-form").addEventListener("submit", (event) => {
  event.preventDefault();
  runProcess(event.currentTarget, "esign-notice", processSignatures);
});
document.getElementById("vrp-reset").addEventListener("click", () => {
  window.setTimeout(() => {
    updateVrpUploadMode();
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
