import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";
import vm from "node:vm";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");

async function runFieldWorker(rows, sheetName = "RESULT") {
  const messages = [];
  const context = vm.createContext({
    ArrayBuffer,
    Blob,
    Date,
    Error,
    Intl,
    JSON,
    Map,
    Math,
    Number,
    Object,
    Promise,
    RegExp,
    Set,
    String,
    TextDecoder,
    TextEncoder,
    Uint8Array,
    clearTimeout,
    console,
    setTimeout
  });
  context.self = context;
  for (const library of ["xlsx.full.min.js", "jszip.min.js"]) {
    const source = readFileSync(path.join(ROOT, "vendor", library), "utf8");
    vm.runInContext(source, context, { filename: library });
  }
  context.self.postMessage = (message) => messages.push(message);
  context.importScripts = () => {};
  const workerSource = readFileSync(path.join(ROOT, "field-result-worker.js"), "utf8");
  vm.runInContext(workerSource, context, { filename: "field-result-worker.js" });
  const workbook = context.XLSX.utils.book_new();
  context.XLSX.utils.book_append_sheet(workbook, context.XLSX.utils.aoa_to_sheet(rows), sheetName);
  const fileData = context.XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  return context.self.onmessage({ data: { file: { arrayBuffer: async () => fileData } } })
    .then(() => messages);
}

test("processes large Field Result row sets off the main thread with a bounded preview", async () => {
  const rows = [Array(42).fill("Header")];
  for (let index = 1; index < 105; index += 1) {
    const row = Array(42).fill("");
    row[41] = `REF-${index}`;
    row[26] = `Field ${index}`;
    rows.push(row);
  }
  const messages = await runFieldWorker(rows);
  const failure = messages.find((message) => message.type === "error");
  assert.equal(failure, undefined, failure?.message);

  const result = messages.find((message) => message.type === "complete");
  assert.ok(result);
  assert.equal(result.totalRows, 105);
  assert.equal(result.previewRows.length, 100);
  assert.equal(
    JSON.stringify(result.previewRows[1]),
    JSON.stringify(["REF-1", "", "", "", "", "", "", "Field 1"])
  );
  const csv = await result.extraCsvBlob.text();
  assert.match(csv, /"Ref Code","Field Name"\r\nREF-1,Field 1/);
  assert.match(csv, /REF-104,Field 104\r\n$/);
});

test("finds RESULT worksheet names with surrounding spaces", async () => {
  const rows = [Array(42).fill("")];
  const row = Array(42).fill("");
  row[41] = "REF-1";
  rows.push(row);

  const messages = await runFieldWorker(rows, " RESULT ");
  const failure = messages.find((message) => message.type === "error");
  assert.equal(failure, undefined, failure?.message);
  assert.equal(messages.find((message) => message.type === "complete")?.totalRows, 2);
});

test("parses a multi-sheet workbook only once to avoid duplicating large workbook memory", async () => {
  const messages = [];
  const context = vm.createContext({
    ArrayBuffer, Blob, Date, Error, Intl, JSON, Map, Math, Number, Object,
    Promise, RegExp, Set, String, TextDecoder, TextEncoder, Uint8Array,
    clearTimeout, console, setTimeout
  });
  context.self = context;
  for (const library of ["xlsx.full.min.js", "jszip.min.js"]) {
    vm.runInContext(
      readFileSync(path.join(ROOT, "vendor", library), "utf8"),
      context,
      { filename: library }
    );
  }
  const originalRead = context.XLSX.read;
  let readCount = 0;
  context.XLSX.read = (...args) => {
    readCount += 1;
    return originalRead(...args);
  };
  context.self.postMessage = (message) => messages.push(message);
  context.importScripts = () => {};
  vm.runInContext(
    readFileSync(path.join(ROOT, "field-result-worker.js"), "utf8"),
    context,
    { filename: "field-result-worker.js" }
  );
  const workbook = context.XLSX.utils.book_new();
  context.XLSX.utils.book_append_sheet(workbook, context.XLSX.utils.aoa_to_sheet([["other"]]), "Summary");
  const resultRows = [Array(42).fill("")];
  const resultRow = Array(42).fill("");
  resultRow[41] = "REF-1";
  resultRows.push(resultRow);
  context.XLSX.utils.book_append_sheet(workbook, context.XLSX.utils.aoa_to_sheet(resultRows), "RESULT");
  const fileData = context.XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  await context.self.onmessage({ data: { file: { arrayBuffer: async () => fileData } } });

  assert.equal(messages.find((message) => message.type === "error"), undefined);
  assert.ok(messages.find((message) => message.type === "complete"));
  assert.equal(readCount, 1);
});

test("reports a missing RESULT sheet instead of a worksheet property error", async () => {
  const messages = await runFieldWorker([Array(42).fill("")], "SUMMARY");
  assert.match(messages.find((message) => message.type === "error")?.message || "", /sheet named RESULT was not found/i);
});
