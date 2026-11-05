import assert from "node:assert/strict";
import { once } from "node:events";
import { readFileSync } from "node:fs";
import path from "node:path";
import test from "node:test";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import { createAppServer, parseEnvFile } from "../server.mjs";

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const env = {
  TUNNEL_ACCESS_MODE: "private",
  LARK_APP_ID: "test-app-id",
  LARK_APP_SECRET: "test-app-secret",
  LARK_WIKI_TOKEN: "test-wiki-token",
  LARK_TABLE_ID: "test-table-id",
  LARK_VIEW_ID: "test-view-id",
  PUBLIC_APP_ORIGIN: "https://private.example.test"
};

async function withServer(action) {
  const server = createAppServer(env);
  server.listen(0, "127.0.0.1");
  await once(server, "listening");
  const address = server.address();
  try {
    await action(`http://127.0.0.1:${address.port}`);
  } finally {
    server.close();
    await once(server, "close");
  }
}

async function withMockedFetch(mock, action) {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = mock;
  try {
    await action(originalFetch);
  } finally {
    globalThis.fetch = originalFetch;
  }
}

function json(body) {
  return new Response(JSON.stringify(body), {
    status: 200,
    headers: { "Content-Type": "application/json" }
  });
}

function makeXlsxWorkbook() {
  const context = vm.createContext({});
  vm.runInContext(
    readFileSync(path.join(ROOT, "vendor", "xlsx.full.min.js"), "utf8"),
    context,
    { filename: "xlsx.full.min.js" }
  );
  const workbook = context.XLSX.utils.book_new();
  context.XLSX.utils.book_append_sheet(
    workbook,
    context.XLSX.utils.aoa_to_sheet([["Other sheet"]]),
    "Summary"
  );
  const sourceRow = Array(42).fill("");
  sourceRow[41] = "REF-100";
  sourceRow[26] = "Field User";
  const resultRows = [Array(42).fill("Header"), sourceRow];
  context.XLSX.utils.book_append_sheet(
    workbook,
    context.XLSX.utils.aoa_to_sheet(resultRows),
    " RESULT "
  );
  return Buffer.from(context.XLSX.write(workbook, { bookType: "xlsx", type: "array" }));
}

test("parses local environment values without requiring a dotenv dependency", () => {
  assert.deepEqual(parseEnvFile([
    "# local-only settings",
    "TUNNEL_ACCESS_MODE=private",
    "LARK_WIKI_TOKEN=wiki-token"
  ].join("\n")), {
    TUNNEL_ACCESS_MODE: "private",
    LARK_WIKI_TOKEN: "wiki-token"
  });
  assert.throws(() => parseEnvFile("invalid entry"), /line 1/);
});

test("refuses to start if the tunnel is not explicitly configured as private", async () => {
  const server = createAppServer({ ...env, TUNNEL_ACCESS_MODE: "public" });
  server.listen(0, "127.0.0.1");
  await once(server, "listening");
  const address = server.address();
  try {
    const response = await fetch(`http://127.0.0.1:${address.port}`);
    assert.equal(response.status, 503);
    assert.match(await response.text(), /TUNNEL_ACCESS_MODE=private/);
  } finally {
    server.close();
    await once(server, "close");
  }
});

test("serves the app without an app-level sign-in when tunnel mode is private", async () => {
  await withServer(async (baseUrl) => {
    const response = await fetch(baseUrl);
    assert.equal(response.status, 200);
    assert.match(await response.text(), /Workspace Tools V2/);

    const appScript = await fetch(`${baseUrl}/app.js?v=19`);
    assert.equal(appScript.status, 200);
    assert.equal(appScript.headers.get("Cache-Control"), "no-store");
    assert.match(appScript.headers.get("Content-Type"), /^text\/javascript/);
    assert.match(await appScript.text(), /new Worker\("field-result-worker\.js\?v=4"\)/);

    const xlsxLibrary = await fetch(`${baseUrl}/vendor/xlsx.full.min.js`);
    assert.equal(xlsxLibrary.status, 200);
    assert.match(xlsxLibrary.headers.get("Content-Type"), /^text\/javascript/);
    assert.match(await xlsxLibrary.text(), /XLSX/);

    const fieldWorker = await fetch(`${baseUrl}/field-result-worker.js`);
    assert.equal(fieldWorker.status, 200);
    assert.match(fieldWorker.headers.get("Content-Type"), /^text\/javascript/);
    assert.equal(fieldWorker.headers.get("Cache-Control"), "no-store");
    assert.match(await fieldWorker.text(), /PREVIEW_ROW_LIMIT = 100/);
  });
});

test("fails closed when local Lark credentials are placeholders", async () => {
  const server = createAppServer({
    ...env,
    LARK_APP_SECRET: "PASTE_NEW_ROTATED_SECRET_HERE"
  });
  server.listen(0, "127.0.0.1");
  await once(server, "listening");
  const address = server.address();
  try {
    const response = await fetch(`http://127.0.0.1:${address.port}`);
    assert.equal(response.status, 503);
    assert.match(await response.text(), /\.dev\.vars/);
  } finally {
    server.close();
    await once(server, "close");
  }
});

test("serves the app without exposing the local secrets file", async () => {
  await withServer(async (baseUrl) => {
    const home = await fetch(baseUrl);
    assert.equal(home.status, 200);
    assert.match(await home.text(), /Workspace Tools V2/);

    const secretFile = await fetch(`${baseUrl}/.dev.vars`);
    assert.equal(secretFile.status, 404);
  });
});

test("serves the Lark API without a browser login and returns no secret configuration", async () => {
  await withMockedFetch(async (input) => {
    const url = new URL(input);
    if (url.pathname.endsWith("/auth/v3/tenant_access_token/internal")) {
      return json({ code: 0, tenant_access_token: "mock-token" });
    }
    if (url.pathname.endsWith("/wiki/v2/spaces/get_node")) {
      return json({ code: 0, data: { node: { obj_type: "bitable", obj_token: "mock-base" } } });
    }
    if (url.pathname.endsWith("/records/search")) {
      return json({ code: 0, data: { items: [{ fields: { Account: "123" } }], has_more: false } });
    }
    throw new Error(`Unexpected Lark URL: ${url}`);
  }, async (realFetch) => {
    await withServer(async (baseUrl) => {
      const response = await realFetch(`${baseUrl}/api/notifications`);
      assert.equal(response.status, 200);
      assert.equal((await response.json()).count, 1);
    });
  });
});

test("streams Field Result processing on the server and serves its output files", async () => {
  const workbook = makeXlsxWorkbook();
  await withServer(async (baseUrl) => {
    const response = await fetch(`${baseUrl}/api/field-result`, {
      method: "POST",
      headers: {
        "Content-Length": String(workbook.length),
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      },
      body: workbook
    });
    if (!response.ok) {
      throw new Error(`Field Result API returned ${response.status}: ${await response.text()}`);
    }
    const result = await response.json();
    assert.equal(result.rows, 2);
    assert.equal(result.previewRows[1][0], "REF-100");
    assert.equal(result.previewRows[1][7], "Field User");
    assert.match(result.workbookName, /^FIELD RESULT \d{2}-\d{2}-\d{4}\.xlsx$/);

    const csvResponse = await fetch(`${baseUrl}/api/field-result/${result.id}/csv`);
    assert.equal(csvResponse.status, 200);
    assert.equal(csvResponse.headers.get("Content-Disposition"), 'attachment; filename="fieldman.csv"');
    assert.match(await csvResponse.text(), /REF-100,Field User/);

    const workbookResponse = await fetch(`${baseUrl}/api/field-result/${result.id}/workbook`);
    assert.equal(workbookResponse.status, 200);
    assert.match(workbookResponse.headers.get("Content-Type"), /spreadsheetml\.sheet/);
    const outputBytes = Buffer.from(await workbookResponse.arrayBuffer());
    assert.ok(outputBytes.byteLength > 0);
    const outputContext = vm.createContext({});
    vm.runInContext(
      readFileSync(path.join(ROOT, "vendor", "xlsx.full.min.js"), "utf8"),
      outputContext,
      { filename: "xlsx.full.min.js" }
    );
    const outputBook = outputContext.XLSX.read(outputBytes, { type: "buffer" });
    const outputRows = outputContext.XLSX.utils.sheet_to_json(
      outputBook.Sheets.RESULT,
      { header: 1, defval: "" }
    );
    assert.equal(outputRows[1][0], "REF-100");
    assert.equal(outputRows[1][7], "Field User");
  });
});

test("rejects a non-XLSX upload with a clear client error", async () => {
  await withServer(async (baseUrl) => {
    const response = await fetch(`${baseUrl}/api/field-result`, {
      method: "POST",
      headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
      body: "not-an-xlsx"
    });
    assert.equal(response.status, 400);
    assert.match((await response.json()).error, /not a zip file/i);
  });
});
