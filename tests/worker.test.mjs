import assert from "node:assert/strict";
import test from "node:test";
import worker from "../worker.mjs";

const origin = "https://example.github.io";
const env = {
  ALLOWED_ORIGINS: `http://localhost:8001,${origin}`,
  LARK_APP_ID: "test-app-id",
  LARK_APP_SECRET: "test-app-secret",
  LARK_WIKI_TOKEN: "test-wiki-token",
  LARK_TABLE_ID: "test-table-id",
  LARK_VIEW_ID: "test-view-id",
  DASHBOARD_ACCESS_KEY: "test-dashboard-key"
};

function request(path, { key = env.DASHBOARD_ACCESS_KEY, requestOrigin = origin } = {}) {
  const headers = { Origin: requestOrigin };
  if (key) headers.Authorization = `Bearer ${key}`;
  return new Request(`https://worker.example${path}`, { headers });
}

async function withMockedFetch(mock, action) {
  const originalFetch = globalThis.fetch;
  globalThis.fetch = mock;
  try {
    return await action();
  } finally {
    globalThis.fetch = originalFetch;
  }
}

function json(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { "Content-Type": "application/json" }
  });
}

test("rejects unauthorized calls before contacting Lark", async () => {
  let fetchCalls = 0;
  await withMockedFetch(() => {
    fetchCalls += 1;
    throw new Error("Lark should not be called");
  }, async () => {
    const response = await worker.fetch(request("/api/notifications", { key: "wrong" }), env);
    assert.equal(response.status, 401);
    assert.equal((await response.json()).error, "Access key is invalid or expired.");
  });
  assert.equal(fetchCalls, 0);
});

test("counts non-empty records across the configured Lark view", async () => {
  await withMockedFetch(async (input, init = {}) => {
    const url = new URL(input);
    if (url.pathname.endsWith("/auth/v3/tenant_access_token/internal")) {
      return json({ code: 0, tenant_access_token: "test-tenant-token" });
    }
    if (url.pathname.endsWith("/wiki/v2/spaces/get_node")) {
      assert.equal(url.searchParams.get("token"), env.LARK_WIKI_TOKEN);
      return json({ code: 0, data: { node: { obj_type: "bitable", obj_token: "test-app-token" } } });
    }
    if (url.pathname.endsWith("/records/search")) {
      assert.equal(init.method, "POST");
      assert.deepEqual(JSON.parse(init.body), { view_id: env.LARK_VIEW_ID });
      assert.equal(url.searchParams.get("page_size"), "500");
      if (!url.searchParams.has("page_token")) {
        return json({
          code: 0,
          data: {
            total: 3,
            has_more: true,
            page_token: "next-page",
            items: [
              { fields: { Name: "Ana" } },
              { fields: { Name: " ", Note: { type: "text", text: "" } } }
            ]
          }
        });
      }
      assert.equal(url.searchParams.get("page_token"), "next-page");
      return json({
        code: 0,
        data: { total: 3, has_more: false, items: [{ fields: { Name: "Ben" } }] }
      });
    }
    throw new Error(`Unexpected Lark URL: ${url}`);
  }, async () => {
    const response = await worker.fetch(request("/api/notifications"), env);
    assert.equal(response.status, 200);
    const result = await response.json();
    assert.equal(result.count, 2);
    assert.ok(result.checkedAt);
  });
});

test("exports every page as quoted, formula-safe CSV", async () => {
  const pageTokens = [];
  const expectedHeaders = [
    "FIELD REMARKS", "ACCOUNT NUMBER", "TYPE OF ACCOUNT", "BANK", "PLACEMENT",
    "CH CODE", "CH NAME", "ADD TYPE", "ADDRESS", "AREA", "MUNICIPALITY",
    "DL TYPE", "REF CODE", "AUTOFIELD DATE", "OB/PRINCIPAL", "ENDO DATE",
    "PULLOUT DATE", "PRINT STATUS", "RELEASE STATUS", "RELEASED TO",
    "UNIQ ID V2", "OPEN CI V2 UPLOAD", "DATE UPLOADED", "RESULT",
    "RESULT DATE", "REPORTED IN", "FINAL AREA", "AREA CLUSTER", "RELEASE DATE",
    "PRINTED DATE", "PULLOUT REMARKS", "REPORTED", "UNIQ", "FINAL FINAL AREA"
  ];
  const csvRow = (fields) => expectedHeaders.map((header) => {
    let value = fields[header] || "";
    if (/^[\t\r ]*[=+@]/.test(value) ||
        /^[\t\r ]*-(?!\d+(?:\.\d+)?(?:[eE][+-]?\d+)?$)/.test(value)) {
      value = `'${value}`;
    }
    return `"${value.replace(/"/g, '""')}"`;
  }).join(",");
  const firstFields = {
    "ACCOUNT NUMBER": [{ text: "367002981510", type: "text" }],
    "FIELD REMARKS": [{ text: "REVISIT-REPROCESS", type: "text" }],
    BANK: "PIF HOME LOAN",
    "CH CODE": [{ text: "04PCRL2609-4", type: "text" }],
    "CH NAME": [{ text: "IMPAS, GIOVANNI MOSES CARTAGENA", type: "text" }],
    ADDRESS: [{ text: "LISTON CMPD ML QUEZON ST BRGY. CABANCALAN MANDAUE CITY CEBU 6014", type: "text" }],
    AREA: [{ text: "CEBU", type: "text" }],
    "REF CODE": [{ text: "p2DL2611-814", type: "text" }],
    "AUTOFIELD DATE": { type: 5, value: [1793635200000] },
    "PULLOUT DATE": { type: 5, value: [1796659200000] },
    "PRINT STATUS": { type: 1, value: [{ text: "PRINTED", type: "text" }] },
    "RELEASE STATUS": { type: 1, value: [{ text: "RELEASED", type: "text" }] },
    "RELEASED TO": { type: 11, value: [{ name: "Robert Jay Miego" }] },
    "UNIQ ID V2": { type: 1, value: [{ text: "7be16dc8-0ed4-441e-9009-3c0c8a411ff0" }] },
    "FINAL AREA": "CEBU",
    "AREA CLUSTER": { type: 3, value: ["VISAYAS"] },
    "RELEASE DATE": { type: 5, value: [1793721600000] },
    "PRINTED DATE": { type: 5, value: [1793721600000] },
    REPORTED: { type: 1, value: [{ text: "CEBUMANDAUE" }] },
    "FINAL FINAL AREA": { type: 1, value: [{ text: "CEBU" }] },
    "EXTRA FIELD": "not in upload template"
  };
  const expectedFirstFields = {
    "ACCOUNT NUMBER": "367002981510",
    "FIELD REMARKS": "REVISIT-REPROCESS",
    BANK: "PIF HOME LOAN",
    "CH CODE": "04PCRL2609-4",
    "CH NAME": "IMPAS, GIOVANNI MOSES CARTAGENA",
    ADDRESS: "LISTON CMPD ML QUEZON ST BRGY. CABANCALAN MANDAUE CITY CEBU 6014",
    AREA: "CEBU",
    "REF CODE": "p2DL2611-814",
    "AUTOFIELD DATE": "2026/11/03 00:00",
    "PULLOUT DATE": "2026/12/08",
    "PRINT STATUS": "PRINTED",
    "RELEASE STATUS": "RELEASED",
    "RELEASED TO": "Robert Jay Miego",
    "UNIQ ID V2": "7be16dc8-0ed4-441e-9009-3c0c8a411ff0",
    "FINAL AREA": "CEBU",
    "AREA CLUSTER": "VISAYAS",
    "RELEASE DATE": "11/04/2026",
    "PRINTED DATE": "11/04/2026",
    REPORTED: "CEBUMANDAUE",
    "FINAL FINAL AREA": "CEBU"
  };
  const secondFields = {
    "ACCOUNT NUMBER": "367002981511",
    "CH NAME": "=1+1",
    ADDRESS: 'said "hello"'
  };
  await withMockedFetch(async (input) => {
    const url = new URL(input);
    if (url.pathname.endsWith("/auth/v3/tenant_access_token/internal")) {
      return json({ code: 0, tenant_access_token: "test-tenant-token" });
    }
    if (url.pathname.endsWith("/wiki/v2/spaces/get_node")) {
      return json({ code: 0, data: { node: { obj_type: "bitable", obj_token: "test-app-token" } } });
    }
    if (url.pathname.endsWith("/records/search")) {
      const pageToken = url.searchParams.get("page_token");
      pageTokens.push(pageToken);
      if (!pageToken) {
        return json({
          code: 0,
          data: {
            total: 2,
            has_more: true,
            page_token: "next-page",
            items: [{ fields: firstFields }]
          }
        });
      }
      return json({
        code: 0,
        data: {
          total: 2,
          has_more: false,
          items: [{ fields: secondFields }]
        }
      });
    }
    throw new Error(`Unexpected Lark URL: ${url}`);
  }, async () => {
    const response = await worker.fetch(request("/api/export"), env);
    assert.equal(response.status, 200);
    assert.match(response.headers.get("Content-Type"), /text\/csv/);
    assert.equal(
      response.headers.get("Content-Disposition"),
      'attachment; filename="NEW VRP ACCOUNTS FOR TAG_PIF WITH DL_JOREM FOR UPLOAD.csv"'
    );
    const csv = await response.text();
    const [headerLine, firstDataLine, secondDataLine] = csv.split("\r\n");
    assert.equal(headerLine, expectedHeaders.map((header) => `"${header}"`).join(","));
    assert.equal(firstDataLine, csvRow(expectedFirstFields));
    assert.equal(secondDataLine, csvRow(secondFields));
  });
  assert.deepEqual(pageTokens, [null, "next-page"]);
});

test("explains how to fix Lark RolePermNotAllow errors", async () => {
  await withMockedFetch(async (input) => {
    const url = new URL(input);
    if (url.pathname.endsWith("/auth/v3/tenant_access_token/internal")) {
      return json({ code: 0, tenant_access_token: "test-tenant-token" });
    }
    if (url.pathname.endsWith("/wiki/v2/spaces/get_node")) {
      return json({ code: 0, data: { node: { obj_type: "bitable", obj_token: "test-app-token" } } });
    }
    if (url.pathname.endsWith("/records/search")) {
      return json({ code: 1254302, msg: "RolePermNotAllow" });
    }
    throw new Error(`Unexpected Lark URL: ${url}`);
  }, async () => {
    const response = await worker.fetch(request("/api/notifications"), env);
    assert.equal(response.status, 502);
    const result = await response.json();
    assert.match(result.error, /Lark app is missing permission for this Bitable/);
    assert.match(result.error, /share the Bitable with the app/);
    assert.match(result.error, /1254302/);
  });
});

test("rejects calls from unlisted origins", async () => {
  const response = await worker.fetch(
    request("/api/notifications", { requestOrigin: "https://attacker.example" }),
    env
  );
  assert.equal(response.status, 403);
});
