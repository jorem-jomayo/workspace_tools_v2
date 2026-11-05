const LARK_API = "https://open.larksuite.com/open-apis";
const PAGE_SIZE = 500;
const MAX_EXPORT_PAGES = 48;
const EXPORT_FILENAME = "NEW VRP ACCOUNTS FOR TAG_PIF WITH DL_JOREM FOR UPLOAD.csv";
const EXPORT_HEADERS = [
  "FIELD REMARKS",
  "ACCOUNT NUMBER",
  "TYPE OF ACCOUNT",
  "BANK",
  "PLACEMENT",
  "CH CODE",
  "CH NAME",
  "ADD TYPE",
  "ADDRESS",
  "AREA",
  "MUNICIPALITY",
  "DL TYPE",
  "REF CODE",
  "AUTOFIELD DATE",
  "OB/PRINCIPAL",
  "ENDO DATE",
  "PULLOUT DATE",
  "PRINT STATUS",
  "RELEASE STATUS",
  "RELEASED TO",
  "UNIQ ID V2",
  "OPEN CI V2 UPLOAD",
  "DATE UPLOADED",
  "RESULT",
  "RESULT DATE",
  "REPORTED IN",
  "FINAL AREA",
  "AREA CLUSTER",
  "RELEASE DATE",
  "PRINTED DATE",
  "PULLOUT REMARKS",
  "REPORTED",
  "UNIQ",
  "FINAL FINAL AREA"
];

function allowedOrigins(env) {
  return String(env.ALLOWED_ORIGINS || "")
    .split(",")
    .map((origin) => origin.trim())
    .filter(Boolean);
}

function corsHeaders(origin) {
  return {
    "Access-Control-Allow-Origin": origin,
    "Access-Control-Allow-Methods": "GET, OPTIONS",
    "Access-Control-Allow-Headers": "Authorization, Content-Type",
    "Access-Control-Max-Age": "86400",
    "Vary": "Origin"
  };
}

function jsonResponse(data, status, origin) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      ...corsHeaders(origin),
      "Cache-Control": "no-store",
      "Content-Type": "application/json; charset=utf-8"
    }
  });
}

function secureEqual(left, right) {
  if (left.length !== right.length) return false;
  let difference = 0;
  for (let index = 0; index < left.length; index += 1) {
    difference |= left.charCodeAt(index) ^ right.charCodeAt(index);
  }
  return difference === 0;
}

function requiredSettings(env) {
  const required = [
    "LARK_APP_ID",
    "LARK_APP_SECRET",
    "LARK_WIKI_TOKEN",
    "LARK_TABLE_ID",
    "LARK_VIEW_ID",
    "DASHBOARD_ACCESS_KEY"
  ];
  return required.filter((name) => !String(env[name] || "").trim());
}

async function readJson(response, action) {
  let body;
  try {
    body = await response.json();
  } catch {
    throw new Error(`${action} returned an invalid response (HTTP ${response.status}).`);
  }
  if (!response.ok || (typeof body.code === "number" && body.code !== 0)) {
    const code = body.code ?? response.status;
    const message = String(body.msg || body.message || "Request failed").slice(0, 300);
    if (String(code) === "1254302" || message.includes("RolePermNotAllow")) {
      throw new Error(
        `${action} failed (${code}): The Lark app is missing permission for this Bitable. In the Lark Developer Console, enable and publish wiki:wiki:readonly and bitable:app:readonly (or base:record:retrieve). Then share the Bitable with the app and grant it read access to the table/view. If advanced permissions are on, grant the app the required Base role too.`
      );
    }
    throw new Error(`${action} failed (${code}): ${message}`);
  }
  return body;
}

async function getTenantAccessToken(env) {
  const response = await fetch(`${LARK_API}/auth/v3/tenant_access_token/internal`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      app_id: env.LARK_APP_ID,
      app_secret: env.LARK_APP_SECRET
    })
  });
  const body = await readJson(response, "Lark authentication");
  if (!body.tenant_access_token) {
    throw new Error("Lark authentication succeeded without an access token.");
  }
  return body.tenant_access_token;
}

async function getBitableAppToken(accessToken, wikiToken) {
  const url = new URL(`${LARK_API}/wiki/v2/spaces/get_node`);
  url.searchParams.set("token", wikiToken);
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` }
  });
  const body = await readJson(response, "Lark wiki lookup");
  const node = body.data?.node;
  if (!node?.obj_token || node.obj_type !== "bitable") {
    throw new Error("The configured Lark wiki link does not resolve to a Bitable.");
  }
  return node.obj_token;
}

async function fetchRecordPage(accessToken, appToken, env, pageToken) {
  const url = new URL(
    `${LARK_API}/bitable/v1/apps/${encodeURIComponent(appToken)}` +
    `/tables/${encodeURIComponent(env.LARK_TABLE_ID)}/records/search`
  );
  url.searchParams.set("page_size", String(PAGE_SIZE));
  if (pageToken) url.searchParams.set("page_token", pageToken);
  const response = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({ view_id: env.LARK_VIEW_ID })
  });
  const body = await readJson(response, "Lark table query");
  if (!Array.isArray(body.data?.items)) {
    throw new Error("Lark table query returned an invalid records list.");
  }
  return body.data;
}

function csvCell(value) {
  let text = value == null
    ? ""
    : typeof value === "string"
      ? value
      : JSON.stringify(value);
  if (/^[\t\r ]*[=+@]/.test(text) ||
      /^[\t\r ]*-(?!\d+(?:\.\d+)?(?:[eE][+-]?\d+)?$)/.test(text)) {
    text = `'${text}`;
  }
  return `"${text.replace(/"/g, '""')}"`;
}

const DATE_FORMATS = {
  "AUTOFIELD DATE": "datetime",
  "ENDO DATE": "slash-date",
  "PULLOUT DATE": "slash-date",
  "DATE UPLOADED": "us-date",
  "RESULT DATE": "us-date",
  "RELEASE DATE": "us-date",
  "PRINTED DATE": "us-date"
};

function formatLarkDate(value, format) {
  const timestamp = typeof value === "number"
    ? value
    : typeof value === "string" && /^\d{13}$/.test(value)
      ? Number(value)
      : NaN;
  if (!Number.isFinite(timestamp)) return String(value ?? "");

  const parts = Object.fromEntries(new Intl.DateTimeFormat("en-GB", {
    timeZone: "Asia/Manila",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hourCycle: "h23"
  }).formatToParts(new Date(timestamp)).map(({ type, value: part }) => [type, part]));
  if (format === "datetime") {
    return `${parts.year}/${parts.month}/${parts.day} ${parts.hour}:${parts.minute}`;
  }
  if (format === "slash-date") return `${parts.year}/${parts.month}/${parts.day}`;
  return `${parts.month}/${parts.day}/${parts.year}`;
}

function larkCellValue(value, header) {
  const dateFormat = DATE_FORMATS[header];
  if (dateFormat && (typeof value === "number" ||
      typeof value === "string" && /^\d{13}$/.test(value))) {
    return formatLarkDate(value, dateFormat);
  }
  if (value == null) return "";
  if (Array.isArray(value)) {
    return value.map((item) => larkCellValue(item, header))
      .filter(Boolean)
      .join(", ");
  }
  if (typeof value === "object") {
    if (dateFormat && Array.isArray(value.value) && value.value.length === 1) {
      return formatLarkDate(value.value[0], dateFormat);
    }
    if (typeof value.text === "string") return value.text;
    if (typeof value.name === "string") return value.name;
    if (typeof value.en_name === "string") return value.en_name;
    if (value.value !== undefined) return larkCellValue(value.value, header);
    if (typeof value.link === "string") return value.link;
    return "";
  }
  return String(value);
}

function recordsToCsv(records) {
  const lines = [
    EXPORT_HEADERS.map(csvCell).join(","),
    ...records.map((record) =>
      EXPORT_HEADERS.map((header) =>
        csvCell(larkCellValue(record.fields?.[header], header))
      ).join(",")
    )
  ];
  return `\uFEFF${lines.join("\r\n")}\r\n`;
}

function hasValue(value) {
  if (value == null) return false;
  if (typeof value === "string") return value.trim().length > 0;
  if (typeof value === "number" || typeof value === "boolean") return true;
  if (Array.isArray(value)) return value.some(hasValue);
  if (typeof value === "object") {
    return Object.entries(value)
      .filter(([key]) => key !== "type" && key !== "field_type")
      .some(([, nestedValue]) => hasValue(nestedValue));
  }
  return false;
}

async function fetchAllRecords(accessToken, appToken, env) {
  const records = [];
  const seenPageTokens = new Set();
  let page = await fetchRecordPage(accessToken, appToken, env);
  records.push(...page.items);
  let pageCount = 1;
  while (page.has_more) {
    const pageToken = page.page_token;
    if (!pageToken || seenPageTokens.has(pageToken)) {
      throw new Error("Lark pagination did not return a valid next-page token.");
    }
    if (pageCount >= MAX_EXPORT_PAGES) {
      throw new Error(`Lark view exceeds the ${MAX_EXPORT_PAGES * PAGE_SIZE}-record Worker limit. Narrow the view and try again.`);
    }
    seenPageTokens.add(pageToken);
    page = await fetchRecordPage(accessToken, appToken, env, pageToken);
    records.push(...page.items);
    pageCount += 1;
  }
  return records;
}

async function handleApiRequest(request, env, origin, path) {
  const missingSettings = requiredSettings(env);
  if (missingSettings.length) {
    return jsonResponse({
      error: `Worker configuration is incomplete: ${missingSettings.join(", ")}.`
    }, 503, origin);
  }

  const authorization = request.headers.get("Authorization") || "";
  const providedKey = authorization.startsWith("Bearer ")
    ? authorization.slice(7)
    : "";
  if (!secureEqual(providedKey, env.DASHBOARD_ACCESS_KEY)) {
    return jsonResponse({ error: "Access key is invalid or expired." }, 401, origin);
  }

  try {
    const accessToken = await getTenantAccessToken(env);
    const appToken = await getBitableAppToken(accessToken, env.LARK_WIKI_TOKEN);
    if (path === "/api/notifications") {
      const records = await fetchAllRecords(accessToken, appToken, env);
      const count = records.filter((record) =>
        Object.values(record.fields || {}).some(hasValue)
      ).length;
      return jsonResponse({ count, checkedAt: new Date().toISOString() }, 200, origin);
    }

    const records = await fetchAllRecords(accessToken, appToken, env);
    const csv = recordsToCsv(records);
    return new Response(csv, {
      status: 200,
      headers: {
        ...corsHeaders(origin),
        "Cache-Control": "no-store",
        "Content-Disposition": `attachment; filename="${EXPORT_FILENAME}"`,
        "Content-Type": "text/csv; charset=utf-8"
      }
    });
  } catch (error) {
    return jsonResponse({ error: error.message || "Lark request failed." }, 502, origin);
  }
}

export default {
  async fetch(request, env) {
    const origin = request.headers.get("Origin") || "";
    if (!origin || !allowedOrigins(env).includes(origin)) {
      return new Response("Origin is not allowed.", { status: 403 });
    }
    if (request.method === "OPTIONS") {
      return new Response(null, { status: 204, headers: corsHeaders(origin) });
    }
    if (request.method !== "GET") {
      return jsonResponse({ error: "Method not allowed." }, 405, origin);
    }
    const path = new URL(request.url).pathname;
    if (path !== "/api/notifications" && path !== "/api/export") {
      return jsonResponse({ error: "Not found." }, 404, origin);
    }
    return handleApiRequest(request, env, origin, path);
  }
};
