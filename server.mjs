import { randomBytes, randomUUID } from "node:crypto";
import { spawn } from "node:child_process";
import { createServer } from "node:http";
import { createReadStream, createWriteStream } from "node:fs";
import { mkdtemp, readFile, rm, stat } from "node:fs/promises";
import { Transform } from "node:stream";
import { pipeline } from "node:stream/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import worker from "./worker.mjs";

const ROOT = path.dirname(fileURLToPath(import.meta.url));
const DEFAULT_PUBLIC_ORIGIN = "https://cfbbqbh6-8001.asse.devtunnels.ms";
const INTERNAL_ACCESS_KEY = randomBytes(32).toString("hex");
const MAX_FIELD_RESULT_UPLOAD_BYTES = 512 * 1024 * 1024;
const FIELD_RESULT_JOB_TTL_MS = 30 * 60 * 1000;
const STATIC_FILES = new Map([
  ["/", ["index.html", "text/html; charset=utf-8"]],
  ["/index.html", ["index.html", "text/html; charset=utf-8"]],
  ["/app.js", ["app.js", "text/javascript; charset=utf-8"]],
  ["/field-result-worker.js", ["field-result-worker.js", "text/javascript; charset=utf-8"]],
  ["/demand_letter_template.csv", ["demand_letter_template.csv", "text/csv; charset=utf-8"]],
  ["/cmd_id.xlsx", ["cmd_id.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"]],
  ["/amounts.csv", ["amounts.csv", "text/csv; charset=utf-8"]]
]);
const MIME_TYPES = new Map([
  [".js", "text/javascript; charset=utf-8"],
  [".svg", "image/svg+xml"],
  [".woff2", "font/woff2"],
  [".txt", "text/plain; charset=utf-8"],
]);

export function parseEnvFile(source) {
  const values = {};
  for (const [index, rawLine] of source.split(/\r?\n/).entries()) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) continue;
    const match = line.match(/^([A-Z][A-Z0-9_]*)=(.*)$/);
    if (!match) throw new Error(`Invalid .dev.vars entry at line ${index + 1}.`);
    let value = match[2].trim();
    if ((value.startsWith('"') && value.endsWith('"')) ||
        (value.startsWith("'") && value.endsWith("'"))) {
      value = value.slice(1, -1);
    }
    values[match[1]] = value;
  }
  return values;
}

export async function loadLocalEnv(root = ROOT) {
  try {
    const source = await readFile(path.join(root, ".dev.vars"), "utf8");
    return parseEnvFile(source);
  } catch (error) {
    if (error.code === "ENOENT") return {};
    throw error;
  }
}

function isConfigured(env) {
  const required = [
    "TUNNEL_ACCESS_MODE",
    "LARK_APP_ID",
    "LARK_APP_SECRET",
    "LARK_WIKI_TOKEN",
    "LARK_TABLE_ID",
    "LARK_VIEW_ID"
  ];
  return required.every((name) => {
    const value = String(env[name] || "").trim();
    return value && !/^(?:replace-with|paste_|choose_)/i.test(value);
  }) && env.TUNNEL_ACCESS_MODE === "private";
}

function sendText(response, status, message, headers = {}) {
  response.writeHead(status, {
    "Cache-Control": "no-store",
    "Content-Type": "text/plain; charset=utf-8",
    "X-Content-Type-Options": "nosniff",
    ...headers
  });
  response.end(message);
}

function sendJson(response, status, value) {
  response.writeHead(status, {
    "Cache-Control": "no-store",
    "Content-Type": "application/json; charset=utf-8",
    "X-Content-Type-Options": "nosniff"
  });
  response.end(JSON.stringify(value));
}

async function saveFieldResultUpload(request, targetPath) {
  let bytesWritten = 0;
  const limit = new Transform({
    transform(chunk, encoding, callback) {
      bytesWritten += chunk.length;
      if (bytesWritten > MAX_FIELD_RESULT_UPLOAD_BYTES) {
        const error = new Error("The workbook is larger than the 512 MB upload limit.");
        error.statusCode = 413;
        callback(error);
      } else {
        callback(null, chunk);
      }
    }
  });
  await pipeline(request, limit, createWriteStream(targetPath, { flags: "wx" }));
}

function runFieldResultProcessor(inputPath, outputDir) {
  const python = process.env.PYTHON_EXECUTABLE || "python";
  const scriptPath = path.join(ROOT, "field_result_processor.py");
  return new Promise((resolve, reject) => {
    const child = spawn(python, [scriptPath, inputPath, outputDir], {
      cwd: ROOT,
      windowsHide: true,
      stdio: ["ignore", "pipe", "pipe"]
    });
    let stdout = "";
    let stderr = "";
    child.stdout.setEncoding("utf8");
    child.stderr.setEncoding("utf8");
    child.stdout.on("data", (chunk) => {
      stdout = `${stdout}${chunk}`.slice(-1024 * 1024);
    });
    child.stderr.on("data", (chunk) => {
      stderr = `${stderr}${chunk}`.slice(-16 * 1024);
    });
    child.once("error", (error) => {
      reject(new Error(`Could not start Python for Field Result processing: ${error.message}`));
    });
    child.once("close", (code) => {
      if (code !== 0) {
        reject(new Error(stderr.trim() || `Field Result processor exited with code ${code}.`));
        return;
      }
      try {
        resolve(JSON.parse(stdout.trim()));
      } catch {
        reject(new Error("Field Result processor returned invalid result metadata."));
      }
    });
  });
}

async function cleanupExpiredFieldResultJobs(jobs) {
  const expiredBefore = Date.now() - FIELD_RESULT_JOB_TTL_MS;
  for (const [id, job] of jobs) {
    if (job.lastAccessedAt >= expiredBefore) continue;
    jobs.delete(id);
    clearTimeout(job.cleanupTimer);
    await rm(job.directory, { recursive: true, force: true });
  }
}

async function processFieldResultUpload(request, response, jobs) {
  const declaredLength = Number(request.headers["content-length"] || 0);
  if (declaredLength > MAX_FIELD_RESULT_UPLOAD_BYTES) {
    request.resume();
    sendJson(response, 413, { error: "The workbook is larger than the 512 MB upload limit." });
    return;
  }

  const directory = await mkdtemp(path.join(os.tmpdir(), "workspace-field-result-"));
  const inputPath = path.join(directory, "input.xlsx");
  try {
    await saveFieldResultUpload(request, inputPath);
    const result = await runFieldResultProcessor(inputPath, directory);
    await rm(inputPath, { force: true });
    const id = randomUUID();
    const job = {
      directory,
      lastAccessedAt: Date.now(),
      workbookName: result.workbookName
    };
    job.cleanupTimer = setTimeout(() => {
      if (jobs.get(id) !== job) return;
      jobs.delete(id);
      rm(directory, { recursive: true, force: true })
        .catch((error) => console.error("Could not clean up expired Field Result output:", error.message));
    }, FIELD_RESULT_JOB_TTL_MS);
    job.cleanupTimer.unref?.();
    jobs.set(id, job);
    sendJson(response, 200, {
      id,
      rows: result.rows,
      previewRows: result.previewRows,
      workbookName: result.workbookName
    });
  } catch (error) {
    await rm(directory, { recursive: true, force: true });
    const status = error.statusCode || (
      /RESULT|worksheet|workbook|zip file|not a zip/i.test(error.message) ? 400 : 500
    );
    sendJson(response, status, {
      error: status === 500
        ? `Field Result could not process this workbook: ${error.message}`
        : error.message
    });
  }
}

async function downloadFieldResultFile(response, jobs, id, type) {
  const job = jobs.get(id);
  if (!job) {
    sendJson(response, 404, { error: "This Field Result output has expired. Process the workbook again." });
    return;
  }
  const isWorkbook = type === "workbook";
  const filename = isWorkbook ? job.workbookName : "fieldman.csv";
  const filePath = path.join(job.directory, filename);
  try {
    const fileInfo = await stat(filePath);
    job.lastAccessedAt = Date.now();
    response.writeHead(200, {
      "Cache-Control": "no-store",
      "Content-Disposition": `attachment; filename="${filename}"`,
      "Content-Length": fileInfo.size,
      "Content-Type": isWorkbook
        ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        : "text/csv; charset=utf-8",
      "X-Content-Type-Options": "nosniff"
    });
    const stream = createReadStream(filePath);
    stream.on("error", (error) => {
      console.error("Field Result download failed:", error.message);
      response.destroy(error);
    });
    stream.pipe(response);
  } catch (error) {
    if (error.code === "ENOENT") {
      sendJson(response, 404, { error: "The requested Field Result output is not available." });
      return;
    }
    throw error;
  }
}

function resolveStaticPath(urlPath) {
  const configured = STATIC_FILES.get(urlPath);
  if (configured) return { filename: configured[0], contentType: configured[1] };

  if (!urlPath.startsWith("/assets/") && !urlPath.startsWith("/vendor/")) return null;
  let decodedPath;
  try {
    decodedPath = decodeURIComponent(urlPath);
  } catch {
    return null;
  }
  const relativePath = decodedPath.slice(1);
  const filename = path.resolve(ROOT, relativePath);
  if (!filename.startsWith(`${ROOT}${path.sep}`)) return null;
  if (relativePath.startsWith("vendor/") &&
      !["vendor/xlsx.full.min.js", "vendor/jszip.min.js"].includes(relativePath)) return null;
  if (relativePath.startsWith("assets/") &&
      !relativePath.startsWith("assets/fonts/") &&
      relativePath !== "assets/workspace-tools-icon.svg") return null;
  const extension = path.extname(filename).toLowerCase();
  const contentType = MIME_TYPES.get(extension) || "application/octet-stream";
  return { filename, contentType };
}

async function serveStatic(request, response, pathname) {
  const asset = resolveStaticPath(pathname);
  if (!asset) {
    sendText(response, 404, "Not found.");
    return;
  }
  let fileInfo;
  let content;
  try {
    [fileInfo, content] = await Promise.all([stat(asset.filename), readFile(asset.filename)]);
  } catch (error) {
    if (error.code === "ENOENT") {
      sendText(response, 404, "Not found.");
      return;
    }
    throw error;
  }
  if (!fileInfo.isFile()) {
    sendText(response, 404, "Not found.");
    return;
  }
  response.writeHead(200, {
    "Cache-Control": pathname === "/" || pathname === "/index.html" ||
      pathname === "/app.js" || pathname === "/field-result-worker.js"
      ? "no-store"
      : "private, max-age=300",
    "Content-Length": content.length,
    "Content-Type": asset.contentType,
    "X-Content-Type-Options": "nosniff"
  });
  if (request.method === "HEAD") response.end();
  else response.end(content);
}

function createWorkerEnv(env) {
  const publicOrigin = env.PUBLIC_APP_ORIGIN || DEFAULT_PUBLIC_ORIGIN;
  return {
    ...env,
    ALLOWED_ORIGINS: [...new Set([publicOrigin, "http://localhost:8001"])].join(","),
    DASHBOARD_ACCESS_KEY: INTERNAL_ACCESS_KEY
  };
}

async function proxyApi(request, response, env) {
  const headers = new Headers();
  headers.set("Origin", request.headers.origin || env.PUBLIC_APP_ORIGIN);
  headers.set("Authorization", `Bearer ${INTERNAL_ACCESS_KEY}`);
  const workerRequest = new Request(`http://localhost${request.url}`, {
    method: request.method,
    headers
  });
  const workerResponse = await worker.fetch(workerRequest, env);
  response.writeHead(workerResponse.status, Object.fromEntries(workerResponse.headers));
  response.end(Buffer.from(await workerResponse.arrayBuffer()));
}

export function createAppServer(env) {
  const workerEnv = createWorkerEnv(env);
  const fieldResultJobs = new Map();
  return createServer(async (request, response) => {
    try {
      if (!isConfigured(env)) {
        sendText(response, 503, "Set TUNNEL_ACCESS_MODE=private and complete the local Lark settings in the ignored .dev.vars file. The server will not start otherwise.");
        return;
      }

      const url = new URL(request.url, "http://localhost");
      if (url.pathname === "/api/field-result") {
        if (request.method !== "POST") {
          sendJson(response, 405, { error: "Method not allowed." });
          return;
        }
        await cleanupExpiredFieldResultJobs(fieldResultJobs);
        await processFieldResultUpload(request, response, fieldResultJobs);
        return;
      }
      const fieldResultDownload = url.pathname.match(/^\/api\/field-result\/([0-9a-f-]{36})\/(workbook|csv)$/i);
      if (fieldResultDownload) {
        if (request.method !== "GET") {
          sendJson(response, 405, { error: "Method not allowed." });
          return;
        }
        await cleanupExpiredFieldResultJobs(fieldResultJobs);
        const [, id, type] = fieldResultDownload;
        await downloadFieldResultFile(response, fieldResultJobs, id, type);
        return;
      }
      if (url.pathname === "/api/notifications" || url.pathname === "/api/export") {
        if (request.method !== "GET" && request.method !== "OPTIONS") {
          sendText(response, 405, "Method not allowed.");
          return;
        }
        await proxyApi(request, response, workerEnv);
        return;
      }
      if (request.method !== "GET" && request.method !== "HEAD") {
        sendText(response, 405, "Method not allowed.");
        return;
      }
      await serveStatic(request, response, url.pathname);
    } catch (error) {
      console.error("Request failed:", error.message);
      if (!response.headersSent) sendText(response, 500, "The local server could not complete the request.");
      else response.destroy();
    }
  });
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const env = await loadLocalEnv();
  const port = Number(process.env.PORT || 8001);
  const host = process.env.HOST || "127.0.0.1";
  const server = createAppServer(env);
  server.listen(port, host, () => {
    console.log(`Workspace Tools V2 is listening at http://${host}:${port}`);
    console.log("The Dev Tunnel login gate is enabled; credentials and Lark secrets are not logged.");
  });
}
