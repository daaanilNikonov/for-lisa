/**
 * Прокси отправки заявок в Яндекс.Форму (CSRF + cookie jar).
 * Раздаёт статику лендинга и принимает POST /api/yandex-submit
 *
 * Запуск из корня репозитория:
 *   node test-kabinet/proxy/server.mjs
 * или из test-kabinet:
 *   node proxy/server.mjs
 */
import http from "node:http";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, "..");
const PORT = Number(process.env.PORT || 8787);
const DEFAULT_SURVEY_ID = "6a83b54feb614605f98e10ee";

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".svg": "image/svg+xml",
  ".webp": "image/webp",
  ".ico": "image/x-icon",
  ".md": "text/markdown; charset=utf-8",
  ".csv": "text/csv; charset=utf-8",
};

function sendJson(res, status, body) {
  const data = JSON.stringify(body);
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST,OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Cache-Control": "no-store",
  });
  res.end(data);
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", (c) => chunks.push(c));
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

function parseSetCookies(res) {
  if (typeof res.headers.getSetCookie === "function") {
    return res.headers.getSetCookie();
  }
  const single = res.headers.get("set-cookie");
  return single ? [single] : [];
}

function storeCookies(jar, setCookieHeaders) {
  for (const raw of setCookieHeaders) {
    const part = String(raw).split(";")[0];
    const eq = part.indexOf("=");
    if (eq <= 0) continue;
    jar.set(part.slice(0, eq).trim(), part.slice(eq + 1).trim());
  }
}

function cookieHeader(jar) {
  return [...jar.entries()].map(([k, v]) => `${k}=${v}`).join("; ");
}

async function yandexFetch(url, { method = "GET", body, jar, csrf } = {}) {
  const headers = {
    Accept: "application/json, text/plain, */*",
    Referer: `https://forms.yandex.ru/u/${DEFAULT_SURVEY_ID}/`,
  };
  if (jar.size) headers.Cookie = cookieHeader(jar);
  if (csrf) headers["x-csrf-token"] = csrf;
  if (body !== undefined) {
    headers["Content-Type"] = "application/json";
  }
  const res = await fetch(url, {
    method,
    headers,
    body: body !== undefined ? JSON.stringify(body) : undefined,
    redirect: "manual",
  });
  storeCookies(jar, parseSetCookies(res));
  const text = await res.text();
  let json = null;
  try {
    json = text ? JSON.parse(text) : null;
  } catch {
    json = null;
  }
  return {
    status: res.status,
    csrf: res.headers.get("x-csrf-token") || csrf || null,
    json,
    text,
  };
}

async function submitToYandex({ surveyId, values }) {
  const id = surveyId || DEFAULT_SURVEY_ID;
  const jar = new Map();
  let csrf = null;

  await yandexFetch(`https://forms.yandex.ru/u/${id}/`, { jar });

  let probe = await yandexFetch(
    "https://forms.yandex.ru/u/gateway/root/form/getSurvey",
    { method: "POST", body: { surveyId: id }, jar, csrf }
  );
  csrf = probe.csrf || csrf;

  if (probe.status === 419 || probe.status >= 400) {
    probe = await yandexFetch(
      "https://forms.yandex.ru/u/gateway/root/form/getSurvey",
      { method: "POST", body: { surveyId: id }, jar, csrf }
    );
    csrf = probe.csrf || csrf;
  }

  if (probe.status >= 400) {
    throw new Error(`getSurvey failed: ${probe.status} ${probe.text.slice(0, 200)}`);
  }

  const posted = await yandexFetch(
    "https://forms.yandex.ru/u/gateway/root/form/postSurvey",
    {
      method: "POST",
      body: {
        surveyId: id,
        values,
        parent: "",
        dryRun: false,
        timestamp: new Date().toISOString(),
      },
      jar,
      csrf,
    }
  );

  if (posted.status >= 400) {
    throw new Error(`postSurvey failed: ${posted.status} ${posted.text.slice(0, 200)}`);
  }

  return posted.json || { ok: true };
}

function safePath(urlPath) {
  const decoded = decodeURIComponent(urlPath.split("?")[0]);
  const rel = decoded === "/" ? "/index.html" : decoded;
  const full = path.normalize(path.join(ROOT, rel));
  if (!full.startsWith(ROOT)) return null;
  return full;
}

const server = http.createServer(async (req, res) => {
  const url = new URL(req.url || "/", `http://${req.headers.host || "localhost"}`);

  if (req.method === "OPTIONS" && url.pathname === "/api/yandex-submit") {
    sendJson(res, 204, {});
    return;
  }

  if (req.method === "POST" && url.pathname === "/api/yandex-submit") {
    try {
      const raw = await readBody(req);
      const payload = JSON.parse(raw.toString("utf8") || "{}");
      const values = payload.values || {};
      if (!values || typeof values !== "object") {
        sendJson(res, 400, { error: "values required" });
        return;
      }
      const result = await submitToYandex({
        surveyId: payload.surveyId || DEFAULT_SURVEY_ID,
        values,
      });
      sendJson(res, 200, { ok: true, result });
    } catch (err) {
      sendJson(res, 502, { ok: false, error: String(err.message || err) });
    }
    return;
  }

  if (req.method !== "GET" && req.method !== "HEAD") {
    res.writeHead(405).end("Method Not Allowed");
    return;
  }

  const filePath = safePath(url.pathname);
  if (!filePath) {
    res.writeHead(400).end("Bad path");
    return;
  }
  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404).end("Not found");
      return;
    }
    const ext = path.extname(filePath).toLowerCase();
    res.writeHead(200, { "Content-Type": MIME[ext] || "application/octet-stream" });
    if (req.method === "HEAD") res.end();
    else res.end(data);
  });
});

server.listen(PORT, "0.0.0.0", () => {
  console.log(`Forus quiz + Yandex proxy: http://127.0.0.1:${PORT}/`);
  console.log(`Submit API: POST http://127.0.0.1:${PORT}/api/yandex-submit`);
});
