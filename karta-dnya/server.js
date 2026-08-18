#!/usr/bin/env node
"use strict";

const http = require("http");
const fs = require("fs");
const path = require("path");
const { URL } = require("url");
const crypto = require("crypto");

const ROOT = __dirname;
const PUBLIC = path.join(ROOT, "public");
const DATA_DIR = path.join(ROOT, "data");
const DB_PATH = path.join(DATA_DIR, "db.json");
const PORT = Number(process.env.PORT || 3847);

const DEFAULT_MANAGERS = [
  { id: "mgr-1", name: "Блохина Елизавета" },
  { id: "mgr-2", name: "Оглоблина Софья" },
  { id: "mgr-3", name: "Кургузов Данил" },
  { id: "mgr-4", name: "Юнусова Юлиана" },
];

const SAMPLE_STICKERS = {
  "mgr-1": [
    {
      id: "stk-a1",
      text: "База №12 — активные лиды",
      x: 28,
      y: 42,
      color: "cyan",
      rotation: -2,
    },
  ],
  "mgr-2": [
    {
      id: "stk-d1",
      text: "Приоритет: демо на этой неделе",
      x: 40,
      y: 50,
      color: "amber",
      rotation: 2,
    },
  ],
};

function uid(prefix) {
  return `${prefix}-${crypto.randomBytes(4).toString("hex")}`;
}

function todayISO() {
  const d = new Date();
  const offset = d.getTimezoneOffset();
  const local = new Date(d.getTime() - offset * 60 * 1000);
  return local.toISOString().slice(0, 10);
}

function emptyDb() {
  const managers = DEFAULT_MANAGERS.map((m) => ({
    ...m,
    createdAt: new Date().toISOString(),
  }));
  const boards = {};
  for (const m of managers) {
    boards[m.id] = (SAMPLE_STICKERS[m.id] || []).map((s) => ({
      ...s,
      updatedAt: new Date().toISOString(),
    }));
  }
  return { managers, forms: [], boards };
}

function migrateDb(db, opts) {
  let changed = false;
  if (!db.managers) {
    db.managers = emptyDb().managers;
    changed = true;
  }
  if (!Array.isArray(db.forms)) {
    db.forms = [];
    changed = true;
  }
  if (!db.boards || typeof db.boards !== "object") {
    db.boards = {};
    changed = true;
  }
  // migrate old shared stickers → first manager board
  if (Array.isArray(db.stickers)) {
    const firstId = db.managers[0]?.id;
    if (firstId && !db.boards[firstId]?.length && db.stickers.length) {
      db.boards[firstId] = db.stickers.map((s) => ({
        ...s,
        updatedAt: s.updatedAt || new Date().toISOString(),
      }));
    }
    delete db.stickers;
    changed = true;
  }
  if (db.checklist) {
    delete db.checklist;
    changed = true;
  }
  for (const m of db.managers) {
    if (!Array.isArray(db.boards[m.id])) {
      db.boards[m.id] = [];
      changed = true;
    }
  }
  if (changed && (!opts || opts.persist !== false)) {
    // fire-and-forget local/github write
    Promise.resolve(writeDb(db)).catch((err) => console.error(err));
  }
  return db;
}

const GH_TOKEN = process.env.GITHUB_TOKEN || process.env.GH_TOKEN || "";
const GH_REPO = process.env.GITHUB_REPO || "daaanilNikonov/for-lisa";
const GH_BRANCH = process.env.GITHUB_BRANCH || "cursor/karta-dnya-produktovogo-zapuska-ed4c";
const GH_DB_PATH = process.env.GITHUB_DB_PATH || "karta-dnya/data/db.json";
const USE_GITHUB = Boolean(GH_TOKEN && GH_REPO);

let githubSha = null;
let writeQueue = Promise.resolve();

function ensureDbLocal() {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
  if (!fs.existsSync(DB_PATH)) {
    writeDbLocal(emptyDb());
  }
}

function writeDbLocal(db) {
  ensureDbLocal();
  const tmp = `${DB_PATH}.${process.pid}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(db, null, 2), "utf8");
  fs.renameSync(tmp, DB_PATH);
}

async function githubRequest(method, apiPath, body) {
  const res = await fetch(`https://api.github.com${apiPath}`, {
    method,
    headers: {
      Accept: "application/vnd.github+json",
      Authorization: `Bearer ${GH_TOKEN}`,
      "User-Agent": "forus-karta-dnya",
      "X-GitHub-Api-Version": "2022-11-28",
      ...(body ? { "Content-Type": "application/json" } : {}),
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  const text = await res.text();
  let data = null;
  try {
    data = text ? JSON.parse(text) : null;
  } catch {
    data = { raw: text };
  }
  if (!res.ok) {
    const msg = data && data.message ? data.message : `GitHub API ${res.status}`;
    const err = new Error(msg);
    err.status = res.status;
    err.data = data;
    throw err;
  }
  return data;
}

async function readDbFromGithub() {
  const encPath = GH_DB_PATH.split("/").map(encodeURIComponent).join("/");
  try {
    const data = await githubRequest(
      "GET",
      `/repos/${GH_REPO}/contents/${encPath}?ref=${encodeURIComponent(GH_BRANCH)}`
    );
    githubSha = data.sha;
    const raw = Buffer.from(data.content.replace(/\n/g, ""), "base64").toString("utf8");
    return migrateDb(JSON.parse(raw), { persist: false });
  } catch (err) {
    if (err.status === 404) {
      const db = emptyDb();
      await writeDbToGithub(db);
      return db;
    }
    throw err;
  }
}

async function writeDbToGithub(db) {
  const encPath = GH_DB_PATH.split("/").map(encodeURIComponent).join("/");
  const content = Buffer.from(JSON.stringify(db, null, 2) + "\n", "utf8").toString("base64");
  const payload = {
    message: `chore(karta-dnya): sync db ${new Date().toISOString()}`,
    content,
    branch: GH_BRANCH,
  };
  if (githubSha) payload.sha = githubSha;
  try {
    const data = await githubRequest(
      "PUT",
      `/repos/${GH_REPO}/contents/${encPath}`,
      payload
    );
    githubSha = data.content && data.content.sha ? data.content.sha : githubSha;
    writeDbLocal(db);
    return;
  } catch (err) {
    // sha conflict — refetch and retry once
    if (err.status === 409 || err.status === 422) {
      const latest = await githubRequest(
        "GET",
        `/repos/${GH_REPO}/contents/${encPath}?ref=${encodeURIComponent(GH_BRANCH)}`
      );
      githubSha = latest.sha;
      payload.sha = githubSha;
      const data = await githubRequest(
        "PUT",
        `/repos/${GH_REPO}/contents/${encPath}`,
        payload
      );
      githubSha = data.content && data.content.sha ? data.content.sha : githubSha;
      writeDbLocal(db);
      return;
    }
    throw err;
  }
}

async function readDb() {
  if (USE_GITHUB) {
    const db = await readDbFromGithub();
    writeDbLocal(db);
    return db;
  }
  ensureDbLocal();
  const raw = fs.readFileSync(DB_PATH, "utf8");
  return migrateDb(JSON.parse(raw), { persist: true });
}

function writeDb(db) {
  // serialize writes so GitHub sha stays consistent
  writeQueue = writeQueue.then(async () => {
    if (USE_GITHUB) {
      await writeDbToGithub(db);
    } else {
      writeDbLocal(db);
    }
  }).catch((err) => {
    console.error("writeDb failed:", err);
    throw err;
  });
  return writeQueue;
}

function sendJson(res, status, payload) {
  const body = JSON.stringify(payload);
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
  });
  res.end(body);
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    req.on("data", (c) => chunks.push(c));
    req.on("end", () => {
      const raw = Buffer.concat(chunks).toString("utf8");
      if (!raw) return resolve({});
      try {
        resolve(JSON.parse(raw));
      } catch (err) {
        reject(err);
      }
    });
    req.on("error", reject);
  });
}

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".json": "application/json; charset=utf-8",
  ".ico": "image/x-icon",
  ".woff2": "font/woff2",
};

function serveStatic(req, res, pathname) {
  let rel = pathname === "/" ? "/index.html" : pathname;
  rel = decodeURIComponent(rel).replace(/\0/g, "");
  const filePath = path.normalize(path.join(PUBLIC, rel));
  if (!filePath.startsWith(PUBLIC)) {
    res.writeHead(403);
    return res.end("Forbidden");
  }
  if (!fs.existsSync(filePath) || fs.statSync(filePath).isDirectory()) {
    res.writeHead(404);
    return res.end("Not found");
  }
  const ext = path.extname(filePath).toLowerCase();
  res.writeHead(200, { "Content-Type": MIME[ext] || "application/octet-stream" });
  fs.createReadStream(filePath).pipe(res);
}

function formTitle(managerName, date) {
  return `${managerName} ${date}`;
}

function normalizeStickers(list) {
  return (list || []).map((s) => ({
    id: s.id || uid("stk"),
    text: String(s.text ?? ""),
    x: Number(s.x) || 0,
    y: Number(s.y) || 0,
    color: ["cyan", "amber", "mint", "rose", "violet"].includes(s.color)
      ? s.color
      : "cyan",
    rotation: Number(s.rotation) || 0,
    updatedAt: new Date().toISOString(),
  }));
}

async function handleApi(req, res, pathname) {
  const method = req.method || "GET";

  if (method === "GET" && pathname === "/api/state") {
    const db = await readDb();
    return sendJson(res, 200, { ...db, today: todayISO() });
  }

  if (method === "POST" && pathname === "/api/managers") {
    const body = await readBody(req);
    const name = String(body.name || "").trim();
    if (!name) return sendJson(res, 400, { error: "Укажите имя менеджера" });
    const db = await readDb();
    const manager = {
      id: uid("mgr"),
      name,
      createdAt: new Date().toISOString(),
    };
    db.managers.push(manager);
    db.boards[manager.id] = [];
    await writeDb(db);
    return sendJson(res, 201, { manager });
  }

  if (method === "PATCH" && pathname.startsWith("/api/managers/")) {
    const id = pathname.split("/").pop();
    const body = await readBody(req);
    const db = await readDb();
    const manager = db.managers.find((m) => m.id === id);
    if (!manager) return sendJson(res, 404, { error: "Менеджер не найден" });
    if (body.name != null) {
      const name = String(body.name).trim();
      if (!name) return sendJson(res, 400, { error: "Имя не может быть пустым" });
      manager.name = name;
      for (const form of db.forms) {
        if (form.managerId === id) {
          form.managerName = name;
          if (form.status === "completed") {
            form.title = formTitle(name, form.date);
          } else {
            form.title = formTitle(name, form.date);
          }
        }
      }
    }
    await writeDb(db);
    return sendJson(res, 200, { manager });
  }

  if (method === "DELETE" && pathname.startsWith("/api/managers/")) {
    const id = pathname.split("/").pop();
    const db = await readDb();
    const before = db.managers.length;
    db.managers = db.managers.filter((m) => m.id !== id);
    if (db.managers.length === before) {
      return sendJson(res, 404, { error: "Менеджер не найден" });
    }
    delete db.boards[id];
    await writeDb(db);
    return sendJson(res, 200, { ok: true });
  }

  // Personal sticker boards
  if (method === "PUT" && pathname.startsWith("/api/boards/")) {
    const managerId = pathname.split("/").pop();
    const body = await readBody(req);
    const db = await readDb();
    if (!db.managers.some((m) => m.id === managerId)) {
      return sendJson(res, 404, { error: "Менеджер не найден" });
    }
    if (!Array.isArray(body.stickers)) {
      return sendJson(res, 400, { error: "Ожидается массив стикеров" });
    }
    db.boards[managerId] = normalizeStickers(body.stickers);
    await writeDb(db);
    return sendJson(res, 200, { stickers: db.boards[managerId] });
  }

  if (method === "POST" && /^\/api\/boards\/[^/]+\/stickers$/.test(pathname)) {
    const managerId = pathname.split("/")[3];
    const body = await readBody(req);
    const db = await readDb();
    if (!db.managers.some((m) => m.id === managerId)) {
      return sendJson(res, 404, { error: "Менеджер не найден" });
    }
    if (!Array.isArray(db.boards[managerId])) db.boards[managerId] = [];
    const sticker = {
      id: uid("stk"),
      text: String(body.text || "Новый приоритет"),
      x: Number(body.x) || 40 + Math.random() * 120,
      y: Number(body.y) || 40 + Math.random() * 80,
      color: body.color || "cyan",
      rotation: body.rotation ?? Math.round((Math.random() * 8 - 4) * 10) / 10,
      updatedAt: new Date().toISOString(),
    };
    db.boards[managerId].push(sticker);
    await writeDb(db);
    return sendJson(res, 201, { sticker });
  }

  if (method === "DELETE" && /^\/api\/boards\/[^/]+\/stickers\/[^/]+$/.test(pathname)) {
    const parts = pathname.split("/");
    const managerId = parts[3];
    const stickerId = parts[5];
    const db = await readDb();
    if (!Array.isArray(db.boards[managerId])) {
      return sendJson(res, 404, { error: "Доска не найдена" });
    }
    db.boards[managerId] = db.boards[managerId].filter((s) => s.id !== stickerId);
    await writeDb(db);
    return sendJson(res, 200, { ok: true });
  }

  // Morning draft
  if (method === "POST" && pathname === "/api/forms/morning") {
    const body = await readBody(req);
    const managerId = body.managerId;
    const date = body.date || todayISO();
    const tasks = Array.isArray(body.tasks)
      ? body.tasks
          .map((t) => ({
            id: t.id || uid("task"),
            text: String(t.text || "").trim(),
            done: false,
          }))
          .filter((t) => t.text)
      : [];
    if (!managerId) return sendJson(res, 400, { error: "Не выбран менеджер" });
    if (!tasks.length) {
      return sendJson(res, 400, { error: "Добавьте хотя бы одну задачу" });
    }
    const db = await readDb();
    const manager = db.managers.find((m) => m.id === managerId);
    if (!manager) return sendJson(res, 404, { error: "Менеджер не найден" });

    let form = db.forms.find(
      (f) => f.managerId === managerId && f.date === date && f.status !== "completed"
    );
    if (!form) {
      form = {
        id: uid("form"),
        managerId,
        managerName: manager.name,
        date,
        title: formTitle(manager.name, date),
        status: "morning",
        tasks,
        createdAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
        completedAt: null,
      };
      db.forms.push(form);
    } else {
      form.managerName = manager.name;
      form.title = formTitle(manager.name, date);
      form.tasks = tasks;
      form.status = "morning";
      form.updatedAt = new Date().toISOString();
    }
    await writeDb(db);
    return sendJson(res, 200, {
      form,
      message: "Чеклист на день сохранён. В архив попадёт после вечерних галочек.",
    });
  }

  // Evening → archive
  if (method === "POST" && pathname === "/api/forms/evening") {
    const body = await readBody(req);
    const managerId = body.managerId;
    const date = body.date || todayISO();
    const tasks = Array.isArray(body.tasks) ? body.tasks : null;
    if (!managerId) return sendJson(res, 400, { error: "Не выбран менеджер" });

    const db = await readDb();
    const manager = db.managers.find((m) => m.id === managerId);
    if (!manager) return sendJson(res, 404, { error: "Менеджер не найден" });

    let form = db.forms.find(
      (f) => f.managerId === managerId && f.date === date && f.status !== "completed"
    );
    if (!form) {
      return sendJson(res, 400, {
        error: "Сначала сохраните утренний чеклист задач",
      });
    }
    if (tasks) {
      form.tasks = tasks.map((t, idx) => ({
        id: t.id || form.tasks[idx]?.id || uid("task"),
        text: String(t.text || form.tasks[idx]?.text || "").trim(),
        done: Boolean(t.done),
      }));
    }
    if (!form.tasks.length) {
      return sendJson(res, 400, { error: "В форме нет задач" });
    }

    form.managerName = manager.name;
    form.title = formTitle(manager.name, date);
    form.status = "completed";
    form.updatedAt = new Date().toISOString();
    form.completedAt = new Date().toISOString();
    await writeDb(db);
    return sendJson(res, 200, {
      form,
      message: `Форма «${form.title}» сохранена в архив`,
    });
  }

  if (method === "DELETE" && pathname.startsWith("/api/forms/")) {
    const id = pathname.split("/").pop();
    const db = await readDb();
    const before = db.forms.length;
    db.forms = db.forms.filter((f) => f.id !== id);
    if (db.forms.length === before) {
      return sendJson(res, 404, { error: "Форма не найдена" });
    }
    await writeDb(db);
    return sendJson(res, 200, { ok: true });
  }

  if (method === "GET" && pathname === "/api/archive") {
    const db = await readDb();
    const archive = db.forms
      .filter((f) => f.status === "completed")
      .sort(
        (a, b) =>
          String(b.date).localeCompare(String(a.date)) ||
          String(a.title).localeCompare(String(b.title))
      );
    return sendJson(res, 200, { archive });
  }

  // legacy shared stickers endpoints removed
  return sendJson(res, 404, { error: "Не найдено" });
}

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url || "/", `http://${req.headers.host || "localhost"}`);
    let pathname = url.pathname;
    if (
      pathname === "/karta-dnya" ||
      pathname === "/karta-dnya/" ||
      pathname === "/gruppa-zapuska/karta-dnya" ||
      pathname === "/gruppa-zapuska/karta-dnya/"
    ) {
      pathname = "/";
    }
    if (pathname.startsWith("/api/")) {
      return await handleApi(req, res, pathname);
    }
    return serveStatic(req, res, pathname);
  } catch (err) {
    console.error(err);
    sendJson(res, 500, { error: "Внутренняя ошибка сервера" });
  }
});

server.listen(PORT, "0.0.0.0", () => {
  const mode = USE_GITHUB ? `GitHub ${GH_REPO}@${GH_BRANCH}` : `file ${DB_PATH}`;
  console.log(`Карта дня Форус → http://localhost:${PORT}/karta-dnya`);
  console.log(`Storage: ${mode}`);
});

// warm local cache
readDb().catch((err) => console.error("Initial DB load failed:", err));
