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
  { id: "mgr-1", name: "Анна" },
  { id: "mgr-2", name: "Дмитрий" },
  { id: "mgr-3", name: "Елена" },
  { id: "mgr-4", name: "Сергей" },
];

const DEFAULT_CHECKLIST = [
  { id: "chk-1", text: "Проверить очередь лидов и назначить приоритеты", done: false },
  { id: "chk-2", text: "Обработать входящие заявки по продуктовому запуску", done: false },
  { id: "chk-3", text: "Обновить статусы по базам и передать эстафету", done: false },
  { id: "chk-4", text: "Зафиксировать итоги дня в карте менеджера", done: false },
];

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
  return {
    managers: DEFAULT_MANAGERS.map((m) => ({
      ...m,
      createdAt: new Date().toISOString(),
    })),
    forms: [],
    stickers: [
      {
        id: "stk-welcome",
        text: "База №12 — активные лиды",
        x: 24,
        y: 36,
        color: "cyan",
        rotation: -2,
        updatedAt: new Date().toISOString(),
      },
      {
        id: "stk-welcome-2",
        text: "Приоритет: демо на этой неделе",
        x: 180,
        y: 90,
        color: "amber",
        rotation: 3,
        updatedAt: new Date().toISOString(),
      },
    ],
    checklist: {
      date: todayISO(),
      items: DEFAULT_CHECKLIST.map((i) => ({ ...i })),
    },
  };
}

function ensureDb() {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
  if (!fs.existsSync(DB_PATH)) {
    writeDb(emptyDb());
  }
}

function readDb() {
  ensureDb();
  const raw = fs.readFileSync(DB_PATH, "utf8");
  const db = JSON.parse(raw);
  if (!db.checklist || db.checklist.date !== todayISO()) {
    db.checklist = {
      date: todayISO(),
      items: DEFAULT_CHECKLIST.map((i) => ({ ...i, done: false })),
    };
    writeDb(db);
  }
  return db;
}

function writeDb(db) {
  ensureDb();
  const tmp = `${DB_PATH}.${process.pid}.tmp`;
  fs.writeFileSync(tmp, JSON.stringify(db, null, 2), "utf8");
  fs.renameSync(tmp, DB_PATH);
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

async function handleApi(req, res, pathname) {
  const method = req.method || "GET";

  if (method === "GET" && pathname === "/api/state") {
    const db = readDb();
    return sendJson(res, 200, { ...db, today: todayISO() });
  }

  if (method === "POST" && pathname === "/api/managers") {
    const body = await readBody(req);
    const name = String(body.name || "").trim();
    if (!name) return sendJson(res, 400, { error: "Укажите имя менеджера" });
    const db = readDb();
    const manager = {
      id: uid("mgr"),
      name,
      createdAt: new Date().toISOString(),
    };
    db.managers.push(manager);
    writeDb(db);
    return sendJson(res, 201, { manager });
  }

  if (method === "PATCH" && pathname.startsWith("/api/managers/")) {
    const id = pathname.split("/").pop();
    const body = await readBody(req);
    const db = readDb();
    const manager = db.managers.find((m) => m.id === id);
    if (!manager) return sendJson(res, 404, { error: "Менеджер не найден" });
    if (body.name != null) {
      const name = String(body.name).trim();
      if (!name) return sendJson(res, 400, { error: "Имя не может быть пустым" });
      manager.name = name;
    }
    writeDb(db);
    return sendJson(res, 200, { manager });
  }

  if (method === "DELETE" && pathname.startsWith("/api/managers/")) {
    const id = pathname.split("/").pop();
    const db = readDb();
    const before = db.managers.length;
    db.managers = db.managers.filter((m) => m.id !== id);
    if (db.managers.length === before) {
      return sendJson(res, 404, { error: "Менеджер не найден" });
    }
    writeDb(db);
    return sendJson(res, 200, { ok: true });
  }

  if (method === "PUT" && pathname === "/api/checklist") {
    const body = await readBody(req);
    const db = readDb();
    if (!Array.isArray(body.items)) {
      return sendJson(res, 400, { error: "Ожидается список пунктов" });
    }
    db.checklist = {
      date: todayISO(),
      items: body.items.map((item, idx) => ({
        id: item.id || `chk-${idx + 1}`,
        text: String(item.text || "").trim() || `Пункт ${idx + 1}`,
        done: Boolean(item.done),
      })),
    };
    writeDb(db);
    return sendJson(res, 200, { checklist: db.checklist });
  }

  if (method === "POST" && pathname === "/api/checklist/items") {
    const body = await readBody(req);
    const text = String(body.text || "").trim();
    if (!text) return sendJson(res, 400, { error: "Введите текст пункта" });
    const db = readDb();
    const item = { id: uid("chk"), text, done: false };
    db.checklist.items.push(item);
    writeDb(db);
    return sendJson(res, 201, { item, checklist: db.checklist });
  }

  if (method === "PUT" && pathname === "/api/stickers") {
    const body = await readBody(req);
    const db = readDb();
    if (!Array.isArray(body.stickers)) {
      return sendJson(res, 400, { error: "Ожидается массив стикеров" });
    }
    db.stickers = body.stickers.map((s) => ({
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
    writeDb(db);
    return sendJson(res, 200, { stickers: db.stickers });
  }

  if (method === "POST" && pathname === "/api/stickers") {
    const body = await readBody(req);
    const db = readDb();
    const sticker = {
      id: uid("stk"),
      text: String(body.text || "Новый приоритет"),
      x: Number(body.x) || 40 + Math.random() * 120,
      y: Number(body.y) || 40 + Math.random() * 80,
      color: body.color || "cyan",
      rotation: body.rotation ?? Math.round((Math.random() * 8 - 4) * 10) / 10,
      updatedAt: new Date().toISOString(),
    };
    db.stickers.push(sticker);
    writeDb(db);
    return sendJson(res, 201, { sticker });
  }

  if (method === "DELETE" && pathname.startsWith("/api/stickers/")) {
    const id = pathname.split("/").pop();
    const db = readDb();
    db.stickers = db.stickers.filter((s) => s.id !== id);
    writeDb(db);
    return sendJson(res, 200, { ok: true });
  }

  // Save morning draft (first fill) — not archived yet
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
    const db = readDb();
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
    writeDb(db);
    return sendJson(res, 200, { form, message: "Задачи на день сохранены. Архив обновится после вечерней отметки." });
  }

  // Evening completion — only then form enters archive storage
  if (method === "POST" && pathname === "/api/forms/evening") {
    const body = await readBody(req);
    const managerId = body.managerId;
    const date = body.date || todayISO();
    const tasks = Array.isArray(body.tasks) ? body.tasks : null;
    if (!managerId) return sendJson(res, 400, { error: "Не выбран менеджер" });

    const db = readDb();
    const manager = db.managers.find((m) => m.id === managerId);
    if (!manager) return sendJson(res, 404, { error: "Менеджер не найден" });

    let form = db.forms.find(
      (f) => f.managerId === managerId && f.date === date && f.status !== "completed"
    );
    if (!form) {
      return sendJson(res, 400, {
        error: "Сначала заполните утреннюю форму с задачами на день",
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
    // Move conceptually to archive: keep in forms with status completed
    writeDb(db);
    return sendJson(res, 200, {
      form,
      message: `Форма «${form.title}» сохранена в хранилище`,
    });
  }

  if (method === "GET" && pathname === "/api/archive") {
    const db = readDb();
    const archive = db.forms
      .filter((f) => f.status === "completed")
      .sort((a, b) => String(b.date).localeCompare(String(a.date)) || String(a.title).localeCompare(String(b.title)));
    return sendJson(res, 200, { archive });
  }

  return sendJson(res, 404, { error: "Не найдено" });
}

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url || "/", `http://${req.headers.host || "localhost"}`);
    let pathname = url.pathname;
    // Pretty URL aliases
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

ensureDb();
server.listen(PORT, () => {
  console.log(`Карта дня Форус → http://localhost:${PORT}/karta-dnya`);
});
