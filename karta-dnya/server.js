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

const DEFAULT_KPI_DEFS = [
  { id: "kpi-calls", name: "Успешные звонки", defaultTarget: 20, unit: "звонков" },
  { id: "kpi-demos", name: "Демонстрации", defaultTarget: 2, unit: "демо" },
  { id: "kpi-leads", name: "Обработанные лиды", defaultTarget: 15, unit: "лидов" },
];

const SAMPLE_STICKERS = {
  "mgr-1": [
    { id: "stk-a1", text: "База №12 — активные лиды", x: 28, y: 42, color: "cyan", rotation: -2, kind: "note" },
  ],
  "mgr-2": [
    { id: "stk-d1", text: "Приоритет: демо на этой неделе", x: 40, y: 50, color: "amber", rotation: 2, kind: "note" },
  ],
};

const STICKER_PACK = [
  { packId: "umnichka", text: "умничка", emoji: "💅", vibe: "pink" },
  { packId: "dobryak", text: "добряк", emoji: "🫶", vibe: "mint" },
  { packId: "molodec", text: "молодец", emoji: "🔥", vibe: "amber" },
  { packId: "eshche", text: "еще че скажешь?", emoji: "🤨", vibe: "violet" },
  { packId: "imba", text: "имба", emoji: "⚡", vibe: "cyan" },
  { packId: "topchik", text: "топчик", emoji: "🏆", vibe: "amber" },
  { packId: "zaoshlo", text: "зашло", emoji: "✨", vibe: "mint" },
  { packId: "vibe", text: "вайб", emoji: "🎧", vibe: "violet" },
  { packId: "ril-tok", text: "рил ток", emoji: "🗣️", vibe: "pink" },
  { packId: "pognalli", text: "погнали", emoji: "🚀", vibe: "cyan" },
  { packId: "zhiza", text: "жиза", emoji: "💀", vibe: "rose" },
  { packId: "krush", text: "краш дня", emoji: "💘", vibe: "pink" },
  { packId: "norm", text: "норм тема", emoji: "😎", vibe: "mint" },
  { packId: "respect", text: "респект", emoji: "🫡", vibe: "amber" },
  { packId: "cap", text: "без кэпа", emoji: "🧢", vibe: "violet" },
  { packId: "slay", text: "slay", emoji: "👑", vibe: "rose" },
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

function addDaysISO(iso, days) {
  const [y, m, d] = iso.split("-").map(Number);
  const dt = new Date(y, m - 1, d);
  dt.setDate(dt.getDate() + days);
  const yy = dt.getFullYear();
  const mm = String(dt.getMonth() + 1).padStart(2, "0");
  const dd = String(dt.getDate()).padStart(2, "0");
  return `${yy}-${mm}-${dd}`;
}

function weekDates(fromISO, days = 7) {
  const start = fromISO || todayISO();
  return Array.from({ length: days }, (_, i) => addDaysISO(start, i));
}

function findForm(db, managerId, date) {
  return db.forms.find((f) => f.managerId === managerId && f.date === date) || null;
}

function upsertMorningPlan(db, manager, date, tasks) {
  let form = db.forms.find(
    (f) => f.managerId === manager.id && f.date === date && f.status !== "completed"
  );
  if (!form) {
    const existing = findForm(db, manager.id, date);
    if (existing && existing.status === "completed") {
      return { error: `День ${date} уже закрыт в архиве`, form: existing };
    }
    form = {
      id: uid("form"),
      managerId: manager.id,
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
    return { form };
  }
  const prevById = Object.fromEntries((form.tasks || []).map((t) => [t.id, t]));
  form.tasks = tasks.map((t) => ({
    ...t,
    carriedFrom: t.carriedFrom || prevById[t.id]?.carriedFrom || null,
    carriedTo: prevById[t.id]?.carriedTo || null,
  }));
  form.managerName = manager.name;
  form.title = formTitle(manager.name, date);
  form.status = "morning";
  form.updatedAt = new Date().toISOString();
  return { form };
}

function buildTomorrowPreview(db, managerId) {
  const today = todayISO();
  const tomorrow = addDaysISO(today, 1);
  const managers = managerId
    ? db.managers.filter((m) => m.id === managerId)
    : db.managers;
  return managers.map((m) => {
    const form = findForm(db, m.id, tomorrow);
    const tasks = (form?.tasks || []).map((t) => ({
      id: t.id,
      text: t.text,
      target: t.target,
      unit: t.unit,
      mandatory: Boolean(t.mandatory),
      carriedFrom: t.carriedFrom || null,
      doneCount: t.doneCount || 0,
    }));
    return {
      managerId: m.id,
      managerName: m.name,
      date: tomorrow,
      status: form?.status || "idle",
      taskCount: tasks.length,
      tasks,
    };
  });
}

function buildWeekPreview(db, managerId, fromISO) {
  const dates = weekDates(fromISO || todayISO(), 7);
  return dates.map((date) => {
    const form = findForm(db, managerId, date);
    return {
      date,
      status: form?.status || "idle",
      taskCount: (form?.tasks || []).length,
      tasks: (form?.tasks || []).map((t) => ({
        id: t.id,
        text: t.text,
        target: t.target,
        unit: t.unit,
        mandatory: Boolean(t.mandatory),
        doneCount: t.doneCount || 0,
        carriedFrom: t.carriedFrom || null,
      })),
    };
  });
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
  return {
    managers,
    forms: [],
    boards,
    kpiDefs: DEFAULT_KPI_DEFS.map((k) => ({ ...k })),
    transfers: [],
  };
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
  if (Array.isArray(db.stickers)) {
    const firstId = db.managers[0]?.id;
    if (firstId && !db.boards[firstId]?.length && db.stickers.length) {
      db.boards[firstId] = db.stickers.map((s) => ({
        ...s,
        kind: s.kind || "note",
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
  if (!Array.isArray(db.kpiDefs) || !db.kpiDefs.length) {
    db.kpiDefs = DEFAULT_KPI_DEFS.map((k) => ({ ...k }));
    changed = true;
  }
  if (!Array.isArray(db.transfers)) {
    db.transfers = [];
    changed = true;
  }
  for (const m of db.managers) {
    if (!Array.isArray(db.boards[m.id])) {
      db.boards[m.id] = [];
      changed = true;
    }
  }
  for (const form of db.forms) {
    for (const t of form.tasks || []) {
      if (t.target == null) {
        t.target = 1;
        changed = true;
      }
      if (t.doneCount == null) {
        t.doneCount = t.done ? Number(t.target) || 1 : 0;
        changed = true;
      }
      if (t.unit == null) t.unit = "";
      if (t.mandatory == null) t.mandatory = false;
      if (!t.kind) t.kind = t.mandatory ? "kpi" : "task";
    }
  }
  if (changed && (!opts || opts.persist !== false)) {
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
    const data = await githubRequest("PUT", `/repos/${GH_REPO}/contents/${encPath}`, payload);
    githubSha = data.content && data.content.sha ? data.content.sha : githubSha;
    writeDbLocal(db);
  } catch (err) {
    if (err.status === 409 || err.status === 422) {
      const latest = await githubRequest(
        "GET",
        `/repos/${GH_REPO}/contents/${encPath}?ref=${encodeURIComponent(GH_BRANCH)}`
      );
      githubSha = latest.sha;
      payload.sha = githubSha;
      const data = await githubRequest("PUT", `/repos/${GH_REPO}/contents/${encPath}`, payload);
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
  writeQueue = writeQueue
    .then(async () => {
      if (USE_GITHUB) await writeDbToGithub(db);
      else writeDbLocal(db);
    })
    .catch((err) => {
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

function normalizeTask(t, idx = 0) {
  const target = Math.max(1, Number(t.target) || 1);
  let doneCount = Number(t.doneCount);
  if (Number.isNaN(doneCount)) doneCount = t.done ? target : 0;
  doneCount = Math.max(0, Math.min(target, doneCount));
  const mandatory = Boolean(t.mandatory);
  return {
    id: t.id || uid("task"),
    text: String(t.text || "").trim() || `Задача ${idx + 1}`,
    target,
    doneCount,
    unit: String(t.unit || "").trim(),
    mandatory,
    kpiId: t.kpiId || null,
    kind: mandatory ? "kpi" : t.kind || "task",
    done: doneCount >= target,
    carriedFrom: t.carriedFrom || null,
    carriedTo: t.carriedTo || null,
    transferDate: t.transferDate || null,
  };
}

function normalizeStickers(list) {
  return (list || []).map((s) => ({
    id: s.id || uid("stk"),
    text: String(s.text ?? ""),
    x: Number(s.x) || 0,
    y: Number(s.y) || 0,
    color: ["cyan", "amber", "mint", "rose", "violet", "pink"].includes(s.color)
      ? s.color
      : "cyan",
    rotation: Number(s.rotation) || 0,
    kind: s.kind === "pack" ? "pack" : "note",
    packId: s.packId || null,
    emoji: s.emoji || null,
    vibe: s.vibe || null,
    updatedAt: new Date().toISOString(),
  }));
}

function taskProgress(t) {
  const target = Math.max(1, Number(t.target) || 1);
  const doneCount = Math.max(0, Number(t.doneCount) || 0);
  return Math.round((doneCount / target) * 100);
}

function buildAnalytics(db) {
  const byManager = {};
  for (const m of db.managers) {
    byManager[m.id] = {
      managerId: m.id,
      managerName: m.name,
      tasksTotal: 0,
      tasksFullyDone: 0,
      tasksPartial: 0,
      tasksCarried: 0,
      unitsPlanned: 0,
      unitsDone: 0,
      conversion: 0,
    };
  }

  for (const form of db.forms) {
    if (form.status !== "completed") continue;
    const bucket = byManager[form.managerId];
    if (!bucket) continue;
    for (const t of form.tasks || []) {
      bucket.tasksTotal += 1;
      const target = Math.max(1, Number(t.target) || 1);
      const doneCount = Math.max(0, Number(t.doneCount) || 0);
      bucket.unitsPlanned += target;
      bucket.unitsDone += doneCount;
      if (doneCount >= target) bucket.tasksFullyDone += 1;
      else if (doneCount > 0) bucket.tasksPartial += 1;
    }
  }

  const carriedByMgr = {};
  for (const tr of db.transfers || []) {
    carriedByMgr[tr.managerId] = (carriedByMgr[tr.managerId] || 0) + 1;
  }
  for (const id of Object.keys(byManager)) {
    byManager[id].tasksCarried = carriedByMgr[id] || 0;
    const b = byManager[id];
    b.conversion = b.unitsPlanned ? Math.round((b.unitsDone / b.unitsPlanned) * 1000) / 10 : 0;
  }

  return Object.values(byManager);
}

function buildDashboard(db, days = 14) {
  const end = todayISO();
  const dates = [];
  for (let i = days - 1; i >= 0; i -= 1) dates.push(addDaysISO(end, -i));

  const series = db.managers.map((m) => {
    const points = dates.map((date) => {
      const form = db.forms.find((f) => f.managerId === m.id && f.date === date);
      if (!form) return { date, conversion: null, planned: 0, done: 0 };
      let planned = 0;
      let done = 0;
      for (const t of form.tasks || []) {
        if (!t.mandatory && t.kind !== "kpi") continue;
        planned += Math.max(1, Number(t.target) || 1);
        done += Math.max(0, Number(t.doneCount) || 0);
      }
      // if no mandatory marked, use all tasks
      if (planned === 0) {
        for (const t of form.tasks || []) {
          planned += Math.max(1, Number(t.target) || 1);
          done += Math.max(0, Number(t.doneCount) || 0);
        }
      }
      return {
        date,
        conversion: planned ? Math.round((done / planned) * 1000) / 10 : null,
        planned,
        done,
      };
    });
    return { managerId: m.id, managerName: m.name, points };
  });

  // KPI table for today (or latest day with data)
  const today = end;
  const table = [];
  for (const m of db.managers) {
    const form = db.forms.find((f) => f.managerId === m.id && f.date === today);
    const defs = db.kpiDefs || DEFAULT_KPI_DEFS;
    for (const def of defs) {
      const task = (form?.tasks || []).find(
        (t) => t.kpiId === def.id || (t.mandatory && t.text === def.name)
      );
      const target = task ? Math.max(1, Number(task.target) || def.defaultTarget) : def.defaultTarget;
      const doneCount = task ? Math.max(0, Number(task.doneCount) || 0) : 0;
      table.push({
        managerId: m.id,
        managerName: m.name,
        kpiId: def.id,
        kpiName: def.name,
        unit: def.unit,
        target,
        doneCount,
        conversion: target ? Math.round((doneCount / target) * 1000) / 10 : 0,
        date: today,
        hasData: Boolean(task),
      });
    }
  }

  return { dates, series, table, today };
}

async function ensureDayForm(db, manager, date) {
  let form = db.forms.find(
    (f) => f.managerId === manager.id && f.date === date && f.status !== "completed"
  );
  if (!form) {
    // also allow appending to completed? No — create new morning draft if none open
    form = db.forms.find((f) => f.managerId === manager.id && f.date === date);
  }
  if (!form) {
    form = {
      id: uid("form"),
      managerId: manager.id,
      managerName: manager.name,
      date,
      title: formTitle(manager.name, date),
      status: "morning",
      tasks: [],
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      completedAt: null,
    };
    db.forms.push(form);
  }
  return form;
}

async function handleApi(req, res, pathname) {
  const method = req.method || "GET";

  if (method === "GET" && (pathname === "/healthz" || pathname === "/api/health")) {
    return sendJson(res, 200, { ok: true });
  }

  if (method === "GET" && pathname === "/api/state") {
    const db = await readDb();
    return sendJson(res, 200, {
      ...db,
      today: todayISO(),
      tomorrow: addDaysISO(todayISO(), 1),
      stickerPack: STICKER_PACK,
      analytics: buildAnalytics(db),
      dashboard: buildDashboard(db, 14),
      tomorrowPreview: buildTomorrowPreview(db),
    });
  }

  if (method === "GET" && pathname === "/api/week") {
    const url = new URL(req.url || "/", `http://${req.headers.host || "localhost"}`);
    const managerId = url.searchParams.get("managerId");
    const from = url.searchParams.get("from") || todayISO();
    if (!managerId) return sendJson(res, 400, { error: "Нужен managerId" });
    const db = await readDb();
    if (!db.managers.some((m) => m.id === managerId)) {
      return sendJson(res, 404, { error: "Менеджер не найден" });
    }
    return sendJson(res, 200, {
      managerId,
      from,
      days: buildWeekPreview(db, managerId, from),
      tomorrow: buildTomorrowPreview(db, managerId)[0] || null,
    });
  }

  if (method === "GET" && pathname === "/api/sticker-pack") {
    return sendJson(res, 200, { stickerPack: STICKER_PACK });
  }

  if (method === "GET" && pathname === "/api/analytics") {
    const db = await readDb();
    return sendJson(res, 200, {
      analytics: buildAnalytics(db),
      dashboard: buildDashboard(db, 14),
      transfers: db.transfers || [],
    });
  }

  if (method === "POST" && pathname === "/api/managers") {
    const body = await readBody(req);
    const name = String(body.name || "").trim();
    if (!name) return sendJson(res, 400, { error: "Укажите имя менеджера" });
    const db = await readDb();
    const manager = { id: uid("mgr"), name, createdAt: new Date().toISOString() };
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
          form.title = formTitle(name, form.date);
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
    if (db.managers.length === before) return sendJson(res, 404, { error: "Менеджер не найден" });
    delete db.boards[id];
    await writeDb(db);
    return sendJson(res, 200, { ok: true });
  }

  if (method === "PUT" && pathname === "/api/kpi-defs") {
    const body = await readBody(req);
    const db = await readDb();
    if (!Array.isArray(body.kpiDefs)) return sendJson(res, 400, { error: "Ожидается kpiDefs" });
    db.kpiDefs = body.kpiDefs.map((k, i) => ({
      id: k.id || uid("kpi"),
      name: String(k.name || `KPI ${i + 1}`).trim(),
      defaultTarget: Math.max(1, Number(k.defaultTarget) || 1),
      unit: String(k.unit || "").trim(),
    }));
    await writeDb(db);
    return sendJson(res, 200, { kpiDefs: db.kpiDefs });
  }

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
    let sticker;
    if (body.packId) {
      const pack = STICKER_PACK.find((p) => p.packId === body.packId);
      if (!pack) return sendJson(res, 400, { error: "Неизвестный стикерпак" });
      sticker = {
        id: uid("stk"),
        text: pack.text,
        emoji: pack.emoji,
        vibe: pack.vibe,
        color: pack.vibe === "pink" ? "pink" : pack.vibe === "rose" ? "rose" : pack.vibe || "cyan",
        kind: "pack",
        packId: pack.packId,
        x: Number(body.x) || 40 + Math.random() * 160,
        y: Number(body.y) || 40 + Math.random() * 120,
        rotation: body.rotation ?? Math.round((Math.random() * 10 - 5) * 10) / 10,
        updatedAt: new Date().toISOString(),
      };
    } else {
      sticker = {
        id: uid("stk"),
        text: String(body.text || "Новый приоритет"),
        x: Number(body.x) || 40 + Math.random() * 120,
        y: Number(body.y) || 40 + Math.random() * 80,
        color: body.color || "cyan",
        kind: "note",
        rotation: body.rotation ?? Math.round((Math.random() * 8 - 4) * 10) / 10,
        updatedAt: new Date().toISOString(),
      };
    }
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

  // Morning: save plan (+ optional seed mandatory KPIs)
  if (method === "POST" && pathname === "/api/forms/morning") {
    const body = await readBody(req);
    const managerId = body.managerId;
    const date = body.date || todayISO();
    const seedKpis = body.seedKpis !== false;
    let tasks = Array.isArray(body.tasks) ? body.tasks.map(normalizeTask).filter((t) => t.text) : [];
    if (!managerId) return sendJson(res, 400, { error: "Не выбран менеджер" });

    const db = await readDb();
    const manager = db.managers.find((m) => m.id === managerId);
    if (!manager) return sendJson(res, 404, { error: "Менеджер не найден" });

    if (!tasks.length && seedKpis) {
      tasks = (db.kpiDefs || DEFAULT_KPI_DEFS).map((k, i) =>
        normalizeTask(
          {
            text: k.name,
            target: k.defaultTarget,
            doneCount: 0,
            unit: k.unit,
            mandatory: true,
            kpiId: k.id,
          },
          i
        )
      );
    }
    if (!tasks.length) return sendJson(res, 400, { error: "Добавьте хотя бы одну задачу" });

    const result = upsertMorningPlan(db, manager, date, tasks);
    if (result.error) return sendJson(res, 400, { error: result.error });
    await writeDb(db);
    return sendJson(res, 200, {
      form: result.form,
      tomorrowPreview: buildTomorrowPreview(db),
      message:
        date === todayISO()
          ? "Чеклист на день сохранён. Вечером укажите прогресс и переносы."
          : `План на ${date} сохранён. Задачи будут ждать в этот день.`,
    });
  }

  // Week plan: save morning checklists for several days ahead
  if (method === "POST" && pathname === "/api/forms/week") {
    const body = await readBody(req);
    const managerId = body.managerId;
    const days = Array.isArray(body.days) ? body.days : [];
    if (!managerId) return sendJson(res, 400, { error: "Не выбран менеджер" });
    if (!days.length) return sendJson(res, 400, { error: "Нет дней для сохранения" });

    const db = await readDb();
    const manager = db.managers.find((m) => m.id === managerId);
    if (!manager) return sendJson(res, 404, { error: "Менеджер не найден" });

    const saved = [];
    for (const day of days) {
      const date = day.date;
      if (!date) continue;
      const tasks = Array.isArray(day.tasks)
        ? day.tasks.map(normalizeTask).filter((t) => t.text)
        : [];
      if (!tasks.length) continue;
      const result = upsertMorningPlan(db, manager, date, tasks);
      if (result.error) return sendJson(res, 400, { error: result.error });
      saved.push(result.form);
    }
    if (!saved.length) return sendJson(res, 400, { error: "Добавьте задачи хотя бы на один день" });

    await writeDb(db);
    return sendJson(res, 200, {
      forms: saved,
      days: buildWeekPreview(db, managerId, body.from || todayISO()),
      tomorrowPreview: buildTomorrowPreview(db),
      message: `Сохранён план на ${saved.length} дн.`,
    });
  }

  // Evening: progress + carry leftovers to chosen dates
  if (method === "POST" && pathname === "/api/forms/evening") {
    const body = await readBody(req);
    const managerId = body.managerId;
    const date = body.date || todayISO();
    if (!managerId) return sendJson(res, 400, { error: "Не выбран менеджер" });

    const db = await readDb();
    const manager = db.managers.find((m) => m.id === managerId);
    if (!manager) return sendJson(res, 404, { error: "Менеджер не найден" });

    let form = db.forms.find(
      (f) => f.managerId === managerId && f.date === date && f.status !== "completed"
    );
    if (!form) {
      return sendJson(res, 400, { error: "Сначала сохраните утренний чеклист задач" });
    }

    const incoming = Array.isArray(body.tasks) ? body.tasks.map(normalizeTask) : form.tasks.map(normalizeTask);
    const createdTransfers = [];

    form.tasks = [];
    for (const t of incoming) {
      const target = Math.max(1, Number(t.target) || 1);
      const doneCount = Math.max(0, Math.min(target, Number(t.doneCount) || 0));
      const remainder = target - doneCount;
      const transferDate = t.transferDate || (remainder > 0 ? addDaysISO(date, 1) : null);

      const saved = {
        ...t,
        target,
        doneCount,
        done: doneCount >= target,
        transferDate: remainder > 0 ? transferDate : null,
      };

      if (remainder > 0 && transferDate) {
        if (transferDate <= date) {
          return sendJson(res, 400, {
            error: `Дата переноса для «${t.text}» должна быть позже ${date}`,
          });
        }
        const destForm = await ensureDayForm(db, manager, transferDate);
        if (destForm.status === "completed") {
          destForm.status = "morning";
          destForm.completedAt = null;
        }
        const newTask = normalizeTask({
          text: t.text,
          target: remainder,
          doneCount: 0,
          unit: t.unit,
          mandatory: t.mandatory,
          kpiId: t.kpiId,
          carriedFrom: {
            date,
            taskId: saved.id,
            amount: remainder,
            doneCount,
            originalTarget: target,
          },
        });
        destForm.tasks.push(newTask);
        destForm.updatedAt = new Date().toISOString();
        destForm.managerName = manager.name;
        destForm.title = formTitle(manager.name, transferDate);

        saved.carriedTo = {
          date: transferDate,
          taskId: newTask.id,
          amount: remainder,
        };

        const transfer = {
          id: uid("tr"),
          managerId,
          managerName: manager.name,
          fromDate: date,
          toDate: transferDate,
          fromTaskId: saved.id,
          toTaskId: newTask.id,
          text: t.text,
          amount: remainder,
          doneCount,
          originalTarget: target,
          unit: t.unit || "",
          createdAt: new Date().toISOString(),
        };
        db.transfers.push(transfer);
        createdTransfers.push(transfer);
      }

      form.tasks.push(saved);
    }

    form.managerName = manager.name;
    form.title = formTitle(manager.name, date);
    form.status = "completed";
    form.updatedAt = new Date().toISOString();
    form.completedAt = new Date().toISOString();
    await writeDb(db);

    return sendJson(res, 200, {
      form,
      transfers: createdTransfers,
      message: createdTransfers.length
        ? `День закрыт. Перенесено задач: ${createdTransfers.length}`
        : `Форма «${form.title}» сохранена в архив`,
      analytics: buildAnalytics(db),
      dashboard: buildDashboard(db, 14),
    });
  }

  // Manual transfer of selected tasks without full evening close
  if (method === "POST" && pathname === "/api/forms/transfer") {
    const body = await readBody(req);
    const managerId = body.managerId;
    const fromDate = body.fromDate || todayISO();
    const toDate = body.toDate;
    const taskIds = Array.isArray(body.taskIds) ? body.taskIds : [];
    if (!managerId || !toDate) return sendJson(res, 400, { error: "Нужны managerId и toDate" });
    if (toDate <= fromDate) return sendJson(res, 400, { error: "Дата переноса должна быть позже" });

    const db = await readDb();
    const manager = db.managers.find((m) => m.id === managerId);
    if (!manager) return sendJson(res, 404, { error: "Менеджер не найден" });
    const form = db.forms.find((f) => f.managerId === managerId && f.date === fromDate);
    if (!form) return sendJson(res, 404, { error: "Форма исходного дня не найдена" });

    const created = [];
    const destForm = await ensureDayForm(db, manager, toDate);
    if (destForm.status === "completed") {
      destForm.status = "morning";
      destForm.completedAt = null;
    }

    for (const task of form.tasks) {
      if (taskIds.length && !taskIds.includes(task.id)) continue;
      if (task.carriedTo) continue;
      const target = Math.max(1, Number(task.target) || 1);
      const doneCount = Math.max(0, Math.min(target, Number(task.doneCount) || 0));
      const remainder = body.amount != null ? Number(body.amount) : target - doneCount;
      if (remainder <= 0) continue;

      const newTask = normalizeTask({
        text: task.text,
        target: remainder,
        doneCount: 0,
        unit: task.unit,
        mandatory: task.mandatory,
        kpiId: task.kpiId,
        carriedFrom: {
          date: fromDate,
          taskId: task.id,
          amount: remainder,
          doneCount,
          originalTarget: target,
        },
      });
      destForm.tasks.push(newTask);
      task.doneCount = doneCount;
      task.done = doneCount >= target;
      task.carriedTo = { date: toDate, taskId: newTask.id, amount: remainder };
      const transfer = {
        id: uid("tr"),
        managerId,
        managerName: manager.name,
        fromDate,
        toDate,
        fromTaskId: task.id,
        toTaskId: newTask.id,
        text: task.text,
        amount: remainder,
        doneCount,
        originalTarget: target,
        unit: task.unit || "",
        createdAt: new Date().toISOString(),
      };
      db.transfers.push(transfer);
      created.push(transfer);
    }
    form.updatedAt = new Date().toISOString();
    destForm.updatedAt = new Date().toISOString();
    await writeDb(db);
    return sendJson(res, 200, {
      transfers: created,
      form,
      destForm,
      message: `Перенесено: ${created.length}`,
    });
  }

  if (method === "DELETE" && pathname.startsWith("/api/forms/")) {
    const id = pathname.split("/").pop();
    const db = await readDb();
    const before = db.forms.length;
    db.forms = db.forms.filter((f) => f.id !== id);
    if (db.forms.length === before) return sendJson(res, 404, { error: "Форма не найдена" });
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
    return sendJson(res, 200, { archive, transfers: db.transfers || [] });
  }

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
    if (pathname === "/healthz" || pathname === "/api/health") {
      return sendJson(res, 200, { ok: true });
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

readDb().catch((err) => console.error("Initial DB load failed:", err));
