/**
 * Cloudflare Worker: POST /api/yandex-submit
 * Body: { surveyId?, values: { answer_short_text_...: "..." } }
 */
const DEFAULT_SURVEY_ID = "6a83b54feb614605f98e10ee";

function parseSetCookies(res) {
  if (typeof res.headers.getSetCookie === "function") return res.headers.getSetCookie();
  const single = res.headers.get("set-cookie");
  return single ? [single] : [];
}

function storeCookies(jar, headers) {
  for (const raw of headers) {
    const part = String(raw).split(";")[0];
    const eq = part.indexOf("=");
    if (eq > 0) jar.set(part.slice(0, eq).trim(), part.slice(eq + 1).trim());
  }
}

function cookieHeader(jar) {
  return [...jar.entries()].map(([k, v]) => `${k}=${v}`).join("; ");
}

async function yandexFetch(url, { method = "GET", body, jar, csrf, surveyId } = {}) {
  const headers = {
    Accept: "application/json, text/plain, */*",
    Referer: `https://forms.yandex.ru/u/${surveyId}/`,
  };
  if (jar.size) headers.Cookie = cookieHeader(jar);
  if (csrf) headers["x-csrf-token"] = csrf;
  if (body !== undefined) headers["Content-Type"] = "application/json";
  const res = await fetch(url, {
    method,
    headers,
    body: body !== undefined ? JSON.stringify(body) : undefined,
  });
  storeCookies(jar, parseSetCookies(res));
  const text = await res.text();
  let json = null;
  try {
    json = text ? JSON.parse(text) : null;
  } catch {}
  return { status: res.status, csrf: res.headers.get("x-csrf-token") || csrf || null, json, text };
}

async function submitToYandex(surveyId, values) {
  const jar = new Map();
  let csrf = null;
  await yandexFetch(`https://forms.yandex.ru/u/${surveyId}/`, { jar, surveyId });
  let probe = await yandexFetch("https://forms.yandex.ru/u/gateway/root/form/getSurvey", {
    method: "POST",
    body: { surveyId },
    jar,
    csrf,
    surveyId,
  });
  csrf = probe.csrf || csrf;
  if (probe.status === 419 || probe.status >= 400) {
    probe = await yandexFetch("https://forms.yandex.ru/u/gateway/root/form/getSurvey", {
      method: "POST",
      body: { surveyId },
      jar,
      csrf,
      surveyId,
    });
    csrf = probe.csrf || csrf;
  }
  if (probe.status >= 400) throw new Error(`getSurvey ${probe.status}`);
  const posted = await yandexFetch("https://forms.yandex.ru/u/gateway/root/form/postSurvey", {
    method: "POST",
    body: {
      surveyId,
      values,
      parent: "",
      dryRun: false,
      timestamp: new Date().toISOString(),
    },
    jar,
    csrf,
    surveyId,
  });
  if (posted.status >= 400) throw new Error(`postSurvey ${posted.status}: ${posted.text.slice(0, 200)}`);
  return posted.json;
}

export default {
  async fetch(request) {
    const url = new URL(request.url);
    if (request.method === "OPTIONS") {
      return new Response(null, {
        status: 204,
        headers: {
          "Access-Control-Allow-Origin": "*",
          "Access-Control-Allow-Methods": "POST,OPTIONS",
          "Access-Control-Allow-Headers": "Content-Type",
        },
      });
    }
    if (request.method === "POST" && url.pathname.endsWith("/api/yandex-submit")) {
      try {
        const payload = await request.json();
        const surveyId = payload.surveyId || DEFAULT_SURVEY_ID;
        const result = await submitToYandex(surveyId, payload.values || {});
        return Response.json(
          { ok: true, result },
          { headers: { "Access-Control-Allow-Origin": "*" } }
        );
      } catch (err) {
        return Response.json(
          { ok: false, error: String(err.message || err) },
          { status: 502, headers: { "Access-Control-Allow-Origin": "*" } }
        );
      }
    }
    return new Response("Not found", { status: 404 });
  },
};
