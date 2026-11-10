import { createReadStream } from "node:fs";
import { readFile, stat } from "node:fs/promises";
import { extname, join, normalize, sep } from "node:path";
import { gzipSync } from "node:zlib";
import {
  hashPassword,
  hashToken,
  newSessionToken,
  normalizeEmail,
  sessionExpiry,
  sessionMaxAgeSeconds,
  validEmail,
  validPassword,
  verifyPassword,
} from "./auth.mjs";

const FILE_LIMIT = 8_000_000;
const JSON_LIMIT = 12_000_000;
const SLIDES_LIMIT = 4_000_000;
const COOKIE = "myboss_session";
const SECURITY_HEADERS = {
  "x-content-type-options": "nosniff",
  "referrer-policy": "no-referrer",
  "x-frame-options": "DENY",
  "content-security-policy": "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'; base-uri 'self'; form-action 'self'",
};
const gzipped = new Map();

let dummyHash = "";

async function dummyPasswordHash() {
  if (!dummyHash) dummyHash = await hashPassword("not-a-real-password");
  return dummyHash;
}

function httpError(status, message) {
  const error = new Error(message);
  error.status = status;
  return error;
}

function sendJson(res, status, body, extraHeaders) {
  const payload = body === undefined ? "" : JSON.stringify(body);
  const headers = {
    ...SECURITY_HEADERS,
    "content-type": "application/json; charset=utf-8",
    "cache-control": "no-store",
    ...extraHeaders,
  };
  if (!payload) delete headers["content-type"];
  res.writeHead(status, headers);
  res.end(payload);
}

function readCookie(req, name) {
  const raw = req.headers.cookie || "";
  for (const part of raw.split(";")) {
    const text = part.trim();
    const eq = text.indexOf("=");
    if (eq === -1) continue;
    if (text.slice(0, eq) === name) return decodeURIComponent(text.slice(eq + 1));
  }
  return "";
}

// The same app is served at / on its own host and at /myboss on me.dadkakdad.ru.
const MOUNT_PREFIX = "/myboss";

function sessionCookie(token, { secure, maxAge, path = "/" }) {
  const parts = [
    `${COOKIE}=${token}`,
    "HttpOnly",
    `Path=${path}`,
    "SameSite=Lax",
    `Max-Age=${maxAge}`,
  ];
  if (secure) parts.push("Secure");
  return parts.join("; ");
}

function forwardedPrefix(req) {
  return req.headers["x-forwarded-prefix"] === MOUNT_PREFIX ? MOUNT_PREFIX : "";
}

function pathIsMounted(pathname) {
  return pathname === MOUNT_PREFIX || pathname.startsWith(`${MOUNT_PREFIX}/`);
}

function stripMount(url) {
  const path = url.pathname === MOUNT_PREFIX ? "/" : url.pathname.slice(MOUNT_PREFIX.length);
  return new URL(path + url.search, url);
}

function injectBase(html, prefix) {
  if (!prefix || html.includes("<base ")) return html;
  const tag = `<base href="${prefix}/">`;
  if (html.includes("<head>")) return html.replace("<head>", `<head>\n  ${tag}`);
  return `${tag}\n${html}`;
}

async function readJson(req) {
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > JSON_LIMIT) throw httpError(413, "That upload is too large.");
    chunks.push(chunk);
  }
  if (!size) return {};
  try {
    return JSON.parse(Buffer.concat(chunks).toString("utf8"));
  } catch {
    throw httpError(400, "The request was not valid JSON.");
  }
}

function cleanName(name) {
  const base = String(name || "").split(/[/\\]/).pop().replace(/[\u0000-\u001f]/g, "").trim();
  return base.slice(0, 180) || "workbook.xlsx";
}

function cleanSheet(name) {
  return String(name || "").replace(/[\u0000-\u001f]/g, "").trim().slice(0, 200);
}

function normalizeSlides(slides) {
  if (slides == null) return null;
  if (!Array.isArray(slides)) throw httpError(400, "Slides must be a list.");
  if (slides.length > 400) throw httpError(413, "That deck has too many slides.");
  let json = "";
  try {
    json = JSON.stringify(slides);
  } catch {
    throw httpError(400, "Slides must be a list.");
  }
  if (json.length > SLIDES_LIMIT) throw httpError(413, "That deck is too large.");
  return slides;
}

function cleanActive(value, slides) {
  const number = Number(value);
  if (!Number.isInteger(number) || number < 0) return 0;
  if (Array.isArray(slides) && number >= slides.length) return Math.max(0, slides.length - 1);
  return number;
}

function decodeFile(dataBase64) {
  if (typeof dataBase64 !== "string" || !dataBase64.trim()) throw httpError(400, "The workbook is missing.");
  const bytes = Buffer.from(dataBase64, "base64");
  if (!bytes.length) throw httpError(400, "The workbook is empty.");
  if (bytes.length > FILE_LIMIT) throw httpError(413, "That workbook is too large.");
  return bytes;
}

function workJson(work, { withBytes = false } = {}) {
  const body = {
    id: work.id,
    name: work.name,
    sheetName: work.sheetName || "",
    active: work.active || 0,
    updatedAt: work.updatedAt instanceof Date ? work.updatedAt.toISOString() : work.updatedAt,
    slides: work.slides ?? null,
  };
  if (withBytes) body.dataBase64 = Buffer.from(work.bytes).toString("base64");
  return body;
}

async function currentUser(req, store, now) {
  const token = readCookie(req, COOKIE);
  if (!token) return null;
  const session = await store.findSession(hashToken(token), now);
  return session?.user || null;
}

async function startSession(res, store, user, now, secure, cookiePath) {
  const { token, tokenHash } = newSessionToken();
  const expiresAt = sessionExpiry(now);
  await store.createSession({ tokenHash, userId: user.id, expiresAt });
  return sessionCookie(token, { secure, maxAge: sessionMaxAgeSeconds(), path: cookiePath });
}

const STATIC_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
  ".webp": "image/webp",
};

// The HTML always carries a content hash for these files, so a new deploy is a new URL.
const FINGERPRINTED = new Set([
  "app.bundle.js",
  "css/app.css",
  "node_modules/xlsx/dist/xlsx.full.min.js",
  "node_modules/pptxgenjs/dist/pptxgen.bundle.js",
]);

function cacheControlFor(requested) {
  if (FINGERPRINTED.has(requested)) return "public, max-age=31536000, immutable";
  return "no-cache";
}

function etagFor(info) {
  return `"${info.size.toString(16)}-${Math.round(info.mtimeMs).toString(16)}"`;
}

function noneMatch(header, etag) {
  if (!header) return false;
  return header.split(",").some((part) => {
    const value = part.trim().replace(/^W\//, "");
    return value === "*" || value === etag;
  });
}

function staticPath(webRoot, urlPath) {
  let path = urlPath.split("?")[0];
  try {
    path = decodeURIComponent(path);
  } catch {
    return null;
  }
  if (path.includes("\0") || path.includes("\\") || path.split("/").includes("..")) return null;
  const requested = path === "/" ? "index.html" : path.replace(/^\/+/, "");
  const allowed =
    requested === "index.html" ||
    requested === "app.bundle.js" ||
    requested.startsWith("css/") ||
    requested.startsWith("js/") ||
    requested.startsWith("icons/") ||
    requested === "node_modules/xlsx/dist/xlsx.full.min.js" ||
    requested === "node_modules/pptxgenjs/dist/pptxgen.bundle.js";
  if (!allowed) return null;
  const root = normalize(webRoot);
  const full = normalize(join(root, requested));
  const prefix = root.endsWith(sep) ? root : root + sep;
  if (full !== root && !full.startsWith(prefix)) return null;
  return full;
}

function wantsGzip(req, type) {
  const text = type.startsWith("text/") || type.includes("javascript") || type.includes("json") || type.includes("svg");
  return text && /\bgzip\b/.test(req.headers["accept-encoding"] || "");
}

function gzipBuffer(full, info, raw) {
  const key = `${full}\0${info.mtimeMs}\0${info.size}`;
  let hit = gzipped.get(key);
  if (!hit) {
    hit = gzipSync(raw, { level: 6 });
    gzipped.set(key, hit);
  }
  return hit;
}

function serveStatic(req, res, webRoot, urlPath, mountPrefix = "") {
  const requested = staticRequest(urlPath);
  const full = requested && staticPath(webRoot, urlPath);
  if (!full) {
    sendJson(res, 404, { error: "Not found." });
    return;
  }
  stat(full)
    .then(async (info) => {
      if (!info.isFile()) {
        sendJson(res, 404, { error: "Not found." });
        return;
      }
      const type = STATIC_TYPES[extname(full)] || "application/octet-stream";
      const cacheControl = cacheControlFor(requested);
      let etag = etagFor(info);
      let body = null;
      if (mountPrefix && requested === "index.html") {
        body = Buffer.from(injectBase(await readFile(full, "utf8"), mountPrefix));
        etag = `"${info.size.toString(16)}-${Math.round(info.mtimeMs).toString(16)}-mount"`;
      }
      const headers = { ...SECURITY_HEADERS, etag, "cache-control": cacheControl, "content-type": type };
      if (requested === "index.html") headers.vary = "x-forwarded-prefix";
      if (noneMatch(req.headers["if-none-match"], etag)) {
        res.writeHead(304, headers);
        res.end();
        return;
      }
      if (!wantsGzip(req, type)) {
        if (!body) {
          res.writeHead(200, { ...headers, "content-length": info.size });
          createReadStream(full).pipe(res);
          return;
        }
        res.writeHead(200, { ...headers, "content-length": body.length });
        res.end(body);
        return;
      }
      const gz = body ? gzipSync(body, { level: 6 }) : gzipBuffer(full, info, await readFile(full));
      const vary = headers.vary ? `${headers.vary}, accept-encoding` : "accept-encoding";
      res.writeHead(200, {
        ...headers,
        vary,
        "content-encoding": "gzip",
        "content-length": gz.length,
      });
      res.end(gz);
    })
    .catch(() => sendJson(res, 404, { error: "Not found." }));
}

function staticRequest(urlPath) {
  let path = urlPath.split("?")[0];
  try {
    path = decodeURIComponent(path);
  } catch {
    return null;
  }
  if (path === "/") return "index.html";
  return path.replace(/^\/+/, "");
}

async function routeApi(req, res, url, store, now, secure, loginLimited, cookiePath = "/") {
  const user = await currentUser(req, store, now);
  const path = url.pathname;

  if (req.method === "POST" && (path === "/api/register" || path === "/api/login")) {
    if (loginLimited(req, now)) throw httpError(429, "Too many attempts. Wait a few minutes and try again.");
    const body = await readJson(req);
    const email = normalizeEmail(body.email);
    const password = String(body.password || "");
    if (!validEmail(email)) throw httpError(400, "Enter a valid email address.");
    if (!validPassword(password)) throw httpError(400, "Use a password of at least 8 characters.");
    const existing = await store.findUserByEmail(email);
    if (path === "/api/register") {
      if (existing) throw httpError(409, "An account with that email already exists.");
      const created = await store.createUser({ email, passwordHash: await hashPassword(password) });
      const cookie = await startSession(res, store, created, now, secure, cookiePath);
      sendJson(res, 201, { email: created.email }, { "set-cookie": cookie });
      return;
    }
    const matches = await verifyPassword(password, existing ? existing.passwordHash : await dummyPasswordHash());
    if (!existing || !matches) throw httpError(401, "Email or password is wrong.");
    const cookie = await startSession(res, store, existing, now, secure, cookiePath);
    sendJson(res, 200, { email: existing.email }, { "set-cookie": cookie });
    return;
  }

  if (req.method === "POST" && path === "/api/logout") {
    const token = readCookie(req, COOKIE);
    if (token) await store.deleteSession(hashToken(token));
    res.writeHead(204, {
      ...SECURITY_HEADERS,
      "set-cookie": sessionCookie("", { secure, maxAge: 0, path: cookiePath }),
      "cache-control": "no-store",
    });
    res.end();
    return;
  }

  if (req.method === "GET" && path === "/api/me") {
    if (!user) throw httpError(401, "Log in first.");
    sendJson(res, 200, { email: user.email });
    return;
  }

  if (path === "/api/works" || path.startsWith("/api/works/")) {
    if (!user) throw httpError(401, "Log in first.");
    const id = path.slice("/api/works/".length);
    if (req.method === "GET" && path === "/api/works") {
      const works = await store.listWorks(user.id);
      sendJson(res, 200, {
        works: works.map((work) => {
          const row = workJson(work);
          delete row.slides;
          return row;
        }),
      });
      return;
    }
    if (req.method === "POST" && path === "/api/works") {
      const body = await readJson(req);
      const slides = body.slides === undefined ? null : normalizeSlides(body.slides);
      const work = await store.createWork({
        userId: user.id,
        name: cleanName(body.name),
        bytes: decodeFile(body.dataBase64),
        sheetName: cleanSheet(body.sheetName),
        slides,
        active: cleanActive(body.active, slides),
      });
      sendJson(res, 201, workJson(work));
      return;
    }
    if (!/^[0-9a-f-]{36}$/i.test(id)) throw httpError(404, "That file was not found.");
    if (req.method === "GET") {
      const work = await store.getWork(user.id, id);
      if (!work) throw httpError(404, "That file was not found.");
      sendJson(res, 200, workJson(work, { withBytes: true }));
      return;
    }
    if (req.method === "PUT") {
      const current = await store.getWork(user.id, id);
      if (!current) throw httpError(404, "That file was not found.");
      const body = await readJson(req);
      const patch = {};
      if (body.name !== undefined) patch.name = cleanName(body.name);
      if (body.dataBase64 !== undefined) patch.bytes = decodeFile(body.dataBase64);
      if (body.sheetName !== undefined) patch.sheetName = cleanSheet(body.sheetName);
      const slides = body.slides === undefined ? current.slides : normalizeSlides(body.slides);
      if (body.slides !== undefined) patch.slides = slides;
      if (body.active !== undefined || body.slides !== undefined) patch.active = cleanActive(body.active ?? current.active, slides);
      const work = await store.updateWork(user.id, id, patch);
      sendJson(res, 200, workJson(work));
      return;
    }
  }

  throw httpError(404, "Not found.");
}

function clientKey(req) {
  const forwarded = req.headers["x-forwarded-for"];
  if (typeof forwarded === "string" && forwarded.trim()) {
    const parts = forwarded.split(",").map((part) => part.trim()).filter(Boolean);
    if (parts.length) return parts[parts.length - 1];
  }
  return req.socket?.remoteAddress || "unknown";
}

function createLoginLimiter({ max = 20, windowMs = 15 * 60 * 1000 } = {}) {
  const buckets = new Map();
  return function limited(req, now) {
    const key = `${req.url || ""} ${clientKey(req)}`;
    const time = now.getTime();
    const bucket = buckets.get(key);
    if (!bucket || bucket.reset <= time) {
      buckets.set(key, { count: 1, reset: time + windowMs });
      return false;
    }
    bucket.count += 1;
    return bucket.count > max;
  };
}

export function createHandler({ store, webRoot, now = () => new Date(), cookieSecure = () => false, limits } = {}) {
  const loginLimited = createLoginLimiter(limits);
  return function handler(req, res) {
    const raw = new URL(req.url || "/", "http://localhost");
    const mounted = pathIsMounted(raw.pathname);
    const mountPrefix = mounted || forwardedPrefix(req) ? MOUNT_PREFIX : "";
    if (raw.pathname === MOUNT_PREFIX) {
      res.writeHead(308, {
        ...SECURITY_HEADERS,
        location: `${MOUNT_PREFIX}/${raw.search}`,
        "cache-control": "no-store",
      });
      res.end();
      return;
    }
    const url = mounted ? stripMount(raw) : raw;
    const secure = cookieSecure(req);
    const run = url.pathname.startsWith("/api/")
      ? routeApi(req, res, url, store, now(), secure, loginLimited, mountPrefix || "/")
      : Promise.resolve(serveStatic(req, res, webRoot, url.pathname, mountPrefix));
    run.catch((error) => {
      if (res.headersSent) return;
      const status = error.status || 500;
      if (!error.status) console.error(error);
      sendJson(res, status, { error: error.status ? error.message : "Something went wrong." });
    });
  };
}
