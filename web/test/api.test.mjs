import assert from "node:assert/strict";
import { once } from "node:events";
import http from "node:http";
import test from "node:test";
import { hashPassword, verifyPassword } from "../server/auth.mjs";
import { createHandler } from "../server/app.mjs";
import { createMemoryStore } from "../server/memory-store.mjs";

function start(store = createMemoryStore()) {
  const handler = createHandler({
    store,
    webRoot: new URL("..", import.meta.url).pathname,
    cookieSecure: (req) => req.headers["x-forwarded-proto"] === "https",
  });
  const server = http.createServer(handler);
  server.listen(0, "127.0.0.1");
  return new Promise((resolve) => {
    server.once("listening", () => {
      const { port } = server.address();
      resolve({ server, base: `http://127.0.0.1:${port}` });
    });
  });
}

async function stop(server) {
  server.close();
  await once(server, "close");
}

async function send(base, path, { method = "GET", body, cookie, proto } = {}) {
  const headers = {};
  if (body !== undefined) headers["content-type"] = "application/json";
  if (cookie) headers.cookie = cookie;
  if (proto) headers["x-forwarded-proto"] = proto;
  const response = await fetch(base + path, {
    method,
    headers,
    body: body === undefined ? undefined : JSON.stringify(body),
  });
  const text = await response.text();
  let json = null;
  if (text) json = JSON.parse(text);
  const setCookie = response.headers.get("set-cookie") || "";
  const token = /myboss_session=([^;]+)/.exec(setCookie)?.[1] || "";
  return { status: response.status, json, setCookie, token };
}

function cookieOf(token) {
  return token ? `myboss_session=${token}` : "";
}

test("password hash verifies the same password only", async () => {
  const stored = await hashPassword("correct horse");
  assert.equal(stored.includes("correct horse"), false);
  assert.equal(await verifyPassword("correct horse", stored), true);
  assert.equal(await verifyPassword("wrong horse", stored), false);
  assert.equal(await verifyPassword("correct horse", "not-a-hash"), false);
});

test("register rejects a bad email and a short password", async () => {
  const { server, base } = await start();
  try {
    const email = await send(base, "/api/register", { method: "POST", body: { email: "not-an-email", password: "long-enough" } });
    assert.equal(email.status, 400);
    const password = await send(base, "/api/register", { method: "POST", body: { email: "ada@example.com", password: "short" } });
    assert.equal(password.status, 400);
  } finally {
    await stop(server);
  }
});

test("register logs in and treats email case as the same account", async () => {
  const { server, base } = await start();
  try {
    const created = await send(base, "/api/register", {
      method: "POST",
      body: { email: "Ada@Example.com", password: "correct horse" },
      proto: "https",
    });
    assert.equal(created.status, 201);
    assert.equal(created.json.email, "ada@example.com");
    assert.match(created.setCookie, /Secure/);
    assert.match(created.setCookie, /HttpOnly/);
    const me = await send(base, "/api/me", { cookie: cookieOf(created.token) });
    assert.equal(me.status, 200);
    assert.equal(me.json.email, "ada@example.com");
    const again = await send(base, "/api/register", {
      method: "POST",
      body: { email: "ada@example.com", password: "another password" },
    });
    assert.equal(again.status, 409);
  } finally {
    await stop(server);
  }
});

test("login fails the same way for an unknown email and a wrong password", async () => {
  const { server, base } = await start();
  try {
    await send(base, "/api/register", { method: "POST", body: { email: "ada@example.com", password: "correct horse" } });
    const wrong = await send(base, "/api/login", { method: "POST", body: { email: "ada@example.com", password: "nope nope" } });
    const missing = await send(base, "/api/login", { method: "POST", body: { email: "nope@example.com", password: "correct horse" } });
    assert.equal(wrong.status, 401);
    assert.equal(missing.status, 401);
    assert.equal(wrong.json.error, missing.json.error);
    const ok = await send(base, "/api/login", { method: "POST", body: { email: "ada@example.com", password: "correct horse" } });
    assert.equal(ok.status, 200);
    assert.equal(ok.json.email, "ada@example.com");
  } finally {
    await stop(server);
  }
});

test("logout ends the session", async () => {
  const { server, base } = await start();
  try {
    const created = await send(base, "/api/register", { method: "POST", body: { email: "ada@example.com", password: "correct horse" } });
    const out = await send(base, "/api/logout", { method: "POST", cookie: cookieOf(created.token) });
    assert.equal(out.status, 204);
    const me = await send(base, "/api/me", { cookie: cookieOf(created.token) });
    assert.equal(me.status, 401);
  } finally {
    await stop(server);
  }
});

test("a saved workbook and its slides belong only to that account", async () => {
  const { server, base } = await start();
  try {
    const ada = await send(base, "/api/register", { method: "POST", body: { email: "ada@example.com", password: "correct horse" } });
    const grace = await send(base, "/api/register", { method: "POST", body: { email: "grace@example.com", password: "correct horse" } });
    const locked = await send(base, "/api/works");
    assert.equal(locked.status, 401);
    const slides = [{ id: "s1", shapes: [{ id: "c1", type: "card", text: "Ada" }] }];
    const saved = await send(base, "/api/works", {
      method: "POST",
      cookie: cookieOf(ada.token),
      body: {
        name: "team.xlsx",
        dataBase64: Buffer.from("excel-bytes").toString("base64"),
        sheetName: "Advanced",
        slides,
        active: 0,
      },
    });
    assert.equal(saved.status, 201);
    assert.equal(saved.json.name, "team.xlsx");
    const list = await send(base, "/api/works", { cookie: cookieOf(ada.token) });
    assert.equal(list.status, 200);
    assert.equal(list.json.works.length, 1);
    assert.equal(list.json.works[0].name, "team.xlsx");
    assert.equal(list.json.works[0].sheetName, "Advanced");
    assert.equal("dataBase64" in list.json.works[0], false);
    const opened = await send(base, `/api/works/${saved.json.id}`, { cookie: cookieOf(ada.token) });
    assert.equal(Buffer.from(opened.json.dataBase64, "base64").toString(), "excel-bytes");
    assert.equal(opened.json.slides[0].shapes[0].text, "Ada");
    const hidden = await send(base, `/api/works/${saved.json.id}`, { cookie: cookieOf(grace.token) });
    assert.equal(hidden.status, 404);
    const moved = await send(base, `/api/works/${saved.json.id}`, {
      method: "PUT",
      cookie: cookieOf(ada.token),
      body: { slides: [{ id: "s1", shapes: [] }, { id: "s2", shapes: [] }], active: 1 },
    });
    assert.equal(moved.status, 200);
    const again = await send(base, `/api/works/${saved.json.id}`, { cookie: cookieOf(ada.token) });
    assert.equal(again.json.active, 1);
    assert.equal(again.json.slides.length, 2);
    assert.equal(Buffer.from(again.json.dataBase64, "base64").toString(), "excel-bytes");
  } finally {
    await stop(server);
  }
});

test("the app page is served and private files are not", async () => {
  const { server, base } = await start();
  try {
    const page = await fetch(base + "/");
    const html = await page.text();
    assert.equal(page.status, 200);
    assert.match(page.headers.get("content-type"), /text\/html/);
    assert.match(html, /MyBoss/);
    const secret = await fetch(base + "/server/schema.sql");
    assert.equal(secret.status, 404);
    const escape = await fetch(base + "/css/../server/schema.sql");
    assert.equal(escape.status, 404);
  } finally {
    await stop(server);
  }
});

test("a card note is stored with the account's slides", async () => {
  const { server, base } = await start();
  try {
    const ada = await send(base, "/api/register", { method: "POST", body: { email: "ada@example.com", password: "correct horse" } });
    const slides = [{
      id: "s1",
      shapes: [
        { id: "c1", type: "card", key: "olga", x: 10, y: 20, w: 88, h: 42 },
        { id: "n1", type: "note", key: "olga", text: "Василиса Премудрая?", fill: "#E8891A", x: 26, y: 52, w: 110, h: 28, dx: 16, dy: 32 },
      ],
    }];
    const saved = await send(base, "/api/works", {
      method: "POST",
      cookie: cookieOf(ada.token),
      body: { name: "team.xlsx", dataBase64: Buffer.from("excel-bytes").toString("base64"), sheetName: "Advanced", slides, active: 0 },
    });
    assert.equal(saved.status, 201);
    const opened = await send(base, `/api/works/${saved.json.id}`, { cookie: cookieOf(ada.token) });
    assert.equal(opened.status, 200);
    assert.equal(opened.json.slides[0].shapes[1].type, "note");
    assert.equal(opened.json.slides[0].shapes[1].text, "Василиса Премудрая?");
    assert.equal(opened.json.slides[0].shapes[1].key, "olga");
    const grace = await send(base, "/api/register", { method: "POST", body: { email: "grace@example.com", password: "correct horse" } });
    const hidden = await send(base, `/api/works/${saved.json.id}`, { cookie: cookieOf(grace.token) });
    assert.equal(hidden.status, 404);
  } finally {
    await stop(server);
  }
});

test("slides that are not a list are rejected", async () => {
  const { server, base } = await start();
  try {
    const ada = await send(base, "/api/register", { method: "POST", body: { email: "ada@example.com", password: "correct horse" } });
    const bad = await send(base, "/api/works", {
      method: "POST",
      cookie: cookieOf(ada.token),
      body: { name: "team.xlsx", dataBase64: Buffer.from("x").toString("base64"), slides: { no: "pe" } },
    });
    assert.equal(bad.status, 400);
  } finally {
    await stop(server);
  }
});
