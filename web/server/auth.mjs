import { createHash, randomBytes, scrypt, timingSafeEqual } from "node:crypto";
import { promisify } from "node:util";

const scryptAsync = promisify(scrypt);
const SESSION_MS = 30 * 24 * 60 * 60 * 1000;

export function normalizeEmail(email) {
  return String(email || "").trim().toLowerCase();
}

export function validEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(normalizeEmail(email));
}

export function validPassword(password) {
  const value = String(password || "");
  return value.length >= 8 && value.length <= 200;
}

export async function hashPassword(password) {
  const salt = randomBytes(16);
  const hash = await scryptAsync(String(password), salt, 32);
  return `scrypt$${salt.toString("base64")}$${Buffer.from(hash).toString("base64")}`;
}

export async function verifyPassword(password, stored) {
  const parts = String(stored || "").split("$");
  if (parts.length !== 3 || parts[0] !== "scrypt") return false;
  const salt = Buffer.from(parts[1], "base64");
  const expected = Buffer.from(parts[2], "base64");
  if (!salt.length || !expected.length) return false;
  const actual = Buffer.from(await scryptAsync(String(password), salt, expected.length));
  if (actual.length !== expected.length) return false;
  return timingSafeEqual(actual, expected);
}

export function sessionExpiry(now = new Date()) {
  return new Date(now.getTime() + SESSION_MS);
}

export function sessionMaxAgeSeconds() {
  return SESSION_MS / 1000;
}

export function newSessionToken() {
  const token = randomBytes(32).toString("base64url");
  return { token, tokenHash: hashToken(token) };
}

export function hashToken(token) {
  return createHash("sha256").update(String(token)).digest("hex");
}
