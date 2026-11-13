import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import pg from "pg";

const schema = readFileSync(join(dirname(fileURLToPath(import.meta.url)), "schema.sql"), "utf8");

function rowToWork(row, { withBytes = false } = {}) {
  const work = {
    id: row.id,
    name: row.name,
    sheetName: row.sheet_name || "",
    slides: row.slides ?? null,
    active: row.active || 0,
    createdAt: row.created_at instanceof Date ? row.created_at.toISOString() : row.created_at,
    updatedAt: row.updated_at instanceof Date ? row.updated_at.toISOString() : row.updated_at,
  };
  if (withBytes) work.bytes = row.bytes;
  return work;
}

export async function createPgStore(connection) {
  const pool = new pg.Pool(typeof connection === "string" ? { connectionString: connection, max: 5 } : { ...connection, max: 5 });
  await pool.query(schema);
  return {
    async close() {
      await pool.end();
    },

    async createUser({ email, passwordHash }) {
      const result = await pool.query(
        "INSERT INTO users (email, password_hash) VALUES ($1, $2) RETURNING id, email",
        [email, passwordHash],
      );
      return result.rows[0];
    },

    async findUserByEmail(email) {
      const result = await pool.query(
        "SELECT id, email, password_hash FROM users WHERE email = $1",
        [email],
      );
      const row = result.rows[0];
      if (!row) return null;
      return { id: row.id, email: row.email, passwordHash: row.password_hash };
    },

    async createSession({ tokenHash, userId, expiresAt }) {
      await pool.query(
        "INSERT INTO sessions (token_hash, user_id, expires_at) VALUES ($1, $2, $3)",
        [tokenHash, userId, expiresAt],
      );
    },

    async deleteSession(tokenHash) {
      await pool.query("DELETE FROM sessions WHERE token_hash = $1", [tokenHash]);
    },

    async findSession(tokenHash) {
      const result = await pool.query(
        `SELECT u.id, u.email, s.expires_at
         FROM sessions s
         JOIN users u ON u.id = s.user_id
         WHERE s.token_hash = $1 AND s.expires_at > now()`,
        [tokenHash],
      );
      const row = result.rows[0];
      if (!row) return null;
      return { expiresAt: row.expires_at, user: { id: row.id, email: row.email } };
    },

    async listWorks(userId) {
      const result = await pool.query(
        `SELECT id, name, sheet_name, active, created_at, updated_at
         FROM works
         WHERE user_id = $1
         ORDER BY updated_at DESC`,
        [userId],
      );
      return result.rows.map((row) => rowToWork(row));
    },

    async createWork({ userId, name, bytes, sheetName, slides, active }) {
      const result = await pool.query(
        `INSERT INTO works (user_id, name, bytes, sheet_name, slides, active)
         VALUES ($1, $2, $3, $4, $5::jsonb, $6)
         RETURNING id, name, sheet_name, slides, active, created_at, updated_at`,
        [userId, name, bytes, sheetName || "", slides == null ? null : JSON.stringify(slides), active || 0],
      );
      return rowToWork(result.rows[0]);
    },

    async getWork(userId, id) {
      const result = await pool.query(
        `SELECT id, name, bytes, sheet_name, slides, active, created_at, updated_at
         FROM works
         WHERE id = $1 AND user_id = $2`,
        [id, userId],
      );
      const row = result.rows[0];
      return row ? rowToWork(row, { withBytes: true }) : null;
    },

    async updateWork(userId, id, patch) {
      const fields = [];
      const values = [];
      const add = (column, value, cast = "") => {
        values.push(value);
        fields.push(`${column} = $${values.length}${cast}`);
      };
      if (patch.name !== undefined) add("name", patch.name);
      if (patch.bytes !== undefined) add("bytes", patch.bytes);
      if (patch.sheetName !== undefined) add("sheet_name", patch.sheetName);
      if (patch.slides !== undefined) add("slides", patch.slides == null ? null : JSON.stringify(patch.slides), "::jsonb");
      if (patch.active !== undefined) add("active", patch.active);
      if (!fields.length) return this.getWork(userId, id);
      values.push(id, userId);
      const result = await pool.query(
        `UPDATE works
         SET ${fields.join(", ")}, updated_at = now()
         WHERE id = $${values.length - 1} AND user_id = $${values.length}
         RETURNING id, name, sheet_name, slides, active, created_at, updated_at`,
        values,
      );
      const row = result.rows[0];
      return row ? rowToWork(row) : null;
    },

    async deleteWork(userId, id) {
      const result = await pool.query(
        "DELETE FROM works WHERE id = $1 AND user_id = $2 RETURNING id",
        [id, userId],
      );
      return result.rowCount > 0;
    },
  };
}
