import { randomUUID } from "node:crypto";

export function createMemoryStore() {
  const users = [];
  const sessions = new Map();
  const works = [];

  function publicWork(work, { withBytes = false } = {}) {
    const row = {
      id: work.id,
      name: work.name,
      sheetName: work.sheetName,
      active: work.active,
      updatedAt: work.updatedAt,
      slides: work.slides,
    };
    if (withBytes) row.bytes = work.bytes;
    return row;
  }

  return {
    async createUser({ email, passwordHash }) {
      const user = { id: randomUUID(), email, passwordHash };
      users.push(user);
      return { id: user.id, email: user.email };
    },

    async findUserByEmail(email) {
      const user = users.find((item) => item.email === email);
      if (!user) return null;
      return { id: user.id, email: user.email, passwordHash: user.passwordHash };
    },

    async createSession({ tokenHash, userId, expiresAt }) {
      sessions.set(tokenHash, { userId, expiresAt: new Date(expiresAt).toISOString() });
    },

    async deleteSession(tokenHash) {
      sessions.delete(tokenHash);
    },

    async findSession(tokenHash, now = new Date()) {
      const session = sessions.get(tokenHash);
      if (!session) return null;
      if (new Date(session.expiresAt).getTime() <= now.getTime()) {
        sessions.delete(tokenHash);
        return null;
      }
      const user = users.find((item) => item.id === session.userId);
      if (!user) return null;
      return { expiresAt: session.expiresAt, user: { id: user.id, email: user.email } };
    },

    async listWorks(userId) {
      return works
        .filter((work) => work.userId === userId)
        .sort((a, b) => new Date(b.updatedAt) - new Date(a.updatedAt))
        .map((work) => {
          const row = publicWork(work);
          delete row.slides;
          return row;
        });
    },

    async createWork({ userId, name, bytes, sheetName, slides, active }) {
      const now = new Date().toISOString();
      const work = {
        id: randomUUID(),
        userId,
        name,
        bytes: Buffer.from(bytes),
        sheetName: sheetName || "",
        slides: slides ?? null,
        active: active || 0,
        updatedAt: now,
      };
      works.push(work);
      return publicWork(work);
    },

    async getWork(userId, id) {
      const work = works.find((item) => item.id === id && item.userId === userId);
      return work ? publicWork(work, { withBytes: true }) : null;
    },

    async updateWork(userId, id, patch) {
      const work = works.find((item) => item.id === id && item.userId === userId);
      if (!work) return null;
      if (patch.name !== undefined) work.name = patch.name;
      if (patch.bytes !== undefined) work.bytes = Buffer.from(patch.bytes);
      if (patch.sheetName !== undefined) work.sheetName = patch.sheetName;
      if (patch.slides !== undefined) work.slides = patch.slides;
      if (patch.active !== undefined) work.active = patch.active;
      work.updatedAt = new Date().toISOString();
      return publicWork(work);
    },
  };
}
