import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import test from "node:test";

test("deploys ship the current master commit, not some other branch", () => {
  const yaml = readFileSync(new URL("../../.github/workflows/deploy-web.yml", import.meta.url), "utf8");
  assert.equal(yaml.includes("branch_name"), false);
  assert.match(yaml, /ref:\s*\$\{\{\s*github\.sha\s*\}\}/);
  assert.match(yaml, /merge-base --is-ancestor origin\/master HEAD/);
});
