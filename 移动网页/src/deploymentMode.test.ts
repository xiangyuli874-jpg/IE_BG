import assert from "node:assert/strict";
import test from "node:test";
import { requiresAccessCode } from "./deploymentMode";

test("keeps access-code protection unless Vercel explicitly disables it", () => {
  assert.equal(requiresAccessCode(undefined), true);
  assert.equal(requiresAccessCode("true"), true);
  assert.equal(requiresAccessCode("false"), false);
});
