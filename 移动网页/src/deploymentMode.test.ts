import assert from "node:assert/strict";
import test from "node:test";
import { requiresAccessCode } from "./deploymentMode";

test("does not require an access code unless a deployment explicitly enables one", () => {
  assert.equal(requiresAccessCode(undefined), false);
  assert.equal(requiresAccessCode("true"), true);
  assert.equal(requiresAccessCode("false"), false);
});
