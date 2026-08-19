import assert from "node:assert/strict";
import test from "node:test";
import { verifyAccessCode } from "./auth.js";

test("allows Netlify API requests without an access code", () => {
  const saved = process.env.SITE_ACCESS_CODE;
  process.env.SITE_ACCESS_CODE = "a-code-that-must-not-be-required";

  try {
    assert.equal(verifyAccessCode(new Request("https://example.test/api/sessions")), null);
  } finally {
    if (saved === undefined) delete process.env.SITE_ACCESS_CODE;
    else process.env.SITE_ACCESS_CODE = saved;
  }
});
