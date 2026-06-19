import assert from "node:assert/strict";
import test from "node:test";
import { RATE_LIMIT_POLICIES } from "./rateLimit.js";

test("uses the approved wide rate limits for each public operation", () => {
  assert.deepEqual(RATE_LIMIT_POLICIES, {
    create: { requests: 30, window: "1 h" },
    read: { requests: 300, window: "10 m" },
    save: { requests: 600, window: "10 m" },
    export: { requests: 60, window: "1 h" }
  });
});
