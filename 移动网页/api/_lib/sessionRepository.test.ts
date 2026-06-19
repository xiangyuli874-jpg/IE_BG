import assert from "node:assert/strict";
import test from "node:test";
import { createSession } from "../../shared/sessionCore.js";
import { loadOrMigrateSession } from "./sessionRepository.js";

test("returns a Vercel session without calling the Netlify migration source", async () => {
  const session = createSession("a-line", "ordinary-washer-dryer", "Vercel");
  let legacyCalls = 0;

  const result = await loadOrMigrateSession(session.id, {
    load: async () => session,
    save: async (value) => value,
    loadLegacy: async () => {
      legacyCalls += 1;
      return null;
    }
  });

  assert.equal(result?.id, session.id);
  assert.equal(legacyCalls, 0);
});

test("migrates a missing Netlify session into Vercel storage", async () => {
  const legacy = createSession("a-line", "ordinary-washer-dryer", "Legacy");
  let savedId = "";

  const result = await loadOrMigrateSession(legacy.id, {
    load: async () => null,
    save: async (session) => {
      savedId = session.id;
      return session;
    },
    loadLegacy: async () => legacy
  });

  assert.equal(result?.modelName, "Legacy");
  assert.equal(savedId, legacy.id);
});

test("does not save when the Netlify source does not contain the session", async () => {
  let saveCalls = 0;

  const result = await loadOrMigrateSession("wt_1234567890abcdef", {
    load: async () => null,
    save: async (session) => {
      saveCalls += 1;
      return session;
    },
    loadLegacy: async () => null
  });

  assert.equal(result, null);
  assert.equal(saveCalls, 0);
});
