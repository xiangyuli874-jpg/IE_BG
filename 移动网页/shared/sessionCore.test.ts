import assert from "node:assert/strict";
import test from "node:test";
import { createSession, normalizeSession } from "./sessionCore.js";

test("creates the washer-dryer structure with blank process names", () => {
  const session = createSession("a-line", "ordinary-washer-dryer", "XQG100-Test");

  assert.equal(session.modelName, "XQG100-Test");
  assert.deepEqual(session.groups.map((group) => group.processes.length), [19, 24, 36, 27]);
  assert.equal(session.groups.every((group) => group.processes.every((process) => process.name === "")), true);
});

test("normalizes process order and defaults invalid people to one-person export behavior", () => {
  const session = createSession("a-line", "ordinary-washer-dryer", "XQG100-Test");
  session.groups[0].processes[0].order = 99;
  session.groups[0].processes[0].people = 0;

  const normalized = normalizeSession(session);

  assert.equal(normalized.groups[0].processes[0].order, 1);
  assert.equal(normalized.groups[0].processes[0].stationNo, 1);
  assert.equal(normalized.groups[0].processes[0].people, null);
});

test("rejects sessions that exceed the one-megabyte request limit", () => {
  const session = createSession("a-line", "ordinary-washer-dryer", "XQG100-Test");
  session.groups[0].processes[0].remark = "x".repeat(1_100_000);

  assert.throws(() => normalizeSession(session), /请求内容过大/);
});
