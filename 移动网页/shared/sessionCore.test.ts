import assert from "node:assert/strict";
import test from "node:test";
import { LINE_CONFIGS } from "../src/data/presets.js";
import { createSession, normalizeSession } from "./sessionCore.js";

test("enables every requested production line", () => {
  assert.deepEqual(
    LINE_CONFIGS.filter((line) => line.enabled).map((line) => line.name),
    ["A线", "B线", "C线", "D线", "E线", "H线", "底座线"]
  );
});

test("uses the roller group structure for roller lines and adds pan-base after outer tub for pulsator lines", () => {
  for (const lineId of ["a-line", "d-line", "e-line", "h-line"]) {
    const session = createSession(lineId, "ordinary-washer-dryer");
    assert.deepEqual(session.groups.map((group) => group.name), ["外筒班", "前总装", "后总装", "包装班"]);
  }

  for (const lineId of ["b-line", "c-line"]) {
    const session = createSession(lineId, "ordinary-washer-dryer");
    assert.deepEqual(session.groups.map((group) => group.name), ["外筒班", "盘座班", "前总装", "后总装", "包装班"]);
    assert.deepEqual(session.groups.map((group) => group.order), [1, 2, 3, 4, 5]);
    assert.equal(session.groups[1].processes.length, 0);
  }

  const baseLineSession = createSession("base-line", "ordinary-washer-dryer");
  assert.deepEqual(baseLineSession.groups.map((group) => group.name), ["底座线"]);
  assert.equal(baseLineSession.groups[0].processes.length, 0);
});

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
