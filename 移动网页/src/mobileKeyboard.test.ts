import assert from "node:assert/strict";
import test from "node:test";
import { isMobileKeyboardVisible } from "./mobileKeyboard";

test("detects a mobile keyboard when a field is focused and the visual viewport shrinks", () => {
  assert.equal(isMobileKeyboardVisible({
    viewportWidth: 390,
    layoutHeight: 844,
    visualHeight: 510,
    hasFocusedField: true
  }), true);
});

test("does not hide the footer for small browser chrome height changes", () => {
  assert.equal(isMobileKeyboardVisible({
    viewportWidth: 390,
    layoutHeight: 844,
    visualHeight: 770,
    hasFocusedField: true
  }), false);
});

test("does not treat a resized desktop viewport as a mobile keyboard", () => {
  assert.equal(isMobileKeyboardVisible({
    viewportWidth: 1024,
    layoutHeight: 768,
    visualHeight: 520,
    hasFocusedField: true
  }), false);
});

test("requires an editable field to be focused", () => {
  assert.equal(isMobileKeyboardVisible({
    viewportWidth: 390,
    layoutHeight: 844,
    visualHeight: 510,
    hasFocusedField: false
  }), false);
});
