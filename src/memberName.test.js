import test from "node:test";
import assert from "node:assert/strict";
import { getMemberDisplayName } from "./memberName.js";

test("prefers an explicit display name for leaderboard labels", () => {
  assert.equal(
    getMemberDisplayName({ displayName: "Asha Rao" }, "asha"),
    "Asha Rao",
  );
});

test("falls back to the login username when no display name exists", () => {
  assert.equal(getMemberDisplayName({}, "mike42"), "mike42");
});
