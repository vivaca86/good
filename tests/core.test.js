"use strict";

const assert = require("node:assert/strict");
const test = require("node:test");
const core = require("../core.js");

test("parseQuantity accepts only non-negative safe integers", () => {
  assert.deepEqual(core.parseQuantity(""), { ok: true, value: 0, text: "" });
  assert.deepEqual(core.parseQuantity(" 007 "), { ok: true, value: 7, text: "7" });
  assert.equal(core.parseQuantity("abc").ok, false);
  assert.equal(core.parseQuantity("-1").ok, false);
  assert.equal(core.parseQuantity("1.5").ok, false);
});

test("positive inventory requires a code", () => {
  assert.equal(core.validateInventoryRow({ code: "", qty: "1" }).ok, false);
  assert.equal(core.validateInventoryRow({ code: "", qty: "0" }).ok, true);
  assert.equal(core.validateInventoryRow({ code: " ab 12 ", qty: "2" }).code, "AB12");
});

test("aggregateByCode sums duplicate database rows", () => {
  const result = core.aggregateByCode([
    { code: " a-1 ", name: "첫 품목", qty: 2 },
    { code: "A-1", name: "", qty: "3" },
    { code: "B-2", name: "둘째", qty: 0 }
  ]);
  assert.equal(result["A-1"].qty, 5);
  assert.equal(result["A-1"].name, "첫 품목");
  assert.equal(result["B-2"].qty, 0);
});

test("request ids are unique enough for consecutive save batches", () => {
  assert.notEqual(core.createRequestId(), core.createRequestId());
});

test("board fingerprints are stable across numeric quantity formatting", () => {
  const left = [{ code: " a 1 ", name: " 품목 ", spec: " 규격 ", qty: "02" }];
  const right = [{ code: "A1", name: "품목", spec: "규격", qty: 2 }];
  assert.equal(core.boardFingerprint(left), core.boardFingerprint(right));
  assert.notEqual(core.boardFingerprint(left), core.boardFingerprint([{ ...right[0], qty: 3 }]));
  assert.equal(core.normalizeCode("“A"), "A");
  assert.equal(core.normalizeCode("＇A"), "A");
});
