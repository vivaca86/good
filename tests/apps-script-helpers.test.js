"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

function loadAppsScriptHelpers() {
  const values = new Map();
  const properties = {
    getProperty(key) { return values.has(key) ? values.get(key) : null; },
    setProperty(key, value) { values.set(key, String(value)); },
    deleteProperty(key) { values.delete(key); }
  };
  const context = {
    console,
    ContentService: {
      MimeType: { JSON: "json" },
      createTextOutput(value) {
        return { value, setMimeType() { return this; } };
      }
    },
    PropertiesService: { getScriptProperties() { return properties; } }
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.join(__dirname, "..", "apps_script.gs"), "utf8");
  vm.runInContext(source, context, { filename: "apps_script.gs" });
  return context;
}

test("server validation rejects malformed, fractional, negative, and code-less positive qty", () => {
  const app = loadAppsScriptHelpers();
  for (const qty of ["abc", "1.5", "-1"]) {
    assert.equal(app.validateSlotItem({ zone: 1, board: 1, slot_no: 1, code: "A", qty }).ok, false);
  }
  assert.equal(app.validateSlotItem({ zone: 1, board: 1, slot_no: 1, code: "", qty: 2 }).code, "CODE_REQUIRED");
  const valid = app.validateSlotItem({ zone: 1, board: 1, slot_no: 1, code: " a 1 ", qty: "002" });
  assert.equal(valid.ok, true);
  assert.equal(valid.item.code, "A1");
  assert.equal(valid.item.qty, 2);
  assert.equal(app.normalizeCode("“A"), "A");
  assert.equal(app.normalizeCode("＇A"), "A");
});

test("server derives usage direction from current sheet quantity", () => {
  const app = loadAppsScriptHelpers();
  const logs = [];
  app.appendInventoryDiffLogs(logs, 1, 2, "A", 10, "A", 7);
  app.appendInventoryDiffLogs(logs, 1, 2, "A", 7, "B", 3);
  assert.equal(logs.length, 3);
  assert.equal(logs[0][4], "OUT");
  assert.equal(logs[0][5], 3);
  assert.deepEqual(logs.slice(1).map(row => [row[3], row[4], row[5]]), [["A", "OUT", 7], ["B", "IN", 3]]);
});

test("request result store replays a confirmed batch and retains its fingerprint", () => {
  const app = loadAppsScriptHelpers();
  app.rememberBulkResult("request-123", "fingerprint", { status: "OK", logged: 1 });
  const stored = app.getStoredBulkResult("request-123");
  assert.equal(stored.fingerprint, "fingerprint");
  assert.equal(stored.response.status, "OK");
  assert.equal(stored.response.logged, 1);
});

test("bulkSave returns the stored result before touching a spreadsheet", () => {
  const app = loadAppsScriptHelpers();
  const items = [{ zone: 1, board: 1, slot_no: 1, code: "A", name: "", spec: "", qty: 2 }];
  app.rememberBulkResult("request-456", JSON.stringify(items), { status: "OK", updated: 1, created: 0, logged: 1 });
  app.LockService = {
    getScriptLock() {
      return { tryLock() { return true; }, releaseLock() {} };
    }
  };
  const output = app.bulkSave({ requestId: "request-456", items });
  const parsed = JSON.parse(output.value);
  assert.equal(parsed.status, "OK");
  assert.equal(parsed.replayed, true);
  assert.equal(parsed.logged, 1);
});

test("safe board deletion replays a completed request without touching the sheet", () => {
  const app = loadAppsScriptHelpers();
  const requestId = "delete-request-123";
  app.rememberDeleteRequest(requestId, {
    status: "DONE",
    zone: 1,
    board: 2,
    expectedFingerprint: "fingerprint"
  });
  app.LockService = {
    getScriptLock() {
      return { tryLock() { return true; }, releaseLock() {} };
    }
  };
  const output = app.deleteBoardRows({
    requestId,
    zone: 1,
    board: 2,
    expectedFingerprint: "fingerprint"
  });
  const parsed = JSON.parse(output.value);
  assert.equal(parsed.status, "OK");
  assert.equal(parsed.replayed, true);
});

test("bulk log request ids prevent duplicate rows during a retry", () => {
  const app = loadAppsScriptHelpers();
  const rows = [["timestamp", "zone", "board", "code", "type", "qty", "request_id"]];
  rows.push([new Date(), 1, 1, "A", "IN", 2, "save-request-123:0"]);
  const sheet = {
    getLastRow() { return rows.length; },
    getRange(row, column, rowCount, columnCount) {
      return {
        getValue() { return rows[row - 1]?.[column - 1] || ""; },
        setValue(value) {
          if (!rows[row - 1]) rows[row - 1] = [];
          rows[row - 1][column - 1] = value;
        },
        getValues() {
          return Array.from({ length: rowCount }, (_, offset) => [rows[row - 1 + offset]?.[column - 1] || ""]);
        },
        setValues(values) {
          values.forEach((valueRow, offset) => {
            rows[row - 1 + offset] = valueRow.slice(0, columnCount);
          });
        }
      };
    }
  };
  const logged = app.appendMissingBulkLogs(sheet, [
    { at: new Date().toISOString(), zone: 1, board: 1, code: "A", type: "IN", qty: 2 },
    { at: new Date().toISOString(), zone: 1, board: 1, code: "A", type: "OUT", qty: 1 }
  ], "save-request-123");
  assert.equal(logged, 2);
  assert.equal(rows.filter(row => row[6] === "save-request-123:0").length, 1);
  assert.equal(rows.filter(row => row[6] === "save-request-123:1").length, 1);
});
