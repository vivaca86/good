(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  if (root) root.GoodCore = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";

  function normalizeCode(value) {
    return String(value == null ? "" : value)
      .trim()
      .replace(/^\uFEFF+/, "")
      .replace(/^[`'‘’＇“”]+/, "")
      .replace(/\s+/g, "")
      .toUpperCase();
  }

  function parseQuantity(value) {
    const raw = String(value == null ? "" : value).trim();
    if (raw === "") return { ok: true, value: 0, text: "" };
    if (!/^\d+$/.test(raw)) {
      return { ok: false, error: "재고는 0 이상의 정수로 입력하세요." };
    }
    const number = Number(raw);
    if (!Number.isSafeInteger(number) || number < 0) {
      return { ok: false, error: "재고 값이 너무 큽니다." };
    }
    return { ok: true, value: number, text: String(number) };
  }

  function validateInventoryRow(row) {
    const quantity = parseQuantity(row && row.qty);
    if (!quantity.ok) return quantity;
    const code = normalizeCode(row && row.code);
    if (quantity.value > 0 && !code) {
      return { ok: false, error: "재고가 1개 이상이면 코드를 먼저 입력하세요." };
    }
    return { ok: true, value: quantity.value, text: quantity.text, code };
  }

  function aggregateByCode(rows) {
    const byCode = {};
    (Array.isArray(rows) ? rows : []).forEach(function (row) {
      const code = normalizeCode(row && row.code);
      if (!code) return;
      const parsed = parseQuantity(row && row.qty);
      if (!parsed.ok) return;
      if (!byCode[code]) {
        byCode[code] = { code, name: String((row && row.name) || "").trim(), qty: 0 };
      }
      byCode[code].qty += parsed.value;
      if (!byCode[code].name && row && row.name) byCode[code].name = String(row.name).trim();
    });
    return byCode;
  }

  function boardFingerprint(board, slotsPerBoard) {
    const total = Number.isInteger(slotsPerBoard) && slotsPerBoard > 0 ? slotsPerBoard : 12;
    const rows = Array.isArray(board) ? board : [];
    return JSON.stringify(
      Array.from({ length: total }, function (_, index) {
        const row = rows[index] && typeof rows[index] === "object" ? rows[index] : {};
        const quantity = parseQuantity(row.qty);
        return {
          slot_no: index + 1,
          code: normalizeCode(row.code),
          name: String(row.name == null ? "" : row.name).trim(),
          spec: String(row.spec == null ? "" : row.spec).trim(),
          qty: quantity.ok ? quantity.value : String(row.qty == null ? "" : row.qty).trim()
        };
      })
    );
  }

  function createRequestId() {
    if (typeof crypto !== "undefined" && crypto.randomUUID) return crypto.randomUUID();
    return "save-" + Date.now().toString(36) + "-" + Math.random().toString(36).slice(2);
  }

  return {
    aggregateByCode,
    boardFingerprint,
    createRequestId,
    normalizeCode,
    parseQuantity,
    validateInventoryRow
  };
});
