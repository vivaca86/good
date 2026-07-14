// ============================
// GET : SLOTS 전체 데이터 반환
// ============================
function doGet() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('slots_v1');
  if (cached) {
    return ContentService
      .createTextOutput(cached)
      .setMimeType(ContentService.MimeType.JSON);
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return jsonOutput({ status: 'BUSY', message: '다른 저장 작업이 진행 중입니다.' });
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName('SLOTS');
    const data = sheet.getDataRange().getValues();

    if (!data.length) return jsonOutput([]);

    const headers = data.shift();
    const result = data
      .filter(row => row.some(cell => cell !== ''))
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => {
          obj[h] = row[i];
        });
        return obj;
      });

    const json = JSON.stringify(result);
    cache.put('slots_v1', json, 15);
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// ============================
// POST : action 분기
// ============================
function doPost(e) {
  let payload;

  try {
    payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return ContentService.createTextOutput('INVALID_JSON');
  }

  if (payload.action === 'updateSlot') return updateSlot(payload);
  if (payload.action === 'logUsage') return appendLog(payload);
  if (payload.action === 'createBoard') return createBoardRows(payload);
  if (payload.action === 'deleteBoard' || payload.action === 'deleteBoardSafe') return deleteBoardRows(payload);
  if (payload.action === 'getDbDiff') return getDbDiffRows(payload);
  if (payload.action === 'bulkSave') return bulkSave(payload);
  if (payload.action === 'logClientError') return logClientError(payload);

  return ContentService.createTextOutput('NO_ACTION');
}


function getSlotRowIndexMap(slotSheet) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('slot_row_index_v1');
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (err) {
      // ignore cache parse errors and rebuild below
    }
  }

  const slotData = slotSheet.getDataRange().getValues();
  const rowByKey = {};
  for (let i = 1; i < slotData.length; i++) {
    const key = [String(slotData[i][0]), String(slotData[i][1]), String(slotData[i][2])].join('-');
    rowByKey[key] = i + 1;
  }

  cache.put('slot_row_index_v1', JSON.stringify(rowByKey), 120);
  return rowByKey;
}

function jsonOutput(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function parseNonNegativeInteger(value) {
  if (value === '' || value === null || typeof value === 'undefined') {
    return { ok: true, value: 0 };
  }
  const raw = String(value).trim();
  if (!/^\d+$/.test(raw)) return { ok: false, message: '재고는 0 이상의 정수여야 합니다.' };
  const qty = Number(raw);
  if (!Number.isSafeInteger(qty) || qty < 0) return { ok: false, message: '재고 값이 너무 큽니다.' };
  return { ok: true, value: qty };
}

function validateSlotItem(item) {
  item = item || {};
  const zone = Number(item.zone);
  const board = Number(item.board);
  const slotNo = Number(item.slot_no);
  if (!Number.isInteger(zone) || zone < 1 || zone > 4 ||
      !Number.isInteger(board) || board < 1 ||
      !Number.isInteger(slotNo) || slotNo < 1 || slotNo > 12) {
    return { ok: false, code: 'INVALID_SLOT', message: '구역/현황판/슬롯 번호가 올바르지 않습니다.' };
  }

  const quantity = parseNonNegativeInteger(item.qty);
  if (!quantity.ok) return { ok: false, code: 'INVALID_QTY', message: quantity.message };
  const code = normalizeCode(item.code);
  if (quantity.value > 0 && !code) {
    return { ok: false, code: 'CODE_REQUIRED', message: '재고가 1개 이상이면 코드가 필요합니다.' };
  }

  return {
    ok: true,
    item: {
      zone: zone,
      board: board,
      slot_no: slotNo,
      code: code,
      name: String(item.name || '').trim(),
      spec: String(item.spec || '').trim(),
      qty: quantity.value
    }
  };
}

// ============================
// SLOTS 업데이트 (없으면 생성)
// ============================
function updateSlot(payload) {
  const validation = validateSlotItem(payload);
  if (!validation.ok) return ContentService.createTextOutput(validation.code);
  const item = validation.item;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return ContentService.createTextOutput('BUSY');

  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName('SLOTS');
    const key = [String(item.zone), String(item.board), String(item.slot_no)].join('-');
    const rowByKey = getSlotRowIndexMap(sheet);
    const rowNo = rowByKey[key];

    if (rowNo) {
      sheet.getRange(rowNo, 4, 1, 4).setValues([[
        item.code,
        item.name,
        item.spec,
        item.qty
      ]]);
      invalidateRuntimeCaches();
      return ContentService.createTextOutput('UPDATED');
    }

    sheet.appendRow([
      item.zone,
      item.board,
      item.slot_no,
      item.code,
      item.name,
      item.spec,
      item.qty
    ]);

    invalidateRuntimeCaches();
    return ContentService.createTextOutput('CREATED');
  } finally {
    lock.releaseLock();
  }
}

// ============================
// USAGE_LOG 추가
// ============================
function appendLog(payload) {
  const code = normalizeCode(payload && payload.code);
  const quantity = parseNonNegativeInteger(payload && payload.qty);
  const type = String(payload && payload.type || '').toUpperCase();
  if (!code || !quantity.ok || quantity.value <= 0 || (type !== 'IN' && type !== 'OUT')) {
    return ContentService.createTextOutput('INVALID_LOG');
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return ContentService.createTextOutput('BUSY');
  try {
    const logSheet = SpreadsheetApp.getActive().getSheetByName('USAGE_LOG');
    logSheet.appendRow([
      new Date(),
      payload.zone,
      payload.board,
      code,
      type,
      quantity.value
    ]);

    invalidateRuntimeCaches();
    return ContentService.createTextOutput('LOGGED');
  } finally {
    lock.releaseLock();
  }
}

// ============================
// 현황판 생성 (12 슬롯 기본행 생성)
// ============================
function createBoardRows(payload) {
  const zone = Number(payload.zone);
  const board = Number(payload.board);
  if (!Number.isInteger(zone) || zone < 1 || zone > 4 || !Number.isInteger(board) || board < 1) {
    return ContentService.createTextOutput('INVALID_BOARD');
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return ContentService.createTextOutput('BUSY');
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName('SLOTS');
    const data = sheet.getDataRange().getValues();
    const existing = new Set();
    for (let i = 1; i < data.length; i++) {
      if (Number(data[i][0]) === zone && Number(data[i][1]) === board) {
        existing.add(Number(data[i][2]));
      }
    }

    for (let slot = 1; slot <= 12; slot++) {
      if (existing.has(slot)) continue;
      sheet.appendRow([zone, board, slot, '', '', '', 0]);
    }

    invalidateRuntimeCaches();
    return ContentService.createTextOutput('BOARD_CREATED');
  } finally {
    lock.releaseLock();
  }
}

// ============================
// 현황판 삭제 + 뒤 현황판 번호 당김
// ============================
const DELETE_REQUEST_PREFIX = 'delete_board_v1_';
const DELETE_REQUEST_INDEX = 'delete_board_index_v1';

function getDeleteRequest(requestId) {
  const raw = PropertiesService.getScriptProperties().getProperty(DELETE_REQUEST_PREFIX + requestId);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (err) { return null; }
}

function rememberDeleteRequest(requestId, value) {
  const properties = PropertiesService.getScriptProperties();
  const now = Date.now();
  let index = [];
  try { index = JSON.parse(properties.getProperty(DELETE_REQUEST_INDEX) || '[]'); } catch (err) { index = []; }
  index = index.filter(entry => {
    const keep = entry && entry.id !== requestId && now - Number(entry.at || 0) < 7 * 24 * 60 * 60 * 1000;
    if (!keep && entry && entry.id) properties.deleteProperty(DELETE_REQUEST_PREFIX + entry.id);
    return keep;
  });
  index.push({ id: requestId, at: now });
  while (index.length > 100) {
    const removed = index.shift();
    if (removed && removed.id) properties.deleteProperty(DELETE_REQUEST_PREFIX + removed.id);
  }
  properties.setProperty(
    DELETE_REQUEST_PREFIX + requestId,
    JSON.stringify(Object.assign({ savedAt: now }, value))
  );
  properties.setProperty(DELETE_REQUEST_INDEX, JSON.stringify(index));
}

function boardFingerprintFromRows(data, zone, board) {
  const bySlot = {};
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][0]) !== zone || Number(data[i][1]) !== board) continue;
    bySlot[Number(data[i][2])] = data[i];
  }
  return JSON.stringify(Array.from({ length: 12 }, (_, index) => {
    const slotNo = index + 1;
    const row = bySlot[slotNo] || [];
    const quantity = parseNonNegativeInteger(row[6]);
    return {
      slot_no: slotNo,
      code: normalizeCode(row[3]),
      name: String(row[4] == null ? '' : row[4]).trim(),
      spec: String(row[5] == null ? '' : row[5]).trim(),
      qty: quantity.ok ? quantity.value : String(row[6] || '').trim()
    };
  }));
}

function deleteBoardRows(payload) {
  const zone = Number(payload.zone);
  const board = Number(payload.board);
  const requestId = String(payload.requestId || '').trim();
  const expectedFingerprint = String(payload.expectedFingerprint || '');
  if (!Number.isInteger(zone) || zone < 1 || zone > 4 || !Number.isInteger(board) || board < 1) {
    return jsonOutput({ status: 'INVALID', code: 'INVALID_BOARD', message: '현황판 번호가 올바르지 않습니다.' });
  }
  if (!/^[A-Za-z0-9._:-]{8,128}$/.test(requestId) || !expectedFingerprint) {
    return jsonOutput({ status: 'INVALID', code: 'INVALID_REQUEST', message: '안전한 삭제 요청 정보가 없습니다.' });
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return jsonOutput({ status: 'BUSY', code: 'BUSY', message: '다른 작업이 진행 중입니다.' });
  try {
    const stored = getDeleteRequest(requestId);
    if (stored) {
      if (stored.zone !== zone || stored.board !== board || stored.expectedFingerprint !== expectedFingerprint) {
        return jsonOutput({ status: 'INVALID', code: 'REQUEST_MISMATCH', message: '같은 요청 ID에 다른 삭제 내용이 전달되었습니다.' });
      }
      if (stored.status === 'DONE') {
        return jsonOutput({ status: 'OK', replayed: true, zone: zone, board: board });
      }
      return jsonOutput({ status: 'ERROR', code: 'AMBIGUOUS_DELETE', message: '이전 삭제 요청의 처리 상태를 확인해야 합니다.' });
    }

    const sheet = SpreadsheetApp.getActive().getSheetByName('SLOTS');
    const data = sheet.getDataRange().getValues();
    const zoneBoards = new Set();
    const rowsToDelete = [];

    for (let i = 1; i < data.length; i++) {
      if (Number(data[i][0]) !== zone) continue;
      const rowBoard = Number(data[i][1]);
      zoneBoards.add(rowBoard);
      if (rowBoard === board) rowsToDelete.push(i + 1);
    }

    if (!rowsToDelete.length) return jsonOutput({ status: 'INVALID', code: 'BOARD_NOT_FOUND', message: '현황판을 찾을 수 없습니다.' });
    if (zoneBoards.size <= 1) return jsonOutput({ status: 'INVALID', code: 'MIN_BOARD', message: '현황판은 최소 1개가 필요합니다.' });
    if (boardFingerprintFromRows(data, zone, board) !== expectedFingerprint) {
      return jsonOutput({ status: 'INVALID', code: 'STALE_BOARD', message: '현황판 내용이 다른 사용자에 의해 변경되었습니다.' });
    }

    rememberDeleteRequest(requestId, {
      status: 'PROCESSING',
      zone: zone,
      board: board,
      expectedFingerprint: expectedFingerprint
    });

    for (let i = rowsToDelete.length - 1; i >= 0; i--) sheet.deleteRow(rowsToDelete[i]);

    const remaining = sheet.getDataRange().getValues();
    for (let i = 1; i < remaining.length; i++) {
      if (Number(remaining[i][0]) === zone && Number(remaining[i][1]) > board) {
        sheet.getRange(i + 1, 2).setValue(Number(remaining[i][1]) - 1);
      }
    }

    SpreadsheetApp.flush();
    invalidateRuntimeCaches();
    rememberDeleteRequest(requestId, {
      status: 'DONE',
      zone: zone,
      board: board,
      expectedFingerprint: expectedFingerprint
    });
    return jsonOutput({ status: 'OK', zone: zone, board: board });
  } finally {
    lock.releaseLock();
  }
}


// ============================
// DB구역용: DB 시트 데이터 반환(재고 0 제외)
// ============================
function getDbDiffRows(payload) {
  const useCache = !(payload && payload.forceRefresh === true);
  const cache = CacheService.getScriptCache();
  if (useCache) {
    const cached = cache.get('db_diff_v1');
    if (cached) {
      return ContentService
        .createTextOutput(cached)
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  const dbSheet = SpreadsheetApp.getActive().getSheetByName('DB');

  if (!dbSheet) {
    return ContentService
      .createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const dbData = dbSheet.getDataRange().getValues();
  if (!dbData.length) {
    return ContentService
      .createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 요청 기준 고정 매핑: DB 시트 코드=B, 품명=C, 재고=F
  const codeIdx = 1;
  const nameIdx = 2;
  const qtyIdx = 5;

  const byCode = {};
  for (let i = 1; i < dbData.length; i++) {
    const code = normalizeCode(dbData[i][codeIdx]);
    const qty = parseNumber(dbData[i][qtyIdx]);
    if (!code || qty <= 0) continue;

    const name = String(dbData[i][nameIdx] || '').trim();
    if (!byCode[code]) byCode[code] = { code: code, name: name, qty: 0 };
    byCode[code].qty += qty;
    if (!byCode[code].name && name) byCode[code].name = name;
  }

  const rows = Object.keys(byCode).sort().map(code => byCode[code]);

  const json = JSON.stringify(rows);
  cache.put('db_diff_v1', json, 45);

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function findHeaderIndex(headers, candidates, fallback) {
  for (let i = 0; i < candidates.length; i++) {
    const idx = headers.indexOf(String(candidates[i]).toLowerCase());
    if (idx !== -1) return idx;
  }
  return fallback;
}


function parseNumber(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const str = String(value || '').replace(/,/g, '').trim();
  const num = Number(str);
  return Number.isFinite(num) ? num : 0;
}

function normalizeCode(value) {
  return String(value || '')
    .trim()
    .replace(/^\uFEFF+/, '')
    .replace(/^[`'‘’＇“”]+/, '')
    .replace(/\s+/g, '')
    .toUpperCase();
}


const BULK_REQUEST_PREFIX = 'bulk_request_v1_';
const BULK_REQUEST_INDEX = 'bulk_request_index_v1';

function getStoredBulkResult(requestId) {
  if (!requestId) return null;
  const raw = PropertiesService.getScriptProperties().getProperty(BULK_REQUEST_PREFIX + requestId);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (err) { return null; }
}

function rememberBulkResult(requestId, fingerprint, response) {
  if (!requestId) return;
  const properties = PropertiesService.getScriptProperties();
  const now = Date.now();
  let index = [];
  try { index = JSON.parse(properties.getProperty(BULK_REQUEST_INDEX) || '[]'); } catch (err) { index = []; }

  index = index.filter(entry => {
    const keep = entry && entry.id !== requestId && now - Number(entry.at || 0) < 7 * 24 * 60 * 60 * 1000;
    if (!keep && entry && entry.id) properties.deleteProperty(BULK_REQUEST_PREFIX + entry.id);
    return keep;
  });
  index.push({ id: requestId, at: now });
  while (index.length > 100) {
    const removed = index.shift();
    if (removed && removed.id) properties.deleteProperty(BULK_REQUEST_PREFIX + removed.id);
  }

  properties.setProperty(BULK_REQUEST_PREFIX + requestId, JSON.stringify({ fingerprint: fingerprint, response: response }));
  properties.setProperty(BULK_REQUEST_INDEX, JSON.stringify(index));
}

function appendInventoryDiffLogs(values, zone, board, oldCode, oldQty, newCode, newQty) {
  oldCode = normalizeCode(oldCode);
  newCode = normalizeCode(newCode);
  oldQty = parseNumber(oldQty);
  newQty = parseNumber(newQty);
  const timestamp = new Date();

  if (oldCode === newCode) {
    const diff = newQty - oldQty;
    if (newCode && diff !== 0) values.push([timestamp, zone, board, newCode, diff > 0 ? 'IN' : 'OUT', Math.abs(diff)]);
    return;
  }
  if (oldCode && oldQty > 0) values.push([timestamp, zone, board, oldCode, 'OUT', oldQty]);
  if (newCode && newQty > 0) values.push([timestamp, zone, board, newCode, 'IN', newQty]);
}

function serializeBulkLogs(values) {
  return values.map(row => ({
    at: new Date(row[0]).toISOString(),
    zone: row[1],
    board: row[2],
    code: row[3],
    type: row[4],
    qty: row[5]
  }));
}

function appendMissingBulkLogs(logSheet, logs, requestId) {
  const expected = Array.isArray(logs) ? logs : [];
  if (!expected.length) return 0;
  const existingIds = new Set();
  const lastRow = logSheet.getLastRow();
  if (lastRow > 1) {
    const ids = logSheet.getRange(2, 7, lastRow - 1, 1).getValues();
    ids.forEach(row => {
      const id = String(row[0] || '').trim();
      if (id) existingIds.add(id);
    });
  }
  if (!String(logSheet.getRange(1, 7).getValue() || '').trim()) {
    logSheet.getRange(1, 7).setValue('request_id');
  }

  const values = [];
  expected.forEach((log, index) => {
    const logId = requestId + ':' + index;
    if (existingIds.has(logId)) return;
    values.push([
      new Date(log.at),
      log.zone,
      log.board,
      log.code,
      log.type,
      log.qty,
      logId
    ]);
  });
  if (values.length) {
    logSheet.getRange(logSheet.getLastRow() + 1, 1, values.length, 7).setValues(values);
  }
  return expected.length;
}

// ============================
// 일괄 저장: 서버의 현재 수량을 기준으로 슬롯과 사용 로그를 함께 계산
// ============================
function bulkSave(payload) {
  const rawItems = Array.isArray(payload && payload.items) ? payload.items : [];
  const requestId = String(payload && payload.requestId || '').trim();
  if (!/^[A-Za-z0-9._:-]{8,128}$/.test(requestId)) {
    return jsonOutput({ status: 'INVALID', message: '저장 요청 ID가 올바르지 않습니다.' });
  }

  const items = [];
  const seenKeys = {};
  for (let i = 0; i < rawItems.length; i++) {
    const validation = validateSlotItem(rawItems[i]);
    if (!validation.ok) return jsonOutput({ status: 'INVALID', message: validation.message, itemIndex: i });
    const item = validation.item;
    const key = [item.zone, item.board, item.slot_no].join('-');
    if (seenKeys[key]) return jsonOutput({ status: 'INVALID', message: '같은 슬롯이 저장 요청에 중복 포함되었습니다.' });
    seenKeys[key] = true;
    items.push(item);
  }

  const fingerprint = JSON.stringify(items);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return jsonOutput({ status: 'BUSY', message: '다른 저장 작업이 진행 중입니다. 잠시 후 다시 시도하세요.' });

  try {
    const stored = getStoredBulkResult(requestId);
    let intent = null;
    if (stored) {
      if (stored.fingerprint !== fingerprint) {
        return jsonOutput({ status: 'INVALID', message: '같은 요청 ID에 다른 저장 내용이 전달되었습니다.' });
      }
      if (stored.response && stored.response.status === 'OK') {
        const replay = Object.assign({}, stored.response, { replayed: true });
        return jsonOutput(replay);
      }
      if (stored.response && stored.response.status === 'PROCESSING' && stored.response.intent) {
        intent = stored.response.intent;
      } else {
        return jsonOutput({ status: 'ERROR', message: '이전 저장 요청의 처리 상태를 확인할 수 없습니다.' });
      }
    }

    const spreadsheet = SpreadsheetApp.getActive();
    const slotSheet = spreadsheet.getSheetByName('SLOTS');
    const logSheet = spreadsheet.getSheetByName('USAGE_LOG');
    if (!slotSheet || !logSheet) return jsonOutput({ status: 'ERROR', message: 'SLOTS 또는 USAGE_LOG 시트를 찾을 수 없습니다.' });

    if (!intent) {
      const initialData = slotSheet.getDataRange().getValues();
      const initialRows = {};
      for (let i = 1; i < initialData.length; i++) {
        initialRows[[String(initialData[i][0]), String(initialData[i][1]), String(initialData[i][2])].join('-')] = i + 1;
      }
      const logValues = [];
      let updated = 0;
      let created = 0;
      items.forEach(item => {
        const key = [item.zone, item.board, item.slot_no].join('-');
        const rowNo = initialRows[key];
        const oldRow = rowNo ? initialData[rowNo - 1] : null;
        appendInventoryDiffLogs(
          logValues,
          item.zone,
          item.board,
          oldRow ? oldRow[3] : '',
          oldRow ? oldRow[6] : 0,
          item.code,
          item.qty
        );
        if (rowNo) updated++;
        else created++;
      });
      intent = { updated: updated, created: created, logs: serializeBulkLogs(logValues) };
      rememberBulkResult(requestId, fingerprint, { status: 'PROCESSING', intent: intent });
    }

    const currentData = slotSheet.getDataRange().getValues();
    const currentRows = {};
    for (let i = 1; i < currentData.length; i++) {
      currentRows[[String(currentData[i][0]), String(currentData[i][1]), String(currentData[i][2])].join('-')] = i + 1;
    }
    items.forEach(item => {
      const key = [item.zone, item.board, item.slot_no].join('-');
      const rowNo = currentRows[key];
      if (rowNo) {
        slotSheet.getRange(rowNo, 4, 1, 4).setValues([[item.code, item.name, item.spec, item.qty]]);
      } else {
        slotSheet.appendRow([item.zone, item.board, item.slot_no, item.code, item.name, item.spec, item.qty]);
        currentRows[key] = slotSheet.getLastRow();
      }
    });

    const logged = appendMissingBulkLogs(logSheet, intent.logs, requestId);

    SpreadsheetApp.flush();
    invalidateRuntimeCaches();
    const response = {
      status: 'OK',
      updated: Number(intent.updated || 0),
      created: Number(intent.created || 0),
      logged: logged
    };
    rememberBulkResult(requestId, fingerprint, response);
    return jsonOutput(response);
  } catch (err) {
    console.error('bulkSave failed', err);
    return jsonOutput({ status: 'ERROR', message: '저장 처리 중 오류가 발생했습니다.' });
  } finally {
    lock.releaseLock();
  }
}



function logClientError(payload) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('APP_ERROR_LOG');
  if (!sheet) return ContentService.createTextOutput('NO_LOG_SHEET');

  sheet.appendRow([
    new Date(),
    String(payload && payload.area || ''),
    String(payload && payload.message || ''),
    String(payload && payload.detail || ''),
    String(payload && payload.ua || '')
  ]);

  return ContentService.createTextOutput('ERROR_LOGGED');
}

function invalidateRuntimeCaches() {
  const cache = CacheService.getScriptCache();
  cache.removeAll(['slots_v1', 'db_diff_v1', 'slot_row_index_v1']);
}
