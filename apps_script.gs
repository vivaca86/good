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

  const sheet = SpreadsheetApp.getActive().getSheetByName('SLOTS');
  const data = sheet.getDataRange().getValues();

  if (!data.length) {
    return ContentService
      .createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

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
  cache.put('slots_v1', json, 20);

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
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
  if (payload.action === 'deleteBoard') return deleteBoardRows(payload);
  if (payload.action === 'getDbDiff') return getDbDiffRows(payload);
  if (payload.action === 'bulkSave') return bulkSave(payload);

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

  cache.put('slot_row_index_v1', JSON.stringify(rowByKey), 20);
  return rowByKey;
}

// ============================
// SLOTS 업데이트 (없으면 생성)
// ============================
function updateSlot(payload) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('SLOTS');

  const zone = Number(payload.zone) || 0;
  const board = Number(payload.board) || 0;
  const slot = Number(payload.slot_no) || 0;
  if (!zone || !board || !slot) return ContentService.createTextOutput('INVALID_SLOT');

  const key = [String(zone), String(board), String(slot)].join('-');
  const rowByKey = getSlotRowIndexMap(sheet);
  const rowNo = rowByKey[key];

  if (rowNo) {
    sheet.getRange(rowNo, 4, 1, 4).setValues([[
      payload.code || '',
      payload.name || '',
      payload.spec || '',
      Number(payload.qty) || 0
    ]]);
    invalidateRuntimeCaches();
    return ContentService.createTextOutput('UPDATED');
  }

  sheet.appendRow([
    zone,
    board,
    slot,
    payload.code || '',
    payload.name || '',
    payload.spec || '',
    Number(payload.qty) || 0
  ]);

  invalidateRuntimeCaches();
  return ContentService.createTextOutput('CREATED');
}

// ============================
// USAGE_LOG 추가
// ============================
function appendLog(payload) {
  const logSheet = SpreadsheetApp.getActive().getSheetByName('USAGE_LOG');

  logSheet.appendRow([
    new Date(),
    payload.zone,
    payload.board,
    payload.code,
    payload.type,
    Number(payload.qty) || 0
  ]);

  invalidateRuntimeCaches();
  return ContentService.createTextOutput('LOGGED');
}

// ============================
// 현황판 생성 (12 슬롯 기본행 생성)
// ============================
function createBoardRows(payload) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('SLOTS');
  const data = sheet.getDataRange().getValues();
  const zone = Number(payload.zone) || 0;
  const board = Number(payload.board) || 0;

  if (!zone || !board) return ContentService.createTextOutput('INVALID_BOARD');

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
}

// ============================
// 현황판 삭제 + 뒤 현황판 번호 당김
// ============================
function deleteBoardRows(payload) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('SLOTS');
  const data = sheet.getDataRange().getValues();
  const zone = Number(payload.zone) || 0;
  const board = Number(payload.board) || 0;

  if (!zone || !board) return ContentService.createTextOutput('INVALID_BOARD');
  if (data.length <= 1) return ContentService.createTextOutput('BOARD_DELETED');

  const kept = [data[0]];
  for (let i = 1; i < data.length; i++) {
    const row = data[i].slice();
    const rowZone = Number(row[0]);
    const rowBoard = Number(row[1]);

    if (rowZone !== zone) {
      kept.push(row);
      continue;
    }

    if (rowBoard === board) continue;
    if (rowBoard > board) row[1] = rowBoard - 1;
    kept.push(row);
  }

  sheet.clearContents();
  sheet.getRange(1, 1, kept.length, kept[0].length).setValues(kept);

  invalidateRuntimeCaches();
  return ContentService.createTextOutput('BOARD_DELETED');
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

  const rows = [];
  for (let i = 1; i < dbData.length; i++) {
    const code = normalizeCode(dbData[i][codeIdx]);
    const qty = parseNumber(dbData[i][qtyIdx]);
    if (!code || qty <= 0) continue;

    const name = String(dbData[i][nameIdx] || '').trim();

    rows.push({
      code: code,
      name: name,
      qty: qty
    });
  }

  const json = JSON.stringify(rows);
  cache.put('db_diff_v1', json, 30);

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
  return String(value || '').trim();
}


// ============================
// 일괄 저장: 슬롯 업데이트 + 로그 일괄 추가
// ============================
function bulkSave(payload) {
  const items = Array.isArray(payload && payload.items) ? payload.items : [];
  const logs = Array.isArray(payload && payload.logs) ? payload.logs : [];

  const slotSheet = SpreadsheetApp.getActive().getSheetByName('SLOTS');
  const rowByKey = getSlotRowIndexMap(slotSheet);

  let updated = 0;
  let created = 0;

  for (let i = 0; i < items.length; i++) {
    const item = items[i] || {};
    const zone = Number(item.zone) || 0;
    const board = Number(item.board) || 0;
    const slotNo = Number(item.slot_no) || 0;
    if (!zone || !board || !slotNo) continue;

    const key = [String(zone), String(board), String(slotNo)].join('-');
    const rowNo = rowByKey[key];

    if (rowNo) {
      slotSheet.getRange(rowNo, 4, 1, 4).setValues([[
        item.code || '',
        item.name || '',
        item.spec || '',
        Number(item.qty) || 0
      ]]);
      updated++;
    } else {
      slotSheet.appendRow([
        zone,
        board,
        slotNo,
        item.code || '',
        item.name || '',
        item.spec || '',
        Number(item.qty) || 0
      ]);
      created++;
    }
  }

  let logged = 0;
  if (logs.length > 0) {
    const logSheet = SpreadsheetApp.getActive().getSheetByName('USAGE_LOG');
    const values = [];
    for (let i = 0; i < logs.length; i++) {
      const log = logs[i] || {};
      const code = normalizeCode(log.code);
      const qty = parseNumber(log.qty);
      const type = String(log.type || '').toUpperCase();
      if (!code || qty <= 0 || (type !== 'IN' && type !== 'OUT')) continue;
      values.push([
        new Date(),
        Number(log.zone) || '',
        Number(log.board) || '',
        code,
        type,
        qty
      ]);
    }

    if (values.length > 0) {
      const start = logSheet.getLastRow() + 1;
      logSheet.getRange(start, 1, values.length, 6).setValues(values);
      logged = values.length;
    }
  }

  invalidateRuntimeCaches();
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'OK', updated: updated, created: created, logged: logged }))
    .setMimeType(ContentService.MimeType.JSON);
}


function invalidateRuntimeCaches() {
  const cache = CacheService.getScriptCache();
  cache.removeAll(['slots_v1', 'db_diff_v1', 'slot_row_index_v1']);
}
