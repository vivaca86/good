// ============================
// GET : SLOTS 전체 데이터 반환
// ============================
function doGet() {
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

  return ContentService
    .createTextOutput(JSON.stringify(result))
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
  if (payload.action === 'getDbDiff') return getDbDiffRows();

  return ContentService.createTextOutput('NO_ACTION');
}

// ============================
// SLOTS 업데이트 (없으면 생성)
// ============================
function updateSlot(payload) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('SLOTS');
  const data = sheet.getDataRange().getValues();

  const zone = String(payload.zone);
  const board = String(payload.board);
  const slot = String(payload.slot_no);

  for (let i = 1; i < data.length; i++) {
    if (
      String(data[i][0]) === zone &&
      String(data[i][1]) === board &&
      String(data[i][2]) === slot
    ) {
      sheet.getRange(i + 1, 4).setValue(payload.code || '');
      sheet.getRange(i + 1, 5).setValue(payload.name || '');
      sheet.getRange(i + 1, 6).setValue(payload.spec || '');
      sheet.getRange(i + 1, 7).setValue(Number(payload.qty) || 0);
      return ContentService.createTextOutput('UPDATED');
    }
  }

  sheet.appendRow([
    Number(payload.zone) || 0,
    Number(payload.board) || 0,
    Number(payload.slot_no) || 0,
    payload.code || '',
    payload.name || '',
    payload.spec || '',
    Number(payload.qty) || 0
  ]);

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

  return ContentService.createTextOutput('BOARD_DELETED');
}


// ============================
// DB구역용: COMPARE diff>0 + DC 품명 조회
// ============================
function getDbDiffRows() {
  const compareSheet = SpreadsheetApp.getActive().getSheetByName('COMPARE');
  const dcSheet = SpreadsheetApp.getActive().getSheetByName('DC');

  if (!compareSheet) {
    return ContentService
      .createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const compareData = compareSheet.getDataRange().getValues();
  if (!compareData.length) {
    return ContentService
      .createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const compareHeaders = compareData[0].map(h => String(h).trim().toLowerCase());

  const codeIdx = findHeaderIndex(compareHeaders, ['code', '코드'], 0);
  const qtyIdx = findHeaderIndex(compareHeaders, ['e', 'result', '결과', '재고', 'qty', '수량'], 4);

  const nameByCode = {};
  if (dcSheet) {
    const dcData = dcSheet.getDataRange().getValues();
    for (let i = 1; i < dcData.length; i++) {
      const code = normalizeCode(dcData[i][0]);
      const name = String(dcData[i][2] || '').trim();
      if (code) nameByCode[code] = name;
    }
  }

  const rows = [];
  for (let i = 1; i < compareData.length; i++) {
    const code = normalizeCode(compareData[i][codeIdx]);
    const qty = parseNumber(compareData[i][qtyIdx]);
    if (!code || qty <= 0) continue;

    const compareName = String(compareData[i][2] || '').trim();

    rows.push({
      code: code,
      name: compareName || nameByCode[code] || '',
      qty: qty
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify(rows))
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
