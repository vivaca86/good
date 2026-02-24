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
// POST : updateSlot / logUsage 처리
// ============================
function doPost(e) {
  let payload;

  try {
    payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return ContentService.createTextOutput('INVALID_JSON');
  }

  if (payload.action === 'updateSlot') {
    return updateSlot(payload);
  }

  if (payload.action === 'logUsage') {
    return appendLog(payload);
  }

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

  // 기존 행 업데이트
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

  // 없으면 신규 행 생성 (upsert)
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
