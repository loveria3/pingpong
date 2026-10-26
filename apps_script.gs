/* =====================================================================
   오토카 바다 미션 · 결과 받는 구글 Apps Script
   구글 시트 → 확장 프로그램 → Apps Script 에 이 내용을 통째로 붙여 넣고
   '배포 → 새 배포 → 웹 앱 (액세스: 모든 사용자)' 로 배포하세요.
   ===================================================================== */

const SHEET_NAME = '제출';
const MAX_NO = 40;

/* 학생 화면이 '보내기'를 누르면 여기로 들어와요: 시트에 한 줄 추가 */
function doPost(e) {
  const d = JSON.parse(e.postData.contents);

  // 선생님 화면의 '기록 지우기'
  if (d.action === 'clear') {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try { clearAll(); } finally { lock.releaseLock(); }
    return json({ ok: true });
  }

  const no = Number(d.no);
  if (!(no >= 1 && no <= MAX_NO)) return json({ ok: false });
  const path = String(d.path || '').replace(/[^FBLR]/g, '').slice(0, 300);

  const lock = LockService.getScriptLock();   // 여러 명이 동시에 보내도 섞이지 않게
  lock.waitLock(10000);
  try {
    sheet().appendRow([new Date(), no, path, d.reached ? '도착' : '미도착']);
  } finally {
    lock.releaseLock();
  }
  return json({ ok: true });
}

/* 선생님 화면이 몇 초마다 읽어 가요: 번호마다 가장 마지막에 보낸 기록만 */
function doGet() {
  const rows = sheet().getDataRange().getValues().slice(1);
  const latest = {};
  rows.forEach(r => {
    if (r[1]) latest[r[1]] = { no: r[1], path: r[2], reached: r[3] === '도착', time: r[0] };
  });
  return json({ ok: true, list: Object.values(latest) });
}

/* 시트 위쪽 메뉴에 '수업 끝나면 기록 지우기' 추가 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('오토카 미션')
    .addItem('제출 기록 모두 지우기', 'clearAll')
    .addToUi();
}
function clearAll() {
  const s = sheet();
  if (s.getLastRow() > 1) s.deleteRows(2, s.getLastRow() - 1);
}

function sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET_NAME);
  if (!s) {
    s = ss.insertSheet(SHEET_NAME);
    s.appendRow(['보낸 시각', '번호', '블록 순서 (F앞 B뒤 L왼 R오)', '도착']);
  }
  return s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
