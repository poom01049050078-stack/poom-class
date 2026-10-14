/**
 * 품태권도 수업계획 시트 ↔ 홈페이지 연결
 * 1) 아래 PIN을 원하는 숫자로 바꾸세요 (사범님과 공유할 수정 비밀번호)
 * 2) 배포 → 새 배포 → 웹 앱 / 실행: 나 / 액세스: 모든 사용자
 */
const PIN = '0000';

function sheet_() { return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet() {
  return out_({ ok: true, rows: sheet_().getDataRange().getDisplayValues() });
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad' }); }
  if (String(b.pin) !== PIN) return out_({ ok: false, error: 'pin' });
  if (b.check) return out_({ ok: true });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const v = sh.getDataRange().getDisplayValues();
    const head = v[0].map(String);
    const col = name => {
      let i = head.indexOf(name);
      if (i < 0) { head.push(name); i = head.length - 1; sh.getRange(1, i + 1).setValue(name); }
      return i + 1;
    };
    const dayCol = head.indexOf('일');
    const r = v.findIndex((row, i) => i > 0 && Number(row[dayCol]) === Number(b.day));
    if (r < 0) return out_({ ok: false, error: 'day' });
    const row = r + 1;
    const put = (c, val) => sh.getRange(row, c).setNumberFormat('@').setValue(val == null ? '' : String(val));

    if (Array.isArray(b.items)) {
      const n = Math.max(4, b.items.length);
      for (let k = 1; k <= n; k++) put(col('종목' + k), b.items[k - 1] || '');
    }
    put(col('전달사항'), b.notice || '');
    put(col('추가설정'), b.extra ? JSON.stringify(b.extra) : '');
    return out_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
