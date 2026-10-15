/** 품태권도 수업계획 시트 ↔ 홈페이지 연결 (초등부: 첫 번째 탭 / 유치부: '유치부' 탭) */
const PIN = '0000';   // ← 수정용 비밀번호

const HEAD = ['일','요일','구분','종목1','종목2','종목3','종목4','전달사항','추가설정'];
const YU_SEED = [
  [1,'목','수업','명상','매트운동:에어매트 점프운동'],[2,'금','수업','명상','제미타프로그램:재미타 발차기'],
  [3,'토','휴관 · 개천절'],[5,'월','휴관 · 대체공휴일'],
  [6,'화','수업','준비운동:시청각준비운동','유아체육:기초 점프운동','커리큘럼:태극 3장'],
  [7,'수','수업','준비운동:바른체형교정','유아체육:징검다리 훈련','커리큘럼:기본막기'],
  [8,'목','수업','줄넘기:원 줄넘기','음악프로그램:태권체조'],[9,'금','휴관 · 한글날'],[10,'토','행사 · 주말'],
  [12,'월','수업','인성교육','유아체육:지그재그 달리기','커리큘럼:기본발차기 앞차기'],
  [13,'화','수업','준비운동:바른체형교정','유아체육:왕복 점프운동','커리큘럼:태극 3장'],
  [14,'수','수업','준비운동:키성장스트레칭','유아체육:중심잡기 운동','커리큘럼:기본서기'],
  [15,'목','수업','줄넘기:시계 돌리기','음악프로그램:태권체조'],
  [16,'금','수업','유아체육:종합장애물','발표력:태권도를 배우는 목적'],[17,'토','행사 · 1박2일 합숙'],
  [19,'월','수업','인성교육','유아체육:협동심 달리기','커리큘럼:기본발차기 돌려차기'],
  [20,'화','수업','준비운동:키성장스트레칭','유아체육:미션 임파서블','커리큘럼:기본동작'],
  [21,'수','수업','준비운동:폼롤러스트레칭','유아체육:점프 달리기','커리큘럼:태극 3장'],
  [22,'목','수업','명상','매트운동:에어매트 매달리기'],[23,'금','수업','명상','제미타프로그램:격파'],[24,'토','행사 · 오이도 축제 시범공연'],
  [26,'월','수업','인성교육','유아체육:미션 달리기','커리큘럼:기본발차기 내려차기'],
  [27,'화','수업','준비운동:폼롤러스트레칭','유아체육:균형감각 훈련','커리큘럼:손날격파'],
  [28,'수','수업','준비운동:시청각준비운동','유아체육:단계별 점프운동','커리큘럼:태극 3장'],
  [29,'목','수업','명상','레크레이션:공던지기'],[30,'금','수업','유아체육:종합장애물','발표력:태권도를 배우는 목적'],
  [31,'토','행사 · 아트팩토리 원데이클래스']
];

function sheet_(tab) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!tab) return ss.getSheets()[0];
  let sh = ss.getSheetByName(tab);
  if (!sh) {
    sh = ss.insertSheet(tab);
    const rows = [HEAD].concat((tab === '유치부' ? YU_SEED : []).map(r => HEAD.map((_, i) => r[i] == null ? '' : String(r[i]))));
    sh.getRange(1, 1, rows.length, HEAD.length).setNumberFormat('@').setValues(rows);
    sh.setFrozenRows(1);
  }
  return sh;
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet(e) {
  const tab = (e && e.parameter && e.parameter.tab) || '';
  return out_({ ok: true, tab: tab, rows: sheet_(tab).getDataRange().getDisplayValues() });
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad' }); }
  if (String(b.pin) !== PIN) return out_({ ok: false, error: 'pin' });
  if (b.check) return out_({ ok: true });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_(b.tab || '');
    let v = sh.getDataRange().getDisplayValues();
    const head = v[0].map(String);
    const col = name => {
      let i = head.indexOf(name);
      if (i < 0) { head.push(name); i = head.length - 1; sh.getRange(1, i + 1).setValue(name); }
      return i + 1;
    };
    const dayCol = head.indexOf('일');
    let r = v.findIndex((row, i) => i > 0 && Number(row[dayCol]) === Number(b.day));
    if (r < 0) { sh.appendRow([String(b.day)]); r = sh.getLastRow() - 1; }
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
