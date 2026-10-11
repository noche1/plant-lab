/**
 * 식물 초능력 발명소 — 구글 시트 저장소
 *
 * 이 스크립트는 시트에 "번호, 그림, 발명품 내용"만 저장합니다.
 * 학생 이름은 저장하지 않습니다. 이름은 선생님 컴퓨터의 명렬표에만 있습니다.
 *
 * 처음 한 번:
 *  1) 아래 ADMIN_KEY 의 '여기에-비밀번호' 를 원하는 비밀번호로 바꾸고 저장
 *  2) 위쪽 함수 선택에서 setup 을 고르고 [실행] (권한 허용)
 *  3) [배포] → [새 배포] → 유형: 웹 앱 → 실행: 나 / 액세스: 모든 사용자 → 배포
 *  4) 나온 웹 앱 URL 을 앱의 선생님 메뉴에 붙여넣기
 */

const ADMIN_KEY_FOR_SETUP = '여기에-비밀번호';
const SHEET_NAME = 'records';
const HEADERS = ['id', 'time', 'no', 'a', 'b', 'problem', 'name', 'desc', 'heart'];
const MAX_NO = 40;      // 번호는 1~40까지만 받음
const MAX_ROWS = 500;   // 장난으로 시트가 터지지 않게

function setup() {
  PropertiesService.getScriptProperties().setProperty('ADMIN_KEY', ADMIN_KEY_FOR_SETUP);
  sheet_();
}

function doGet() {
  const s = sheet_();
  const values = s.getDataRange().getValues();
  const head = values.shift();
  const items = values.map(r => {
    const o = {};
    head.forEach((k, i) => { o[k] = r[i]; });
    return o;
  });
  return json_({ ok: true, items: items });
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, error: 'bad_json' }); }

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const s = sheet_();

    if (body.action === 'add') {
      if (s.getLastRow() - 1 >= MAX_ROWS) return json_({ ok: false, error: 'full' });
      const d = body.item || {};
      const no = int_(d.no), a = int_(d.a), b = int_(d.b);
      if (!(no >= 1 && no <= MAX_NO)) return json_({ ok: false, error: 'bad_no' });
      if (!(a >= 0 && a < 50 && b >= 0 && b < 50) || a === b) return json_({ ok: false, error: 'bad_plants' });
      const id = Utilities.getUuid().slice(0, 8);
      s.appendRow([id, new Date().toISOString(), no, a, b,
        text_(d.problem, 60), text_(d.name, 30), text_(d.desc, 200), text_(d.heart, 60)]);
      return json_({ ok: true, id: id });
    }

    // 아래는 모두 선생님 비밀번호가 필요한 동작
    if (!keyOk_(body.key)) return json_({ ok: false, error: 'wrong_key' });

    if (body.action === 'check') return json_({ ok: true });

    if (body.action === 'remove') {
      const ids = s.getRange(2, 1, Math.max(s.getLastRow() - 1, 1), 1).getValues();
      for (let i = ids.length - 1; i >= 0; i--) {
        if (String(ids[i][0]) === String(body.id)) { s.deleteRow(i + 2); return json_({ ok: true }); }
      }
      return json_({ ok: false, error: 'not_found' });
    }

    if (body.action === 'clear') {
      if (s.getLastRow() > 1) s.deleteRows(2, s.getLastRow() - 1);
      return json_({ ok: true });
    }

    return json_({ ok: false, error: 'unknown_action' });
  } finally {
    lock.releaseLock();
  }
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET_NAME);
  if (!s) {
    s = ss.insertSheet(SHEET_NAME);
    s.appendRow(HEADERS);
    s.setFrozenRows(1);
  }
  return s;
}

function keyOk_(k) {
  const key = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  return !!key && key !== '여기에-비밀번호' && typeof k === 'string' && k === key;
}

function int_(v) { const n = parseInt(v, 10); return isNaN(n) ? -1 : n; }

// 학생이 쓴 글을 그대로 시트에 넣으면 "=" 로 시작하는 글이 수식으로 실행될 수 있어서 막음
function text_(v, max) {
  let t = String(v == null ? '' : v).slice(0, max);
  if (/^[=+\-@]/.test(t)) t = "'" + t;
  return t;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
