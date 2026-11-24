/**
 * AI 디지털 기반 교실 특강 — 실시간 응답·투표 백엔드 (Google Apps Script)
 *
 * 설치 방법
 * 1. 새 구글 스프레드시트 → 확장 프로그램 → Apps Script
 * 2. 이 코드를 Code.gs에 붙여넣고, 아래 ADMIN_KEY를 원하는 암호로 바꾼 뒤 저장
 * 3. 배포 → 새 배포 → 유형: 웹 앱
 *    - 실행 사용자: 나
 *    - 액세스 권한: 모든 사용자
 * 4. 발급된 웹 앱 URL(…/exec)을 index.html 상단 CONFIG.SCRIPT_URL에 붙여넣기
 *    (코드를 고친 뒤에는 '배포 관리 → 수정 → 새 버전'으로 다시 배포)
 *
 * 시트 구조 (자동 생성)
 *  - sessions  | id | label | created |      ← 강의 회차. 마지막 행이 '진행 중' 강의
 *  - responses | timestamp | session | board | name | text |
 *      board: opening(여는 질문) / poll(투표, text=선택지 id) / etc(투표 기타 응답) / closing(마무리)
 *
 * 리셋 = 새 회차 추가. 이전 회차의 응답은 지워지지 않고 '이전 강의 결과 보기'에서 확인할 수 있습니다.
 */

const ADMIN_KEY = 'change-me';   // ← 리셋용 관리자 암호로 바꾸세요
const MAX_TEXT = 300;
const MAX_NAME = 20;
const BOARDS = ['opening', 'poll', 'etc', 'closing'];
const POLL_IDS = ['question', 'verify', 'ethics', 'self', 'collab', 'create', 'etc'];

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name, header) {
  let sh = ss_().getSheetByName(name);
  if (!sh) {
    sh = ss_().insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
  }
  return sh;
}
const respSheet_ = () => sheet_('responses', ['timestamp', 'session', 'board', 'name', 'text']);
const sessSheet_ = () => sheet_('sessions', ['id', 'label', 'created']);

function sessions_() {
  const sh = sessSheet_();
  const last = sh.getLastRow();
  if (last < 2) {
    sh.appendRow(['s1', '1회차', new Date()]);
    return [{ id: 's1', label: '1회차', created: new Date().toISOString() }];
  }
  return sh.getRange(2, 1, last - 1, 3).getValues().map(function (r) {
    return { id: String(r[0]), label: String(r[1]), created: r[2] instanceof Date ? r[2].toISOString() : String(r[2]) };
  });
}
function current_() { const s = sessions_(); return s[s.length - 1]; }

function rows_() {
  const sh = respSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 5).getValues();
}
function group_(rows, sid) {
  const b = {};
  BOARDS.forEach(function (k) { b[k] = []; });
  rows.forEach(function (r) {
    if (String(r[1]) !== sid) return;
    const board = String(r[2]);
    if (!b[board]) return;
    b[board].push({
      t: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      name: String(r[3] || ''),
      text: String(r[4] || '')
    });
  });
  return b;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * GET ?action=list    → { ok, session, boards }            (진행 중 강의)
 * GET ?action=history → { ok, current, sessions:[{id,label,created,boards}] }
 */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'list';
  const rows = rows_();
  if (action === 'history') {
    const list = sessions_();
    return json_({
      ok: true,
      current: list[list.length - 1].id,
      sessions: list.map(function (s) {
        return { id: s.id, label: s.label, created: s.created, boards: group_(rows, s.id) };
      })
    });
  }
  const cur = current_();
  return json_({ ok: true, session: { id: cur.id, label: cur.label }, boards: group_(rows, cur.id) });
}

/**
 * POST { action:'add', board, name, text }
 * POST { action:'reset', key, label }
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    const action = data.action || 'add';
    const lock = LockService.getScriptLock();
    lock.waitLock(5000);
    try {
      if (action === 'reset') {
        if (String(data.key || '') !== ADMIN_KEY) return json_({ ok: false, error: '관리자 암호가 맞지 않습니다' });
        const n = sessions_().length + 1;
        const id = 's' + n + '_' + Date.now().toString(36);
        const label = String(data.label || '').trim().slice(0, 60) || (n + '회차');
        sessSheet_().appendRow([id, label, new Date()]);
        return json_({ ok: true, session: { id: id, label: label } });
      }

      const board = String(data.board || '');
      if (BOARDS.indexOf(board) < 0) return json_({ ok: false, error: 'bad board' });
      let text = String(data.text || '').trim().slice(0, MAX_TEXT);
      let name = String(data.name || '').trim().slice(0, MAX_NAME);
      if (!text) return json_({ ok: false, error: 'empty' });
      if (board === 'poll' && POLL_IDS.indexOf(text) < 0) return json_({ ok: false, error: 'bad option' });
      // 수식 주입 방지
      if (/^[=+\-@]/.test(text)) text = "'" + text;
      if (/^[=+\-@]/.test(name)) name = "'" + name;

      respSheet_().appendRow([new Date(), current_().id, board, name, text]);
      return json_({ ok: true });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}
