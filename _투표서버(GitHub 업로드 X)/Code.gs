/**
 * @OnlyCurrentDoc  — 권한을 '이 스프레드시트 1개'로만 제한합니다.
 */

/**
 * 한화오션 현장책임자 특강 — QR 실시간 투표 백엔드 (Google Apps Script)
 * Data Thinking Lab
 *
 * 배포: 배포 > 새 배포 > 유형 "웹 앱"
 *   - 다음 사용자로 실행: 나(본인)
 *   - 액세스 권한이 있는 사용자: 모든 사용자
 *
 * 엔드포인트 (모두 GET)
 *   ?action=counts                         → 현재 집계 {ok, counts:[4], total, session}
 *   ?action=vote&choice=1~4&cid=기기ID      → 투표(같은 기기는 재투표 시 선택만 변경)
 *   ?action=reset&key=ADMIN_KEY            → 투표·아이디어 초기화(새 세션)
 *   ?action=ideas                          → 아이디어(포스트잇) 목록 {ok, ideas:[{id,text,who,likes,ts}]}
 *   ?action=idea&text=…&who=…&cid=기기ID    → 아이디어 등록
 *   ?action=like&id=아이디어ID&cid=기기ID    → 공감 1회(기기당 1회)
 *   ?action=deleteIdea&id=…&key=ADMIN_KEY  → 아이디어 삭제(강사용)
 */

// ▼ 초기화용 비밀번호. config.js 의 POLL_ADMIN_KEY 와 동일하게 맞춰 주세요.
const ADMIN_KEY = 'hanwha-2026';

const OPTION_COUNT = 4;
const SHEET_NAME = '투표기록';
const IDEA_SHEET = '아이디어';
const IDEA_MAX_TEXT = 160;
const IDEA_MAX_WHO = 20;

function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = String(p.action || 'counts');
  try {
    if (action === 'vote') return json_(vote_(p.choice, p.cid));
    if (action === 'reset') return json_(reset_(p.key));
    if (action === 'ideas') return json_(listIdeas_());
    if (action === 'idea') return json_(addIdea_(p.text, p.who, p.cid));
    if (action === 'like') return json_(likeIdea_(p.id, p.cid));
    if (action === 'deleteIdea') return json_(deleteIdea_(p.id, p.key));
    return json_(readCounts_());
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function vote_(choiceRaw, cidRaw) {
  const choice = parseInt(choiceRaw, 10);
  const cid = String(cidRaw || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 40);
  if (!(choice >= 1 && choice <= OPTION_COUNT)) return { ok: false, error: 'invalid choice' };
  if (cid.length < 8) return { ok: false, error: 'invalid client id' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties();
    const state = getState_(props);
    const key = 'v_' + state.session + '_' + cid;
    const prev = parseInt(props.getProperty(key) || '0', 10);
    if (prev === choice) return Object.assign({ ok: true, changed: false, choice }, publicState_(state));

    if (prev >= 1) state.counts[prev - 1] = Math.max(0, state.counts[prev - 1] - 1);
    state.counts[choice - 1] += 1;
    props.setProperties({ [key]: String(choice), STATE: JSON.stringify(state) });
    log_(state.session, cid, choice, prev);
    CacheService.getScriptCache().put('STATE', JSON.stringify(state), 60);
    return Object.assign({ ok: true, changed: true, choice }, publicState_(state));
  } finally {
    lock.releaseLock();
  }
}

function reset_(key) {
  if (String(key || '') !== ADMIN_KEY) return { ok: false, error: 'unauthorized' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties();
    const old = getState_(props);
    // 이전 세션의 기기별 투표 키 정리
    Object.keys(props.getProperties())
      .filter((k) => k.indexOf('v_' + old.session + '_') === 0)
      .forEach((k) => props.deleteProperty(k));
    const state = { session: old.session + 1, counts: new Array(OPTION_COUNT).fill(0) };
    props.setProperty('STATE', JSON.stringify(state));
    CacheService.getScriptCache().put('STATE', JSON.stringify(state), 60);
    CacheService.getScriptCache().remove('IDEAS_' + old.session);
    log_(state.session, 'RESET', 0, 0);
    return Object.assign({ ok: true }, publicState_(state));
  } finally {
    lock.releaseLock();
  }
}

function readCounts_() {
  const cached = CacheService.getScriptCache().get('STATE');
  const state = cached ? JSON.parse(cached) : getState_(PropertiesService.getScriptProperties());
  return Object.assign({ ok: true }, publicState_(state));
}

function getState_(props) {
  const raw = props.getProperty('STATE');
  if (raw) return JSON.parse(raw);
  return { session: 1, counts: new Array(OPTION_COUNT).fill(0) };
}

function publicState_(state) {
  const total = state.counts.reduce((a, b) => a + b, 0);
  return { counts: state.counts, total, session: state.session };
}

function log_(session, cid, choice, prev) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return;
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) {
      sh = ss.insertSheet(SHEET_NAME);
      sh.appendRow(['시각', '세션', '기기ID', '선택', '이전 선택']);
    }
    sh.appendRow([new Date(), session, cid, choice, prev || '']);
  } catch (err) {
    // 기록 실패는 집계에 영향 없음
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}


/* ===== 아이디어(포스트잇) 보드 ===== */

function ideaSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(IDEA_SHEET);
  if (!sh) {
    sh = ss.insertSheet(IDEA_SHEET);
    sh.appendRow(['ID', '시각', '세션', '기기ID', '내용', '작성자', '공감', '삭제']);
  }
  return sh;
}

function currentSession_() {
  return getState_(PropertiesService.getScriptProperties()).session;
}

function cleanText_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001f<>]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, max);
}

function listIdeas_() {
  const session = currentSession_();
  const cache = CacheService.getScriptCache();
  const cached = cache.get('IDEAS_' + session);
  if (cached) return { ok: true, session, ideas: JSON.parse(cached) };

  const sh = ideaSheet_();
  const last = sh.getLastRow();
  const ideas = [];
  if (last > 1) {
    const rows = sh.getRange(2, 1, last - 1, 8).getValues();
    rows.forEach((r) => {
      if (Number(r[2]) !== session || r[7] === true || r[7] === 'Y') return;
      ideas.push({ id: String(r[0]), text: String(r[4]), who: String(r[5] || ''), likes: Number(r[6]) || 0, ts: Number(r[1] && r[1].getTime ? r[1].getTime() : 0) });
    });
  }
  ideas.sort((a, b) => b.ts - a.ts);
  cache.put('IDEAS_' + session, JSON.stringify(ideas), 4);
  return { ok: true, session, ideas };
}

function addIdea_(textRaw, whoRaw, cidRaw) {
  const text = cleanText_(textRaw, IDEA_MAX_TEXT);
  const who = cleanText_(whoRaw, IDEA_MAX_WHO);
  const cid = String(cidRaw || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 40);
  if (text.length < 2) return { ok: false, error: 'empty text' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const session = currentSession_();
    const id = 'i' + Date.now().toString(36) + Math.floor(Math.random() * 1000).toString(36);
    ideaSheet_().appendRow([id, new Date(), session, cid || 'presenter', text, who, 0, '']);
    CacheService.getScriptCache().remove('IDEAS_' + session);
    return Object.assign({ ok: true, id }, listIdeas_());
  } finally {
    lock.releaseLock();
  }
}

function likeIdea_(idRaw, cidRaw) {
  const id = String(idRaw || '').replace(/[^A-Za-z0-9]/g, '').slice(0, 30);
  const cid = String(cidRaw || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 40) || 'presenter';
  if (!id) return { ok: false, error: 'invalid id' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const props = PropertiesService.getScriptProperties();
    const session = currentSession_();
    const key = 'l_' + session + '_' + id + '_' + cid;
    if (props.getProperty(key)) return Object.assign({ ok: true, already: true }, listIdeas_());

    const sh = ideaSheet_();
    const last = sh.getLastRow();
    if (last > 1) {
      const ids = sh.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === id) {
          const cell = sh.getRange(i + 2, 7);
          cell.setValue((Number(cell.getValue()) || 0) + 1);
          props.setProperty(key, '1');
          CacheService.getScriptCache().remove('IDEAS_' + session);
          break;
        }
      }
    }
    return Object.assign({ ok: true }, listIdeas_());
  } finally {
    lock.releaseLock();
  }
}

function deleteIdea_(idRaw, key) {
  if (String(key || '') !== ADMIN_KEY) return { ok: false, error: 'unauthorized' };
  const id = String(idRaw || '').replace(/[^A-Za-z0-9]/g, '').slice(0, 30);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = ideaSheet_();
    const last = sh.getLastRow();
    if (last > 1) {
      const ids = sh.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === id) { sh.getRange(i + 2, 8).setValue('Y'); break; }
      }
    }
    CacheService.getScriptCache().remove('IDEAS_' + currentSession_());
    return Object.assign({ ok: true }, listIdeas_());
  } finally {
    lock.releaseLock();
  }
}
