/**
 * K-Church 911 · 123 기도 (사역자 기도 언약) — Google Apps Script 백엔드
 *
 * 사역자들이 kchurch911.com/123prayer/ 에서 보낸 기도 제목을 Google Sheet에 모으고,
 * 김성수(코디네이터)가 kchurch911.com/123prayer/pray.html 에서 목록을 보며 "기도함"을 기록합니다.
 *
 * 설정 (한 번만):
 * 1. sheets.google.com → 새 스프레드시트 → 이름 "123 사역자 기도"
 * 2. 확장 프로그램 → Apps Script → 이 코드 전체 붙여넣기 → 아래 PRAY_KEY 를 원하는 비밀번호로 변경 → Ctrl+S
 * 3. 배포 → 새 배포 → 웹 앱 / 실행: 나 / 액세스: 모든 사용자 → 배포 → 권한 허용
 * 4. 웹 앱 URL(…/exec) 복사 → 123prayer/index.html 과 pray.html 의 API_URL 에 붙여넣기
 * 5. 코드 수정 시: 배포 → 배포 관리 → 연필 → 버전: 새 버전 → 배포 (URL 동일)
 *
 * API
 *   POST {action:'submit', ...form}                 → 새 기도 제목 저장
 *   POST {action:'prayed', key, id, date}           → 기도 기록 추가 (김성수)
 *   POST {action:'status', key, id, status}         → 상태 변경 (김성수): 기도중 | 응답됨 | 부분응답 | 종료
 *   GET  ?list=1&key=PRAY_KEY                        → 전체 목록 JSON (김성수)
 */

const SHEET_NAME  = '123 사역자 기도';
const PRAY_KEY    = 'kingdom123';          // ← pray.html 에서 입력하는 비밀번호 (꼭 바꾸세요)
const NOTIFY_EMAIL = '';                   // 새 기도 제목이 오면 알림 받을 이메일 (비우면 알림 없음)
const SPREADSHEET_ID = '13RaexIVdty9uZL91nOBDY4IcvIR5PQcWQ8hfJdznwXA'; // 123 사역자 기도 (pastor@ijiguchon.org 소유, gmc.hc300 편집자)

const HEADERS = ['번호','접수일시','이름','소속','연락처','기도 제목 1','기도 제목 2','긴급','123 언약','응답 소식','공개 범위','언어','상태','기도 기록'];

function doPost(e) {
  try {
    const data = JSON.parse((e.postData && e.postData.contents) || '{}');
    const sheet = getSheet_();
    const action = data.action || 'submit';

    if (action === 'submit') {
      const seq = Math.max(sheet.getLastRow() - 1, 0) + 1;
      const now = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyy-MM-dd HH:mm');
      sheet.appendRow([
        seq, now,
        data.name || '', data.group || '', data.contact || '',
        data.req1 || '', data.req2 || '',
        data.urgent ? 'Y' : '', data.covenant ? 'Y' : '', data.notify ? 'Y' : '',
        data.share === 'class' ? '동료 공유' : '김성수만',
        data.lang === 'en' ? 'EN' : 'KO',
        '기도중', '[]'
      ]);
      styleRow_(sheet, sheet.getLastRow(), !!data.urgent);
      if (NOTIFY_EMAIL) notify_(data, now);
      return json_({ ok: true, id: seq });
    }

    if (!data.key || data.key !== PRAY_KEY) return json_({ ok: false, error: 'unauthorized' });

    const row = findRow_(sheet, data.id);
    if (!row) return json_({ ok: false, error: 'not found' });

    if (action === 'prayed') {
      const cell = sheet.getRange(row, 14);
      let log = [];
      try { log = JSON.parse(cell.getValue() || '[]'); } catch (err) { log = []; }
      const d = data.date || Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyy-MM-dd');
      if (!log.includes(d)) log.push(d);
      cell.setValue(JSON.stringify(log));
      return json_({ ok: true, log });
    }
    if (action === 'status') {
      sheet.getRange(row, 13).setValue(data.status || '기도중');
      return json_({ ok: true });
    }
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!p.list) return ContentService.createTextOutput('123 Prayer API is live');
  if (p.key !== PRAY_KEY) return json_({ ok: false, error: 'unauthorized' });
  const sheet = getSheet_();
  const last = sheet.getLastRow();
  if (last < 2) return json_({ ok: true, items: [] });
  const rows = sheet.getRange(2, 1, last - 1, HEADERS.length).getValues();
  const items = rows.filter(r => r[0] !== '').map(r => {
    let log = []; try { log = JSON.parse(r[13] || '[]'); } catch (err) {}
    return { id: r[0], at: String(r[1]), name: r[2], group: r[3], contact: r[4], req1: r[5], req2: r[6],
             urgent: r[7] === 'Y', covenant: r[8] === 'Y', notify: r[9] === 'Y', share: r[10], lang: r[11],
             status: r[12] || '기도중', log };
  });
  return json_({ ok: true, items });
}

// ── helpers ────────────────────────────────────────────────────────
function getSheet_() {
  // 시트에 연결된 스크립트면 그 시트를, 독립 스크립트면 "123 사역자 기도" 스프레드시트를 자동 생성해 사용
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    const props = PropertiesService.getScriptProperties();
    const id = props.getProperty('SSID');
    if (id) { try { ss = SpreadsheetApp.openById(id); } catch (err) { ss = null; } }
    if (!ss) { ss = SpreadsheetApp.create(SHEET_NAME); props.setProperty('SSID', ss.getId()); }
  }
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.getRange(1, 1, 1, HEADERS.length).setBackground('#1a1a2e').setFontColor('#f0c040').setFontWeight('bold');
    [50, 130, 90, 150, 160, 300, 300, 50, 70, 70, 80, 50, 80, 220].forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.setFrozenRows(1);
  }
  return sh;
}
function findRow_(sheet, id) {
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return null;
}
function styleRow_(sheet, row, urgent) {
  const rg = sheet.getRange(row, 1, 1, HEADERS.length);
  rg.setBackground(urgent ? '#fff1f2' : (row % 2 ? '#ffffff' : '#f8f7fb'));
  sheet.getRange(row, 6, 1, 2).setWrap(true);
}
function notify_(d, now) {
  try {
    MailApp.sendEmail({ to: NOTIFY_EMAIL, subject: '[123 기도] 새 기도 제목 — ' + (d.name || ''),
      body: ['접수: ' + now, '이름: ' + (d.name || ''), '소속: ' + (d.group || ''), '', '기도 제목 1: ' + (d.req1 || ''), d.req2 ? '기도 제목 2: ' + d.req2 : '', d.urgent ? '※ 긴급' : ''].filter(Boolean).join('\n') });
  } catch (err) { Logger.log('mail failed: ' + err); }
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function test_() { Logger.log(doGet({ parameter: { list: '1', key: PRAY_KEY } }).getContent()); }
// 에디터에서 실행: 시트 생성·권한 확인
function setupCheck() { const sh = getSheet_(); Logger.log('Sheet URL: ' + sh.getParent().getUrl()); return sh.getParent().getUrl(); }
