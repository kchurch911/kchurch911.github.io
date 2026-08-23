/**
 * K-Church 911 · 123 Prayer (기도 언약) — Google Apps Script 백엔드  v5
 *
 * 김성수 개인의 기도 언약 사역: 성도·외부 사역자 모두 kchurch911.com/123prayer/ 한 곳으로.
 *   - 시트 1 "123 사역자 기도": 기도 제목 (제출 건마다 1행)
 *   - 시트 2 "사역자 명단":     사람마다 1행 (토큰·등록일·갱신일·상태·응답 수·메모)
 *   - 김목사 기도제목: 스크립트 속성(PASTOR_REQS)에 JSON 저장, 기도방에서 수정
 *
 * 엔드포인트
 *   POST {action:'submit', ...form}                       → 기도 제목 저장 + 명단 upsert + 확인 메일(이메일일 때) → {ok,id,token}
 *   GET  ?me=TOKEN                                        → 내 정보·내 기도 제목·김목사 기도제목 (개인 페이지용)
 *   POST {action:'answered', token, id, note}            → 본인이 "응답 받았어요" 표시
 *   POST {action:'add', token, req1, req2, urgent}       → 본인이 새 기도 제목 추가 (= 갱신)
 *   POST {action:'renew', token}                         → "올해도 계속 기도해 주세요" (갱신일 연장)
 *   POST {action:'contact', token, contact}              → 연락처 수정
 *   GET  ?pastor=1                                        → 김목사 기도제목 (공개)
 *   POST {action:'setPastor', key, items:[...]}          → 김목사 기도제목 저장 (기도방)
 *   GET  ?list=1&key=KEY                                  → 기도 제목 전체 (기도방·HebronAltar)
 *   GET  ?people=1&key=KEY                                → 명단 전체 (기도방 명단 탭)
 *   POST {action:'prayed', key, id, date}                → 기도함 기록 (사람 누적 기도 횟수 +1)
 *   POST {action:'status', key, id, status, note}        → 상태 변경 (응답됨이면 본인에게 응답 메일)
 *   POST {action:'memo', key, token, memo, pstatus}      → 명단 메모·상태 수정
 *   GET  ?cron=1&key=KEY                                  → 연 1회 갱신 안내 메일 일괄 처리 (기도방 열 때 자동 호출)
 */

const SHEET_NAME   = '123 사역자 기도';
const PEOPLE_SHEET = '사역자 명단';
const PRAY_KEY     = 'CHANGE_ME';            // ← 실제 비밀번호는 Apps Script 편집기에서만 설정 (공개 저장소에 넣지 말 것)
const NOTIFY_EMAIL = '';                     // 새 기도 제목이 오면 알림 받을 이메일 (비우면 알림 없음)
const SPREADSHEET_ID = '13RaexIVdty9uZL91nOBDY4IcvIR5PQcWQ8hfJdznwXA'; // 123 사역자 기도 (pastor@ijiguchon.org 소유, gmc.hc300 편집자)
const SITE = 'https://kchurch911.com/123prayer/';
const SENDER_KO = '김성수 · 123 Prayer', SENDER_EN = 'Paul Kim · 123 Prayer';

const HEADERS = ['번호','접수일시','이름','소속','연락처','기도 제목 1','기도 제목 2','긴급','123 언약','응답 소식','공개 범위','언어','상태','기도 기록','토큰','응답일','응답 메모','기도 대상','소중한 분'];
const PHEADERS = ['토큰','이름','연락처','섬기는 자리','소속 교회·기관','언약','언어','첫 등록일','최근 갱신일','다음 갱신일','상태','받은 기도','응답 건수','갱신 알림','메모','최근 제목'];
const COL = {}; HEADERS.forEach((h, i) => COL[h] = i + 1);
const PCOL = {}; PHEADERS.forEach((h, i) => PCOL[h] = i + 1);

// ────────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse((e.postData && e.postData.contents) || '{}');
    const action = data.action || 'submit';
    if (action === 'submit') return json_(submit_(data));

    // 본인(토큰) 액션
    if (['answered', 'add', 'renew', 'contact'].includes(action)) {
      const p = findPerson_(data.token); if (!p) return json_({ ok: false, error: 'unauthorized' });
      if (action === 'answered') return json_(answered_(p, data));
      if (action === 'add')      return json_(addRequest_(p, data));
      if (action === 'renew')    return json_(renew_(p));
      if (action === 'contact')  return json_(updateContact_(p, data));
    }

    // 김목사(비밀번호) 액션
    if (!data.key || data.key !== PRAY_KEY) return json_({ ok: false, error: 'unauthorized' });
    if (action === 'setPastor') { PropertiesService.getScriptProperties().setProperty('PASTOR_REQS', JSON.stringify((data.items || []).map(String).filter(Boolean).slice(0, 6))); return json_({ ok: true, items: pastorReqs_() }); }
    if (action === 'memo') { const p = findPerson_(data.token); if (!p) return json_({ ok:false, error:'not found' }); const ps = getPeople_(); if (data.memo !== undefined) ps.getRange(p.row, PCOL['메모']).setValue(String(data.memo||'')); if (data.pstatus) ps.getRange(p.row, PCOL['상태']).setValue(String(data.pstatus)); return json_({ ok: true }); }

    const sheet = getSheet_();
    const row = findRow_(sheet, data.id);
    if (!row) return json_({ ok: false, error: 'not found' });
    if (action === 'prayed') {
      const cell = sheet.getRange(row, COL['기도 기록']);
      let log = []; try { log = JSON.parse(cell.getValue() || '[]'); } catch (err) { log = []; }
      const d = data.date || today_();
      if (!log.includes(d)) { log.push(d); cell.setValue(JSON.stringify(log)); bumpPerson_(sheet.getRange(row, COL['토큰']).getValue(), '받은 기도', 1); }
      return json_({ ok: true, log });
    }
    if (action === 'status') {
      const st = data.status || '기도중';
      sheet.getRange(row, COL['상태']).setValue(st);
      if (st === '응답됨' || st === '부분응답') {
        if (!sheet.getRange(row, COL['응답일']).getValue()) sheet.getRange(row, COL['응답일']).setValue(txt_(today_()));
        if (data.note) sheet.getRange(row, COL['응답 메모']).setValue(String(data.note));
        if (st === '응답됨') { bumpPerson_(sheet.getRange(row, COL['토큰']).getValue(), '응답 건수', 1); answeredMail_(rowObj_(sheet, row), 'pastor'); }
      }
      return json_({ ok: true });
    }
    return json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.me)     { const person = findPerson_(p.me); if (!person) return json_({ ok:false, error:'unauthorized' }); return json_({ ok:true, me: personObj_(person), requests: requestsOf_(person.token), pastor: pastorReqs_() }); }
    if (p.pastor) return json_({ ok: true, items: pastorReqs_() });
    if (p.tz) return json_({ ok: true, sheetTz: sheetTz_(), scriptTz: Session.getScriptTimeZone(), today: today_() });
    if (!p.list && !p.people && !p.cron) return ContentService.createTextOutput('123 Prayer API v5 is live');
    if (p.key !== PRAY_KEY) return json_({ ok: false, error: 'unauthorized' });
    if (p.people) return json_({ ok: true, items: allPeople_() });
    if (p.cron)   return json_(renewalSweep_());
    const items = allRequests_();
    return json_({ ok: true, items, pastor: pastorReqs_() });
  } catch (err) { return json_({ ok: false, error: String(err) }); }
}

// ── 제출 ─────────────────────────────────────────────────────────
function submit_(data) {
  const sheet = getSheet_();
  const person = upsertPerson_(data);
  const seq = nextSeq_(sheet);
  const now = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyy-MM-dd HH:mm');
  sheet.appendRow([
    seq, txt_(now), data.name || '', data.group || '', data.contact || '',
    data.req1 || '', data.req2 || '',
    data.urgent ? 'Y' : '', data.covenant ? 'Y' : '', data.notify ? 'Y' : '',
    data.share === 'class' ? '동역자 공유' : '김목사',
    data.lang === 'en' ? 'EN' : 'KO',
    '기도중', '[]', person.token, '', '',
    data.whoType === 'vip' ? '소중한 분' : '나', data.vipName || ''
  ]);
  styleRow_(sheet, sheet.getLastRow(), !!data.urgent);
  getPeople_().getRange(person.row, PCOL['최근 제목']).setValue(String(data.req1 || '').slice(0, 80));
  if (NOTIFY_EMAIL) notify_(data, now);
  confirm_(data, person.token);
  return { ok: true, id: seq, token: person.token };
}
function nextSeq_(sheet) {
  const last = sheet.getLastRow(); if (last < 2) return 1;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues().map(r => Number(r[0]) || 0);
  return Math.max.apply(null, ids.concat([0])) + 1;
}

// ── 본인 액션 ────────────────────────────────────────────────────
function answered_(p, data) {
  const sheet = getSheet_(); const row = findRow_(sheet, data.id);
  if (!row || String(sheet.getRange(row, COL['토큰']).getValue()) !== p.token) return { ok: false, error: 'not found' };
  sheet.getRange(row, COL['상태']).setValue('응답됨');
  sheet.getRange(row, COL['응답일']).setValue(txt_(today_()));
  if (data.note) sheet.getRange(row, COL['응답 메모']).setValue(String(data.note).slice(0, 300));
  bumpPerson_(p.token, '응답 건수', 1);
  answeredMail_(rowObj_(sheet, row), 'self');
  return { ok: true };
}
function addRequest_(p, data) {
  const ps = getPeople_();
  const name = ps.getRange(p.row, PCOL['이름']).getValue(), contact = ps.getRange(p.row, PCOL['연락처']).getValue();
  const group = ps.getRange(p.row, PCOL['섬기는 자리']).getValue(), lang = ps.getRange(p.row, PCOL['언어']).getValue();
  const sheet = getSheet_(); const seq = nextSeq_(sheet);
  const now = Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyy-MM-dd HH:mm');
  sheet.appendRow([seq, txt_(now), name, group, contact, data.req1 || '', data.req2 || '', data.urgent ? 'Y' : '', 'Y', '', '김목사', lang || 'KO', '기도중', '[]', p.token, '', '', '나', '']);
  styleRow_(sheet, sheet.getLastRow(), !!data.urgent);
  ps.getRange(p.row, PCOL['최근 제목']).setValue(String(data.req1 || '').slice(0, 80));
  touchRenew_(p.row);
  return { ok: true, id: seq };
}
function renew_(p) { touchRenew_(p.row); return { ok: true, next: getPeople_().getRange(p.row, PCOL['다음 갱신일']).getValue() }; }
function updateContact_(p, data) {
  const ps = getPeople_();
  if (data.contact !== undefined) ps.getRange(p.row, PCOL['연락처']).setValue(String(data.contact || '').trim());
  return { ok: true };
}
function touchRenew_(row) {
  const ps = getPeople_(); const t = today_();
  ps.getRange(row, PCOL['최근 갱신일']).setValue(txt_(t));
  ps.getRange(row, PCOL['다음 갱신일']).setValue(txt_(addDays_(t, 365)));
  ps.getRange(row, PCOL['상태']).setValue('활동');
  ps.getRange(row, PCOL['갱신 알림']).setValue('[]');
}

// ── 명단 ─────────────────────────────────────────────────────────
function personKey_(name, contact) {
  const c = String(contact || '').trim().toLowerCase();
  if (/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(c)) return 'e:' + c;
  const digits = c.replace(/\D/g, '');
  return 'n:' + String(name || '').trim().replace(/\s+/g, '') + (digits ? ':' + digits : '');
}
function upsertPerson_(data) {
  const ps = getPeople_(); const last = ps.getLastRow();
  const key = personKey_(data.name, data.contact);
  if (last >= 2) {
    const rows = ps.getRange(2, 1, last - 1, PHEADERS.length).getValues();
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (personKey_(r[PCOL['이름']-1], r[PCOL['연락처']-1]) === key) {
        const row = i + 2;
        if (data.group) ps.getRange(row, PCOL['섬기는 자리']).setValue(data.group);
        if (data.org) ps.getRange(row, PCOL['소속 교회·기관']).setValue(data.org);
        if (data.covenant) ps.getRange(row, PCOL['언약']).setValue('Y');
        ps.getRange(row, PCOL['언어']).setValue(data.lang === 'en' ? 'EN' : 'KO');
        touchRenew_(row);
        return { row, token: String(r[PCOL['토큰']-1]) };
      }
    }
  }
  const token = Utilities.getUuid().replace(/-/g, '').slice(0, 20);
  const t = today_();
  ps.appendRow([token, data.name || '', data.contact || '', data.group || '', data.org || '', data.covenant ? 'Y' : '', data.lang === 'en' ? 'EN' : 'KO', txt_(t), txt_(t), txt_(addDays_(t, 365)), '활동', 0, 0, '[]', '', '']);
  return { row: ps.getLastRow(), token };
}
function findPerson_(token) {
  if (!token) return null;
  const ps = getPeople_(); const last = ps.getLastRow(); if (last < 2) return null;
  const toks = ps.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < toks.length; i++) if (String(toks[i][0]) === String(token)) return { row: i + 2, token: String(token) };
  return null;
}
function bumpPerson_(token, col, n) {
  const p = findPerson_(token); if (!p) return;
  const ps = getPeople_(); const c = ps.getRange(p.row, PCOL[col]); c.setValue((Number(c.getValue()) || 0) + n);
}
function personObj_(p) {
  const ps = getPeople_(); const r = ps.getRange(p.row, 1, 1, PHEADERS.length).getValues()[0];
  const o = {}; PHEADERS.forEach((h, i) => o[h] = (r[i] instanceof Date) ? fmt_(r[i]) : r[i]);
  return { token: o['토큰'], name: o['이름'], contact: o['연락처'], group: o['섬기는 자리'], org: o['소속 교회·기관'], covenant: o['언약'] === 'Y', lang: o['언어'], joined: o['첫 등록일'], renewed: o['최근 갱신일'], next: o['다음 갱신일'], status: o['상태'], prayed: Number(o['받은 기도'])||0, answered: Number(o['응답 건수'])||0, memo: o['메모'], lastReq: o['최근 제목'] };
}
function allPeople_() {
  const ps = getPeople_(); const last = ps.getLastRow(); if (last < 2) return [];
  return ps.getRange(2, 1, last - 1, PHEADERS.length).getValues().filter(r => r[0]).map((r, i) => personObj_({ row: i + 2, token: String(r[0]) }));
}

// ── 기도 제목 조회 ────────────────────────────────────────────────
function rowObj_(sheet, row) {
  const r = sheet.getRange(row, 1, 1, HEADERS.length).getValues()[0];
  let log = []; try { log = JSON.parse(r[COL['기도 기록']-1] || '[]'); } catch (err) {}
  return { id: r[0], at: fmt_(r[1]), name: r[2], group: r[3], contact: r[4], req1: r[5], req2: r[6],
           urgent: r[7] === 'Y', covenant: r[8] === 'Y', notify: r[9] === 'Y', share: r[10], lang: r[11],
           status: r[12] || '기도중', log, token: r[14] || '', answeredAt: fmt_(r[15]), answerNote: r[16] || '', whoType: r[17] || '나', vipName: r[18] || '' };
}
function allRequests_() {
  const sheet = getSheet_(); const last = sheet.getLastRow(); if (last < 2) return [];
  const out = []; for (let row = 2; row <= last; row++) { const o = rowObj_(sheet, row); if (o.id !== '') out.push(o); }
  return out;
}
function requestsOf_(token) { return allRequests_().filter(r => String(r.token) === String(token)); }
function pastorReqs_() {
  try { const v = PropertiesService.getScriptProperties().getProperty('PASTOR_REQS'); if (v) return JSON.parse(v); } catch (err) {}
  return ['말씀과 기도, 종의 리더십으로 사역하게 하소서', '가정과 사역마다 복음의 열매 맺게 하소서'];
}

// ── 연 1회 갱신 안내 (기도방 열 때 ?cron=1 로 호출) ─────────────────
function renewalSweep_() {
  const ps = getPeople_(); const last = ps.getLastRow(); if (last < 2) return { ok: true, sent: 0 };
  const rows = ps.getRange(2, 1, last - 1, PHEADERS.length).getValues();
  let sent = 0, archived = 0; const t = today_();
  rows.forEach((r, i) => {
    const row = i + 2; const status = r[PCOL['상태']-1]; if (status === '보관') return;
    const email = String(r[PCOL['연락처']-1] || '').trim(); const isEmail = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
    const next = fmt_(r[PCOL['다음 갱신일']-1]); if (!next) return;
    const days = daysBetween_(t, next);   // 음수면 갱신일 지남
    let log = []; try { log = JSON.parse(r[PCOL['갱신 알림']-1] || '[]'); } catch (err) {}
    const stage = days <= -14 ? 'final' : days <= 0 ? 'due' : days <= 14 ? 'soon' : null;
    if (!stage || log.includes(stage)) return;
    if (isEmail) { renewalMail_(personObj_({ row, token: String(r[0]) }), stage); sent++; }
    log.push(stage); ps.getRange(row, PCOL['갱신 알림']).setValue(JSON.stringify(log));
    if (stage === 'due')   ps.getRange(row, PCOL['상태']).setValue('갱신 대기');
    if (stage === 'final') { ps.getRange(row, PCOL['상태']).setValue('보관'); archived++; }
  });
  return { ok: true, sent, archived };
}

// ── 메일 ─────────────────────────────────────────────────────────
function isEmail_(s) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s || '').trim()); }
function myLink_(token) { return SITE + 'my.html?t=' + token; }
function confirm_(d, token) {
  try {
    const to = String(d.contact || '').trim(); if (!isEmail_(to)) return;
    const name = String(d.name || '').trim(); const en = d.lang === 'en'; const pr = pastorReqs_();
    const subject = en ? '[123 Prayer] Your prayer request has arrived' : '[123 Prayer] ' + name + '님, 기도 제목이 잘 도착했습니다';
    const body = en
      ? [name + ', your prayer request has arrived.', 'Starting this week I will call your name in prayer — twice a week, at least three seconds each.', '', 'Your request:', '- ' + (d.req1 || ''), d.req2 ? '- ' + d.req2 : '', '', "Pastor Kim's prayer requests (please pray twice a week, 3+ seconds):", ...pr.map(x => '- ' + x), '', 'Your prayer page (update requests, mark answered):', myLink_(token), '', '“Pray without ceasing; give thanks in all circumstances.” (1 Thess 5:17-18)', '', SENDER_EN, SITE].filter(x => x !== '').join('\n')
      : [name + '님, 기도 제목이 잘 도착했습니다.', '이번 주부터 ' + name + '님의 이름을 부르며 일주일에 두 번, 삼 초 이상 기도하겠습니다.', '', '보내주신 기도 제목:', '- ' + (d.req1 || ''), d.req2 ? '- ' + d.req2 : '', '', '김목사 기도제목 (일주일에 두 번, 3초 이상 함께 기도해 주세요):', ...pr.map(x => '- ' + x), '', '나의 기도 페이지 (기도 제목 추가·응답 표시):', myLink_(token), '', '“쉬지 말고 기도하라. 범사에 감사하라.” (살전 5:17-18)', '', SENDER_KO, SITE].filter(x => x !== '').join('\n');
    MailApp.sendEmail({ to, subject, body, name: en ? SENDER_EN : SENDER_KO });
  } catch (err) { Logger.log('confirm mail failed: ' + err); }
}
function answeredMail_(r, by) {
  try {
    if (!isEmail_(r.contact)) return;
    const en = r.lang === 'EN'; const name = String(r.name || '').trim();
    const subject = en ? '[123 Prayer] Rejoicing with you — a prayer answered' : '[123 Prayer] ' + name + '님, 기도 응답을 함께 기뻐합니다';
    const body = en
      ? [name + ',', by === 'pastor' ? 'I have marked this request as answered. Let us give thanks together:' : 'Thank you for sharing that God answered. I rejoice with you:', '- ' + (r.req1 || ''), r.answerNote ? 'Note: ' + r.answerNote : '', '', 'I keep praying for you twice a week. Share a new request any time:', myLink_(r.token), '', SENDER_EN].filter(x => x !== '').join('\n')
      : [name + '님,', by === 'pastor' ? '이 기도 제목을 "응답됨"으로 기록했습니다. 함께 감사드립니다:' : '하나님이 응답하셨다는 소식, 함께 기뻐합니다:', '- ' + (r.req1 || ''), r.answerNote ? '메모: ' + r.answerNote : '', '', '계속 일주일에 두 번 ' + name + '님의 이름을 부르며 기도합니다. 새 기도 제목은 언제든 여기서:', myLink_(r.token), '', SENDER_KO].filter(x => x !== '').join('\n');
    MailApp.sendEmail({ to: r.contact, subject, body, name: en ? SENDER_EN : SENDER_KO });
  } catch (err) { Logger.log('answered mail failed: ' + err); }
}
function renewalMail_(p, stage) {
  try {
    const en = p.lang === 'EN'; const name = p.name; const link = myLink_(p.token); const pr = pastorReqs_();
    const S = {
      soon:  en ? ['[123 Prayer] One year of prayer — let’s renew in two weeks', 'It has been almost a year since I began praying for you by name. In two weeks the covenant comes up for renewal.', 'If you would like me to keep praying, open your page and tap “Keep praying for me this year” — and update your requests if they have changed.']
                : ['[123 Prayer] 기도한 지 1년 — 2주 뒤 갱신합니다', name + '님의 이름을 부르며 기도한 지 곧 1년이 됩니다. 2주 뒤에 언약을 갱신합니다.', '계속 기도받기를 원하시면 나의 기도 페이지에서 "올해도 계속 기도해 주세요"를 눌러 주시고, 바뀐 기도 제목이 있으면 새로 적어 주세요.'],
      due:   en ? ['[123 Prayer] Today: renew our prayer covenant', 'Today marks one year. Thank you for walking this year in prayer together.', 'Tap “Keep praying for me this year” on your page to continue (and share a fresh request). If I don’t hear back in two weeks, I will gently set your page aside — you can come back any time.']
                : ['[123 Prayer] 오늘, 기도 언약을 갱신합니다', '오늘로 1년이 되었습니다. 한 해 동안 함께 기도해 주셔서 감사합니다.', '나의 기도 페이지에서 "올해도 계속 기도해 주세요"를 누르시면 계속됩니다(새 기도 제목도 적어 주세요). 2주 안에 응답이 없으면 조용히 보관 상태로 두겠습니다 — 언제든 다시 오실 수 있습니다.'],
      final: en ? ['[123 Prayer] Last note — your page will rest for now', 'I have not heard back, so I will set your prayer page aside for now — with gratitude for this year.', 'Whenever you wish, open your page and tap “Keep praying for me” or send a new request, and we begin again.']
                : ['[123 Prayer] 마지막 안내 — 잠시 보관합니다', '응답이 없어 기도 페이지를 잠시 보관 상태로 둡니다 — 한 해 동안 감사했습니다.', '언제든 페이지에서 "올해도 계속 기도해 주세요"를 누르거나 새 기도 제목을 보내시면 다시 시작합니다.']
    }[stage];
    const body = [name + (en ? ',' : '님,'), '', S[1], S[2], '', en ? 'Your page:' : '나의 기도 페이지:', link, '', en ? "Pastor Kim's prayer requests:" : '김목사 기도제목:', ...pr.map(x => '- ' + x), '', en ? SENDER_EN : SENDER_KO, SITE].join('\n');
    MailApp.sendEmail({ to: p.contact, subject: S[0], body, name: en ? SENDER_EN : SENDER_KO });
  } catch (err) { Logger.log('renewal mail failed: ' + err); }
}
function notify_(d, now) {
  try { MailApp.sendEmail({ to: NOTIFY_EMAIL, subject: '[123 기도] 새 기도 제목 — ' + (d.name || ''), body: ['접수: ' + now, '이름: ' + (d.name || ''), '섬기는 자리: ' + (d.group || ''), '', '기도 제목 1: ' + (d.req1 || ''), d.req2 ? '기도 제목 2: ' + d.req2 : '', d.urgent ? '※ 긴급' : ''].filter(Boolean).join('\n') }); }
  catch (err) { Logger.log('notify mail failed: ' + err); }
}

// ── 시트 ─────────────────────────────────────────────────────────
function ss_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (err) { ss = null; }
  if (!ss && SPREADSHEET_ID) { try { ss = SpreadsheetApp.openById(SPREADSHEET_ID); } catch (err) { ss = null; } }
  if (!ss) {
    const props = PropertiesService.getScriptProperties(); const id = props.getProperty('SSID');
    if (id) { try { ss = SpreadsheetApp.openById(id); } catch (err) { ss = null; } }
    if (!ss) { ss = SpreadsheetApp.create(SHEET_NAME); props.setProperty('SSID', ss.getId()); }
  }
  return ss;
}
function getSheet_() {
  const ss = ss_(); let sh = ss.getSheetByName(SHEET_NAME); if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) { sh.appendRow(HEADERS); sh.setFrozenRows(1); }
  // 헤더 보강 (v5: 토큰·응답일·응답 메모·기도 대상·소중한 분)
  const cur = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  if (cur.length < HEADERS.length || cur[HEADERS.length - 1] !== HEADERS[HEADERS.length - 1]) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  sh.getRange(1, 1, 1, HEADERS.length).setBackground('#1a1a2e').setFontColor('#f0c040').setFontWeight('bold');
  return sh;
}
function getPeople_() {
  const ss = ss_(); let sh = ss.getSheetByName(PEOPLE_SHEET);
  if (!sh) { sh = ss.insertSheet(PEOPLE_SHEET); sh.appendRow(PHEADERS); sh.setFrozenRows(1); sh.getRange(1, 1, 1, PHEADERS.length).setBackground('#1a1a2e').setFontColor('#f0c040').setFontWeight('bold'); [160,90,170,150,150,50,50,95,95,95,80,70,70,120,220,260].forEach((w,i) => sh.setColumnWidth(i+1, w)); }
  return sh;
}
function findRow_(sheet, id) {
  const last = sheet.getLastRow(); if (last < 2) return null;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return null;
}
function styleRow_(sheet, row, urgent) {
  const rg = sheet.getRange(row, 1, 1, HEADERS.length);
  rg.setBackground(urgent ? '#fff1f2' : (row % 2 ? '#ffffff' : '#f8f7fb'));
  sheet.getRange(row, 6, 1, 2).setWrap(true);
}
// ── 유틸 ─────────────────────────────────────────────────────────
function today_() { return Utilities.formatDate(new Date(), 'America/Los_Angeles', 'yyyy-MM-dd'); }
function txt_(s) { return "'" + String(s); }   // 시트에 텍스트로 저장 (날짜 자동 변환 방지)
function fmt_(v) { if (!v) return ''; if (v instanceof Date) return Utilities.formatDate(v, sheetTz_(), 'yyyy-MM-dd'); const s = String(v); return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.slice(0,10) : s; }
let _tz = null; function sheetTz_() { if (_tz) return _tz; try { _tz = ss_().getSpreadsheetTimeZone(); } catch (err) { _tz = 'America/Los_Angeles'; } return _tz; }
function addDays_(ymd, n) { const d = new Date(ymd + 'T12:00:00'); d.setDate(d.getDate() + n); return Utilities.formatDate(d, 'America/Los_Angeles', 'yyyy-MM-dd'); }
function daysBetween_(a, b) { return Math.round((new Date(b + 'T12:00:00') - new Date(a + 'T12:00:00')) / 86400000); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
// Drive 권한 확보용 (SpreadsheetApp.create 에 필요) — 호출되지 않음
function ensureDriveScope_() { const f = DriveApp.createFile('scope-check.txt', 'x'); f.setTrashed(true); return DriveApp.getRootFolder().getName(); }
// 에디터에서 실행: 시트·명단 생성/헤더 보강 확인
function setupCheck() { const sh = getSheet_(); getPeople_(); Logger.log('Sheet URL: ' + sh.getParent().getUrl()); return sh.getParent().getUrl(); }
