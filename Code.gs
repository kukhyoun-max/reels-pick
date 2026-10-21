/**
 * 요즘 본 릴스 제보 — Apps Script 백엔드
 * 구글 시트에 바인딩된 스크립트로 사용 (시트 > 확장 프로그램 > Apps Script)
 * 목적: 일반 학생이 실제로 반응한 콘텐츠 수집 → 분석은 관리자가 시트에서
 */

const SHEET_NAME = '릴스제보';
const PROFILE_SHEET = '조사자 정보';            // 순서 | 별명 | 이름 | 성별 | 나이
const PROFILE_HEADERS = ['이름', '성별', '나이']; // 릴스제보 탭 맨 오른쪽에 자동으로 채워짐
const HEADERS = [
  '제출일시', '별명', '학년', 'URL', '플랫폼',
  '추천 점수(1~5)', '어떤 영상?', '왜 끌렸어?', '나의 반응', '대학 계정이 해도 볼래?',
  '한마디', '관리자 메모'
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('요즘 본 릴스 알려줘')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** GitHub Pages(index.html)에서 보내는 제보·교차 평가 받기 (action 없음 = 제보) */
function doPost(e) {
  let res;
  try {
    const data = JSON.parse(e.postData.contents);
    res = data.action === 'queue' ? getQueue(data)
      : data.action === 'vote' ? submitVote(data)
      : submitReel(data);
  } catch (err) {
    res = { ok: false, msg: '저장 중 오류가 났어요. 잠시 후 다시 시도해 줘요.' };
  }
  return ContentService.createTextOutput(JSON.stringify(res))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  // 미리 만들어 둔 시트(첫 탭에 헤더가 있는 경우)는 그 탭을 그대로 사용
  if (!sh) {
    const first = ss.getSheets()[0];
    if (first && first.getRange(1, 1).getValue() === HEADERS[0]) {
      sh = first.setName(SHEET_NAME);
      sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, HEADERS.length)
        .setFontWeight('bold').setBackground('#1f3a5f').setFontColor('#ffffff');
      return sh;
    }
  }
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold').setBackground('#1f3a5f').setFontColor('#ffffff');
    sh.setColumnWidth(4, 260);
    sh.setColumnWidth(8, 200);
    sh.setColumnWidth(9, 200);
    sh.setColumnWidth(11, 260);
  }
  return sh;
}

function detectPlatform_(url) {
  if (/instagram\.com/i.test(url)) return '인스타그램';
  if (/youtube\.com|youtu\.be/i.test(url)) return '유튜브';
  if (/tiktok\.com/i.test(url)) return '틱톡';
  return '기타';
}

/** 공유 파라미터(?igsh= 등)를 떼서 중복 판별이 되게 정리 */
function normalizeUrl_(url) {
  url = String(url || '').trim();
  if (/instagram\.com|tiktok\.com|youtube\.com\/shorts|youtu\.be/i.test(url)) url = url.split('?')[0];
  return url.replace(/\/+$/, '');
}

/** 수식 주입 방지 + 배열은 콤마로 */
function clean_(v) {
  const s = Array.isArray(v) ? v.join(', ') : String(v == null ? '' : v).trim();
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function submitReel(data) {
  if (!data || !data.url || !data.name) {
    return { ok: false, msg: '별명이랑 링크는 꼭 넣어줘!' };
  }
  const sc = Number(data.score);
  if (!(sc >= 1 && sc <= 5)) {
    return { ok: false, msg: '별점을 매겨줘요! ⭐' };
  }
  const url = normalizeUrl_(data.url);
  if (!/^https?:\/\//i.test(url)) {
    return { ok: false, msg: '링크가 이상해요. 공유 → 링크 복사로 다시 붙여넣어 줘!' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_();
    const last = sh.getLastRow();
    if (last > 1) {
      const urls = sh.getRange(2, 4, last - 1, 1).getValues().flat();
      if (urls.indexOf(url) !== -1) {
        return { ok: false, msg: '이 영상은 이미 누가 올렸어요. 다른 거 있으면 알려줘!' };
      }
    }
    sh.appendRow([
      new Date(),
      clean_(data.name),
      clean_(data.grade),
      url,
      detectPlatform_(url),
      Number(data.score) || '',
      clean_(data.kind),
      clean_(data.why),
      clean_(data.action),
      clean_(data.univ),
      clean_(data.comment),
      ''
    ]);
    syncProfiles_(sh);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 조사자 정보 탭의 별명 → 이름·성별·나이를 릴스제보 탭 오른쪽 열에 채움
 * 제보가 들어올 때마다 전체를 다시 맞추므로, 조사자 정보에 별명을 나중에 적어도 다음 제보 때 과거 줄까지 반영됨
 * 수동으로 바로 반영하려면 편집기에서 syncProfilesNow 실행
 */
function syncProfiles_(sh) {
  const ss = sh.getParent();
  const ps = ss.getSheetByName(PROFILE_SHEET);
  if (!ps) return;

  const map = {};
  const plast = ps.getLastRow();
  if (plast > 1) {
    ps.getRange(2, 2, plast - 1, 4).getValues().forEach(r => {
      const nick = String(r[0]).trim();
      if (nick) map[nick] = [r[1], r[2], r[3]];
    });
  }

  const startCol = HEADERS.length + 1;
  sh.getRange(1, startCol, 1, PROFILE_HEADERS.length).setValues([PROFILE_HEADERS])
    .setFontWeight('bold').setBackground('#4a6a8f').setFontColor('#ffffff');

  const last = sh.getLastRow();
  if (last < 2) return;
  const nicks = sh.getRange(2, 2, last - 1, 1).getValues();
  const out = nicks.map(r => map[String(r[0]).trim()] || ['미등록', '', '']);
  sh.getRange(2, startCol, out.length, PROFILE_HEADERS.length).setValues(out);
}

function syncProfilesNow() {
  syncProfiles_(getSheet_());
}

/* =====================================================================
 *  🗳️ 교차 평가 — 다른 조사자가 제보한 릴스를 "우리 계정에서 보고 싶다 / 아니다"로 평가
 *  릴스마다 여러 명의 판단이 쌓이게, 평가 수가 적은 영상부터 내려준다
 * ===================================================================== */
const VOTE_SHEET = '교차평가';
const VOTE_HEADERS = ['평가일시', '평가자', 'URL', '판단', '어울리는 계정'];
const VOTE_YES = '보고 싶다', VOTE_NO = '아니다', VOTE_SKIP = '못 봄';
const QUEUE_SIZE = 10;
const TOP_PICKS = 3;  // 제보자가 하루에 낸 영상 중 평가에 올리는 개수
const MIN_VOTES = 3;  // 평가 3명 미만은 분석 탭에 "표본 적음"

function getVoteSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(VOTE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(VOTE_SHEET);
    sh.appendRow(VOTE_HEADERS);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, VOTE_HEADERS.length)
      .setFontWeight('bold').setBackground('#1f3a5f').setFontColor('#ffffff');
    sh.setColumnWidth(3, 260);
  }
  return sh;
}

/** 별명 비교용: 앞뒤·중간 공백, 보이지 않는 글자, 한글 조합 방식 차이를 없앤다 */
function nickKey_(v) {
  return String(v == null ? '' : v).normalize('NFC').replace(/[\s​-‏⁠﻿]/g, '');
}

function bodyRows_(sh, ncol) {
  const last = sh.getLastRow();
  return last > 1 ? sh.getRange(2, 1, last - 1, ncol).getValues() : [];
}

/**
 * 제보자가 하루에 낸 영상 중 상위 TOP_PICKS개만 평가에 올린다 (스토리 링크는 순위 전에 뺌)
 * 순위: 별점 → 강한 반응(돌려봄·공유·저장·팔로우) → 대학 계정 "볼 것 같다" → 먼저 낸 것
 */
function topPicks_(rows) {
  const groups = {};
  rows.forEach(r => {   // 제출일시 | 별명 | 학년 | URL | 플랫폼 | 별점 | 유형 | 이유 | 반응 | 대학 계정
    const url = String(r[3]), nick = nickKey_(r[1]);
    if (!url || !nick || /\/stories\//i.test(url)) return;
    const t = r[0] instanceof Date ? r[0] : new Date(r[0]);
    const day = isNaN(t) ? '' : Utilities.formatDate(t, 'Asia/Seoul', 'yyyy-MM-dd');
    const acts = String(r[8]).split(',').map(s => s.trim());
    (groups[nick + '|' + day] = groups[nick + '|' + day] || []).push({
      nick: nick, url: url, t: isNaN(t) ? 0 : t.getTime(),
      score: Number(r[5]) || 0,
      strong: acts.some(a => STRONG_ACTIONS.indexOf(a) !== -1) ? 1 : 0,
      ok: r[9] === '볼 것 같다' ? 1 : 0
    });
  });
  return Object.keys(groups).reduce((out, k) => out.concat(groups[k]
    .sort((a, b) => b.score - a.score || b.strong - a.strong || b.ok - a.ok || a.t - b.t)
    .slice(0, TOP_PICKS)), []);
}

/** 평가할 영상 목록: 제보자별 하루 상위 3개 중 내가 제보했거나 이미 평가한 영상, 스토리 링크(24시간 뒤 사라짐)는 뺀다 */
function getQueue(data) {
  const name = nickKey_(data && data.name);
  if (!name) return { ok: false, msg: '별명을 먼저 정해줘요!' };

  const cnt = {}, mine = {};
  let own = 0;  // 내가 제보해서 빠진 영상 수
  bodyRows_(getVoteSheet_(), 4).forEach(v => {
    const u = String(v[2]);
    if (nickKey_(v[1]) === name) mine[u] = 1;
    if (v[3] === VOTE_YES || v[3] === VOTE_NO) cnt[u] = (cnt[u] || 0) + 1;
  });

  const seen = {};
  const pool = topPicks_(bodyRows_(getSheet_(), 10))
    .filter(r => {
      if (r.nick === name) { own++; return false; }
      if (seen[r.url] || mine[r.url]) return false;
      return (seen[r.url] = true);
    })
    .map(r => ({ url: r.url, n: cnt[r.url] || 0, k: Math.random() }))
    .sort((a, b) => a.n - b.n || a.k - b.k);

  return { ok: true, items: pool.slice(0, QUEUE_SIZE).map(r => r.url), left: pool.length, own: own };
}

function submitVote(data) {
  const name = nickKey_(data && data.name);
  const url = normalizeUrl_(data && data.url);
  const verdict = [VOTE_YES, VOTE_NO, VOTE_SKIP].indexOf(data && data.verdict) !== -1 ? data.verdict : '';
  if (!name || !url || !verdict) {
    return { ok: false, msg: '평가가 저장되지 않았어요. 다시 눌러줘요.' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getVoteSheet_();
    // 같은 사람이 같은 영상을 두 번 평가하면 처음 것만 남김 (버튼 연타·재전송 대비)
    const dup = bodyRows_(sh, 3).some(v => nickKey_(v[1]) === name && String(v[2]) === url);
    if (!dup) {
      sh.appendRow([new Date(), clean_(name), url, verdict, verdict === VOTE_YES ? clean_(data.account) : '']);
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* =====================================================================
 *  📊 분석 탭 — 메뉴 [📊 릴스 분석 → 분석 새로고침] 또는 매일 아침 자동 갱신
 * ===================================================================== */
const REPORT_SHEET = '분석';
const STRONG_ACTIONS = ['여러 번 돌려봄', '친구 태그/공유', '저장', '계정 팔로우']; // "강한 반응"

function onOpen() {
  SpreadsheetApp.getUi().createMenu('📊 릴스 분석')
    .addItem('분석 새로고침', 'buildReport')
    .addItem('매일 아침 8시 자동 갱신 켜기', 'enableDailyReport')
    .addToUi();
}

function enableDailyReport() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'buildReport')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('buildReport').timeBased().everyDays(1).atHour(8).create();
  try { SpreadsheetApp.getUi().alert('매일 아침 8시에 분석 탭이 자동으로 갱신돼요.'); } catch (e) {}
}

/** 순수 계산부: 시트 값(2차원 배열) → 표 블록들. votes는 교차평가 탭 값(없으면 빈 배열) */
function computeReport_(values, votes) {
  const head = values[0].map(String);
  const col = name => head.indexOf(name);
  const split = v => String(v == null ? '' : v).split(',').map(s => s.trim()).filter(Boolean);
  const C = {
    nick: col('별명'), grade: col('학년'), url: col('URL'), score: col('추천 점수(1~5)'),
    kind: col('어떤 영상?'), why: col('왜 끌렸어?'), act: col('나의 반응'),
    univ: col('대학 계정이 해도 볼래?'), comment: col('한마디'), gender: col('성별')
  };
  const recs = values.slice(1).filter(r => r[C.url]).map(r => {
    const acts = split(r[C.act]);
    return {
      nick: String(r[C.nick] || ''), grade: r[C.grade] ? r[C.grade] + '학년' : '미입력',
      gender: C.gender >= 0 && r[C.gender] && r[C.gender] !== '미등록' ? String(r[C.gender]) : '미등록',
      url: r[C.url], score: Number(r[C.score]) || 0,
      kinds: split(r[C.kind]), whys: split(r[C.why]), acts: acts,
      univYes: r[C.univ] === '볼 것 같다',
      strong: acts.some(a => STRONG_ACTIONS.indexOf(a) !== -1),
      comment: String(r[C.comment] || '')
    };
  });
  const N = recs.length;
  const pct = (a, b) => b ? a / b : 0;

  function group(getList) {
    const m = {};
    recs.forEach(r => getList(r).forEach(k => {
      const g = m[k] || (m[k] = { n: 0, s: 0, u: 0, st: 0 });
      g.n++; g.s += r.score; g.u += r.univYes ? 1 : 0; g.st += r.strong ? 1 : 0;
    }));
    return Object.keys(m).map(k => [k, m[k].n, pct(m[k].n, N), pct(m[k].s, m[k].n), pct(m[k].u, m[k].n), pct(m[k].st, m[k].n)])
      .sort((a, b) => b[1] - a[1] || b[3] - a[3]);
  }
  function cross(segOf, getList) {
    const segs = {}, cnt = {};
    recs.forEach(r => {
      const s = segOf(r); segs[s] = (segs[s] || 0) + 1;
      getList(r).forEach(k => { cnt[k] = cnt[k] || {}; cnt[k][s] = (cnt[k][s] || 0) + 1; });
    });
    const segKeys = Object.keys(segs).sort();
    const rows = Object.keys(cnt).map(k => [k].concat(segKeys.map(s => pct(cnt[k][s] || 0, segs[s]))))
      .sort((a, b) => b.slice(1).reduce((x, y) => x + y, 0) - a.slice(1).reduce((x, y) => x + y, 0));
    return { headers: ['영상 유형'].concat(segKeys.map(s => `${s} (${segs[s]}건)`)), rows: rows };
  }
  const statHead = ['항목', '건수', '전체 대비', '평균 별점', '대학계정 OK', '강한 반응'];
  const statFmt = [null, '0', '0%', '0.0', '0%', '0%'];
  const nicks = {}; recs.forEach(r => nicks[r.nick] = 1);

  const top = recs.filter(r => r.score >= 4 && r.univYes)
    .sort((a, b) => b.score - a.score || (b.strong - a.strong))
    .slice(0, 20)
    .map(r => [r.score, r.kinds.join(', '), r.whys.join(', '), r.acts.join(', '), r.comment, r.url]);

  // 교차 평가 집계: 영상별 보고 싶다 / 아니다 (못 봄은 세지 않음)
  const V = {};
  let voteN = 0;
  (votes || []).slice(1).forEach(v => {
    const u = String(v[2] || '');
    if (!u || (v[3] !== VOTE_YES && v[3] !== VOTE_NO)) return;
    const g = V[u] || (V[u] = { y: 0, n: 0, acc: {} });
    voteN++;
    if (v[3] === VOTE_YES) { g.y++; split(v[4]).forEach(a => g.acc[a] = (g.acc[a] || 0) + 1); }
    else g.n++;
  });
  const voted = recs.filter(r => V[r.url]).map(r => {
    const g = V[r.url], t = g.y + g.n;
    const acc = Object.keys(g.acc).sort((a, b) => g.acc[b] - g.acc[a])[0] || '';
    return [pct(g.y, t), t, g.y, acc, r.score, r.kinds.join(', '), r.comment, r.url, t < MIN_VOTES ? '표본 적음' : ''];
  }).sort((a, b) => (b[1] >= MIN_VOTES) - (a[1] >= MIN_VOTES) || b[0] - a[0] || b[1] - a[1]);

  return [
    { title: '① 한눈에 보기', headers: ['지표', '값'], fmt: [null, null], rows: [
      ['총 제보', N + '건'],
      ['참여 조사자', Object.keys(nicks).length + '명'],
      ['평균 별점', N ? (recs.reduce((s, r) => s + r.score, 0) / N).toFixed(2) + ' / 5' : '-'],
      ['"대학 계정이 해도 볼 것 같다"', Math.round(pct(recs.filter(r => r.univYes).length, N) * 100) + '%'],
      ['강한 반응(돌려봄·공유·저장·팔로우)', Math.round(pct(recs.filter(r => r.strong).length, N) * 100) + '%'],
      ['교차 평가', `${voteN}표 · 평가받은 영상 ${voted.length}개 · ${MIN_VOTES}명 이상 평가 ${voted.filter(r => r[1] >= MIN_VOTES).length}개`]
    ] },
    { title: '② 어떤 영상에 반응했나 (영상 유형)', headers: statHead, fmt: statFmt, rows: group(r => r.kinds), chart: '영상 유형별 제보 수' },
    { title: '③ 왜 끝까지 봤나 (끌린 이유)', headers: statHead, fmt: statFmt, rows: group(r => r.whys), chart: '끌린 이유별 제보 수' },
    { title: '④ 보고 나서 한 행동', headers: statHead, fmt: statFmt, rows: group(r => r.acts) },
    Object.assign({ title: '⑤ 성별로 좋아하는 영상 유형 (각 성별 제보 중 비율)', fmtAll: '0%' }, cross(r => r.gender, r => r.kinds)),
    Object.assign({ title: '⑥ 학년별로 좋아하는 영상 유형 (각 학년 제보 중 비율)', fmtAll: '0%' }, cross(r => r.grade, r => r.kinds)),
    { title: '⑦ 벤치마킹 후보 (별점 4+ & 대학 계정이 해도 볼 것 같다)', headers: ['별점', '영상 유형', '끌린 이유', '반응', '한마디', 'URL'], fmt: ['0', null, null, null, null, null], rows: top },
    { title: `⑧ 교차 평가 순위 (제보자 말고 다른 조사자가 "보고 싶다"를 누른 비율 · ${MIN_VOTES}명 미만은 표본 적음)`,
      headers: ['보고 싶다 비율', '평가 수', '보고 싶다', '어울리는 계정', '제보 별점', '영상 유형', '한마디', 'URL', '비고'],
      fmt: ['0%', '0', '0', null, '0', null, null, null, null], rows: voted }
  ];
}

function buildReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = getSheet_();
  syncProfiles_(src);
  const vs = ss.getSheetByName(VOTE_SHEET);
  const blocks = computeReport_(src.getDataRange().getValues(), vs ? vs.getDataRange().getValues() : []);

  let sh = ss.getSheetByName(REPORT_SHEET) || ss.insertSheet(REPORT_SHEET, 0);
  sh.getCharts().forEach(c => sh.removeChart(c));
  sh.clear();
  sh.setConditionalFormatRules([]);
  sh.getRange('A1').setValue('📊 요즘 본 릴스 — 학생 반응 분석').setFontSize(16).setFontWeight('bold');
  sh.getRange('A2').setValue('마지막 갱신: ' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm') +
    '   ·   대학계정 OK = "볼 것 같다" 비율   ·   강한 반응 = 여러 번 돌려봄/공유/저장/팔로우 중 하나 이상').setFontColor('#6b7686');

  let row = 4;
  blocks.forEach(b => {
    const w = b.headers.length;
    sh.getRange(row, 1).setValue(b.title).setFontWeight('bold').setFontSize(12).setFontColor('#1f3a5f');
    row++;
    sh.getRange(row, 1, 1, w).setValues([b.headers])
      .setFontWeight('bold').setBackground('#1f3a5f').setFontColor('#ffffff');
    const body = b.rows.length ? b.rows : [[ '데이터 없음' ].concat(new Array(w - 1).fill(''))];
    const rng = sh.getRange(row + 1, 1, body.length, w);
    rng.setValues(body);
    for (let i = 0; i < w; i++) {
      const f = b.fmtAll && i > 0 ? b.fmtAll : (b.fmt ? b.fmt[i] : null);
      if (f) sh.getRange(row + 1, i + 1, body.length, 1).setNumberFormat(f);
    }
    if (b.fmtAll && b.rows.length) {  // 비율표는 진할수록 높게
      const heat = sh.getRange(row + 1, 2, body.length, w - 1);
      sh.setConditionalFormatRules(sh.getConditionalFormatRules().concat(
        SpreadsheetApp.newConditionalFormatRule().setGradientMinpoint('#ffffff').setGradientMaxpoint('#7fa8e8').setRanges([heat]).build()));
    }
    if (b.chart && b.rows.length) {
      sh.insertChart(sh.newChart().setChartType(Charts.ChartType.BAR)
        .addRange(sh.getRange(row + 1, 1, body.length, 2))
        .setPosition(row, 8, 0, 0)
        .setOption('title', b.chart).setOption('legend', { position: 'none' })
        .setOption('colors', ['#2f6fde']).setOption('width', 460).setOption('height', Math.max(200, body.length * 30 + 60))
        .build());
    }
    row += body.length + 3;
  });
  sh.setColumnWidth(1, 230);
  sh.setColumnWidths(2, 5, 110);
  sh.setColumnWidth(5, 220);
  sh.setFrozenRows(2);
}
