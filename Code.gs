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

/** GitHub Pages(index.html)에서 보내는 제보 받기 */
function doPost(e) {
  let res;
  try {
    res = submitReel(JSON.parse(e.postData.contents));
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
