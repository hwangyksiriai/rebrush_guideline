/**
 * 리브러쉬 12월 포도듀오 올영 어워즈 캠페인 신청폼 → 구글 시트 저장용 Apps Script.
 *
 * 대상 시트 : https://docs.google.com/spreadsheets/d/1af9We7Wkr5UyCe201SUM4D8Qxfj6tA3dNj-rE_S-O-A/edit
 * 대상 탭   : 시트1
 *
 * 하는 일
 *  - 휴대폰 · 우편번호를 텍스트로 저장해 앞자리 0 이 지워지지 않게 한다.
 *  - 인스타그램은 @handle / handle / instagram.com/handle 어떻게 넣어도
 *    https://www.instagram.com/handle/ 형태로 정리한다.
 *  - 주소 + 상세주소를 한 칸에 합쳐 넣는다.
 *  - 한 탭 안에서 고료별(A → E Type)로 묶이도록, 새 신청을 같은 티어의 마지막 줄 바로
 *    아래에 끼워 넣는다. 티어 칸은 티어마다 다른 색으로 칠해 경계가 눈에 띄게 한다.
 *
 * ─────────────────────────────────────────────────────────────
 * 설정 방법 (9월 스크립트와는 별개의 새 프로젝트로 만드세요)
 *
 * 1. 위 시트를 열고 [확장 프로그램] > [Apps Script].
 * 2. 기본 코드(myFunction)를 전부 지우고 이 파일 내용을 통째로 붙여넣은 뒤 저장.
 * 3. 함수 목록에서 testAllTiers 를 고르고 [▶ 실행] → 권한 승인.
 *    시트1 에 TEST 5행이 A~E 순서로 생기면 정상입니다. 확인 후 TEST 행은 지워주세요.
 * 4. [배포] > [새 배포]
 *    - 유형 선택(톱니바퀴) : 웹 앱          ← '라이브러리' 가 아니라 웹 앱
 *    - 실행 계정          : 나
 *    - 액세스 권한        : 모든 사용자     ← '나만' 이면 신청이 안 들어옵니다
 * 5. 배포 후 나오는 '웹 앱 URL'(…/exec 로 끝남)을 전달해 주세요.
 *    12월 페이지(december/*) 의 <form action="..."> 에 넣으면 연결됩니다.
 *
 * ※ 코드를 고친 뒤에는 [배포] > [배포 관리] > 연필 > 버전 '새 버전' > [배포].
 *    이렇게 하면 URL 이 유지됩니다. [새 배포] 를 하면 URL 이 바뀝니다.
 * ─────────────────────────────────────────────────────────────
 */

var SPREADSHEET_ID = '1af9We7Wkr5UyCe201SUM4D8Qxfj6tA3dNj-rE_S-O-A';
var TARGET_SHEET_NAME = '시트1';

var HEADERS = [
  '타임스탬프', '티어', '고료', '이름', '인스타그램',
  '휴대폰', '이메일', '우편번호', '배송지 주소', '요청사항', '캠페인'
];
var TIER_COL = HEADERS.indexOf('티어') + 1;
var PHONE_COL = HEADERS.indexOf('휴대폰') + 1;
var ZIPCODE_COL = HEADERS.indexOf('우편번호') + 1;

// 시트 안에서 묶이는 순서와 티어 칸 색
var TIER_ORDER = ['A Type', 'B Type', 'C Type', 'D Type', 'E Type'];
var TIER_COLOR = {
  'A Type': '#EDE4F8',
  'B Type': '#E2F0FA',
  'C Type': '#E6F4E1',
  'D Type': '#FFF3D6',
  'E Type': '#FBE4E4'
};

// "+82 10-1234-5678", "01012345678" → "010-1234-5678"
function normalizePhone(v) {
  var d = String(v || '').replace(/[^0-9]/g, '').replace(/^0*82(?=\d)/, '0').slice(0, 11);
  if (d.length === 11) return d.slice(0, 3) + '-' + d.slice(3, 7) + '-' + d.slice(7);
  if (d.length === 10) return d.slice(0, 3) + '-' + d.slice(3, 6) + '-' + d.slice(6);
  return String(v || '').trim();
}

// 숫자만 남기고, 4자리로 들어온 우편번호는 앞에 0 을 채워 5자리로
function normalizeZip(v) {
  var d = String(v || '').replace(/[^0-9]/g, '');
  if (d.length === 4) d = '0' + d;
  return d || String(v || '').trim();
}

// @handle · handle · instagram.com/handle · https://www.instagram.com/handle/?igsh=...
//   → https://www.instagram.com/handle/
function normalizeInstagram(v) {
  var s = String(v || '').trim();
  if (!s) return '';
  var m = s.match(/instagram\.com\/@?([A-Za-z0-9._]+)/i);
  var handle = m ? m[1] : s.replace(/^@+/, '').split(/[\/?#\s]/)[0];
  if (!/^[A-Za-z0-9._]+$/.test(handle)) return s; // 형식이 이상하면 원문 그대로 둔다
  return 'https://www.instagram.com/' + handle + '/';
}

// 티어 순서상 몇 번째인지. 목록에 없는 값(TEST 등)은 맨 뒤로.
function tierRank_(tier) {
  var i = TIER_ORDER.indexOf(String(tier || '').trim());
  return i < 0 ? TIER_ORDER.length : i;
}

/**
 * 새 행을 넣을 시트 행 번호를 구한다.
 * tiers : 2행부터 마지막 행까지의 티어 값 배열 (헤더 제외)
 * 같은 티어 이하 순위를 가진 마지막 행 바로 아래 → 티어끼리 묶이고 A→E 순서가 유지된다.
 */
function findInsertRow_(tiers, tier) {
  var rank = tierRank_(tier);
  var after = 1; // 헤더 바로 아래부터
  for (var i = 0; i < tiers.length; i++) {
    if (tierRank_(tiers[i]) <= rank) after = i + 2;
  }
  return after + 1;
}

function getSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(TARGET_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
      .setFontWeight('bold').setBackground('#2E2A3A').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doPost(e) {
  // 동시에 여러 명이 신청해도 같은 줄에 덮어쓰지 않도록 잠근다.
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sheet = getSheet_();
    var p = e.parameter;
    var tier = String(p.tier || '').trim();
    var fullAddress = [p.address, p.addressDetail]
      .map(function (v) { return String(v || '').trim(); })
      .filter(function (v) { return v; })
      .join(' ');

    var row = [
      new Date(),
      tier,
      p.fee || '',
      p.name || '',
      normalizeInstagram(p.instagram),
      normalizePhone(p.phone),
      p.email || '',
      normalizeZip(p.zipcode),
      fullAddress,
      p.request || '',
      p.campaign || ''
    ];

    var lastRow = sheet.getLastRow();
    var tiers = lastRow > 1
      ? sheet.getRange(2, TIER_COL, lastRow - 1, 1).getValues().map(function (r) { return r[0]; })
      : [];
    var target = findInsertRow_(tiers, tier);

    if (target <= lastRow) {
      sheet.insertRowBefore(target);
    }
    // 텍스트 서식을 먼저 지정해야 앞자리 0 이 살아남는다
    sheet.getRange(target, PHONE_COL).setNumberFormat('@');
    sheet.getRange(target, ZIPCODE_COL).setNumberFormat('@');
    var range = sheet.getRange(target, 1, 1, row.length);
    range.setBackground(null).setFontWeight('normal').setFontColor('#000000');
    range.setValues([row]);
    sheet.getRange(target, 1).setNumberFormat('yyyy-mm-dd hh:mm');
    sheet.getRange(target, TIER_COL).setBackground(TIER_COLOR[tier] || null).setFontWeight('bold');

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 고료 5종 테스트. ★ 배포 전에 이걸 먼저 실행하세요.
 * 일부러 E → A 역순으로 넣습니다. 시트에서 A → E 순서로 정렬돼 들어가면 묶기도 정상입니다.
 * 확인 후 TEST 행은 지워주세요.
 */
function testAllTiers() {
  var cases = [
    ['E Type', '고료 조정'],
    ['C Type', '500,000원'],
    ['A Type', '200,000원'],
    ['D Type', '1,000,000원'],
    ['B Type', '400,000원']
  ];
  cases.forEach(function (c, i) {
    doPost({ parameter: {
      campaign: '리브러쉬 포도듀오 어워즈 (12월)',
      tier: c[0],
      fee: c[1],
      name: 'TEST ' + c[0],
      instagram: i % 2 ? '@test_handle' + i : 'instagram.com/test.handle' + i,
      phone: '0101234567' + i,
      email: 'test' + i + '@test.com',
      zipcode: '03900',
      address: '서울 마포구 가양대로 1',
      addressDetail: (101 + i) + '호',
      request: '고료별 테스트'
    }});
  });
  Logger.log('5종 저장 완료 / 데이터 ' + Math.max(0, getSheet_().getLastRow() - 1) + '행');
}

// 웹 앱 URL 을 브라우저로 열면 시트 연결 상태를 보여준다.
//   ok / 시트1 / 데이터 N행  -> 정상
function doGet() {
  try {
    var sheet = getSheet_();
    return ContentService.createTextOutput(
      'ok / ' + sheet.getName() + ' / 데이터 ' + Math.max(0, sheet.getLastRow() - 1) + '행');
  } catch (err) {
    return ContentService.createTextOutput('error: ' + err.message);
  }
}
