/**
 * 리브러쉬 과즙 구강스프레이 캠페인 신청폼 → 구글 시트 저장용 Apps Script.
 *
 * 대상 시트: [응답]리브러쉬 26년 9-10월
 * https://docs.google.com/spreadsheets/d/1dJGNZ9LAzlC_kldI8DzJv8z1TuWgmMBKwJ9XMJMDtw8/edit
 *
 * ─────────────────────────────────────────────────────────────
 * 설정 방법
 *
 * 1. Apps Script 편집기를 엽니다. 시트의 [확장 프로그램] > [Apps Script] 로 열어도 되고,
 *    script.google.com 에서 만든 독립 프로젝트여도 됩니다.
 *    (위 SPREADSHEET_ID 로 시트를 직접 열기 때문에 어느 쪽이든 동작합니다.)
 * 2. 기본 코드(myFunction)를 전부 지우고 이 파일 내용을 통째로 붙여넣습니다.
 * 3. 저장(디스크 아이콘)합니다.
 * 4. 우측 상단 [배포] > [새 배포] 를 클릭합니다.
 *    - 유형 선택(톱니바퀴) : 웹 앱   ← '라이브러리' 를 고르면 신청이 들어오지 않습니다
 *    - 설명               : 아무거나 (예: 리브러쉬 신청폼)
 *    - 실행 계정          : 나
 *    - 액세스 권한        : 모든 사용자   ← 반드시 이걸로. '나만'이면 신청이 안 들어옵니다
 * 5. [배포] 를 누르면 권한 승인 창이 뜹니다. 승인해 주세요.
 *    ("이 앱은 확인되지 않았습니다" 가 뜨면 [고급] > [OOO(안전하지 않음)으로 이동])
 * 6. 배포 후 나오는 '웹 앱 URL'(…/exec 로 끝남)을 복사해서 전달해 주세요.
 *    그 주소를 a~d-type 네 페이지의 <form action="..."> 에 넣어야 연결됩니다.
 *
 * ※ 헤더는 첫 신청이 들어올 때 자동으로 만들어집니다. 미리 입력하지 않으셔도 됩니다.
 * ※ 코드를 고친 뒤에는 [배포] > [배포 관리] > 연필 아이콘 > 버전 '새 버전' > [배포]
 *    를 해야 반영됩니다. 이때 URL 은 그대로 유지됩니다.
 * ─────────────────────────────────────────────────────────────
 */

// 응답을 쌓을 스프레드시트 ID.
// getActiveSpreadsheet() 를 쓰면 '시트에 붙은 스크립트'일 때만 동작하고,
// script.google.com 에서 새로 만든 독립 스크립트에서는 null 이 되어 저장이 실패한다.
// ID 로 직접 열면 스크립트가 어디에 있든 같은 시트에 쌓인다.
var SPREADSHEET_ID = '1dJGNZ9LAzlC_kldI8DzJv8z1TuWgmMBKwJ9XMJMDtw8';

// 응답을 쌓을 탭 이름
var TARGET_SHEET_NAME = '시트1';

// 컬럼 순서. 고료·티어가 앞쪽에 있어 타입별로 정렬·필터하기 쉽습니다.
var HEADERS = [
  '타임스탬프', '캠페인', '티어', '고료', '제품 선택',
  '이름', '인스타그램', '휴대폰', '이메일',
  '우편번호', '배송지 주소', '요청사항'
];

// 시트가 "010-1234-5678", "06236" 같은 값을 숫자로 오인해 앞자리 0을 지우는 것을 막기 위해
// 휴대폰/우편번호 칸은 항상 텍스트 서식으로 먼저 지정한 뒤 값을 입력한다.
var PHONE_COL = HEADERS.indexOf('휴대폰') + 1;
var ZIPCODE_COL = HEADERS.indexOf('우편번호') + 1;

// 자동완성으로 "+82 10-0000-0000" 형태가 들어와도 010-0000-0000 형태로 정규화
function normalizePhone(v) {
  var d = String(v || '').replace(/[^0-9]/g, '').replace(/^0*82(?=\d)/, '0').slice(0, 11);
  if (d.length === 11) return d.slice(0, 3) + '-' + d.slice(3, 7) + '-' + d.slice(7);
  if (d.length === 10) return d.slice(0, 3) + '-' + d.slice(3, 6) + '-' + d.slice(6);
  return v || '';
}

function normalizeInstagram(v) {
  var s = String(v || '').trim();
  if (!s) return '';
  if (/^https?:\/\//i.test(s)) return s;
  s = s.replace(/^@/, '').replace(/^www\./i, '');
  if (/^instagram\.com/i.test(s)) return 'https://' + s;
  return 'https://instagram.com/' + s.replace(/^\/+/, '');
}

// 탭을 찾고, 비어 있으면 헤더를 먼저 깔아둔다.
function getSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(TARGET_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(TARGET_SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
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
    var fullAddress = [p.address, p.addressDetail].filter(function (v) { return v; }).join(' ');

    var row = [
      new Date(),
      p.campaign || '',
      p.tier || '',
      p.fee || '',
      p.flavor || '',
      p.name || '',
      normalizeInstagram(p.instagram),
      normalizePhone(p.phone),
      p.email || '',
      p.zipcode || '',
      fullAddress,
      p.request || ''
    ];

    var targetRow = sheet.getLastRow() + 1;
    sheet.getRange(targetRow, PHONE_COL).setNumberFormat('@');
    sheet.getRange(targetRow, ZIPCODE_COL).setNumberFormat('@');
    sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 고료 타입 6종을 한 번에 저장해 보는 테스트. ★ 이걸 실행하세요.
 *
 * 상단 함수 목록에서 testAllTiers 를 고르고 [▶ 실행].
 * 실제 저장 경로(doPost)를 그대로 지나므로, 신청폼으로 들어올 때와 같은 값이 쌓입니다.
 *
 *   시트에 TEST 6행이 생김 -> 스크립트와 시트는 정상.
 *                            신청이 안 들어온다면 원인은 배포나 웹페이지 쪽입니다.
 *   빨간 오류가 뜸         -> 그 메시지가 원인입니다. 그대로 알려주세요.
 *
 * 확인 후 TEST 행 6개는 지우시면 됩니다.
 */
function testAllTiers() {
  var cases = [
    ['A Type', '50,000원',  '🍇 머루포도 (PURE)'],
    ['B Type', '100,000원', '🍏 청사과 (COOLING)'],
    ['C Type', '150,000원', '🍑 납작복숭아 (CLEAN)'],
    ['D Type', '200,000원', '🍇 머루포도 (PURE)'],
    ['E Type', '300,000원', '🍏 청사과 (COOLING)'],
    ['F Type', '고료 조정',  '🍑 납작복숭아 (CLEAN)']
  ];
  cases.forEach(function (c, i) {
    doPost({ parameter: {
      campaign: '리브러쉬 과즙 구강스프레이',
      tier: c[0],
      fee: c[1],
      flavor: c[2],
      name: 'TEST ' + c[0],
      instagram: '@test' + (i + 1),
      phone: '0100000000' + (i + 1),
      email: 'test' + (i + 1) + '@test.com',
      zipcode: '03900',
      address: '서울 마포구 가양대로 1',
      addressDetail: (101 + i) + '호',
      request: '고료별 테스트'
    }});
  });
  Logger.log('6종 저장 완료 / 데이터 ' + Math.max(0, getSheet_().getLastRow() - 1) + '행');
}

/**
 * 편집기에서 바로 돌려보는 자가진단.
 *
 * 상단 함수 목록에서 testSave 를 고르고 [▶ 실행] 을 누르세요.
 * 웹페이지·배포와 무관하게 '스크립트가 이 시트에 쓸 수 있는지'만 확인합니다.
 *
 *   시트에 TEST 행이 생김  -> 스크립트와 시트는 정상.
 *                            문제는 배포(웹 앱/버전) 아니면 웹페이지 쪽입니다.
 *   실행 로그에 빨간 오류  -> 그 메시지가 원인입니다. 그대로 알려주세요.
 *
 * 확인한 뒤 시트에서 TEST 행은 지우시면 됩니다.
 */
function testSave() {
  var res = doPost({
    parameter: {
      campaign: '리브러쉬 과즙 구강스프레이',
      tier: 'TEST',
      fee: '0원',
      flavor: '\uD83C\uDF4F 청사과 (COOLING)',
      name: '편집기테스트',
      instagram: '@test',
      phone: '01000000000',
      email: 'test@test.com',
      zipcode: '03900',
      address: '서울 마포구 가양대로 1',
      addressDetail: '101호',
      request: '편집기에서 실행한 테스트'
    }
  });
  Logger.log('결과: ' + res.getContent());
  Logger.log('시트: ' + getSheet_().getName() + ' / 데이터 ' +
             Math.max(0, getSheet_().getLastRow() - 1) + '행');
}

// 배포가 살아 있는지 브라우저로 확인할 때 쓰는 용도.
// 웹 앱 URL 을 주소창에 그대로 열면 시트에 실제로 닿는지까지 보여준다.
//   ok / 시트1 / 데이터 N행   -> 정상
//   error: ...                -> 시트를 못 여는 상태 (ID 나 권한 확인)
function doGet() {
  try {
    var sheet = getSheet_();
    var rows = Math.max(0, sheet.getLastRow() - 1);
    return ContentService.createTextOutput(
      'ok / ' + sheet.getName() + ' / 데이터 ' + rows + '행');
  } catch (err) {
    return ContentService.createTextOutput('error: ' + err.message);
  }
}
