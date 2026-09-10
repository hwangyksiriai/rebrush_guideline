/**
 * 리브러쉬 과즙 구강스프레이 캠페인 신청 폼 → 구글 시트 저장용 Apps Script.
 *
 * 설정 방법
 * 1. sheet.new 로 새 구글 시트를 만들고, 1행에 아래 헤더를 순서대로 입력:
 *    타임스탬프 | 캠페인 | 티어 | 고료 | 이름 | 인스타그램 | 휴대폰 | 이메일 | 우편번호 | 배송지 주소 | 요청사항 | 선택 일정 | 제품 선택
 *    (배송지 주소와 상세주소는 시트에 한 칸으로 합쳐서 들어감)
 *    (선택 일정: 참여 가능한 캠페인/일정을 고르는 폼에서만 값이 들어오며, 없는 폼은 빈 칸으로 저장됨)
 *    (제품 선택: 과즙 스프레이처럼 향/종류를 고르는 폼에서만 값이 들어오며, 없는 폼은 빈 칸으로 저장됨)
 *    (고료: 폼마다 고료가 다르게 설정돼 있어도 이 컬럼 값으로 한 시트 안에서 구분됨 — 별도 탭 분리 불필요)
 * 2. 시트 메뉴 [확장 프로그램] > [Apps Script] 를 열고, 기본 코드를 지운 뒤 이 파일 내용을 전체 붙여넣기.
 * 3. 아래 TARGET_SHEET_NAME 값을 실제 응답을 쌓을 탭 이름으로 맞춰주세요. (기본값: 응답시트)
 * 4. 우측 상단 [배포] > [새 배포] 클릭.
 *    - 유형: 웹 앱
 *    - 실행 계정: 나
 *    - 액세스 권한이 있는 사용자: 전체 (익명 사용자 포함)
 * 5. [배포] 클릭 후 나오는 웹 앱 URL(…/exec 로 끝남)을 복사.
 * 6. 각 페이지의 <form id="applyForm" ... action="https://script.google.com/macros/s/…/exec">
 *    부분의 URL을 방금 복사한 웹 앱 URL로 교체.
 */
var TARGET_SHEET_NAME = '응답시트';

// 시트가 "010-1234-5678", "06236" 같은 값을 숫자로 오인해 앞자리 0을 지우는 것을 막기 위해
// 휴대폰/우편번호 칸은 항상 텍스트 서식으로 먼저 지정한 뒤 값을 입력한다.
var PHONE_COL = 7;
var ZIPCODE_COL = 9;

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

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // 탭 이름을 직접 지정해서 항상 같은 탭에 쌓이도록 함 (getActiveSheet()는
  // 마지막으로 화면에 열려 있던 탭에 저장돼서, 고료별로 폼이 여러 개면
  // 신청이 엉뚱한 탭으로 흩어질 수 있음).
  var sheet = ss.getSheetByName(TARGET_SHEET_NAME) || ss.getActiveSheet();
  var p = e.parameter;
  var fullAddress = [p.address, p.addressDetail].filter(function (v) { return v; }).join(' ');

  var row = [
    new Date(),
    p.campaign || '',
    p.tier || '',
    p.fee || '',
    p.name || '',
    normalizeInstagram(p.instagram),
    normalizePhone(p.phone),
    p.email || '',
    p.zipcode || '',
    fullAddress,
    p.request || '',
    p.schedule || '',
    p.flavor || ''
  ];

  var targetRow = sheet.getLastRow() + 1;
  sheet.getRange(targetRow, PHONE_COL).setNumberFormat('@');
  sheet.getRange(targetRow, ZIPCODE_COL).setNumberFormat('@');
  sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}
