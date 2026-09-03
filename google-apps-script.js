// =========================================================================================
// [금옥여자고등학교 대입 모의면접 신청 구글 시트 연동 스크립트]
// 
// 사용 방법:
// 1. 새 구글 스프레드시트(Google Sheets)를 하나 만듭니다.
//    (예: "2026학년도 금옥여고 수시 모의면접 신청 명단")
// 2. 상단 메뉴에서 [확장 프로그램] -> [Apps Script]를 클릭합니다.
// 3. 기존의 코드를 모두 지우고 이 파일의 내용을 복사해서 붙여넣습니다.
// 4. 오른쪽 상단 [배포] -> [새 배포] 클릭
// 5. 유형 선택(톱니바퀴) -> [웹 앱] 선택
//    - 설명: 금옥여고 모의면접 접수 API
//    - 다음 사용자 권한으로 실행: "나(선생님 계정)"
//    - 액세스 권한이 있는 사용자: "모든 사용자(Anyone)"  <-- ★ 필수! (학생이 구글 로그인 없이 즉시 제출 가능)
// 6. [배포] 버튼 클릭 후 [액세스 승인] 절차 진행
// 7. 발급된 "웹 앱 URL"(https://script.google.com/macros/s/.../exec)을 복사하여
//    index.html 상단의 GOOGLE_SCRIPT_URL 에 넣어주시면 됩니다.
// =========================================================================================

function doPost(e) {
  var lock = LockService.getScriptLock();
  // 동시 제출 충돌 방지 (최대 10초 대기)
  lock.tryLock(10000);

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // 시트가 비어있는 경우 첫 행 헤더 자동 생성
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "접수일시",
        "신청ID",
        "반",
        "번호",
        "이름",
        "지망순번",
        "지원 대학교",
        "지원 학과",
        "전형명",
        "계열",
        "면접 일정(단일/기간)"
      ]);

      // 헤더 스타일링 (에메랄드/그린 테마 및 볼드 처리)
      var headerRange = sheet.getRange(1, 1, 1, 11);
      headerRange.setBackground("#E6F4EA");
      headerRange.setFontColor("#137333");
      headerRange.setFontWeight("bold");
      headerRange.setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
    }

    var requestData = JSON.parse(e.postData.contents);
    var timestamp = new Date();
    var submissionId = "KM-" + Utilities.formatDate(timestamp, "Asia/Seoul", "yyyyMMdd-HHmmss") + "-" + Math.floor(Math.random() * 899 + 100);
    var formattedDate = Utilities.formatDate(timestamp, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");

    var studentClass = requestData.studentClass;
    var studentNumber = requestData.studentNumber;
    var studentName = requestData.studentName;
    var applications = requestData.applications || [];

    var rowsToAdd = [];

    // 학생이 신청한 지원 대학별로 1개 행(Row)씩 깔끔하게 정규화하여 저장
    for (var i = 0; i < applications.length; i++) {
      var app = applications[i];
      rowsToAdd.push([
        formattedDate,
        submissionId,
        studentClass + "반",
        studentNumber + "번",
        studentName,
        "지원 " + (i + 1) + "지망",
        app.school || "",
        app.department || "",
        app.admissionType || "",
        app.track || "",
        app.interviewDate || ""
      ]);
    }

    if (rowsToAdd.length > 0) {
      var startRow = sheet.getLastRow() + 1;
      var numRows = rowsToAdd.length;
      var numCols = rowsToAdd[0].length;
      sheet.getRange(startRow, 1, numRows, numCols).setValues(rowsToAdd);
      
      // 주요 열(일시, ID, 반, 번호, 이름, 순번, 계열, 면접일정) 가운데 정렬
      sheet.getRange(startRow, 1, numRows, 6).setHorizontalAlignment("center");
      sheet.getRange(startRow, 10, numRows, 2).setHorizontalAlignment("center");
    }

    return ContentService.createTextOutput(JSON.stringify({
      result: "success",
      submissionId: submissionId,
      count: applications.length
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      result: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return ContentService.createTextOutput("금옥여자고등학교 모의면접 접수 API 정상 가동 중").setMimeType(ContentService.MimeType.TEXT);
}
