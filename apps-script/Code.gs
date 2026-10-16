// 학급 수시 달력 만들기 — 구글 Apps Script 웹 앱
// 화면(Index.html)을 보여 주기만 합니다. 올린 엑셀은 브라우저 안에서만 읽고 이 스크립트로 보내지 않습니다.
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('학급 수시 달력 만들기')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
