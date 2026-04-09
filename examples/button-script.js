function abrirGenerador() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetId = ss.getId();
  var gid = ss.getActiveSheet().getSheetId();

  var baseUrl = "WEB_APP_URL_DEMO";
  var url =
    baseUrl +
    "?id=" +
    encodeURIComponent(spreadsheetId) +
    "&gid=" +
    encodeURIComponent(gid);

  var html = HtmlService.createHtmlOutput(
    "<!DOCTYPE html>" +
      "<html>" +
      '<head><base target="_blank"></head>' +
      '<body style="font-family:Arial, sans-serif;padding:16px;">' +
      "<p><strong>Hacé clic para abrir el generador:</strong></p>" +
      '<p><a href="' +
      url +
      '" target="_blank" ' +
      'style="display:inline-block;background:#1a73e8;color:#fff;padding:10px 14px;text-decoration:none;border-radius:6px;font-weight:bold;">' +
      "🚀 ABRIR GENERADOR" +
      "</a></p>" +
      "</body></html>",
  ).setTitle("Generador");

  SpreadsheetApp.getUi().showSidebar(html);
}
