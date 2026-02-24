function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  template.isAdmin = (e && e.parameter && e.parameter.admin === 'true');
  return template.evaluate()
    .setTitle('UNIFORM BUILDER')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('★管理').addItem('起動', 'openApp').addToUi();
}

function openApp() {
  var url = ScriptApp.getService().getUrl();
  var html = HtmlService.createHtmlOutput('<html><script>window.open("' + url + '", "_blank");google.script.host.close();</script></html>').setWidth(300).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, '起動中...');
}

function getDesignLibraryBySport(sportName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sportName);
  if (!sheet) throw new Error('シート「' + sportName + '」が見つかりません');
  var data = sheet.getDataRange().getValues(), lib = {};
  for (var i = 1; i < data.length; i++) {
    if (!lib[data[i][1]]) lib[data[i][1]] = { variants: [] };
    lib[data[i][1]].variants.push({ id: data[i][0], collar: data[i][2], svg: data[i][3] });
  }
  return lib;
}

function saveOrder(d) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('注文一覧') || SpreadsheetApp.getActiveSpreadsheet().insertSheet('注文一覧');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["日時", "ID", "デザイン", "競技", "襟", "番号", "名前", "身頃", "袖", "襟色", "ライン1", "ライン2", "番色", "名色"]);
  }
  sheet.appendRow([new Date(), d.designId, d.designName, d.sportType, d.collarType, d.number, d.nameText, d.colorBody, d.colorSleeve, d.colorCollar, d.colorLine1, d.colorLine2, d.colorNum, d.colorName]);
  return "SUCCESS";
}