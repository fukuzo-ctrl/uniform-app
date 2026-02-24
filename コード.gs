function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  // 管理者モード判定 (?admin=true でアクセス)
  template.isAdmin = (e && e.parameter && e.parameter.admin === 'true');
  return template.evaluate()
    .setTitle('UNIFORM BUILDER PRO')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('★デザイン管理').addItem('シミュレーター起動', 'openApp').addToUi();
}

function openApp() {
  var url = ScriptApp.getService().getUrl();
  var html = HtmlService.createHtmlOutput('<html><script>window.open("' + url + '", "_blank");google.script.host.close();</script></html>').setWidth(300).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, '起動中...');
}

// 競技別デザイン取得
function getDesignLibraryBySport(sportName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sportName);
  if (!sheet) throw new Error('シート「' + sportName + '」が見つかりません');
  var data = sheet.getDataRange().getValues(), lib = {};
  for (var i = 1; i < data.length; i++) {
    var id = data[i][0], name = data[i][1], collar = data[i][2], svg = data[i][3];
    if (!name || !svg) continue;
    if (!lib[name]) lib[name] = { variants: [] };
    lib[name].variants.push({ id: id, collar: collar, svg: svg });
  }
  return lib;
}

// 注文保存（色名で保存）
function saveOrder(d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('注文一覧') || ss.insertSheet('注文一覧');
  if (sheet.getLastRow() === 0) {
    var h = ["日時", "ID", "デザイン", "競技", "襟", "番号", "名前", "身頃色", "袖色", "襟色", "ライン1色", "ライン2色", "文字色"];
    sheet.appendRow(h);
    sheet.getRange(1,1,1,h.length).setBackground("#e2efda").setFontWeight("bold");
  }
  sheet.appendRow([new Date(), d.designId, d.designName, d.sportType, d.collarType, d.number, d.nameText, d.colorBody, d.colorSleeve, d.colorCollar, d.colorLine1, d.colorLine2, d.colorNum]);
  return "SUCCESS";
}