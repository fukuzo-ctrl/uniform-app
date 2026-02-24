function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  template.isAdmin = false;
  return template.evaluate().setTitle('UNIFORM BUILDER').addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function onOpen() { SpreadsheetApp.getUi().createMenu('★管理').addItem('起動', 'openApp').addToUi(); }
function openApp() { var html = HtmlService.createHtmlOutput('<html><script>window.open("' + ScriptApp.getService().getUrl() + '", "_blank");google.script.host.close();</script></html>').setWidth(300).setHeight(100); SpreadsheetApp.getUi().showModalDialog(html, '起動中...'); }

function getDesignLibraryBySport(s) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(s);
  if (!sheet) throw new Error('シート「' + s + '」が見つかりません');
  var data = sheet.getDataRange().getValues(), lib = {};
  for (var i = 1; i < data.length; i++) {
    if (!lib[data[i][1]]) lib[data[i][1]] = { variants: [] };
    lib[data[i][1]].variants.push({ id: data[i][0], collar: data[i][2], svg: data[i][3] });
  }
  return lib;
}

// 競技設定の保存（名前の幅 m_w を含む）
function saveSportSettings(d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('システム設定') || ss.insertSheet('システム設定');
  if (sheet.getLastRow() === 0) sheet.appendRow(['競技', 'Scale', 'bgUrl', 'n_size', 'n_x', 'n_y', 'm_size', 'm_x', 'm_y', 'm_w']);
  var rows = sheet.getDataRange().getValues(), f = -1;
  for (var i = 1; i < rows.length; i++) { if (rows[i][0] === d.sportName) { f = i + 1; break; } }
  var row = [d.sportName, d.scale, d.bgUrl, d.n_size, d.n_x, d.n_y, d.m_size, d.m_x, d.m_y, d.m_w];
  if (f !== -1) sheet.getRange(f, 1, 1, 10).setValues([row]);
  else sheet.appendRow(row);
  return "設定を保存しました";
}

// カラーマスターの取得・保存
function getColorPalette() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('カラー定義');
  if (!sheet) return null; // 初期値を使用
  var data = sheet.getDataRange().getValues();
  return data.slice(1).map(function(r){ return {n:r[0], c:r[1]}; });
}

function saveColorPalette(palette) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('カラー定義') || ss.insertSheet('カラー定義');
  sheet.clear();
  sheet.appendRow(['色名', 'カラーコード']);
  palette.forEach(function(p){ sheet.appendRow([p.n, p.c]); });
  return "カラーパレットを更新しました";
}

function getSportSettings() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('システム設定');
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues(), s = {};
  for (var i = 1; i < data.length; i++) s[data[i][0]] = { scale:data[i][1], bgUrl:data[i][2], n_size:data[i][3], n_x:data[i][4], n_y:data[i][5], m_size:data[i][6], m_x:data[i][7], m_y:data[i][8], m_w:data[i][9] };
  return s;
}

function saveOrder(d) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('注文一覧') || SpreadsheetApp.getActiveSpreadsheet().insertSheet('注文一覧');
  if (sheet.getLastRow() === 0) sheet.appendRow(["日時", "ID", "デザイン", "競技", "襟", "番号", "名前", "身頃1", "身頃2", "身頃3", "袖", "襟色", "番色", "名色"]);
  sheet.appendRow([new Date(), d.designId, d.designName, d.sportType, d.collarType, d.number, d.nameText, d.colorBody1, d.colorBody2, d.colorBody3, d.colorSleeve, d.colorCollar, d.colorNum, d.colorName]);
  return "SUCCESS";
}