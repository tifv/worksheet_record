var UploadRecord = function() { // begin namespace

const sheet_name = "uploads";

const required_keys = [
  "group", "category", "title", "date", "uploader", "author",
  "id",
  "pdf", "src", "initial_pdf", "initial_src",
  "stable_pdf", "stable_src", "filename",
  "archive_pdf", "archive_src", "archive_target", "archive_mtime",
  "request", "argument", "status" ];
const key_format = {
  "group"    : {width:  50, name: "группа", frozen: true},
  "category" : {width:  50, name: "катег.", frozen: true},
  "title"    : {width: 300, name: "заголовок", frozen: true},
  "date"     : {width: 100, name: "дата", frozen: true},
  "uploader" : {width: 150, hidden: true},
  "author"   : {width: 150, name: "автор", frozen: true},
  "id"             : {width: 175, hidden: true},
  "pdf"            : {width: 175},
  "src"            : {width: 175, hidden: true},
  "initial_pdf"    : {width: 175},
  "initial_src"    : {width: 175, hidden: true},
  "stable_pdf"     : {width: 175},
  "stable_src"     : {width: 175, hidden: true},
  "filename"       : {width: 175},
  "archive_pdf"    : {width: 175},
  "archive_src"    : {width: 175, hidden: true},
  "archive_target" : {width: 175},
  "archive_mtime"  : {width: 175},
  "request"  : {width: 100},
  "argument" : {width: 100},
  "status"   : {width: 100},
}

function exists(spreadsheet) {
  if (spreadsheet == null)
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var uploads_sheet = spreadsheet.getSheetByName(sheet_name);
  return (uploads_sheet != null);
}

function get(spreadsheet, mode = "full") {
  if (spreadsheet == null)
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var uploads_sheet = spreadsheet.getSheetByName(sheet_name);
  if (uploads_sheet == null)
    return null;
  return new DataTable(uploads_sheet, {required_keys: required_keys, name: sheet_name, mode: mode});
}

function create(spreadsheet) {
  if (spreadsheet == null)
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var uploads_sheet = spreadsheet.insertSheet(sheet_name);
  DataTable.init(uploads_sheet, required_keys, 3);
  uploads_sheet.getRange(1, 1, uploads_sheet.getMaxRows(), uploads_sheet.getMaxColumns())
    .setVerticalAlignment("middle")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  uploads_sheet.getRange("A1:1")
    .setFontSize(8)
    .setFontFamily("monospace")
    .protect().setWarningOnly(true);
  uploads_sheet.getRange("A2:2")
    .setNumberFormat('@STRING@')
    .setFontWeight('bold')
    .setFontFamily("Times New Roman,serif")
    .setFontSize(12);
  uploads_sheet.getRange("A2")
    .setValue("Реестр загрузок")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
  uploads_sheet.getRange("A3:3")
    .setNumberFormat('@STRING@')
    .setFontWeight('bold')
    .setFontFamily("Times New Roman,serif")
    .setVerticalAlignment("bottom");
  uploads_sheet.getRange(3, 1, 1, required_keys.length)
    .setValues([required_keys.map(key => (key_format[key] || {name: "???"}).name || key)]);
  var frozen_columns = 0;
  for (let i = 0; i < required_keys.length; ++i) {
    let column = i + 1;
    let key = required_keys[i];
    let format = key_format[key] || {};
    if (format.width != null)
      uploads_sheet.setColumnWidth(column, format.width);
    if (format.hidden)
      uploads_sheet.hideColumns(column);
    if (format.frozen)
      frozen_columns = frozen_columns > column ? frozen_columns : column;
  }
  uploads_sheet.setFrozenColumns(frozen_columns);
  // XXX add more formatting
  // XXX color categories
  // XXX color titles when pairs group+date coincide
  // XXX color filenames when they coincide
  // XXX color groups
}

function recreate_group_cf(spreadsheet) {
  if (spreadsheet == null)
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var uploads = get(spreadsheet);
  var uploads_sheet = uploads.sheet;
  // XXX remove any cf rules that were in this range before.
  // XXX switch to CFormatting
  // XXX if the color is too dark, switch to white text
  var cf_rules = uploads_sheet.getConditionalFormatRules();
  var group_range = uploads.get_range(null, "group");
  for (let group of StudyGroup.list(spreadsheet)) {
    var name = group.name;
    var color = group.sheet.getTabColor();
    if (color == null)
      continue;
    cf_rules.push(SpreadsheetApp.newConditionalFormatRule()
      .setRanges([group_range])
      .whenTextEqualTo(name)
      .setBackground(color)
      .build());
  }
  uploads_sheet.setConditionalFormatRules(cf_rules)
}

return {
  exists: exists, get: get,
  create: create,
  recreate_group_cf: recreate_group_cf, };
}(); // end UploadRecord namespace


