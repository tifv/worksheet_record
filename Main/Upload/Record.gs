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
  "id"             : {width: 175},
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
  var uploads = DataTable.init(uploads_sheet, required_keys, 3);
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
  var cfrules = ConditionalFormatting.RuleList.load(uploads_sheet);
  add_helper_cf(uploads, cfrules);
  cfrules.save(uploads_sheet);
}

function add_helper_cf(uploads, cfrules) {
  var first_R1 = "R" + uploads.first_row;
  var id_C1 = "C" + uploads.key_columns.get("id");
  var group_C1 = "C" + uploads.key_columns.get("group");
  var date_C1 = "C" + uploads.key_columns.get("date");
  var request_C1 = "C" + uploads.key_columns.get("request");
  var stable_C1 = "C" + uploads.key_columns.get("stable_pdf");
  var pdf_C1 = "C" + uploads.key_columns.get("initial_pdf");
  var src_C1 = "C" + uploads.key_columns.get("initial_src");
  var filename_C1 = "C" + uploads.key_columns.get("filename");
  var title_cfrange = ConditionalFormatting.Range
    .from_range(uploads.get_range("max", "title"));
  var pdf_cfrange = ConditionalFormatting.Range
    .from_range(uploads.get_range("max", "initial_pdf"));
  var filename_cfrange = ConditionalFormatting.Range
    .from_range(uploads.get_range("max", "filename"));
  cfrules.insert({
    type: "boolean",
    condition: {
      type: SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA,
      values: ["".concat(
        "=and(not(isblank(", "R[0]", id_C1, ")),countifs(",
          "R[0]", group_C1, ":", group_C1, ",", "R[0]", group_C1, ",",
          "R[0]",  date_C1, ":",  date_C1, ",", "R[0]",  date_C1,
        ")>1)"
      )],
    },
    effect: {background: "#ffaaaa"},
    ranges: [title_cfrange],
  });
  cfrules.insert({
    type: "boolean",
    condition: {
      type: SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA,
      values: ["".concat(
        "=and(not(isblank(", "R[0]", id_C1, ")),countifs(",
          first_R1, group_C1, ":", "R[0]", group_C1, ",", "R[0]", group_C1, ",",
          first_R1,  date_C1, ":", "R[0]",  date_C1, ",", "R[0]",  date_C1,
        ")>1)"
      )],
    },
    effect: {background: "#aaffaa"},
    ranges: [title_cfrange],
  });
  cfrules.insert({
    type: "boolean",
    condition: {
      type: SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA,
      values: ["".concat(
        "=or(", "R[0]", request_C1, "=\"stabilize\",not(isblank(", "R[0]", stable_C1, ")))"
      )],
    },
    effect: {background: "#aaaaff"},
    ranges: [title_cfrange],
  });
  cfrules.insert({
    type: "boolean",
    condition: {
      type: SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA,
      values: ["".concat(
        "=and(",
          "not(isblank(", "R[0]", pdf_C1, ")),",
          "or(",
            "isblank(", "R[0]", src_C1, "),",
            "iferror(find(\"overleaf.com\",", "R[0]", src_C1, ")>=0,FALSE)",
          ")",
        ")"
      )],
    },
    effect: {background: "#ffaaaa"},
    ranges: [pdf_cfrange],
  });
  cfrules.insert({
    type: "boolean",
    condition: {
      type: SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA,
      values: ["".concat(
        "=and(not(isblank(", "R[0]", filename_C1, ")),",
          "countif(", first_R1, filename_C1, ":", filename_C1, ",", "R[0]", filename_C1, ")>1",
        ")"
      )],
    },
    effect: {background: "#ffaaaa"},
    ranges: [filename_cfrange],
  });
}

function recreate_category_cf(spreadsheet) {
  if (spreadsheet == null)
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var uploads = get(spreadsheet);
  var uploads_sheet = uploads.sheet;
  var cfrules = ConditionalFormatting.RuleList.load(uploads_sheet);
  var cfrange = ConditionalFormatting.Range
    .from_range(uploads.get_range("max", "category"));
  cfrules.remove({
    type: "boolean",
    condition: {
      match: (cfcondition) => (
        cfcondition.type == SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA &&
        cfcondition.values[0].startsWith("=exact")
      ) },
    locations: [cfrange],
  });
  var categories = Categories.get(spreadsheet);
  var category_R1C1 = "R[0]C" + uploads.key_columns.get("category");
  for (let code in categories) {
    var category = categories[code];
    if (category.color == null)
      continue;
    var color = HSL.to_hex(category.color);
    cfrules.insert({
      type: "boolean",
      condition: {
        type: SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA,
        values: ["".concat(
          "=exact(",
            "\"", code.replace('"', '""'), "\",",
            category_R1C1,
          ")" )],
      },
      effect: {background: color},
      ranges: [cfrange],
    });
  }
  cfrules.save(uploads_sheet);
}

function recreate_group_cf(spreadsheet) {
  if (spreadsheet == null)
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var uploads = get(spreadsheet);
  var uploads_sheet = uploads.sheet;
  var cfrules = ConditionalFormatting.RuleList.load(uploads_sheet);
  var cfrange = ConditionalFormatting.Range
    .from_range(uploads.get_range("max", "group"));
  cfrules.remove({
    type: "boolean",
    condition: {
      match: (cfcondition) => (
        cfcondition.type == SpreadsheetApp.BooleanCriteria.TEXT_EQUAL_TO
      ) },
    locations: [cfrange],
  });
  for (let group of StudyGroup.list(spreadsheet)) {
    var name = group.name;
    var color = group.sheet.getTabColor();
    if (color == null)
      continue;
    // XXX if the color is too dark, switch to white text
    cfrules.insert({
      type: "boolean",
      condition: {
        type: SpreadsheetApp.BooleanCriteria.TEXT_EQUAL_TO,
        values: [name] },
      effect: {background: color},
      ranges: [cfrange],
    });
  }
  cfrules.save(uploads_sheet);
}

return {
  exists: exists, get: get,
  create: create,
  recreate_group_cf: recreate_group_cf,
  recreate_category_cf: recreate_category_cf,
};
}(); // end UploadRecord namespace

function upload_record_create() {
  UploadRecord.create();
}

function upload_record_recreate_cf() {
  UploadRecord.recreate_group_cf(SpreadsheetApp.getActiveSpreadsheet());
  UploadRecord.recreate_category_cf(SpreadsheetApp.getActiveSpreadsheet());
}

