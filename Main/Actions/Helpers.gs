var ActionHelpers = function() { // begin namespace

function get_active_sheet(spreadsheet, sheet) {
  if (sheet == null) {
    if (spreadsheet == null) {
      spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    }
    sheet = spreadsheet.getActiveSheet();
  } else {
    if (spreadsheet == null) {
      spreadsheet = sheet.getParent();
    }
  }
  return [spreadsheet, sheet];
}

function get_active_range(spreadsheet, sheet, range) {
  if (range == null) {
    [spreadsheet, sheet] = get_active_sheet(spreadsheet, sheet);
    range = sheet.getActiveRange();
  } else {
    if (sheet == null)
      sheet = range.getSheet();
    if (spreadsheet == null)
      spreadsheet = sheet.getParent();
  }
  return [spreadsheet, sheet, range];
}

function get_active_group(spreadsheet, sheet) {
  [spreadsheet, sheet] = get_active_sheet(spreadsheet, sheet);
  try {
    var group = new StudyGroup(sheet);
    group.check();
    return group;
  } catch (error) {
    report_error(error);
//    report_error(
//      "Не удалось определить учебную группу. " +
//      "Выберите вкладку таблицы, соответствующую группе." );
    return null;
  }
}

function get_active_worksheet(spreadsheet, sheet, range) {
  [spreadsheet, sheet, range] = get_active_range(spreadsheet, sheet, range);
  var group = get_active_group(spreadsheet, sheet);
  if (group == null)
    return null;
  try {
    var worksheet = Worksheet.surrounding(group, range);
    return worksheet;
  } catch (error) {
    report_error(error);
//    report_error(
//      "Не удалось определить листочек. " +
//      "Выберите диапазон целиком внутри листочка (например, заголовок)." );
//    return;
  }
}

function get_active_section(spreadsheet, sheet, range) {
  [spreadsheet, sheet, range] = get_active_range(spreadsheet, sheet, range);
  var worksheet = get_active_worksheet(spreadsheet, sheet, range);
  if (worksheet == null)
    return null;
  try {
    var section = Worksheet.surrounding_section(worksheet.group, worksheet, range);
    return section;
  } catch (error) {
    report_error(error);
//    report_error(
//      "Не удалось определить раздел листочка. " +
//      "Выберите диапазон внутри одного раздела (например, заголовок)." );
//    return;
  }
}

return {
  get_active_group: get_active_group,
  get_active_worksheet: get_active_worksheet,
  get_active_section: get_active_section,
};
}(); // end ActionHelpers namespace

