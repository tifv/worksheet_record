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
    console.error(error);
    if (error instanceof StudyGroupDetectionError) {
      throw new ReportError(
        "Не удалось определить учебную группу. " +
        "Выберите вкладку таблицы, соответствующую группе." );
    }
    throw error;
  }
}

function get_active_worksheet(spreadsheet, sheet, range) {
  [spreadsheet, sheet, range] = get_active_range(spreadsheet, sheet, range);
  var group = get_active_group(spreadsheet, sheet);
  try {
    var worksheet = Worksheet.surrounding(group, range);
    return worksheet;
  } catch (error) {
    console.error(error);
    if (error instanceof WorksheetDetectionError) {
      throw new ReportError(
        "Не удалось определить листочек. " +
        "Выберите диапазон целиком внутри листочка (например, заголовок)." );
    }
    throw error;
  }
}

function get_active_section(spreadsheet, sheet, range) {
  [spreadsheet, sheet, range] = get_active_range(spreadsheet, sheet, range);
  var worksheet = get_active_worksheet(spreadsheet, sheet, range);
  try {
    var section = Worksheet.Section.surrounding(worksheet.group, worksheet, range);
    return section;
  } catch (error) {
    console.error(error);
    if (error instanceof WorksheetSectionDetectionError) {
      throw new ReportError(
        "Не удалось определить раздел листочка. " +
        "Выберите диапазон внутри одного раздела (например, заголовок)." );
    }
    throw error;
  }
}

// Lock policy:
// Actions that add or remove columns should definetly acquire lock.
// Actions that can modify spreadsheet in other ways should acquire lock.
// (this does include all worksheet.get_location() calls with validation).
// In those cases, lock should be acquired before calls to sheetbuf.
// Sidebar contents initial scan does not acquire lock, even though
// it can, in some cases, modify spreadsheet (it runs too often to acquire lock).
function acquire_lock() {
  var lock = LockService.getDocumentLock();
  var success = lock.tryLock(300);
  if (success)
    return lock;
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Ожидание завершения других операций." );
  var success = lock.tryLock(5000);
  if (success)
    return lock;
  throw new ReportError(
    "Не удалось получить доступ к таблице. " +
    "Другие операции (возможно, запущенные другими редакторами) " +
    "создают опасность одновременного редактирования." );
}

return {
  get_active_group: get_active_group,
  get_active_worksheet: get_active_worksheet,
  get_active_section: get_active_section,
  acquire_lock: acquire_lock,
};
}(); // end ActionHelpers namespace

