var ActionHelpers = function() { // begin namespace

function get_active_sheet({spreadsheet, sheet} = {}) {
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

function get_active_range({spreadsheet, sheet, range} = {}) {
  if (range == null) {
    if (sheet == null) {
      range = SpreadsheetApp.getActiveRange();
      sheet = range.getSheet();
    } else {
      range = sheet.getActiveRange();
    }
  } else {
    if (sheet == null)
      sheet = range.getSheet();
  }
  if (spreadsheet == null)
    spreadsheet = sheet.getParent();
  return [spreadsheet, sheet, range];
}

function get_active_group({spreadsheet, sheet, lock: lockopt} = {}) {
  if (sheet == null)
    [spreadsheet, sheet] = get_active_sheet({spreadsheet, sheet});
  var group;
  try {
    group = new StudyGroup(sheet);
    group.check();
  } catch (error) {
    console.error(error);
    if (!(error instanceof StudyGroupDetectionError))
      throw error;
    if (group == null)
      throw error;
    try {
      group.check({metadata: false, dim: true});
      const ui = SpreadsheetApp.getUi();
      let response = ui.alert( "Ошибка",
        "Выбранная вкладка несёт структуру учебной группы, но не отмечена как таковая. " +
        "Отметить эту вкладку как группу?",
        ui.ButtonSet.YES_NO );
      if (response == ui.Button.YES) {
        group.add_metadatum();
        if (lockopt == "preserve")
          throw new ReportError("Перезапустите функцию.");
        if (lockopt == "acquire") {
          return [group, acquire_lock()];
        }
        return group;
      }
      throw new ReportError(
        "Выбранная вкладка не отмеченка как учебная группа." );
    } catch (error) {
      if (error instanceof StudyGroupDetectionError) {
        throw new ReportError(
          "Выбранная вкладка не несёт структуры учебной группы. " +
          "Выберите вкладку таблицы, соответствующую группе." );
      }
      throw error;
    }
  }
  if (lockopt == "acquire")
    return [group, acquire_lock()];
  return group;
}

function get_active_worksheet({spreadsheet, sheet, range, group, lock: lockopt} = {}) {
  if (range == null || (sheet == null && group == null)) {
    [spreadsheet, sheet, range] = get_active_range({spreadsheet, sheet, range});
  }
  var lock;
  if (group == null) {
    if (lockopt == "acquire") {
      [group, lock] = get_active_group({ spreadsheet, sheet,
        lock: "acquire" });
    } else {
      group = get_active_group({ spreadsheet, sheet,
        lock: lockopt });
    }
  }
  try {
    var worksheet = Worksheet.surrounding(group, range);
  } catch (error) {
    console.error(error);
    if (error instanceof WorksheetDetectionError) {
      throw new ReportError(
        "Не удалось определить листочек. " +
        "Выберите диапазон целиком внутри листочка (например, заголовок)." );
    }
    throw error;
  }
  if (lockopt == "acquire")
    return [worksheet, lock];
  return worksheet;
}

function get_active_section({spreadsheet, sheet, range, group, worksheet, lock: lockopt} = {}) {
  if (range == null) {
    [spreadsheet, sheet, range] = get_active_range({spreadsheet, sheet, range});
  }
  var lock;
  if (worksheet == null) {
    if (lockopt == "acquire") {
      [worksheet, lock] = get_active_worksheet({ spreadsheet, sheet, range, group,
        lock: "acquire" });
    } else {
      worksheet = get_active_worksheet({ spreadsheet, sheet, range, group,
        lock: lockopt });
    }
  }
  try {
    var section = Worksheet.Section.surrounding(worksheet.group, worksheet, range);
  } catch (error) {
    console.error(error);
    if (error instanceof WorksheetSectionDetectionError) {
      throw new ReportError(
        "Не удалось определить раздел листочка. " +
        "Выберите диапазон внутри одного раздела (например, заголовок)." );
    }
    throw error;
  }
  if (lockopt == "acquire")
    return [section, lock];
  return section;
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

