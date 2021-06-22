var Active = function() { // begin namespace

function get_range({spreadsheet, sheet, range} = {}) {
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
  return {spreadsheet, sheet, range};
}

class ActiveGroupDetectionError extends Error {
  constructor(group, original_error) {
    super();
    this.group = group;
    this.original = original_error;
  }
}

function with_group_norecover(operator) {
  var {spreadsheet, sheet, range} = get_range();
  return ActionLock.with_lock(() => {
    var group;
    try {
      group = new StudyGroup(sheet);
      group.check();
    } catch (error) {
      if (!(error instanceof StudyGroupDetectionError))
        throw error;
      throw new ActiveGroupDetectionError(group, error);
    }

    return operator(group, {spreadsheet, sheet, range});
  });
}

function recover_metadata(group) {
  try {
    group.check({metadata: false, dim: true});
  } catch (error) {
    if (!(error instanceof StudyGroupDetectionError))
      throw error;
    return false;
  }
  const ui = SpreadsheetApp.getUi();
  let response = ui.alert( "Восстановление",
    "Выбранная вкладка несёт структуру учебной группы, но не отмечена как таковая. " +
    "Отметить эту вкладку как группу?",
    ui.ButtonSet.YES_NO );
  if (response != ui.Button.YES)
    throw new ReportError.standard.StudyGroupNoMetadata();
  group.recover_metadata();
  ui.alert( "Восстановление",
    "Вкладка отмечена как учебная группа. " +
    "Перезапустите функцию.",
    ui.ButtonSet.OK );
  return true;
}

function recover_frozen_rows(group) {
  try {
    group.check({metadata: true, dim: false});
  } catch (error) {
    if (!(error instanceof StudyGroupDetectionError))
      throw error;
    return false;
  }
  var max_rows = group.sheet.getMaxRows();
  if (max_rows == 0)
    return false;
  if (max_rows > 100)
    max_rows = 100;
  var frozen_rows = group.sheet.getFrozenRows();
  var max_row_marker = 0;
  var marker_range = group.sheet.getRange(1, 1, max_rows, 1);
  var markers = marker_range.getValues().map(([v]) => v);
  for (let i = 0; i < markers.length; ++i) {
    if (markers[i] != "")
      max_row_marker = i + 1;
  }
  if (max_row_marker <= frozen_rows)
    return false;
  var new_frozen_rows = max_row_marker;
  const ui = SpreadsheetApp.getUi();
  let response = ui.alert( "Восстановление",
    "Выбранная вкладка отмечена как учебная группа, но у неё не хватает закреплленных строк. " +
    "Закрепить " + new_frozen_rows + " строк?",
    ui.ButtonSet.YES_NO );
  if (response != ui.Button.YES)
    throw new ReportError.standard.StudyGroupNoFrozenRows();
  group.sheet.setFrozenRows(new_frozen_rows);
  ui.alert( "Восстановление",
    "Во вкладке закреплены " + new_frozen_rows + " строк. " +
    "Перезапустите функцию.",
    ui.ButtonSet.OK );
  return true;
}

function with_group(operator) {
  try {
    return with_group_norecover(operator);
  } catch (error) {
    if (!(error instanceof ActiveGroupDetectionError))
      throw error;
    var group = error.group;
    if (recover_frozen_rows(group))
      throw ReportError.ALREADY_REPORTED;
    if (recover_metadata(group))
      throw ReportError.ALREADY_REPORTED;
    console.error(error.original);
    throw ReportError.standard.StudyGroupDetection();
  }
}

function with_worksheet(operator) {
  return with_group((group, {spreadsheet, sheet, range}) => {
    try {
      var worksheet = Worksheet.surrounding(group, range);
    } catch (error) {
      if (!(error instanceof WorksheetDetectionError))
        throw error;
      console.error(error);
      throw ReportError.standard.WorksheetDetection();
    }

    return operator(worksheet, {spreadsheet, sheet, range});
  });
}

function with_section(operator) {
  return with_worksheet((worksheet, {spreadsheet, sheet, range}) => {
    try {
      var section = Worksheet.Section.surrounding(worksheet.group, worksheet, range);
    } catch (error) {
      if (!(error instanceof WorksheetSectionDetectionError))
        throw error;
      console.error(error);
      throw ReportError.standard.WorksheetSectionDetection();
    }

    return operator(section, {spreadsheet, sheet, range});
  });
}

return {
  with_group,
  with_worksheet,
  with_section,
};
}(); // end Active namespace

