function test_set_timetable_for_today() {
  const spreadsheet = MainSpreadsheet.get();
  var now_s = WorksheetDate.today().format();
  for (let group of StudyGroup.list(spreadsheet)) {
    console.log(group.name);
    let timetable = {};
    for (var i = 1; i < 10; ++i) {
      timetable[i] = {time: 1030 + i * 10, duration: 5};
    }
    group.set_timetable({[now_s]: timetable});
    console.log(group.get_timetable());
    console.log(group.get_today_timetable());
  }
}

function test_set_worksheet_plan_for_today() {
  const spreadsheet = MainSpreadsheet.get();
  var now_s = WorksheetDate.today().format();
  for (let group of StudyGroup.list(spreadsheet)) {
    console.log(group.name);
    let worksheet_plan = [{period: "1", category: "a"}, {period: "2", category: "g"}, {period: "3", category: "c"}];
    group.set_worksheet_plan({[now_s]: worksheet_plan});
    console.log(group.get_worksheet_plan());
    console.log(group.get_today_worksheet_plan());
  }
}

function test_stupid_trigger() {
  var now = new Date();
  now.setSeconds(now.getSeconds() + 15);
  //ScriptApp.newTrigger("worksheet_planned_add.sheet$test_04_ZCmw")
  ScriptApp.newTrigger("worksheet_planned_add.gid$1544770151")
    .timeBased()
    .at(now)
    .create();
}

function import_timetable() {
  const ui = SpreadsheetApp.getUi();
  let response = ui.prompt( "Импорт из расписания",
    "Введите ID или URL расписания:",
    ui.ButtonSet.OK_CANCEL );
  if (response.getSelectedButton() != ui.Button.OK) {
    return;
  }
  var ref = response.getResponseText();
  var timetable;
  if (/\//.exec(ref) != null) {
    timetable = SpreadsheetApp.openByUrl(ref);
  } else {
    timetable = SpreadsheetApp.openById(ref);
  }
  const spreadsheet = MainSpreadsheet.get();
  const sheet = timetable.getSheetByName("timetable");
  const frozen_rows = sheet.getFrozenRows();
  const frozen_cols = sheet.getFrozenColumns();
  const max_rows = sheet.getMaxRows();
  const max_cols = sheet.getMaxColumns();
  var group_names = sheet.getRange(1, frozen_cols + 1, 1, max_cols - frozen_cols)
    .getValues()[0];
  var row_labels = sheet.getRange(frozen_rows + 1, 1, max_rows - frozen_rows, 1)
    .getValues().map(([x]) => x);
  var values = sheet.getRange(frozen_rows + 1, frozen_cols + 1, max_rows - frozen_rows, max_cols - frozen_cols)
    .getValues();
  var merges = sheet.getRange(frozen_rows + 1, frozen_cols + 1, max_rows - frozen_rows, max_cols - frozen_cols)
    .getMergedRanges();
  for (let merged_range of merges) {
    let start_row = merged_range.getRow()    - (frozen_rows + 1);
    let start_col = merged_range.getColumn() - (frozen_cols + 1);
    let end_row = merged_range.getLastRow()    - (frozen_rows + 1);
    let end_col = merged_range.getLastColumn() - (frozen_cols + 1);
    if (start_row < 0 || end_row < 0)
      continue;
    for (let i = start_row; i <= end_row; ++i) {
      for (let j = start_col; j <= end_col; ++j) {
        values[i][j] = null;
      }
    }
  }
  var date = null;
  var timetables = {};
  var worksheet_plans = {};
  function get_object(root, key) {
    if (key in root)
      return root[key];
    root[key] = {};
    return root[key];
  }
  function get_list(root, key) {
    if (key in root)
      return root[key];
    root[key] = [];
    return root[key];
  }
  function time_to_min(hh, mm) {
    return parseInt(hh, 10) * 60 + parseInt(mm);
  }
  var categories = Object.entries(Categories.get(spreadsheet))
    .map(([code, {name}]) => [code, name]);
  for (let i = 0; i < row_labels.length; ++i) {
    let row_label = row_labels[i];
    if (row_label instanceof Date) {
      date =
        row_label.getFullYear().toString().padStart(2, "0") + "-" +
        (row_label.getMonth() + 1).toString().padStart(2, "0") + "-" +
        row_label.getDate().toString().padStart(2, "0");
      continue;
    }
    if (date == null)
      continue;
    let period_match = /^\d/.exec(row_label)
    if (period_match == null)
      continue;
    let period = period_match[0];
    let time_matches = Array.from(row_label.matchAll(/(\d\d):(\d\d)/g));
    let timetable_item;
    if (time_matches.length == 0) {
      timetable_item = {};
    } else if (time_matches.length == 1) {
      let start = time_to_min(time_matches[0][1], time_matches[0][2]);
      timetable_item = {time: start};
    } else if (time_matches.length == 2) {
      let start = time_to_min(time_matches[0][1], time_matches[0][2]);
      let end = time_to_min(time_matches[1][1], time_matches[1][2]);
      timetable_item = {time: start, duration: end - start};
    } else {
      timetable_item = {};
    }
    for (let j = 0; j < group_names.length; ++j) {
      let group_name = group_names[j];
      let group_timetable = get_object(timetables, group_name);
      let group_worksheet_plan = get_object(worksheet_plans, group_name);
      let today_timetable = get_object(group_timetable, date);
      let today_worksheet_plan = get_list(group_worksheet_plan, date);
      today_timetable[period] = timetable_item;
      let value = values[i][j];
      if (value == null)
        continue;
      let category_codes = categories
        .filter(([code, name]) => value.toLowerCase().startsWith(name.toLowerCase()))
        .map(([code]) => code);
      var category = (category_codes.length == 1) ? category_codes[0] : null;
      today_worksheet_plan.push({period: period, category: category});
    }
  }
  //console.log(JSON.stringify(timetable));
  //console.log(JSON.stringify(worksheet_plans));
  for (let group_name of group_names) {
    var group = StudyGroup.find_by_name(spreadsheet, group_name);
    if (group == null)
      continue;
    group.set_timetable(timetables[group_name]);
    group.set_worksheet_plan(worksheet_plans[group_name]);
    //console.log(JSON.stringify(group.get_timetable()));
    //console.log(JSON.stringify(group.get_worksheet_plan()));
  }
}

