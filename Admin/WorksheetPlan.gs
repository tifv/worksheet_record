function worksheet_planned_add_single(group, today = WorksheetDate.today()) {
  var plan = group.get_today_worksheet_plan(today);
  if (plan == null)
    return;
  var sheet = group.sheet;
  var last_column = sheet.getLastColumn();
  if (last_column == group.dim.sheet_width) {
    throw new Error("Последний столбец вкладки «" + group.name + "» должен быть пустым.");
  }
  for (let plan_item of plan) {
    plan_item.date = today;
    if (plan_item.period != null) {
      plan_item.date.period = parseInt(plan_item.period, "10");
    }
    if (plan_item.title == null) {
      plan_item.title = "{Бланк " + plan_item.date.format() + "}";
    }
    WorksheetBuilder.build(group, sheet.getRange(1, last_column + 1), plan_item);
    last_column = sheet.getLastColumn();
  }
}

function worksheet_planned_add_all() {
  const spreadsheet = MainSpreadsheet.get();
  var today = WorksheetDate.today();
  var errors = [];
  for (let group of StudyGroup.list(spreadsheet)) {
    try {
      worksheet_planned_add_single(group, today);
    } catch (error) {
      console.error(error);
      errors.push(error);
    }
  }
  if (errors.length > 0) {
    throw errors[0];
  }
}

var worksheet_planned_add = new Proxy({}, {get: function(obj, name) {
  if (name.startsWith("sheet$")) {
    const spreadsheet = MainSpreadsheet.get();
    let sheet_name = name.substring(6);
    return () => {
      let group = StudyGroup.find_by_name(spreadsheet, sheet_name);
      worksheet_planned_add_single(group);
    };
  } else if (name.startsWith("gid$")) {
    const spreadsheet = MainSpreadsheet.get();
    let gid = parseInt(name.substring(4));
    let sheets = spreadsheet.getSheets().filter(sheet => sheet.getSheetId() == gid);
    if (sheets.length != 1)
      return null;
    return () => {
      let group = new StudyGroup(sheets[0]);
      group.check();
      worksheet_planned_add_single(group);
    }
  }
}});

function worksheet_planned_add_all_delay() {
  const spreadsheet = MainSpreadsheet.get();
  var today = WorksheetDate.today();
  for (let trigger of ScriptApp.getProjectTriggers()) {
    if (trigger.getHandlerFunction().startsWith("worksheet_planned_add."))
      ScriptApp.deleteTrigger(trigger);
  }
  var date = new Date();
  date.setSeconds(0);
  date.setMinutes(date.getMinutes() + 2);
  for (let group of StudyGroup.list(spreadsheet)) {
    let plan = group.get_today_worksheet_plan(today);
    if (plan == null || plan.length == 0)
      continue;
    ScriptApp.newTrigger("worksheet_planned_add.gid$" + group.sheet.getSheetId())
      .timeBased().at(date)
      .create();
    date.setMinutes(date.getMinutes() + plan.length);
  }
}