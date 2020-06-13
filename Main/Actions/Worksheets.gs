function action_worksheet_insert() {
  try {
    var lock = ActionHelpers.acquire_lock();
    var worksheet = ActionHelpers.get_active_worksheet();
    var note_info = Worksheet.parse_title_note(worksheet.get_title_note());
    // XXX check that the next column exists and is empty
    Worksheet.add(
      worksheet.group,
      worksheet.sheet.getRange(1, worksheet.dim.end + 1),
      {date: note_info.date} );
    lock.releaseLock();
    worksheet.sheet.getParent().toast(
      "Исправьте дату в примечании к заголовку таблички, если требуется." );
  } catch (error) {
    report_error(error);
  }
}

function action_worksheet_add() {
  try {
    var lock = ActionHelpers.acquire_lock();
    var group = ActionHelpers.get_active_group();
    var sheet = group.sheet;
    var last_column = sheet.getLastColumn();
    // XXX also check that the insertion column is not in the frozen range
    if (last_column == group.dim.sheet_width) {
      throw ReportError("Последний столбец вкладки должен быть пустым.");
    }
    var date = WorksheetDate.today();
    date.period = group.get_current_period(7);
    Worksheet.add(group, sheet.getRange(1, last_column + 1), {date: date});
    lock.releaseLock();
    group.sheet.getParent().toast(
      "Дата: " + date.format() + "; " +
      "исправьте её в примечании к заголовку таблички, если требуется." );
  } catch (error) {
    report_error(error);
  }
}

function action_worksheet_recolor() {
  try {
    var group = ActionHelpers.get_active_group();
    var template = HtmlService.createTemplateFromFile(
      "Actions/Worksheets-Recolor" );
    template.group_name = group.name;
    template.color_schemes = ColorSchemes.get(SpreadsheetApp.getActiveSpreadsheet());
    template.color_scheme_group = group.get_color_scheme();
    template.color_scheme_default = ColorSchemes.get_default();
    template.editable = User.admin_is_acquired();
    var output = template.evaluate();
    output.setWidth(400).setHeight(400);
    SpreadsheetApp.getUi().showModelessDialog(output, "Перекрасить листочек");
  } catch (error) {
    report_error(error);
  }
}

function action_worksheet_recolor_finish(group_name, color_scheme, {scope, group: group_options}) {
  var lock = ActionHelpers.acquire_lock();
  if (scope == "worksheet") {
    var worksheet = ActionHelpers.get_active_worksheet();
    worksheet.recolor_cf_rules(color_scheme);
  } else if (scope == "group") {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var group = StudyGroup.find_by_name(spreadsheet, group_name);
    var worksheet_start_col = Worksheet.find_start_col(group);
    var cfrules = ConditionalFormatting.RuleList.load(group.sheet);
    if (group_options.rating) {
      let end_col = (worksheet_start_col != null) ?
        worksheet_start_col : this.sheetbuf.dim.sheet_width;
      cfrules.replace({ type: "gradient",
        condition: group.get_cfcondition_rating(),
        locations: [
          [group.dim.data_row, 1, group.dim.data_height, end_col],
          [group.dim.max_row, 1, 1, end_col],
        ],
      }, group.get_cfeffect_rating(color_scheme));
    }
    if (group_options.worksheets && worksheet_start_col != null) {
      Worksheet.recolor_cf_rules(group, color_scheme, cfrules, worksheet_start_col);
    }
    if (group_options.group) {
      if (color_scheme.origin == "group") {
        // no-op, this should equal to the current scheme
        // XXX no, this is incorrect assumption
      } else if (color_scheme.origin == "default") {
        group.set_color_scheme(null);
      } else {
        group.set_color_scheme(color_scheme);
      }
    }
    cfrules.save(group.sheet);
  }
  lock.releaseLock();
}

function action_worksheet_upload() {
  try {
    if (!upload_enabled_()) {
      throw new ReportError("Загрузка файлов не настроена");
    }
    var lock = ActionHelpers.acquire_lock();
    var section = ActionHelpers.get_active_section();
    if (!section.is_solutions()) {
      upload_start_dialog_(section);
    } else {
      let problems_section = section.get_unsolutions();
      upload_start_dialog_( section,
        action_worksheet_upload_solutions.get_dialog_options(
          problems_section, section )
      );
    }
    lock.releaseLock();
  } catch (error) {
    report_error(error);
  }
}

function action_worksheet_upload_solutions() {
  try {
    if (!upload_enabled_()) {
      throw new ReportError("Загрузка файлов не настроена");
    }
    var lock = ActionHelpers.acquire_lock();
    var section = ActionHelpers.get_active_section();
    var solutions_section;
    if (section.is_solutions()) {
      solutions_section = section;
    } else {
      solutions_section = section.get_solutions({title: "решения"});
    }
    var problems_section = solutions_section.get_unsolutions();
    upload_start_dialog_( solutions_section,
      action_worksheet_upload_solutions.get_dialog_options(
        problems_section, solutions_section )
    );
    lock.releaseLock();
  } catch (error) {
    report_error(error);
  }
}

action_worksheet_upload_solutions.get_dialog_options =
function(problems_section, solutions_section) {
  return {
    filename_suffix: "solutions",
    filename_date: Worksheet.parse_title_note(problems_section.get_title_note()).date,
  };
}

function action_worksheet_planned() {
  var template = HtmlService.createTemplateFromFile(
    "Actions/Worksheets-Timetable" );
  var output = template.evaluate();
  output.setWidth(400).setHeight(400);
  SpreadsheetApp.getUi().showModelessDialog(output, "Добавить листочки по плану");
}


function action_worksheet_planned_load() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var worksheet_plans = {};
  for (let group of StudyGroup.list(spreadsheet)) {
    let plan = group.get_today_worksheet_plan();
    worksheet_plans[group.name] = plan != null ? plan.length : null;
  }
  return worksheet_plans;
}

function action_worksheet_planned_add(group_name) {
  try {
    var lock = ActionHelpers.acquire_lock();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var group = StudyGroup.find_by_name(spreadsheet, group_name);
    var today = WorksheetDate.today();
    var plan = group.get_today_worksheet_plan(today);
    if (plan == null)
      return;
    var sheet = group.sheet;
    var last_column = sheet.getLastColumn();
    if (last_column == group.dim.sheet_width) {
      throw ReportError("Последний столбец вкладки должен быть пустым.");
    }
    for (let plan_item of plan) {
      plan_item.date = today;
      if (plan_item.period != null) {
        plan_item.date.period = parseInt(plan_item.period, "10");
      }
      if (plan_item.title == null) {
        plan_item.title = "{Бланк " + plan_item.date.format() + "}";
      }
      Worksheet.add(group, sheet.getRange(1, last_column + 1), plan_item);
      last_column = sheet.getLastColumn();
    }
    lock.releaseLock();
  } catch (error) {
    report_error(error);
  }
}

