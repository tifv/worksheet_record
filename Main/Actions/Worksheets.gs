function action_worksheet_insert() {
  try {
    var [worksheet, lock] = ActionHelpers.get_active_worksheet({lock: "acquire"});
    var options = worksheet.group.get_worksheet_options();
    var note_data = worksheet.get_title_note_data();
    // XXX check that the next column exists and is empty
    WorksheetBuilder.build(
      worksheet.group,
      worksheet.sheet.getRange(1, worksheet.dim.end + 1),
      Object.assign({}, {date: note_data.get("date")}, options) );
    lock.releaseLock();
    worksheet.sheet.getParent().toast(
      "Исправьте дату в примечании к заголовку таблички, если требуется." );
  } catch (error) {
    report_error(error);
  }
}

function action_worksheet_add() {
  try {
    var [group, lock] = ActionHelpers.get_active_group({lock: "acquire"});
    var sheet = group.sheet;
    var options = group.get_worksheet_options();
    var last_column = sheet.getLastColumn();
    {
      let frozen_columns = sheet.getFrozenColumns();
      if (frozen_columns > last_column)
        last_column = frozen_columns;
    }
    if (last_column >= group.sheetbuf.dim.sheet_width) {
      throw new ReportError("Последний столбец вкладки должен быть пустым.");
    }
    var date = WorksheetDate.today();
    date.period = group.get_current_period(7);
    WorksheetBuilder.build( group,
      sheet.getRange( 1, last_column + 1,
        1, group.sheetbuf.dim.sheet_width - last_column ),
      Object.assign({}, {date: date}, options) );
    lock.releaseLock();
    group.sheet.getParent().toast(
      "Дата: " + date.format() + "; " +
      "исправьте её в примечании к заголовку таблички, если требуется." );
  } catch (error) {
    report_error(error);
  }
}

function action_worksheet_recolor_dialog() {
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

function action_worksheet_recolor_single(color_scheme) {
  var [worksheet, lock] = ActionHelpers.get_active_worksheet({lock: "acquire"});
  worksheet.recolor_cf_rules(color_scheme);
  lock.releaseLock();
}

function action_worksheet_recolor_group(group_name, color_scheme, options = {}) {
  var lock = ActionHelpers.acquire_lock();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var group = StudyGroup.find_by_name(spreadsheet, group_name);
  var worksheet_start_col = Worksheet.find_start_col(group);
  var cfrules = ConditionalFormatting.RuleList.load(group.sheet);
  if (options.rating) {
    let end_col = (worksheet_start_col != null) ?
      worksheet_start_col : group.sheetbuf.dim.sheet_width;
    cfrules.replace({ type: "gradient",
      condition: group.get_cfcondition_rating(),
      locations: [
        [group.dim.data_row, 1, group.dim.data_height, end_col],
        [group.dim.max_row, 1, 1, end_col],
      ],
    }, group.get_cfeffect_rating(color_scheme));
  }
  if (options.worksheets && worksheet_start_col != null) {
    Worksheet.recolor_cf_rules(group, color_scheme, cfrules, worksheet_start_col);
  }
  if (options.group) {
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
  lock.releaseLock();
}

function action_worksheet_upload() {
  try {
    if (!upload_enabled_()) {
      throw new ReportError("Загрузка файлов не настроена");
    }
    var [section, lock] = ActionHelpers.get_active_section({lock: "acquire"});
    if (section.worksheet.is_unused()) {
      throw new ReportError("У листочка нет заголовка.")
    }
    if (!section.is_addendum()) {
      upload_start_dialog_(section);
    } else {
      let original_section = section.get_original();
      upload_start_dialog_( section,
        action_worksheet_upload_addendum.get_dialog_options(
          original_section, section )
      );
    }
    lock.releaseLock();
  } catch (error) {
    report_error(error);
  }
}

function action_worksheet_upload_addendum(options) {
  if (options.type == null)
    throw new Error("internal error: missing option");
  try {
    if (!upload_enabled_()) {
      throw new ReportError("Загрузка файлов не настроена");
    }
    var [section, lock] = ActionHelpers.get_active_section({lock: "acquire"});
    var addendum_section;
    if (section.is_addendum()) {
      if (section.get_addendum_type() == options.type) {
        addendum_section = section;
      } else {
        throw new ReportError("Несовместимый тип доп. материалов")
      }
    } else {
      addendum_section = section.get_addendum(options);
    }
    var original_section = addendum_section.get_original();
    upload_start_dialog_( addendum_section,
      action_worksheet_upload_addendum.get_dialog_options(
        original_section, addendum_section )
    );
    lock.releaseLock();
  } catch (error) {
    report_error(error);
  }
}

action_worksheet_upload_addendum.get_dialog_options =
function(original_section, addendum_section) {
  return {
    filename_suffix: addendum_section.get_addendum_type(),
    filename_date: original_section.get_title_note_data().get("date"),
  };
}

menu_meta_setup_( action_worksheet_upload_addendum,
  addendum_metadata_key );

function addendums_restore_hardcoded() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var menu_data = SpreadsheetMetadata.set_object( spreadsheet, addendum_metadata_key, [
    ["theory",    "теория",    "nut"   ],
    ["hints",     "подсказки", "cookie"],
    ["answers",   "ответы",    "cake"  ],
    ["solutions", "решения",   "pizza" ],
  ].map( ([name, label, emoji_key]) =>
    [ name,
      { type: name, title: label,
        label: (emoji_key != null && emoji_key in emoji) ? emoji[emoji_key] : null },
      (emoji_key != null && emoji_key in emoji) ? emojipad[emoji_key] + label + "…" : label + "…",
    ] ));
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
    var options = group.get_worksheet_options();
    var today = WorksheetDate.today();
    var plan = group.get_today_worksheet_plan(today);
    if (plan == null)
      return;
    var sheet = group.sheet;
    var last_column = sheet.getLastColumn();
    if (last_column == group.dim.sheet_width) {
      throw new ReportError("Последний столбец вкладки должен быть пустым.");
    }
    for (let plan_item of plan) {
      plan_item.date = today;
      if (plan_item.period != null) {
        plan_item.date.period = parseInt(plan_item.period, "10");
      }
      if (plan_item.title == null) {
        plan_item.title = worksheet_blank_namer_(plan_item.date);
      }
      WorksheetBuilder.build( group,
        sheet.getRange(1, last_column + 1),
        Object.assign({}, options, plan_item) );
      last_column = sheet.getLastColumn();
    }
    lock.releaseLock();
  } catch (error) {
    report_error(error);
  }
}

function action_worksheet_convert_to_olympiad() {
  if (!User.admin_is_acquired()) {
    const ui = SpreadsheetApp.getUi();
    let response_btn = ui.alert( "Конвертация в олимпиаду",
      "Это труднообратимая операция.",
      ui.ButtonSet.OK_CANCEL );
    if (response_btn != ui.Button.OK) {
      return;
    }
  }
  try {
    var [worksheet, lock] = ActionHelpers.get_active_worksheet({lock: "acquire"});
    var group = worksheet.group;
    var sheet = group.sheet;
    if (group.dim.weight_row != null) {
      group.sheetbuf.set_values( "weight_row",
        worksheet.dim.data_start, worksheet.dim.data_end,
        null );
      worksheet.set_data_borders(
        worksheet.dim.data_start, worksheet.dim.data_end, {
          weight_row: false,
          max_row: worksheet.has_max_row(),
        } );
    }
    var limit_cell = null;
    if (group.dim.weight_row != null && worksheet.sum_column != null) {
      limit_cell = sheet.getRange(group.dim.weight_row, worksheet.sum_column);
      limit_cell.setFontSize(8).setBorder(true, true, true, true, null, null);
      group.sheetbuf.set_value( "weight_row", worksheet.sum_column,
        4 );
      group.sheetbuf.set_note( "weight_row", worksheet.sum_column,
        "пороговый балл решённой задачи" );
    }
    if (worksheet.sum_column != null) {
      var data_row_sum_R1C1 =
        'R[0]C[' + (worksheet.dim.data_start - 1 - worksheet.sum_column) + ']:' +
        'R[0]C[' + (worksheet.dim.data_end   + 1 - worksheet.sum_column) + ']';
      var limit_R1C1;
      if (limit_cell != null) {
        limit_R1C1 = '">="&R' + group.dim.weight_row + 'C[0]';
      } else {
        limit_R1C1 = '">=4"';
      }
      var sum_formula_R1C1;
      if (group.dim.max_row != null) {
        var max_row_sum_R1C1 =
          'R' + group.dim.max_row +
          'C[' + (worksheet.dim.data_start - 1 - worksheet.sum_column) + ']:' +
          'R' + group.dim.max_row +
          'C[' + (worksheet.dim.data_end   + 1 - worksheet.sum_column) + ']';
        sum_formula_R1C1 = ''.concat(
          '=countifs(',
            max_row_sum_R1C1,  ';' + limit_R1C1 + ';',
            data_row_sum_R1C1, ';' + limit_R1C1 + '',
          ')'
        );
      } else {
          sum_formula_R1C1 = ''.concat(
              '=countif(',
                  data_row_sum_R1C1, ';' + limit_R1C1 + '',
              ')'
          );
      }
      sheet.getRange(group.dim.data_row, worksheet.sum_column, group.dim.data_height, 1)
        .setFormulaR1C1(sum_formula_R1C1);
      if (group.dim.max_row != null) {
        group.sheetbuf.set_formula( "max_row",
          worksheet.sum_column, sum_formula_R1C1 );
      }
    }
    if (worksheet.rating_column != null) {
      var data_row_rating_R1C1 =
        'R[0]C[' + (worksheet.dim.data_start - 1 - worksheet.rating_column) + ']:' +
        'R[0]C[' + (worksheet.dim.data_end   + 1 - worksheet.rating_column) + ']';
      var rating_formula_R1C1 = ''.concat(
        '=sum(',
          data_row_rating_R1C1,
        ')'
      );
      var number_format = "0.#;−0.#";
      sheet.getRange(group.dim.data_row, worksheet.rating_column, group.dim.data_height, 1)
        .setFormulaR1C1(rating_formula_R1C1)
        .setNumberFormat(number_format);
      if (group.dim.max_row != null) {
          group.sheetbuf.set_formula( "max_row",
            worksheet.rating_column, rating_formula_R1C1 );
          sheet.getRange(group.dim.max_row, worksheet.rating_column)
            .setNumberFormat(number_format);
      }
    }
    var color_scheme = group.get_color_scheme();
    var cfrules = ConditionalFormatting.RuleList.load(sheet);
    var data_limit_cfrule;
    if (limit_cell != null) {
      data_limit_cfrule = worksheet.new_cfrule_data_limit(color_scheme);
    } else {
      data_limit_cfrule = worksheet.new_cfrule_data_limit(color_scheme, 4);
    }
    cfrules.insert( data_limit_cfrule,
      { type: "boolean",
        condition: Worksheet.get_cfcondition_data(this.group),
        location: [
          [group.dim.data_row, 1, group.dim.data_height, group.sheetbuf.dim.sheet_width],
          [group.dim.max_row, 1, 1, group.sheetbuf.dim.sheet_width],
        ].filter(([r,]) => r != null)
      },
    );
    cfrules.save(sheet);
    lock.releaseLock();
  } catch (error) {
    report_error(error);
  }
}

