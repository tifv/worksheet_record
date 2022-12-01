function action_worksheet_insert() {
  ReportError.with_reporting(() => {
    var {sheet} = Active.with_worksheet((worksheet) => {
      var options = worksheet.group.get_worksheet_options();
      var note_data = worksheet.get_title_note_data();
      // XXX check that the next column exists and is empty
      WorksheetBuilder.build(
        worksheet.group,
        worksheet.sheet.getRange(1, worksheet.dim.end + 1),
        Object.assign({}, {date: note_data.get("date")}, options) );
      return {sheet: worksheet.sheet};
    });
    sheet.getParent().toast(
      "Исправьте дату в примечании к заголовку таблички, если требуется." );
  });
}

function action_worksheet_add() {
  ReportError.with_reporting(() => {
    var {sheet, date} = Active.with_group((group) => {
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
      date.period = group.get_current_period(/* offset = */ 10);
      WorksheetBuilder.build( group,
        sheet.getRange( 1, last_column + 1,
          1, group.sheetbuf.dim.sheet_width - last_column ),
        Object.assign({}, {date: date}, options) );
      return {sheet, date};
    });
    sheet.getParent().toast(
      "Дата: " + date.format() + "; " +
      "исправьте её в примечании к заголовку таблички, если требуется." );
  });
}

function action_worksheet_recolor_dialog() {
  ReportError.with_reporting(() => {
    var {output} = Active.with_group((group) => {
      var template = HtmlService.createTemplateFromFile(
        "Actions/Worksheets-Recolor" );
      template.group_name = group.name;
      template.color_schemes = ColorSchemes.get(SpreadsheetApp.getActiveSpreadsheet());
      template.color_scheme_group = group.get_color_scheme();
      template.color_scheme_default = ColorSchemes.get_default();
      template.editable = User.admin_is_acquired();
      return {output: template.evaluate()};
    });
    output.setWidth(400).setHeight(400);
    SpreadsheetApp.getUi().showModelessDialog(output, "Перекрасить листочек");
  });
}

function action_worksheet_recolor_single(color_scheme) {
  Active.with_worksheet((worksheet) => {
    worksheet.recolor_cf_rules(color_scheme);
  });
}

function action_worksheet_recolor_group(group_name, color_scheme, options = {}) {
  ActionLock.with_lock(() => {
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
  });
}

function action_worksheet_upload() {
  ReportError.with_reporting(() => {
    if (!upload_enabled_()) {
      throw new ReportError("Загрузка файлов не настроена");
    }
    Active.with_section((section) => {
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
    });
  });
}

function action_worksheet_fake_upload() {
  ReportError.with_reporting(() => {
    if (!upload_enabled_()) {
      throw new ReportError("Загрузка файлов не настроена");
    }
    Active.with_section((section) => {
      if (section.worksheet.is_unused()) {
        throw new ReportError("У листочка нет заголовка.")
      }
      if (!section.is_addendum()) {
        upload_fake_finish_(section);
      } else {
        let original_section = section.get_original();
        upload_fake_finish_( section,
          action_worksheet_upload_addendum.get_dialog_options(
            original_section, section )
        );
      }
    });
  });
}

function action_worksheet_upload_addendum(options) {
  if (options.type == null)
    throw new Error("internal error: missing option");
  ReportError.with_reporting(() => {
    if (!upload_enabled_()) {
      throw new ReportError("Загрузка файлов не настроена");
    }
    Active.with_section((section) => {
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
    });
  });
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

function action_worksheet_upload_show_src_link() {
  ReportError.with_reporting(() => {
    if (!upload_enabled_()) {
      throw new ReportError("Загрузка файлов не настроена");
    }
    Active.with_section((section) => {
      var spreadsheet = section.group.sheet.getParent();
      successful: {
        let title_formula = section.get_title_formula();
        if (title_formula == "")
          break successful;
        let title_formula_decode = decode_hyperlink_formula_(title_formula);
        if (title_formula_decode == null)
          break successful;
        let [{filter = null}, ] = title_formula_decode;
        if (filter == null)
          break successful;
        var response = [];
        for (let datum of UploadRecord.get(spreadsheet, "minimal").find("id", filter)) {
          let response_part = [];
          response_part.push('<p>');
          if (datum.has('pdf')) {
            let pdf_link = datum.get('pdf');
            response_part.push(
              'PDF: <a target="_blank" href="' + pdf_link + '">' +
                pdf_link.substring(pdf_link.lastIndexOf('/')) + '</a>'
            );
          }
          if (datum.has('pdf') && datum.has('src')) {
            response_part.push('<br/>')
          }
          if (datum.has('src')) {
            let src_link = datum.get('src');
            response_part.push(
              'Source: <a target="_blank" href="' + src_link + '">' +
                src_link.substring(src_link.lastIndexOf('/')) + '</a>'
            );
          }
          response_part.push('</p>');
          response.push(response_part.join(''));
        }
        if (response.length == 0)
          break successful;
        SpreadsheetApp.getUi().showModelessDialog(HtmlService.createHtmlOutput(response.join('')), "Загруженный файл");
        return;
      }
      throw new ReportError("Файл не найден в реестре загрузок.")
    });
  });
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
  ReportError.with_reporting(() => {
    ActionLock.with_lock(() => {
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
          plan_item.date.period = parseInt(plan_item.period, 10);
        }
        if (plan_item.title == null) {
          plan_item.title = worksheet_blank_namer_(plan_item.date);
        }
        WorksheetBuilder.build( group,
          sheet.getRange(1, last_column + 1),
          Object.assign({}, options, plan_item) );
        last_column = sheet.getLastColumn();
      }
    });
  });
}

function action_worksheet_convert_to_olympiad() {
  const default_limit = OlympiadSheet.initial.olympiad.limit;
  ReportError.with_reporting(() => {
    Active.with_worksheet((worksheet) => {
      worksheet = OlympiadSheet.from_worksheet(worksheet);
      var group = worksheet.group;
      var sheet = group.sheet;
      worksheet.set_weight_formula_default();
      if (group.dim.weight_row != null) {
        for (let section of worksheet.list_sections()) {
          section.set_data_borders(
            section.dim.data_start, section.dim.data_end, {
              weight_row: false,
              max_row: section.has_max_row(),
            } );
        }
      }
      if (worksheet.has_limit_cell()) {
        worksheet.set_limit(default_limit, {
          set_cell_format: true, set_cell_borders: true });
      }
      if (worksheet.sum_column != null) {
        worksheet.set_sum_formula_default({
          limit_value: worksheet.has_limit_cell() ? null : default_limit,
        });
      }
      if (worksheet.rating_column != null) {
        worksheet.set_rating_formula_default({set_number_format: true});
      }
      var color_scheme = group.get_color_scheme();
      var cfrules = ConditionalFormatting.RuleList.load(sheet);
      cfrules.insert(
        worksheet.new_cfrule_data_limit(color_scheme, {
          limit_value: worksheet.has_limit_cell() ? null : default_limit,
        }),
        worksheet.constructor.get_cffilter_data(group),
      );
      cfrules.save(sheet);
    });
  });
}

