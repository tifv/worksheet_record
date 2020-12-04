function action_add_columns() {
  try {
    var lock = ActionHelpers.acquire_lock();
    var section = ActionHelpers.get_active_section();
    var worksheet = section.worksheet;
    var group = worksheet.group;
    var template = HtmlService.createTemplateFromFile(
      "Actions/Columns-Add" );
    template.standalone = true;
    template.group_name = group.name;
    template.worksheet_location = worksheet.get_location();
    template.section_location = section.get_location({validate: section.dim.offset > 0});
    template.title = section.get_qualified_title();
    template.labels = group.sheetbuf.slice_values( "label_row",
      section.dim.data_start, section.dim.data_end );
    lock.releaseLock();
    var output = template.evaluate();
    output.setWidth(500).setHeight(225);
    SpreadsheetApp.getUi().showModelessDialog(output, "Добавление колонок");
  } catch (error) {
    report_error(error);
  }
}

function action_add_columns_finish(
  group_name,
  worksheet_location,
  section_location,
  data_index, data_width
) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var lock = ActionHelpers.acquire_lock();
  var group = StudyGroup.find_by_name(spreadsheet, group_name);
  var worksheet = Worksheet.find_by_location(group, worksheet_location);
  var section = worksheet.find_section_by_location(section_location);
  section.add_columns(data_index, data_width);
  //group.sheetbuf.test();
}

function action_add_section() {
  try {
    var lock = ActionHelpers.acquire_lock();
    var worksheet = ActionHelpers.get_active_worksheet();
    var template = HtmlService.createTemplateFromFile(
      "Actions/Columns-AddSection" );
    template.standalone = true;
    template.group_name = worksheet.group.name;
    template.worksheet_location = worksheet.get_location();
    template.date = WorksheetDate.today().to_object();
    template.date.period = worksheet.group.get_current_period(7);
    lock.releaseLock();
    var output = template.evaluate();
    output.setWidth(250).setHeight(225);
    SpreadsheetApp.getUi().showModelessDialog(output, "Добавление раздела");
  } catch (error) {
    report_error(error);
  }
}

function action_add_section_finish(
  group_name,
  worksheet_location,
  {
    data_width, title, weight, date: date_obj
  }
) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var lock = ActionHelpers.acquire_lock();
  var group = StudyGroup.find_by_name(spreadsheet, group_name);
  var worksheet = Worksheet.find_by_location(group, worksheet_location);
  var last_section = Worksheet.Section.surrounding( group, worksheet,
    group.sheet.getRange(1, worksheet.dim.end) );
  var date;
  if (date_obj != null) {
    date = WorksheetDate.from_object(date_obj);
  } else {
    date = null;
  }
  var new_section = worksheet.add_section_after( last_section,
    {
      data_width: data_width, title: title,
      title_note_data: new Worksheet.NoteData([["date", date]])
    } );
  new_section.worksheet.set_metaweight(weight, {add: true});
  lock.releaseLock();
  //group.sheetbuf.test();
}

function action_remove_excess_columns() {
  try {
    var lock = ActionHelpers.acquire_lock();
    var section = ActionHelpers.get_active_section();
    var worksheet = section.worksheet;
    var group = worksheet.group;
    var remove_count = section.remove_excess_columns();
    lock.releaseLock();
    //section.group.sheetbuf.test();
    if (remove_count == 0) {
      throw "Колонки не удалены. " +
        "Автоматически удаляются только пустые колонки без номера задачи.";
    } else if (remove_count == -1) {
      throw "Колонки не удалены. " +
        "Раздел или листочек целиком можно удалить только вручную.";
    }
  } catch (error) {
    report_error(error);
  }
}

function action_alloy_subproblems() {
  try {
    var lock = ActionHelpers.acquire_lock();
    var worksheet = ActionHelpers.get_active_worksheet();
    worksheet.alloy_subproblems();
    lock.releaseLock();
  } catch (error) {
    report_error(error);
  }
}

const finished_colours = {
  label: {h:  60, s: 0.70, l: 0.75},
  data:  {h:  60, s: 0.70, l: 0.85},
};

function action_mark_columns_finished() {
  try {
    var lock = ActionHelpers.acquire_lock();
    var group = ActionHelpers.get_active_group();
    var sheet = group.sheet;
    var active_ranges = sheet.getActiveRangeList().getRanges();
    var worksheet = null;
    var worksheet_has_weight_row = null;
    var label_ranges = [];
    var data_ranges = [];
    for (let range of active_ranges) {
      let start = range.getColumn();
      let end = range.getLastColumn();
      let width = end - start + 1;
      if ( worksheet == null ||
        start < worksheet.dim.data_start || end > worksheet.dim.data_end
      ) {
        worksheet = Worksheet.surrounding(group, range);
        // handle possible error when finding worksheet
        worksheet_has_max_row = ( group.dim.max_row != null &&
          worksheet.has_max_row() );
        worksheet_has_weight_row = ( group.dim.weight_row != null &&
          worksheet.has_weight_row() );
        if (start < worksheet.dim.data_start || end > worksheet.dim.data_end) {
          throw new Error("XXX range invalid " + range.getA1Notation());
          // XXX save all invalid ranges and report them at the end instead
        }
      }
      label_ranges.push(sheet.getRange(group.dim.label_row, start, 1, width));
      data_ranges.push(
        sheet.getRange(group.dim.data_row, start, group.dim.data_height, width) );
      if (worksheet_has_max_row) {
        data_ranges.push(sheet.getRange(group.dim.max_row, start, 1, width));
      }
      if (worksheet_has_weight_row) {
        data_ranges.push(sheet.getRange(group.dim.weight_row, start, 1, width));
      }
    }
    const label_colour = HSL.to_hex(finished_colours.label);
    for (let range of label_ranges) {
      range.setBackground(label_colour);
    }
    const data_colour = HSL.to_hex(finished_colours.data);
    for (let range of data_ranges) {
      range.setBackground(data_colour);
    }
    lock.releaseLock();
  } catch (error) {
    report_error(error);
  }
}
