function action_add_columns() {
  var section = ActionHelpers.get_active_section();
  if (section == null)
    return;
  var worksheet = section.worksheet;
  var group = worksheet.group;
  try {
    var template = HtmlService.createTemplateFromFile(
      "Actions/Columns-Add" );
    template.standalone = true;
    template.group_name = group.name;
    template.worksheet_location = worksheet.get_location();
    template.section_location = section.get_location({check_id: section.offset > 0});
    template.title = worksheet.get_title() +
      (section.dim.offset > 0 ? ". " + section.get_title() : "");
    console.log(worksheet.get_title());
    console.log(section.dim.offset);
    console.log(section.get_title());
    template.labels = group.sheetbuf.slice_values( "label_row",
      section.dim.data_start, section.dim.data_end );
    var output = template.evaluate();
    output.setWidth(500).setHeight(225);
    SpreadsheetApp.getUi().showModelessDialog(output, "Добавление колонок");
  } catch (error) {
    report_error(error);
    return;
  }
}

function action_add_columns_finish(
  group_name,
  worksheet_location,
  section_location,
  data_index, data_width
) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var group = StudyGroup.find_by_name(spreadsheet, group_name);
  var worksheet = Worksheet.find_by_location(group, worksheet_location);
  var section = worksheet.find_section_by_location(section_location);
  section.add_columns(data_index, data_width);
  //group.sheetbuf.test();
}

function action_add_section() {
  var worksheet = ActionHelpers.get_active_worksheet();
  if (worksheet == null)
    return;
  try {
    var template = HtmlService.createTemplateFromFile(
      "Actions/Columns-AddSection" );
    template.standalone = true;
    template.group_name = worksheet.group.name;
    template.worksheet_location = worksheet.get_location();
    template.date = WorksheetDate.today().to_object();
    // XXX detect date period
    var output = template.evaluate();
    output.setWidth(250).setHeight(225);
    SpreadsheetApp.getUi().showModelessDialog(output, "Добавление раздела");
  } catch (error) {
    report_error(error);
    return;
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
  var group = StudyGroup.find_by_name(spreadsheet, group_name);
  var worksheet = Worksheet.find_by_location(group, worksheet_location);
  var last_section = Worksheet.surrounding_section( group, worksheet,
    group.sheet.getRange(1, worksheet.dim.end) );
  var date;
  if (date_obj != null) {
    date = WorksheetDate.from_object(date_obj);
  } else {
    date = null;
  }
  var weightcell = worksheet.metaweight_cell;
  var weight = weightcell.getValue() + weight;
  worksheet.add_section_after(last_section, {data_width: data_width, title: title, date: date});
  group.sheetbuf.set_value("weight_row", worksheet.dim.rating, weight);
  //group.sheetbuf.test();
}

function action_remove_excess_columns() {
  var section = ActionHelpers.get_active_section();
  if (section == null)
    return;
  var worksheet = section.worksheet;
  var group = worksheet.group;
  try {
    var remove_count = section.remove_excess_columns();
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
    return;
  }
}

function action_alloy_subproblems() {
  var worksheet = ActionHelpers.get_active_worksheet();
  if (worksheet == null)
    return;
  try {
    worksheet.alloy_subproblems();
  } catch (error) {
    report_error(error);
    return;
  }
}

const finished_colours = {
  label: [60, 0.7, 0.75],
  data:  [60, 0.7, 0.85],
};

function action_mark_columns_finished() {
  var group = ActionHelpers.get_active_group();
  if (group == null)
    return;
  try {
    var sheet = group.sheet;
    var active_ranges = sheet.getActiveRangeList().getRanges();
    console.log(active_ranges.map(range => range.getA1Notation()));
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
        worksheet_has_weight_row = worksheet.has_weight_row();
        // XXX handle possible error?
        if (start < worksheet.dim.data_start || end > worksheet.dim.data_end) {
          throw new Error("XXX range invalid " + range.getA1Notation());
        }
      }
      label_ranges.push(sheet.getRange(group.dim.label_row, start, 1, width));
      data_ranges.push(
        sheet.getRange(group.dim.data_row, start, group.dim.data_height, width),
        sheet.getRange(group.dim.max_row, start, 1, width) );
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
  } catch (error) {
    report_error(error);
    return;
  }
}
