function action_add_columns() {
  ReportError.with_reporting(() => {
    var {output} = Active.with_section((section) => {
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
      return {output: template.evaluate()};
    });
    output.setWidth(500).setHeight(225);
    SpreadsheetApp.getUi().showModelessDialog(output, "Добавление колонок");
  });
}

function action_add_columns_finish(
  group_name,
  worksheet_location,
  section_location,
  data_index, data_width
) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  ActionLock.with_lock(() => {
    var group = StudyGroup.find_by_name(spreadsheet, group_name);
    var worksheet = Worksheet.find_by_location(group, worksheet_location);
    var section = worksheet.find_section_by_location(section_location);
    section.add_columns(data_index, data_width);
  });
}

function action_add_section() {
  ReportError.with_reporting(() => {
    var {output} = Active.with_worksheet((worksheet) => {
      var template = HtmlService.createTemplateFromFile(
        "Actions/Columns-AddSection" );
      template.standalone = true;
      template.group_name = worksheet.group.name;
      template.worksheet_location = worksheet.get_location();
      template.date = WorksheetDate.today().to_object();
      template.date.period = worksheet.group.get_current_period(7);
      return {output: template.evaluate()};
    });
    output.setWidth(250).setHeight(225);
    SpreadsheetApp.getUi().showModelessDialog(output, "Добавление раздела");
  });
}

function action_add_section_finish(
  group_name,
  worksheet_location,
  {
    data_width, title, weight, date: date_obj
  }
) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  ActionLock.with_lock(() => {
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
        title_note_data: new Worksheet.NoteData(date != null ? [["date", date]] : date)
      } );
    new_section.worksheet.set_metaweight(weight, {add: true});
  });
}

function action_remove_excess_columns() {
  ReportError.with_reporting(() => {
    var remove_count = Active.with_section((section) => {
      return section.remove_excess_columns();
    });
    if (remove_count == 0) {
      throw new ReportError("Колонки не удалены. " +
        "Автоматически удаляются только пустые колонки без номера задачи.");
    } else if (remove_count == -1) {
      throw new ReportError("Колонки не удалены. " +
        "Раздел или листочек целиком можно удалить только вручную.");
    }
  });
}

function action_alloy_subproblems() {
  ReportError.with_reporting(() => {
    Active.with_group((group) => {
      var sheet = group.sheet;
      var active_ranges = sheet.getActiveRangeList().getRanges();
      var worksheet = null;
      var had_detection_errors = false;
      for (let range of active_ranges) {
        let inside_previous_worksheet = ( worksheet != null &&
          range.getColumn() >= worksheet.dim.start &&
          range.getLastColumn() <= worksheet.dim.end );
        if (inside_previous_worksheet)
          continue;
        try {
          worksheet = Worksheet.surrounding(group, range);
        } catch (error) {
          if (!(error instanceof WorksheetDetectionError))
            throw error;
          console.error(error);
          had_detection_errors = true;
          continue;
        }
        worksheet.alloy_subproblems();
        // XXX optimize to only collect ranges, and then set borders all at once;
        // factor it out of border_data_set
      }
      if (had_detection_errors) {
        SpreadsheetApp.flush();
        throw ReportError.standard.WorksheetDetection();
      }
    });
  });
}

function action_mark_columns(colours) {
  ReportError.with_reporting(() => {
    Active.with_group((group) => {
      var sheet = group.sheet;
      var active_ranges = sheet.getActiveRangeList().getRanges();
      var worksheet = null;
      var label_ranges = [];
      var data_ranges = [];
      var had_detection_errors = false;
      var had_boundary_errors = false;
      for (let range of active_ranges) {
        let start = range.getColumn();
        let end = range.getLastColumn();
        let width = end - start + 1;
        let inside_previous_worksheet = ( worksheet != null &&
          start >= worksheet.dim.data_start && end <= worksheet.dim.data_end );
        if (!inside_previous_worksheet) {
          try {
            worksheet = Worksheet.surrounding(group, range);
          } catch (error) {
            if (!(error instanceof WorksheetDetectionError))
              throw error;
            console.error(error);
            had_detection_errors = true;
            continue;
          }
          let inside_new_worksheet = (
            start >= worksheet.dim.data_start && end <= worksheet.dim.data_end );
          if (!inside_new_worksheet) {
            console.error( "Range " + range.getA1Notation() +
              " is outside the worksheet data range." )
            had_boundary_errors = true;
            continue;
          }
        }
        label_ranges.push(sheet.getRange(group.dim.label_row, start, 1, width));
        data_ranges.push(
          sheet.getRange(group.dim.data_row, start, group.dim.data_height, width) );
        if (worksheet.has_max_row()) {
          data_ranges.push(sheet.getRange(group.dim.max_row, start, 1, width));
        }
        if (worksheet.has_weight_row()) {
          data_ranges.push(sheet.getRange(group.dim.weight_row, start, 1, width));
        }
      }
      const label_colour = HSL.to_hex(colours.label);
      for (let range of label_ranges) {
        range.setBackground(label_colour);
      }
      const data_colour = HSL.to_hex(colours.data);
      for (let range of data_ranges) {
        range.setBackground(data_colour);
      }
      if (had_detection_errors || had_boundary_errors)
        SpreadsheetApp.flush();
      if (had_detection_errors)
        throw ReportError.standard.WorksheetDetection();
      if (had_boundary_errors)
        throw new ReportError(
          "Закрашиваемые столбцы должны содержать задачи внутри листочков." );
    });
  });
}

function action_mark_columns_finished() {
  action_mark_columns({
    label: {h:  60, s: 0.70, l: 0.75},
    data:  {h:  60, s: 0.70, l: 0.85},
  });
}

function action_mark_columns_burning() {
  action_mark_columns({
    label: {h: -20, s: 0.70, l: 0.80},
    data:  {h: -20, s: 0.70, l: 0.88},
  });
}
