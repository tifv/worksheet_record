function worksheet_cleanup_single(group, {errors: super_errors}) {
  var errors = (super_errors != null) ? super_errors : [];
  let items = [];
  let today = WorksheetDate.today();
  let yesterday = WorksheetDate.today(-1);
  for (let worksheet of Worksheet.list(group)) {
    if (!worksheet.is_unused())
      continue;
    let date = worksheet.get_title_note_data().get("date");
    if (today.compare(date) <= 0)
      continue;
    if (yesterday.compare(date) <= 0) {
      items.push(["minimize", worksheet]);
    } else {
      items.push(["trash", worksheet]);
    }
  }
  items.reverse();
  for (let [method, worksheet] of items) {
    try {
      let data_empty = worksheet.data_range.getValues()
        .every(row => row.every(value => value == ""));
      if (!data_empty) {
        throw new WorksheetError(
          "Worksheet is marked as unused, but data is not empty",
          worksheet.full_range );
      }
      if (method == "minimize") {
        if (worksheet.dim.data_width > 6)
          group.sheetbuf.delete_columns(
            worksheet.dim.data_start + 6, worksheet.dim.data_width - 6 );
      } else if (method == "trash") {
        if (!worksheet.separator_column_range.isBlank())
          throw new WorksheetError(
            "Separator column must be blank",
            worksheet.full_range );
          group.sheetbuf.delete_columns(
            worksheet.dim.start, worksheet.dim.width + 1 );
      }
    } catch (error) {
      console.error(error.toString());
      errors.push(error);
    }
  }
  if (super_errors == null && errors.length > 0) {
    throw errors[0];
  }
}

function worksheet_cleanup_all() {
  const spreadsheet = MainSpreadsheet.get();
  var errors = [];
  for (let group of StudyGroup.list(spreadsheet)) {
    try {
      worksheet_cleanup_single(group, {errors: errors});
    } catch (error) {
      console.error(error.toString());
      errors.push(error);
    }
  }
  if (errors.length > 0) {
    throw errors[0];
  }
}

