var SampleRow = function() { // begin namespace

function is_hidden(group) {
  return group.sheet.isRowHiddenByUser(group.dim.data_row);
}

function warn_not_hidden(group) {
  const ui = SpreadsheetApp.getUi();
  ui.alert( "Строка-образец",
    "Рекомендуется оставить первую строку (после закрепленных строк) пустой и скрытой. " +
    "В ином случае функции добавления и удаления строк могут работать некорректно.",
    ui.ButtonSet.OK );
}

return {is_hidden, warn_not_hidden};
}(); // end SampleRow namespace


function action_add_rows() {
  ReportError.with_reporting(() => {
    var {output, sample_is_hidden} = Active.with_group((group) => {
      var template = HtmlService.createTemplateFromFile(
        "Actions/Rows-Add" );
      var sample_is_hidden = SampleRow.is_hidden(group);
      template.sample_is_hidden = sample_is_hidden;
      template.group_name = group.name;
      template.names = group.sheet.getRange(group.dim.data_row, 1, group.dim.data_height, 2)
        .getValues().map(row => row.filter(x => x).join(" "));
      return {
        output: template.evaluate(),
        sample_is_hidden };
    });
    if (!sample_is_hidden)
      SampleRow.warn_not_hidden();
    output.setWidth(250).setHeight(400);
    SpreadsheetApp.getUi().showModelessDialog(output, "Добавление строк");
  });
}

function action_add_rows_finish(group_name, row_index, row_count) {
  // XXX fix participants counter (maybe just load the formula and then reset it)
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  ActionLock.with_lock(() => {
    var group = StudyGroup.find_by_name(spreadsheet, group_name);
    var sheet = group.sheet;
    function get_row_range(row, height) {
      return sheet.getRange(row, 1, height, group.sheetbuf.dim.sheet_width);
    }
    var sample_range;
    if (row_index > 0) {
      sheet.insertRowsAfter(group.dim.data_row + row_index - 1, row_count);
      sample_range = get_row_range(group.dim.data_row, 1);
    } else {
      sheet.insertRowsBefore(group.dim.data_row, row_count);
      sample_range = get_row_range(group.dim.data_row + row_count, 1);
    }
    var added_range = get_row_range(group.dim.data_row + row_index, row_count);
    sample_range.copyTo( added_range,
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false );
    var formulas = [];
    formulas.length = row_count;
    formulas.fill(
      sample_range.getFormulasR1C1()[0]
        .map(x => (x === "") ? null : x)
    );
    added_range.setValues(formulas);
    added_range
      .setBorder(
        row_index > 0 ? true : null, null,
        row_index < group.dim.data_height ? true : null, null,
        null, true,
        "black", SpreadsheetApp.BorderStyle.DOTTED );
    if (row_index == group.dim.data_height) {
      added_range
        .setBorder(
          null, null, true, null, null, null );
    }
  });
}

function action_sort_rows_by_name() {
  ReportError.with_reporting(() => {
    var {sample_is_hidden} = Active.with_group((group) => {
      var sample_is_hidden = SampleRow.is_hidden(group);
      var sample_offset = sample_is_hidden ? 1 : 0;
      var range = group.sheet.getRange(
        group.dim.data_row + sample_offset, 1,
        group.dim.data_height - sample_offset, group.sheetbuf.dim.sheet_width );
      range.sort([2]);
      var name_range = range.offset(0, 1, range.getNumRows(), 1);
      var names = name_range.getValues().map(([v]) => v);
      var effective_height = -1;
      for (let i = 0; i < names.length; ++i) {
        if (names[i] != "")
          effective_height = i;
      }
      effective_height += 1;
      if (effective_height > 0)
        range.offset(0, 0, effective_height).setBorder(
          null, null, true, null, null, null );
      return {sample_is_hidden};
    });
    if (!sample_is_hidden)
      SampleRow.warn_not_hidden();
  });
}

