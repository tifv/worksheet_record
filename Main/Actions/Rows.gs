function action_add_rows() {
  try {
    var group = ActionHelpers.get_active_group();
    var template = HtmlService.createTemplateFromFile(
      "Actions/Rows-Add" );
    var sample_is_hidden = group.sheet.isRowHiddenByUser(group.dim.data_row);
    if (!sample_is_hidden) {
      const ui = SpreadsheetApp.getUi();
      ui.alert( "Строка-образец",
        "Рекомендуется оставить первую строку (после закрепленных строк) пустой и скрытой. " +
        "В ином случае функции добавления и удаления строк могут работать некорректно.",
        ui.ButtonSet.OK );
    }
    template.sample_is_hidden = sample_is_hidden;
    template.group_name = group.name;
    template.names = group.sheet.getRange(group.dim.data_row, 1, group.dim.data_height, 2)
      .getValues().map(row => row.filter(x => x).join(" "));
    var output = template.evaluate();
    output.setWidth(250).setHeight(400);
    SpreadsheetApp.getUi().showModelessDialog(output, "Добавление строк");
  } catch (error) {
    report_error(error);
  }
}

function action_add_rows_finish(group_name, row_index, row_count) {
  // XXX fix participants counter (maybe just load the formula and then reset it)
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
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
}

function action_remove_excess_rows() {
  try {
    var group = ActionHelpers.get_active_group();
    const ui = SpreadsheetApp.getUi();
    var sample_is_hidden = group.sheet.isRowHiddenByUser(group.dim.data_row);
    if (!sample_is_hidden) {
      ui.alert( "Строка-образец",
        "Рекомендуется оставить первую строку (после строк-заголовков) пустой и скрытой. " +
        "В ином случае функции добавления и удаления строк могут работать некорректно.",
        ui.ButtonSet.OK );
    }
    var name_range = group.sheet.getRange(group.dim.data_row, 1, group.dim.data_height, 2);
    var names = name_range
      .getValues().map(row => row.filter(x => x).join(" "));
    var excess_rows = [];
    for (let i = 0; i < names.length; ++i) {
      if (i == 0 && sample_is_hidden)
        continue;
      if (names[i] != "")
        continue;
      excess_rows.push(i);
    }
    if (excess_rows.length == 0) {
      throw "Отсутствуют строки, которые можно было бы удалить автоматически.";
    }
    if (excess_rows.length == names.length) {
      throw "Нельзя просто взять и удалить все строки.";
    }
    var backgrounds = name_range.getBackgrounds();
    for (let i of excess_rows) {
      group.sheet.getRange(group.dim.data_row + i, 1, 1, 2)
        .setBackground("red");
    }
    SpreadsheetApp.flush();
    var response = ui.alert( "Удаление строк",
      "Строки, выделенные красным, будут удалены, окей?",
      ui.ButtonSet.OK_CANCEL );
    if (response != ui.Button.OK) {
      for (let i of excess_rows) {
        group.sheet.getRange(group.dim.data_row + i, 1, 1, 2)
          .setBackgrounds([
            backgrounds[i].map(colour => colour == "#ffffff" ? null : colour)
          ]);
      }
      return;
    }
    var delete_from = null, delete_length = 0;
    excess_rows.sort((a, b) => b - a);
    excess_rows.push(-2);
    for (let i of excess_rows) {
      if (delete_from == null) {
        delete_from = i;
        delete_length = 1;
        continue;
      }
      if (delete_from - delete_length == i) {
        ++delete_length;
        continue;
      } else {
        group.sheet.deleteRows(
          group.dim.data_row + delete_from - delete_length + 1,
          delete_length );
        delete_from = i;
        delete_length = 1;
      }
    }
    excess_rows.pop();
    if (excess_rows[0] == names.length - 1) {
      group.sheet.getRange(
        group.dim.data_row + group.dim.data_height - 1 - excess_rows.length, 1,
        1, group.sheetbuf.dim.sheet_width
      )
        .setBorder(
          null, null, true, null, null, null );
    }
    if (excess_rows[excess_rows.length - 1] == 0) {
      group.sheet.getRange(
        group.dim.data_row, 1,
        1, group.sheetbuf.dim.sheet_width
      )
        .setBorder(
          true, null, null, null, null, null );
    }
  } catch (error) {
    report_error(error);
  }
}

// XXX add sort_by_name action
