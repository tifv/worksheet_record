function action_regenerate_toc() {
  const toc_name = "toc";
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var toc_sheet = spreadsheet.getSheetByName(toc_name);
  var toc_sheet_index = null;
  if (toc_sheet != null) {
    var toc_sheet_index = toc_sheet.getIndex() - 1;
    spreadsheet.deleteSheet(toc_sheet);
  }
  var min_group_sheet_index = null;
  var formula_pieces = ['={"Оглавление","",""'];
  var first_group = true;
  for (let group of StudyGroup.list(spreadsheet)) {
    if (toc_sheet_index == null) {
      let group_sheet_index = group.sheet.getIndex();
      if (min_group_sheet_index == null || min_group_sheet_index > group_sheet_index)
        min_group_sheet_index = group_sheet_index;
    }
    let start_col = Worksheet.find_start_col(group);
    if (start_col == null)
      continue;
    formula_pieces.push(';"","","";');
    first_group = false;
    let name = group.name;
    let title_ref = "'" + name + "'!R" + group.dim.title_row + "C" + start_col + ":R" + group.dim.title_row;
    let void_ref = "arrayformula(" + title_ref + "+na())";
    let category_ref = group.dim.category_row != null ?
      ("'" + name + "'!R" + group.dim.category_row + "C" + start_col + ":R" + group.dim.category_row) :
      void_ref;
    let data_ref = "iferror({" + void_ref + ";" + category_ref + ";" + title_ref + "})";
    formula_pieces.push(
      '"' + group.name + '","","";',
      "transpose(filter(" +
        data_ref + ",not(isblank(" + title_ref + "))" +
      "))"
    );
  }
  formula_pieces.push("}");
  if (toc_sheet_index == null && min_group_sheet_index != null) {
    toc_sheet_index = min_group_sheet_index - 1;
  }
  if (toc_sheet_index != null) {
    toc_sheet = spreadsheet.insertSheet(toc_name, toc_sheet_index);
  } else {
    toc_sheet = spreadsheet.insertSheet(toc_name);
  }
  toc_sheet.getRange(1, 1).setFormulaR1C1(formula_pieces.join(""));
  var group_range = toc_sheet.getRange("A2:A");
  var title_range = toc_sheet.getRange("C2:C");
  var toc_sheet_height = toc_sheet.getMaxRows();
  toc_sheet.getRange(1,1).setFontSize(20);
  group_range.setFontSize(16);
  title_range.setFontSize(12);
  var cfrules = new ConditionalFormatting.RuleList();
  cfrules.push( ConditionalFormatting.Rule.from_object({ type: "boolean",
    condition: { type: SpreadsheetApp.BooleanCriteria.TEXT_STARTS_WITH,
      values: ["{"] },
    ranges: [[2, 3, toc_sheet_height - 1, 1]],
    effect: {font_color: "#dddddd"},
  }));
  for (let [code, category] of Object.entries(Categories.get(spreadsheet))) {
    let color = category.color;
    if (color == null)
      continue;
    cfrules.push(ConditionalFormatting.Rule.from_object({ type: "boolean",
      condition: { type: SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA,
        values: [ "=exact(" +
          '"' + code.replace('"', '""') + '"' + ",R[0]C2)" ]
      },
      ranges: [[2, 3, toc_sheet_height - 1, 1]],
      effect: {background: HSL.to_hex(color)},
    }));
  }
  cfrules.save(toc_sheet);
  toc_sheet.setFrozenRows(1);
  toc_sheet.setColumnWidth(1, 25);
  toc_sheet.setColumnWidth(2, 25);
  toc_sheet.hideColumns(2);
  toc_sheet.setColumnWidth(3, 400);
  toc_sheet.setHiddenGridlines(true);
  toc_sheet.deleteColumns(4, toc_sheet.getMaxColumns() - 3);
  toc_sheet.protect().setWarningOnly(true);
}

function action_regenerate_uploads() {
  UploadRecord.create();
}

