/*

function test_whatever_spreadsheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = spreadsheet.getActiveSheet();
  const range = sheet.getActiveRange();
}

function test_whatever_range() {
  var range = SpreadsheetApp.getActiveRange();
}

function test_whatever() {
  console.log("whatever");
}

*/


function test_worksheet() {
  test_worksheet_1_();
  test_worksheet_color_();
}

function test_worksheet_clear_(spreadsheet, name) {
  var sheet = spreadsheet.getSheetByName(name);
  if (sheet != null)
    spreadsheet.deleteSheet(sheet);
}

function test_worksheet_1_() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  test_worksheet_clear_(spreadsheet, "test_1");
  var group = StudyGroup.add(spreadsheet, "test_1", {
    rows: {
      mirror_row:   1,
      category_row: 5,
      title_row:    6,
      weight_row:   2,
      max_row:      3,
      label_row:    7,
      data_row:     8,
    },
    data_height: 20,
    rating: true, sum: true,
    categories: [
      {code: "a"},
      {code: "g"},
      {code: "c"},
      {code: "o", rating: false}
    ],
    category_musthave: true,
    attendance: {
      columns: {
        date_lists: [
          {
            title: "I",
            start: new Date(Date.parse("2020-02-01 ")),
            end:   new Date(Date.parse("2020-06-01 ")),
            weekdays: [true, false, false, true, false, false, false]
          },
          {
            title: "II",
            start: new Date(Date.parse("2020-09-01 ")),
            end:   new Date(Date.parse("2021-01-01 ")),
            weekdays: [true, false, false, true, false, false, false]
          }
        ] } },
    color_scheme: ColorSchemes.get(spreadsheet)["lotus"],
  });
  var sheet = group.sheet;
  group = new StudyGroup(sheet);
  { // cache…
    var playsheet = spreadsheet.getSheetByName("play");
    playsheet.getDataRange().getValues();
  }
  test_worksheet_add_random_(group, {
    data_width: 20, title: "Алгебра", category: "a",
  });
  test_worksheet_add_random_(group, {
    data_width: 20, title: "Геометрия", category: "g",
  });
  test_worksheet_add_random_(group, {
    data_width: 20, title: "Комбинаторика", category: "c",
  });
}

function test_worksheet_color_() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  test_worksheet_clear_(spreadsheet, "test_color");
  var color_schemes = ColorSchemes.get(spreadsheet);
  function named_scheme(name) {
    return Object.assign({name: name}, color_schemes[name]);
  }
  var group = StudyGroup.add(spreadsheet, "test_color", {
    rows: {
      mirror_row:   1,
      category_row: 5,
      title_row:    6,
      weight_row:   2,
      max_row:      3,
      label_row:    7,
      data_row:     8,
    },
    data_height: 20,
    rating: true, sum: true,
    categories: [
      {code: "a"},
      {code: "g"},
      {code: "c"},
      {code: "o", rating: {integrate: false}}
    ],
    category_musthave: false,
    //color_scheme: named_scheme("lotus*"),
  });
  var sheet = group.sheet;
  group = new StudyGroup(sheet);
  { // cache…
    var playsheet = spreadsheet.getSheetByName("play");
    playsheet.getDataRange().getValues();
  }
  function add_worksheet(name, shift=0) {
    test_worksheet_add_random_(group, {
      data_width: 20, title: name + (shift != 0 ? " " + shift : ""),
      color_scheme: shift == 0 ? named_scheme(name) : scheme_shift(named_scheme(name), shift),
    });
  }
  function hsl_shift(hsl, i=0) {
    let {h, s, l} = hsl;
    return {h: (h + 6*i) % 360, s: s, l: l};
  }
  function scheme_shift(scheme, i=0) {
    let {mark, rating_mid, rating_top} = scheme;
    return {
      mark: hsl_shift(mark, i),
      rating_mid: hsl_shift(rating_mid, i),
      rating_top: hsl_shift(rating_top, i),
    };
  }
  add_worksheet("default");
  for (let name in color_schemes) {
    add_worksheet(name);
  }
  add_worksheet("default");
}

function test_worksheet_add_(group, options) {
  const sheet = group.sheet;
  var worksheet = Worksheet.add(group, sheet.getRange(1,sheet.getMaxColumns()), options);
  SpreadsheetApp.getActiveSpreadsheet().getRangeByName("D20_data")
    .copyValuesToRange( sheet.getSheetId(), worksheet.dim.data_start, worksheet.dim.data_end, group.dim.data_row + 1, sheet.getMaxRows() );
  SpreadsheetApp.getActiveSpreadsheet().getRangeByName("D20_labels")
    .copyValuesToRange( sheet.getSheetId(), worksheet.dim.data_start, worksheet.dim.data_end, group.dim.label_row, group.dim.label_row );
}

function test_worksheet_add_random_(group, options) {
  const sheet = group.sheet;
  var worksheet = Worksheet.add(group, sheet.getRange(1,sheet.getMaxColumns()), options);
  SpreadsheetApp.getActiveSpreadsheet().getRangeByName("D20_data_random")
    .copyValuesToRange( sheet.getSheetId(), worksheet.dim.data_start, worksheet.dim.data_end, group.dim.data_row + 1, sheet.getMaxRows() );
  SpreadsheetApp.getActiveSpreadsheet().getRangeByName("D20_labels")
    .copyValuesToRange( sheet.getSheetId(), worksheet.dim.data_start, worksheet.dim.data_end, group.dim.label_row, group.dim.label_row );
  worksheet.check();
}

function test_set_upload_config() {
  const ui = SpreadsheetApp.getUi();
  function get_value(label) {
    var response = ui.prompt( "Загрузка (тест)",
      label + ":", ui.ButtonSet.OK_CANCEL );
    if (response.getSelectedButton() == ui.Button.CANCEL)
      throw "wat";
    return response.getResponseText();
  }

  UploadConfig.set({
    access_key: get_value("Access key"),
    secret_key: get_value("Secret key"),
    region:     get_value("Region"),
    bucket_url: get_value("Bucket URL"),
  });
}

function test_add_uploads() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spreadsheet.getSheetByName("uploads");
  if (sheet != null)
    spreadsheet.deleteSheet(sheet);
  UploadRecord.create();
}

function test_formula_parser() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = spreadsheet.getActiveSheet();
  const cell = sheet.getActiveCell();
  const base_cell = [cell.getRow(), cell.getColumn()];
  var formula_R1C1 = cell.getFormulaR1C1();
  var formula_A1 = cell.getFormula();
  function compare_and_log(formula_orig, formula) {
    console.log(formula_orig);
    if (formula_orig != formula) {
      console.error(formula);
    }
  }

  var formulaX = new Formula(formula_R1C1, {notation: "R1C1"});
  for (let token of formulaX.tokens) {
    console.log(token.constructor.name + " " + JSON.stringify(token));
  }
  compare_and_log( formula_R1C1,
    formulaX.toString({notation: "R1C1"}) );
  compare_and_log( formula_A1,
    formulaX.toString({notation: "A1", base_cell: base_cell}) );

  var formulaY = new Formula(formula_A1, {notation: "A1", base_cell: base_cell});
  for (let token of formulaY.tokens) {
    console.log(token.constructor.name + " " + JSON.stringify(token));
  }
  compare_and_log( formula_R1C1,
    formulaY.toString({notation: "R1C1"}) );
  compare_and_log( formula_A1,
    formulaY.toString({notation: "A1", base_cell: base_cell}) );
}

function test_conditional_formatting() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = spreadsheet.getActiveSheet();
  var cfrules = ConditionalFormatting.RuleList.load(sheet);
  cfrules.forEach(cfrule => { console.log(JSON.stringify(cfrule)); });
  cfrules.save(sheet);
}

