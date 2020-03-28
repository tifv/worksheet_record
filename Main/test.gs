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

function test_worksheet_clear_(spreadsheet, name) {
  var sheet = spreadsheet.getSheetByName(name);
  if (sheet != null)
    spreadsheet.deleteSheet(sheet);
}

function test_worksheet() {
  test_worksheet_1();
}

function test_worksheet_1() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  test_worksheet_clear_(spreadsheet, "test1");
  var group = StudyGroup.add(spreadsheet, "test1", {
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
    attendance: {
      columns: {
        date_lists: [
          {
            title: "I",
            start: new Date(2019, 8, 1),
            end:   new Date(2020, 0, 1),
            weekdays: [false, true, false, false, false, true, false]
          },
          {
            title: "II",
            start: new Date(2020, 0,1),
            end:   new Date(2020, 5,1),
            weekdays: [false, true, false, false, false, true, false]
          }
        ] } },
  });
  var sheet = group.sheet;
  group = new StudyGroup(sheet);
  var playsheet = spreadsheet.getSheetByName("play");
  playsheet.getDataRange().getValues();
  test_worksheet_add_(group, {data_width: 20, title: "Алгебра", category: "a"});
  test_worksheet_add_random_(group, {data_width: 20, title: "Геометрия", category: "g"});
  for (var i = 1; i <= 5; ++i) {
    test_worksheet_add_random_(group, {data_width: 20, title: "Геометрия " + i, category: "g"});
  }
//  var worksheet1 = Worksheet.add(group, sheet.getRange(1,sheet.getMaxColumns()), {data_width: 20, title: "Алгебра", category: "a"});
//  worksheet1.data_range.offset(1,0,20,20).setValues(playsheet.getRange("A1:T20").getValues());
//  worksheet1.label_range.setValues(playsheet.getRange("A22:T22").getValues());
//  var worksheet2 = Worksheet.add(group, sheet.getRange(1,sheet.getMaxColumns()), {data_width: 20, title: "Геометрия", category: "g"});
//  worksheet2.data_range.offset(1,0,20,20).setValues(playsheet.getRange("Y1:AR20").getValues());
//  worksheet2.label_range.setValues(playsheet.getRange("A22:T22").getValues());
//  var worksheet3 = Worksheet.add(group, sheet.getRange(1,sheet.getMaxColumns()), {data_width: 20, category: "c"});
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

function test_set_meta_category() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  Categories.set(spreadsheet, {
    a: {name: "алгебра",       filename: "algebra",       color: [-30, .40, .80]},
    g: {name: "геометрия",     filename: "geometry",      color: [210, .40, .80]},
    c: {name: "комбинаторика", filename: "combinatorics", color: [ 90, .40, .80]},
    n: {name: "теория чисел",  filename: "number-theory", color: [ 30, .40, .80]},
    o: {name: "олимпиада",     filename: "olympiad",      color: [ 30,1.00, .80]},
    mixture: {name: "разнобой", filename: "mixture"}
  });
  Logger.log(Categories.get(spreadsheet));
}

function test_set_self_admin() {
  Admin.set_self_admin();
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

function test_worksheet_sections() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = spreadsheet.getSheetByName("proto")
  const range = sheet.getRange("CW:DS");
  var group = new StudyGroup(sheet);
  var worksheet = new Worksheet(group, range);
  worksheet.check();
  var section = Array.from(worksheet.list_sections())[1];
  worksheet.add_section_after(section, {date: new WorksheetDate(2020,5,5)});
  //console.log(worksheet.get_location());
  //for (let section of worksheet.list_sections()) {
  //  console.log(section.get_location());
  //}
}

function test_add_uploads() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var sheet = spreadsheet.getSheetByName("uploads");
  if (sheet != null)
    spreadsheet.deleteSheet(sheet);
  UploadRecord.create();
}

function test_set_group_filename() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var group = StudyGroup.find_by_name(spreadsheet, "proto");
  group.set_filename("proto");
}

