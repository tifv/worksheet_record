function action_add_group() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var template = HtmlService.createTemplateFromFile(
    "Actions/StudyGroups-Add" );
  template.color_schemes = ColorSchemes.get(spreadsheet);
  template.color_scheme_default = ColorSchemes.get_default();
  var categories = Categories.get(spreadsheet)
  template.categories = categories;
  template.category_css = format_category_css_(categories);
  var output = template.evaluate();
  output.setHeight(500).setWidth(300);
  SpreadsheetApp.getUi().showModelessDialog(output, "Добавить группу");
}

function action_add_group_existing(group_name) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(group_name);
  if (sheet == null) {
    throw new ReportError(
      "Вкладка с именем «" + group_name + "» не найдена." );
  }
  var group = new StudyGroup(sheet, group_name);
  group.add_metadatum();
}

function action_add_group_new(group_name, options) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //console.log(JSON.stringify(options.timetable));
  //console.log(JSON.stringify(options.worksheet_plan));
  var group = StudyGroupBuilder.build(spreadsheet, group_name, options);
  //console.log(JSON.stringify(group.get_timetable()));
  //console.log(JSON.stringify(group.get_worksheet_plan()));
}

