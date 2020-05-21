function action_add_group() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var template = HtmlService.createTemplateFromFile(
    "Actions/StudyGroups-Add" );
  template.color_schemes = ColorSchemes.get(spreadsheet);
  template.color_scheme_default = ColorSchemes.get_default();
  template.categories = Categories.get(spreadsheet);
  var output = template.evaluate();
  output.setHeight(500);
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
  var group = StudyGroup.add(spreadsheet, group_name, options);
}

