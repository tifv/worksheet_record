
function metadata_editor() {
  var template = HtmlService.createTemplateFromFile(
    "MetadataEditor/Spreadsheet" );
  template.color_schemes = ColorSchemes.get(SpreadsheetApp.getActiveSpreadsheet());
  template.color_scheme_default = ColorSchemes.get_default();
  var output = template.evaluate();
  output.setHeight(500);
  SpreadsheetApp.getUi().showModelessDialog(output, "Метаданные");
}

function metadata_editor_set_color_schemes(schemes) {
  ColorSchemes.set(SpreadsheetApp.getActiveSpreadsheet(), schemes);
}
