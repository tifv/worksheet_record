function upload_configure() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var template = HtmlService.createTemplateFromFile(
    "Upload/ConfigDialog" );
  template.upload_record_exists = UploadRecord.exists(spreadsheet);
  template.upload_config = UploadConfig.get();
  var output = template.evaluate();
  output.setWidth(500).setHeight(475);
  SpreadsheetApp.getUi().showModelessDialog(output, "Настройка загрузки файлов");
}

function upload_config_set(config, preserve_secret_key = false) {
  UploadConfig.set(config, preserve_secret_key);
}

