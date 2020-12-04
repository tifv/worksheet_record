var MainSpreadsheet = function myFunction() {

const property_key = "main_spreadsheet";

var main_spreadsheet = null;

function get() {
  if (main_spreadsheet != null)
    return main_spreadsheet;
  var id = PropertiesService.getDocumentProperties().getProperty(property_key);
  if (id == null)
    return null;
  try {
    main_spreadsheet = SpreadsheetApp.openById(id);
    return main_spreadsheet;
  } catch (error) {
    return null;
  }
}

function set(spreadsheet) {
  main_spreadsheet = spreadsheet;
  PropertiesService.getDocumentProperties().setProperty(property_key, main_spreadsheet.getId());
}

return {get: get, set: set};
}();
