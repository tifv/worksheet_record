var MainSpreadsheet = function() {

const property_key = "main_spreadsheet";

var spreadsheet = null;

function get() {
  if (spreadsheet != null)
    return spreadsheet;
  var id = PropertiesService.getDocumentProperties()
    .getProperty(property_key);
  if (id == null)
    return null;
  try {
    spreadsheet = SpreadsheetApp.openById(id);
    return spreadsheet;
  } catch (error) {
    return null;
  }
}

function set(spreadsheet) {
  PropertiesService.getDocumentProperties()
    .setProperty(property_key, spreadsheet.getId());
}

function is_set() {
  if (spreadsheet != null)
    return true;
  var id = PropertiesService.getDocumentProperties()
    .getProperty(property_key);
  if (id == null)
    return false;
  return true;
}

return {get: get, set: set, is_set: is_set};
}();
