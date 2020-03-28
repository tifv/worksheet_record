/* SpreadsheetMetadata
 *   .get(spreadsheet, key)
 *   .unset(spreadsheet, key)
 *   .set(spreadsheet, key, value)
 *   .get_object(spreadsheet, key) → JSON
 *   .set_object(spreadsheet, key, value_object)
 */

var SpreadsheetMetadata = function () { // namespace

function get(spreadsheet, key) {
  var metadata = spreadsheet.createDeveloperMetadataFinder()
    .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.SPREADSHEET)
    .withVisibility(SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
    .withKey(key)
    .find();
  var metadatum = metadata[0];
  if (metadatum == null)
    return null;
  return metadatum.getValue();
}

function unset(spreadsheet, key) {
  var metadata = spreadsheet.createDeveloperMetadataFinder()
    .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.SPREADSHEET)
    .withVisibility(SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
    .withKey(key)
    .find();
  for (var i = 0; i < metadata.length; ++i) {
    metadata[i].remove();
  }
}

function set(spreadsheet, key, value) {
  unset(spreadsheet, key);
  spreadsheet.addDeveloperMetadata( key, value,
    SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT );
}

function get_object(spreadsheet, key) {
  var value = get(spreadsheet, key);
  if (value == null)
    return null;
  return JSON.parse(value);
}

function set_object(spreadsheet, key, value_object) {
  var value = JSON.stringify(value_object);
  set(spreadsheet, key, value);
}

return {
  get: get, unset: unset, set: set,
  get_object: get_object, set_object: set_object };
}(); // end SpreadsheetMetadata namespace


var SheetMetadata = function () { // namespace

function get(sheet, key) {
  var metadata = sheet.createDeveloperMetadataFinder()
    .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.SHEET)
    .withVisibility(SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
    .withKey(key)
    .find();
  var metadatum = metadata[0];
  if (metadatum == null)
    return metadatum;
  return metadatum.getValue();
}

function unset(sheet, key) {
  var metadata = sheet.createDeveloperMetadataFinder()
    .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.SHEET)
    .withVisibility(SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
    .withKey(key)
    .find();
  for (var i = 0; i < metadata.length; ++i) {
    metadata[i].remove();
  }
}

function set(sheet, key, value) {
  unset(sheet, key);
  sheet.addDeveloperMetadata( key, value,
    SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT );
}

function get_object(sheet, key) {
  var value = get(sheet, key);
  if (value == null)
    return null;
  return JSON.parse(value);
}

function set_object(sheet, key, value_object) {
  var value = JSON.stringify(value_object);
  set(sheet, key, value);
}

return {
  get: get, unset: unset, set: set,
  get_object: get_object, set_object: set_object };
}(); // end SheetMetadata namespace

