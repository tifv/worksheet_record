/* SpreadsheetMetadata
 *   .get(spreadsheet, key)
 *   .unset(spreadsheet, key)
 *   .set(spreadsheet, key, value)
 *   .get_object(spreadsheet, key) → JSON
 *   .set_object(spreadsheet, key, value_object)
 */

var SpreadsheetMetadata = function () { // namespace

function get(spreadsheet, key) {
  if (typeof key != "string")
    throw new Error("type error: key must be a string");
  var [metadatum] = spreadsheet.createDeveloperMetadataFinder()
    .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.SPREADSHEET)
    .withVisibility(SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
    .withKey(key)
    .find();
  if (metadatum == null)
    return null;
  return metadatum.getValue();
}

function unset(spreadsheet, key) {
  if (typeof key != "string")
    throw new Error("type error: key must be a string");
  spreadsheet.createDeveloperMetadataFinder()
    .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.SPREADSHEET)
    .withVisibility(SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
    .withKey(key)
    .find()
    .forEach(metadatum => { metadatum.remove(); });
}

function set(spreadsheet, key, value) {
  if (typeof key != "string")
    throw new Error("type error: key must be a string");
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
  if (typeof key != "string")
    throw new Error("type error: key must be a string");
  var [metadatum] = sheet.createDeveloperMetadataFinder()
    .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.SHEET)
    .withVisibility(SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
    .withKey(key)
    .find();
  if (metadatum == null)
    return null;
  return metadatum.getValue();
}

function unset(sheet, key) {
  if (typeof key != "string")
    throw new Error("type error: key must be a string");
  sheet.createDeveloperMetadataFinder()
    .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.SHEET)
    .withVisibility(SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
    .withKey(key)
    .find()
    .forEach(metadatum => { metadatum.remove(); });
}

function set(sheet, key, value) {
  if (typeof key != "string")
    throw new Error("type error: key must be a string");
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


var SheetMetacell = function () { // namespace

function get(sheet, key) {
  if (typeof key != "string")
    throw new Error("type error: key must be a string");
  var [row_metadatum] = sheet.createDeveloperMetadataFinder()
    .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.ROW)
    .withVisibility(SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
    .withKey(key)
    .find();
  if (row_metadatum == null)
    return null;
  var row = row_metadatum.getLocation().getRow().getRow();
  var [col_metadatum] = sheet.createDeveloperMetadataFinder()
    .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.COLUMN)
    .withVisibility(SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
    .withKey(key)
    .find();
  if (col_metadatum == null)
    return null;
  var col = col_metadatum.getLocation().getColumn().getColumn();
  return sheet.getRange(row, col);
}

function unset(sheet, key) {
  if (typeof key != "string")
    throw new Error("type error: key must be a string");
  [
    ...sheet.createDeveloperMetadataFinder()
      .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.ROW)
      .withVisibility(SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
      .withKey(key)
      .find(),
    ...sheet.createDeveloperMetadataFinder()
      .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.COLUMN)
      .withVisibility(SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
      .withKey(key)
      .find(),
  ].forEach(metadatum => { metadatum.remove(); });
}

function set(sheet, key, range) {
  if (typeof key != "string")
    throw new Error("type error: key must be a string");
  if (range.getNumRows() > 1 || range.getNumColumns() > 1) {
    throw new Error("range is not a cell");
  }
  unset(sheet, key);
  get_column_range_(sheet, range.getColumn()).addDeveloperMetadata( key,
    SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT );
  get_row_range_(sheet, range.getRow()).addDeveloperMetadata( key,
    SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT );
}

return {
  get: get, unset: unset, set: set };
}(); // end SheetMetadata namespace

