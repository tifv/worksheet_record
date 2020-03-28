/* A sheet-based database.
 * Columns are fields, rows are entries.
 */

/* Sheet structure:
 * first row contains keys (field names) [k1, k2, …, kn]
 * everything below frozen rows are entries like [v1, v2, …, vn]
 * which are traslated into a sequence of maps {k1 => v1, k2 => v2, …, kn => vn, index: index}
 *   where index is the index in the table
*/

var DataTable = function() { // begin namespace

// "minimal" mode is sufficient to append data and to search
// in "full" mode, search is not getting more efficient

function DataTable(sheet, {required_keys = [], name = null, mode = "full"} = {}) {
  // mode = "full" or "minimal"
  this.sheet = sheet;
  if (name != null) {
    Object.defineProperty(this, "name", {value: name});
  }
  this.mode = (mode == "minimal") ? mode : "full";
  this.column_count = sheet.getMaxColumns();
  var full_values, head_values;
  if (this.mode == "full") {
    SpreadsheetFlusher.add_dimensions(
      false, this.sheet.getSheetId(), null, this.column_count,
      1, 1, this.last_row, this.column_count );
    full_values = this.sheet.getRange(
      1, 1, this.last_row, this.column_count
    ).getValues();
    head_values = full_values[0];
  } else {
    SpreadsheetFlusher.add_dimensions(
      false, this.sheet.getSheetId(), null, this.column_count,
      1, 1, 1, this.column_count );
    head_values = this.sheet.getRange(
      1, 1, 1, this.column_count
    ).getValues()[0];
  }
  this.keys = head_values.map(key => (key != "") ? key : null);
  this.key_columns = revert_keys(this.keys);
  for (let key of required_keys) {
    if (this.key_columns.get(key) == null) {
      throw new Error("DataTable: key \"" + key + "\" is missing");
    }
  }
  if (this.mode == "full") {
    full_values.splice(0, this.first_row - 1);
    this._full_data = [];
    for (let i = 0; i < full_values.length; ++i) {
      this._full_data.push(
        decode_datum.call(this, i, full_values[i]) );
    }
  }
}

define_lazy_property_(DataTable.prototype, "name", function() {
  return this.sheet.getName(); });
define_lazy_property_(DataTable.prototype, "first_row", function() {
  return this.sheet.getFrozenRows() + 1; });
define_lazy_property_(DataTable.prototype, "last_row", function() {
  let last_row = this.sheet.getLastRow();
  if (last_row < this.first_row - 1)
    last_row = this.first_row - 1;
  return last_row;
});
define_lazy_property_(DataTable.prototype, "very_last_row", function() {
  return this.sheet.getMaxRows();
});
define_lazy_property_(DataTable.prototype, "row_count", function() {
  return this.last_row - this.first_row + 1; });

DataTable.prototype.get_range = function(index = null, key = null) {
  var row, row_count, column, column_count;
  if (index == null) {
    row = this.first_row;
    row_count = this.row_count;
  } else {
    row = this.first_row + index;
    row_count = 1;
  }
  if (key == null) {
    column = 1;
    column_count = this.keys.length;
  } else {
    column = this.key_columns.get(key);
    column_count = 1;
  }
  return this.sheet.getRange(row, column, row_count, column_count);
}

DataTable.prototype.toString = function() {
  return "[ data table " + this.sheet.getName() + "]";
}

function revert_keys(keys) {
  var key_columns = new Map();
  for (var i = 0; i < keys.length; ++i) {
    if (keys[i] == null)
      continue;
    key_columns.set(keys[i], i + 1);
  }
  return key_columns;
}

function decode_datum(index, value_row) { // applied to DataTable
  let datum = new Map();
  for (let [key, column] of this.key_columns) {
    datum.set(key, value_row[column-1]);
  }
  datum.index = index;
  return datum;
}

DataTable.prototype[Symbol.iterator] = function() {
  if (this._full_data == null) {
    throw new Error("DataTable: full data not available in this mode");
  }
  return this._full_data[Symbol.iterator]();
}

DataTable.prototype.find = function*(key, value, finder_mod = null) {
  var
    start_row = this.first_row,
    row_count = this.very_last_row -  start_row + 1;
  var search_range;
  if (key == null) {
    search_range = this.sheet.getRange(start_row, 1, row_count, this.column_count);
  } else {
    search_range = this.sheet.getRange(start_row, this.key_columns.get(key), row_count, 1);
  }
  var finder = search_range.createTextFinder(value)
    .matchCase(true)
    .matchEntireCell(true);
  if (finder_mod != null)
    finder = finder_mod(finder);
  var indices = new Set();
  for (let cell of finder.findAll()) {
    indices.add(cell.getRow() - this.first_row);
  }
  indices = Array.from(indices).sort((x, y) => (x - y));
  for (let index of indices) {
    yield decode_datum.call( this,
      index, this.get_range(index, null).getValues()[0] );
  }
}

DataTable.prototype.append = function(datum) {
  this.sheet.appendRow(
    this.keys.map(key => (key != null) ? datum.get(key) : null)
  );
  return this;
}

DataTable.prototype.clear = function(index) {
  if (index == null) {
    this.get_range().clearContent();
  } else {
    this.get_range(index).clearContent();
  }
}

DataTable.init = function(sheet, keys, frozen_rows = 1) {  
  rectify_width: {
    let current_width = sheet.getMaxColumns();
    let width = keys.length;
    if (current_width > width) {
      sheet.deleteColumns(width + 1, current_width - width);
    } else if (current_width < width) {
      sheet.insertColumnsAfter(current_width, width - current_width);
    }
  }
  sheet.getRange(1, 1, 1, keys.length)
    .setValues([keys]);
  sheet.setFrozenRows(frozen_rows);
}

/*
function decode_data(values, offset=0) { // applied to DataTable
  var data = [];
  for (let i = 0; i < this.row_count; ++i) {
    let values_row = values[offset + i];
    let datum = {};
    for (let [key, column] of this.key_columns) {
      datum[key] = values_row[column-1];
    }
    data.push(datum);
  }
  return data;
}
*/

/*
DataTable.prototype.get_length = function(index) {
  return this._data.length;
}

DataTable.prototype.get_datum = function(index) {
  return this._data[index];
}

DataTable.prototype.reload = function() {
  this.row_count = this.sheet.getLastRow() - this.first_row + 1;
  var values = this.get_range().getValues();
  this._data = decode_data.call(this, values);
}
*/

return DataTable;
}(); // end DataTable namespace