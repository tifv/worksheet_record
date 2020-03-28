// Get unbounded column range from bounded range.
function get_column_range_(range) {
  var col_ref = range.getA1Notation().replace(/[0-9]+/g, "");
  if (col_ref.indexOf(":") < 0)
    col_ref = col_ref + ":" + col_ref;
  return range.getSheet().getRange(col_ref);
}

// Get unbounded row range from bounded range.
function get_row_range_(range) {
  var row_ref = range.getA1Notation().replace(/[A-Z]+/g, "");
  if (row_ref.indexOf(":") < 0)
    row_ref = row_ref + ":" + row_ref;
  return range.getSheet().getRange(row_ref);
}

function set_marker_guard_(range, value) {
  // XXX remove setting the value from here
  range
    .setValue(value)
    .setDataValidation( SpreadsheetApp.newDataValidation()
      .requireTextEqualTo(value)
      .setHelpText("This cell must contain a marker (" + value + ")")
      .setAllowInvalid(false).build() );
}

function set_blank_guard_(range) {
  range.setDataValidation( SpreadsheetApp.newDataValidation()
    .requireTextEqualTo("")
    .setHelpText("This cell must be blank")
    .setAllowInvalid(false).build() );
}

