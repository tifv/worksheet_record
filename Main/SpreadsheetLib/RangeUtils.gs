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

function set_fixed_value_validation_(range, value) {
  range
    .setDataValidation( SpreadsheetApp.newDataValidation()
      .requireTextEqualTo(value)
      .setHelpText(
        value != "" ?
          "This cell must contain a marker (" + value + ")" :
          "This cell must remain blank"
      )
      .setAllowInvalid(false).build() );
}

