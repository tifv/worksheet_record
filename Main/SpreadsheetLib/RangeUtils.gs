var ACodec = {
  encode: function(n) {
    var pieces = [];
    while (n > 0) {
      var i = (n - 1) % 26;
      pieces.unshift(letters[i]);
      n = (n - i - 1) / 26;
    }
    return pieces.join("");
  },
  decode: function(s) {
    var n = 0;
    while (s.length > 0) {
      n *= 26;
      n += letters.indexOf(s[0]) + 1;
      s = s.substring(1);
    }
    return n;
  },
};

// Get unbounded row range
function get_row_range_(sheet, row, last_row = row) {
  return sheet.getRange(row.toString() + ":" + last_row.toString());
}

// Get unbounded column range
function get_column_range_(sheet, column, width = 1) {
  return sheet.getRange(
    ACodec.encode(column) + ":" + ACodec.encode(column + width - 1) );
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

