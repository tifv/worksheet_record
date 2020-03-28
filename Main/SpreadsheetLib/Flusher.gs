var SpreadsheetFlusher = function() { // begin namespace

var toxicity = 0;
var previous = {};

function add_toxicity(flush, added_toxicity) {
  toxicity += added_toxicity;
  if (flush && toxicity > 10000) {
    SpreadsheetApp.flush();
    toxicity = added_toxicity;
  }
}

function add_dimensions(flush, sheet_id, max_rows, max_cols, start_row, start_col, rows, cols) {
  if (rows > 100 || cols > 100)
    return;
  if (
      sheet_id == previous.sheet_id &&
      start_row >= previous.start_row && start_row + rows <= previous.start_row + 100 &&
      start_col >= previous.start_col && start_col + cols <= previous.start_col + 100 )
  {
    return;
  } else {
    previous = { sheet_id: sheet_id,
      start_row: start_row, start_col: start_col,
      rows: rows, cols: cols };
  }
  var toxicity_rows = max_rows == null ? 100 : max_rows - start_row + 1;
  if (toxicity_rows > 100)
    toxicity_rows = 100;
  var toxicity_cols = max_cols == null ? 100 : max_cols - start_col + 1;
  if (toxicity_cols > 100)
    toxicity_cols = 100;
  add_toxicity(flush, toxicity_rows * toxicity_cols);
}

function add_range(flush, range, sheet) {
  if (sheet == null)
    sheet = range.getSheet();
  add_dimensions( flush,
    sheet.getId(), sheet.getMaxRows(), sheet.getMaxColumns(),
    range.getRow(), range.getColumn(), range.getNumRows(), range.getNumColumns()
  );
}

function reset() {
  toxicity = 0;
}

return {add_dimensions: add_dimensions, add_range: add_range, reset: reset};
}(); // end SpreadsheetFlusher namespace
