function display_error(error) {
  const ui = SpreadsheetApp.getUi();
  ui.alert("Ошибка", error.toString(), ui.ButtonSet.OK);
}

class ReportError extends Error {};

function report_error(error) {
  if (error instanceof ReportError) {
    display_error(error.message);
  } else {
    display_error(error);
    throw error;
  }
}

