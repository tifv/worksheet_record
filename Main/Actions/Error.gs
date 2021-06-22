function display_error(error) {
  const ui = SpreadsheetApp.getUi();
  ui.alert("Ошибка", error.toString(), ui.ButtonSet.OK);
}

class ReportError extends Error {};
ReportError.ALREADY_REPORTED = new ReportError();

ReportError.with_reporting = function(operator) {
  try {
    return operator();
  } catch (error) {
    if (error instanceof ReportError) {
      if (error === ReportError.ALREADY_REPORTED)
        return;
      display_error(error.message);
    } else {
      display_error(error);
      throw error;
    }
  }
}

ReportError.standard = {
  StudyGroupDetection: function() {
    return new ReportError(
      "Выбранная вкладка не соответствует учебной группе. " +
      "Выберите вкладку таблицы, соответствующую группе." );
  },
  StudyGroupNoMetadata: function() {
    return new ReportError(
      "Выбранная вкладка не отмечена как учебная группа." );
  },
  StudyGroupNoFrozenRows: function() {
    return new ReportError(
      "Выбранная вкладка не несёт закрепленных строк, " +
      "необходимых для структуры группы." );
  },
  WorksheetDetection: function() {
    return new ReportError(
      "Не удалось определить листочек. " +
      "Выберите диапазон целиком внутри листочка (например, заголовок)." );
  },
  WorksheetSectionDetection: function() {
    return new ReportError(
      "Не удалось определить раздел листочка. " +
      "Выберите диапазон внутри одного раздела (например, заголовок)." );
  },
};
