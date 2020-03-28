function action_insert_worksheet() {
  var worksheet = ActionHelpers.get_active_worksheet();
  if (worksheet == null) {
    return;
  }
  try {
    var note_info = Worksheet.parse_title_note(worksheet.get_title_note());
    Worksheet.add(
      worksheet.group,
      worksheet.sheet.getRange(1, worksheet.dim.end + 1),
      {date: note_info.date} );
    worksheet.sheet.getParent().toast(
      "Исправьте дату в примечании к заголовку таблички, если требуется." );
  } catch (error) {
    report_error(error);
    return;
  }
}

function action_add_worksheet() {
  var group = ActionHelpers.get_active_group();
  if (group == null) {
    return;
  }
  var sheet = group.sheet;
  try {
    var last_column = sheet.getLastColumn();
    if (last_column == group.dim.sheet_width) {
      throw "Последний столбец вкладки должен быть пустым.";
    }
    Worksheet.add(group, sheet.getRange(1, last_column + 1));
    group.sheet.getParent().toast(
      "Исправьте дату в примечании к заголовку таблички, если требуется." );
  } catch (error) {
    report_error(error);
    return;
  }
}