function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu("Ведомость");
  menu
    .addItem("Включить меню", "create_menu");
  menu.addToUi();
}

const emoji = {
  plus:    "\u2795",
  minus:   "\u2796",
  koala:   "\uD83D\uDC28",
  bat:     "\uD83E\uDD87",
  chicken: "\uD83D\uDC24",
  sun:     "\uD83C\uDF1E",
  devil:   "\uD83D\uDC7F",
}

function create_menu() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu("Ведомость");
  { let test_menu = ui.createMenu(emoji.devil + " Тест");
    test_menu
      .addItem("Редактор метаданных", "metadata_editor")
      .addItem("Обновить меню", "create_menu")
      ;
    menu
//      .addSeparator()
      .addSubMenu(test_menu);
  }
  menu
    .addItem(emoji.koala + " Оглавление (панель)", "sidebar_show")
    .addSeparator();
  { let columns_menu = ui.createMenu("Колонки-задачи");
    columns_menu
      .addItem(emoji.plus + " Добавить колонки", "action_add_columns")
      .addItem(emoji.plus + " Добавить раздел (добавку)", "action_add_section")
      .addItem(emoji.minus + " Удалить лишние колонки", "action_remove_excess_columns")
      .addSeparator()
      .addItem(emoji.bat + " Размыть границы подпунктов", "action_alloy_subproblems")
      .addItem(emoji.chicken + " Отметить задачи как разобранные", "action_mark_columns_finished")
      ;
    menu.addSubMenu(columns_menu);
  }
  { let rows_menu = ui.createMenu("Строки-участники");
    rows_menu
      .addItem(emoji.plus + " Добавить строки", "action_add_rows")
      .addItem(emoji.minus + " Удалить лишние строки", "action_remove_excess_rows")
      ;
    menu.addSubMenu(rows_menu);
  }
  { let worksheets_menu = ui.createMenu("Таблички-листочки");
    worksheets_menu
      .addItem(emoji.plus + " Вставить бланк рядом справа", "action_insert_worksheet")
      .addItem(emoji.plus + " Добавить бланк в конец", "action_add_worksheet")
      ;
    menu.addSubMenu(worksheets_menu);
  }
  menu.addItem(emoji.sun + " Выложить листочек…", "upload_worksheet_init");
  menu.addToUi();
};

