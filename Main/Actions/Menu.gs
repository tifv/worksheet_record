const emoji = {
    plus:    "\u2795",
    minus:   "\u2796",
    koala:   "\uD83D\uDC28",
    bat:     "\uD83E\uDD87",
    chicken: "\uD83D\uDC24",
    sun:     "\uD83C\uDF1E",
    devil:   "\uD83D\uDC7F",
}

function onOpen() {
    if (User.menu_is_enabled()) {
        menu_create();
    } else {
        var ui = SpreadsheetApp.getUi();
        var menu = ui.createMenu("Ведомость");
        menu.addItem("Включить меню", "menu_enable");
        menu.addToUi();
    }
}

function menu_enable() {
    User.menu_enable();
    menu_create();
}

function menu_create() {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu("Ведомость");
    menu
        .addItem( emoji.koala + " Оглавление (панель)",
            "sidebar_show" )
        .addSeparator();
    { let columns_menu = ui.createMenu("Колонки-задачи");
        columns_menu
            .addItem( emoji.plus    + " Добавить колонки",
                "action_add_columns" )
            .addItem( emoji.plus    + " Добавить раздел (добавку)",
                "action_add_section" )
            .addItem( emoji.minus   + " Удалить лишние колонки",
                "action_remove_excess_columns" )
            .addSeparator()
            .addItem( emoji.bat     + " Размыть границы подпунктов",
                "action_alloy_subproblems" )
            .addItem( emoji.chicken + " Отметить задачи как разобр.",
                "action_mark_columns_finished" )
            ;
        menu.addSubMenu(columns_menu);
    }
    { let rows_menu = ui.createMenu("Строки-участники");
        rows_menu
            .addItem( emoji.plus  + " Добавить строки",
                "action_add_rows" )
            .addItem( emoji.minus + " Удалить лишние строки",
                "action_remove_excess_rows" )
            ;
        menu.addSubMenu(rows_menu);
    }
    { let worksheets_menu = ui.createMenu("Таблички-листочки");
        worksheets_menu
            .addItem( emoji.plus + " Вставить бланк рядом справа",
                "action_worksheet_insert" )
            .addItem( emoji.plus + " Добавить бланк в конец",
                "action_worksheet_add" )
            .addItem( "Перекрасить листочек…",
                "action_worksheet_recolor" )
            ;
        menu.addSubMenu(worksheets_menu);
    }
    menu.addItem( emoji.sun + " Выложить листочек…",
        "upload_worksheet_init" );
    menu.addSeparator();
    if (User.admin_is_acquired()) {
        menu_add_admin_(menu);
    } else {
        menu.addItem( emoji.devil + " Функции администратора",
            "user_admin_acquire" );
    }
    menu.addToUi();
};

function menu_add_admin_(menu) {
    menu
        .addItem("Метаданные ведомости…", "metadata_editor")
        .addItem("Обновить меню", "menu_create")
        .addSeparator()
        .addItem( emoji.devil + " Скрыть функции адм-ра",
            "user_admin_relinquish" )
        ;
}

// XXX add function that sets hyperlink color of the spreadsheet to hsl(220, 75%, 40%)

// vim: set fdm=marker sw=4 :
