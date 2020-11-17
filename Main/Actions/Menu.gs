const emoji = {
    plus:    "\u2795",
    minus:   "\u2796",
    koala:   "\uD83D\uDC28",
    bat:     "\uD83E\uDD87", // U+1F987
    chicken: "\uD83D\uDC24",
    sun:     "\uD83C\uDF1E", // U+1F31E
    moon:    "\uD83C\uDF1D",
    devil:   "\uD83D\uDC7F",
    pizza:   "\uD83C\uDF55", // U+1F355
    cookie:  "\uD83C\uDF6A", // U+1F36A
    cake:    "\uD83E\uDD67", // U+1F967
    snake:   "\uD83D\uDC0D", // U+1F40D
};

const emojipad = Object.fromEntries(Object.entries(emoji)
  .map(([name, value]) => ([name, value + " "])) );

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
        .addItem( "Оглавление (панель)",
            "sidebar_show" )
        .addSeparator();
    { let columns_menu = ui.createMenu("Колонки");
        columns_menu
            .addItem( emojipad.plus    + "Добавить колонки…",
                "action_add_columns" )
            .addItem( emojipad.plus    + "Добавить раздел-добавку…",
                "action_add_section" )
            .addItem( emojipad.minus   + "Удалить лишние колонки",
                "action_remove_excess_columns" )
            .addSeparator()
            .addItem( "Размыть границы подпунктов",
                "action_alloy_subproblems" )
            .addItem( "Отметить задачи как разобр.",
                "action_mark_columns_finished" )
            ;
        menu.addSubMenu(columns_menu);
    }
    { let rows_menu = ui.createMenu("Строки");
        rows_menu
            .addItem( emoji.plus  + " Добавить строки…",
                "action_add_rows" )
            .addItem( emoji.minus + " Удалить лишние строки…",
                "action_remove_excess_rows" )
            ;
        menu.addSubMenu(rows_menu);
    }
    { let worksheets_menu = ui.createMenu("Листочки");
        worksheets_menu
            .addItem( emoji.plus + " Вставить бланк рядом справа",
                "action_worksheet_insert" )
            .addItem( emoji.plus + " Добавить бланк в конец",
                "action_worksheet_add" )
            .addItem( "Перекрасить листочки…",
                "action_worksheet_recolor" )
            .addItem( "Конвертировать в олимпиаду",
                "action_worksheet_convert_to_olympiad" )
            ;
        menu.addSubMenu(worksheets_menu);
    }
    if (UploadConfig.is_configured()) {
        menu.addItem( emojipad.sun   + "Выложить листочек…",
            "action_worksheet_upload" );
        let addendum_menu = ui.createMenu("Выложить другое");
        addendum_menu
            .addItem( emojipad.cookie + "подсказки…",
                "action_worksheet_upload_addendum.hints" )
            .addItem( emojipad.cake   + "ответы…",
                "action_worksheet_upload_addendum.answers" )
            .addItem( emojipad.pizza  + "решения…",
                "action_worksheet_upload_addendum.solutions" );
        menu.addSubMenu(addendum_menu);
    }
    menu.addSeparator();
    if (User.admin_is_acquired()) {
        menu_add_admin_(menu);
    } else {
        menu.addItem( emojipad.snake + "Функции администратора",
            "user_admin_acquire" );
    }
    menu.addToUi();
};

function menu_add_admin_(menu) {
    menu
        .addItem("Метаданные ведомости…", "metadata_editor")
        .addItem("(WIP) Добавить группу…", "action_add_group")
        .addItem("(WIP) Листочки по плану…", "action_worksheet_planned")
        .addItem("Воссоздать toc", "action_regenerate_toc")
        .addItem("Воссоздать uploads", "action_regenerate_uploads")
        .addItem("Обновить меню", "menu_create")
        .addSeparator()
        .addItem( emojipad.snake + "Скрыть функции адм-ра",
            "user_admin_relinquish" )
        ;
}

// XXX add function that sets hyperlink color of the spreadsheet to hsl(220, 75%, 40%)

// vim: set fdm=marker sw=4 :
