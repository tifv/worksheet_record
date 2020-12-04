const emoji = {
    plus:    "\u{2795}",
    minus:   "\u{2796}",
    koala:   "\uD83D\uDC28",
    bat:     "\u{1F987}",
    chicken: "\uD83D\uDC24",
    sun:     "\u{1F31E}",
    moon:    "\uD83C\uDF1D",
    devil:   "\uD83D\uDC7F",
    nut:     "\u{1F330}",
    cookie:  "\u{1F36A}",
    cake:    "\u{1F967}",
    pizza:   "\u{1F355}",
    snake:   "\u{1F40D}",
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
            .addItem( emojipad.plus    + "Добавить раздел-добавку…",
                "action_add_section" )
            .addSeparator()
            .addItem( emoji.plus + " Вставить бланк рядом справа",
                "action_worksheet_insert" )
            .addItem( emoji.plus + " Добавить бланк в конец",
                "action_worksheet_add" )
            .addSeparator()
            .addItem( "Перекрасить листочки…",
                "action_worksheet_recolor" )
            .addItem( "Конвертировать в олимпиаду",
                "action_worksheet_convert_to_olympiad" )
            ;
        menu.addSubMenu(worksheets_menu);
    }
    if (UploadConfig.is_configured()) {
        let upload_menu = ui.createMenu("Выложить");
        upload_menu
            .addItem( emojipad.sun   + "листочек…",
                "action_worksheet_upload" );
        action_worksheet_upload_addendum.populate_menu(upload_menu, true);
        menu.addSubMenu(upload_menu);
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

function menu_meta_setup_(mainfunc, metadata_key) {
    var mainname = mainfunc.name;
    function load_list() {
        var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        var menu_data = SpreadsheetMetadata.get_object(spreadsheet, metadata_key);
        if (menu_data == null)
            return [];
        return menu_data;
    }
    function load_object() {
        return Object.fromEntries(load_list().map(
            ([name, arg, label,]) => [name, [arg, label]] ));
    }
    mainfunc.populate_menu = function(menu, add_separator) {
        var menu_data = load_list();
        if (menu_data.length == 0)
            return false;
        if (add_separator)
            menu.addSeparator();
        for (let [name, arg, label,] of menu_data) {
            menu.addItem(label, mainname + ".dispatch." + name);
        }
        return true;
    }
    mainfunc.dispatch = new Proxy(mainfunc, {get: function(mainfunc, name) {
        var menu_data = load_object();
        if (menu_data[name] == null) {
            throw new Error("no such function: " + mainname + ".dispatch." + name);
        }
        let [arg,] = menu_data[name];
        return () => { mainfunc(arg); }
    }});
    return mainfunc;
}

// XXX add function that sets hyperlink color of the spreadsheet to hsl(220, 75%, 40%)

// vim: set fdm=marker sw=4 :
