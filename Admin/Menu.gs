function onOpen() {
    if (MainSpreadsheet.is_set()) {
        menu_create();
    } else {
        var ui = SpreadsheetApp.getUi();
        var menu = ui.createMenu("Ведомость");
        menu.addItem("Подключить ведомость", "connect_main");
        menu.addToUi();
    }
}

function connect_main() {
    const ui = SpreadsheetApp.getUi();
    let response = ui.prompt( "Подключить ведомость",
        "Введите ID или URL ведомости:",
        ui.ButtonSet.OK_CANCEL );
    if (response.getSelectedButton() != ui.Button.OK) {
        return;
    }
    var ref = response.getResponseText();
    var spreadsheet;
    if (/\//.exec(ref) != null) {
        spreadsheet = SpreadsheetApp.openByUrl(ref);
    } else {
        spreadsheet = SpreadsheetApp.openById(ref);
    }
    MainSpreadsheet.set(spreadsheet);
    menu_create();
}

function menu_create() {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu("Ведомость");
    menu
        .addItem("Обновить меню", "menu_create")
    menu.addToUi();
};

// vim: set fdm=marker sw=4 :
