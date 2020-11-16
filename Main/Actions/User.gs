/* Amdministrative priviliges
 * • anyone can acquire administrative priviliges
 *   • the only requirement is to answer the question
 *   • spreadsheet owner can acquire privileges without any questions
 * • anyone can relinquish administrative priviliges
 * • some menu items will be hidden unless priviliges are enabled
 */

var User = function() { // begin namespace

const user_status_key = "user_status";
var user_status = null;

function get_user_status() {
    if (user_status == null) {
        user_status = JSON.parse(
            PropertiesService.getUserProperties()
                .getProperty(user_status_key) ||
            "{}" );
    }
    return user_status;
}

function save_user_status() {
    PropertiesService.getUserProperties()
        .setProperty(user_status_key, JSON.stringify(get_user_status()));
}

function admin_is_acquired() {
    return get_user_status().admin == "true";
}

function admin_acquire() {
    get_user_status().admin = "true";
    save_user_status();
}

function admin_relinquish() {
    get_user_status().admin = "false";
    save_user_status();
}

function menu_is_enabled() {
    return get_user_status().enabled == "true";
}

function menu_enable() {
    get_user_status().enabled = "true";
    save_user_status();
}

return {
    admin_is_acquired: admin_is_acquired,
    admin_acquire: admin_acquire,
    admin_relinquish: admin_relinquish,
    menu_is_enabled: menu_is_enabled,
    menu_enable: menu_enable,
}
}(); // end User namespace

function user_admin_acquire() {
    const password = "я хочу сломать ведомость";
    var user = Session.getActiveUser();
    var email = user.getEmail();
    if ( email == "" ||
        email != SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail()
    ) {
        const ui = SpreadsheetApp.getUi();
        let response_btn = ui.alert( "Функции администратора",
            "Пароль к функциям администратора: " +
            "«" + password + "».",
            ui.ButtonSet.OK_CANCEL );
        if (response_btn != ui.Button.OK) {
            return;
        }
        let response = ui.prompt( "Функции администратора",
            "Введите пароль:",
            ui.ButtonSet.OK_CANCEL );
        if (
            response.getSelectedButton() != ui.Button.OK ||
            response.getResponseText() != password
        ) {
            return;
        }
    }
    User.admin_acquire();
    menu_create();
}

function user_admin_relinquish() {
    User.admin_relinquish();
    menu_create();
}

// vim: set fdm=marker sw=4 :
