/*

var Admin = function() { // begin namespace

const user_is_admin_key_base = "is_admin";
const document_has_admin_key = "has_admin";

function get_user_is_admin_key() {
  return user_is_admin_key_base + "/" + SpreadsheetApp.getActiveSpreadsheet().getId();
}

function set_self_admin() {
  PropertiesService.getUserProperties().setProperty(
    get_user_is_admin_key(), "true" );
  PropertiesService.getDocumentProperties().setProperty(
    document_has_admin_key, "true" );
}

function unset_self_admin() {
  PropertiesService.getUserProperties().deleteProperty(
    get_user_is_admin_key() );
}

function admin_exists() {
  return PropertiesService.getDocumentProperties().getProperty(document_has_admin_key) != null;
}

function self_is_admin() {
  return PropertiesService.getUserProperties().getProperty(
    get_user_is_admin_key() ) != null;
}

function self_is_owner() {
  var user = Session.getActiveUser();
  var email = user.getEmail();
  if (email == "") {
    throw new Error("Admin.self_is_owner: unable to determine");
  }
  return (email == SpreadsheetApp.getActiveSpreadsheet().getOwner().getEmail());
}

return {
  admin_exists: admin_exists, self_is_admin: self_is_admin,
  self_is_owner: self_is_owner,
  set_self_admin: set_self_admin, unset_self_admin: unset_self_admin,
};
}(); // end Admin namespace

*/