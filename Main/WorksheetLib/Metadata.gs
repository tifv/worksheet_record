/* ColorSchemes
 *   .get() → schemes
 *   .get_default() → scheme
 *   .set(schemes)
 * schemes = {`name` : scheme for each `name`}
 * scheme = {mars: [h,s,l], marks_deep: …, rating_mid: …, rating_top: …}
 */

/* Categories
 *   .get() → categories
 *   .set(categories)
 * categories = {`name` : category for each `name`}
 * category = {color: [h,s,l], name: like "алгебра"}
 */


var ColorSchemes = function() { // namespace

const metadata_key = "worksheet_meta-color_schemes";
const colors = ["marks", "marks_deep", "rating_mid", "rating_top"];

function get(spreadsheet) {
  SpreadsheetMetadata.get_object(spreadsheet, metadata_key);
}

function set(schemes) {
  var schemes_object = {};
  for (var key in schemes) {
    var scheme = schemes[key];
    var scheme_object = schemes_object[key] = {};
    scheme_object.marks =
      HSL.to_hsl(scheme.marks) || HSL.to_hsl(default_scheme.marks);
    scheme_object.marks_deep =
      HSL.to_hsl(scheme.marks_deep) || HSL.deepen(scheme_object.marks_deep, 2.0);
    scheme_object.rating_mid =
      HSL.to_hsl(scheme.rating_mid) || HSL.to_hsl(default_scheme.rating_mid);
    scheme_object.rating_top =
      HSL.to_hsl(scheme.rating_top) || HSL.to_hsl(default_scheme.rating_top);
  }
  SpreadsheetMetadata.set_object(spreadsheet, metadata_key, schemes_object);
}

const default_scheme = {
  marks:      [  0, .00, .90],
  marks_deep: [  0, .00, .80],
  rating_mid: [300, .50, .85],
  rating_top: [180, .60, .70]
};

function get_default() {
  var scheme = {};
  for (var i = 0; i < colors.length; ++i) {
    scheme[colors[i]] = HSL.to_hsl(default_scheme[colors[i]]);
  }
  return scheme;
}

return {
  get: get, set: set,
  get_default: get_default };
}(); // end ColorSchemes namespace


var Categories = function() { // namespace

const metadata_key = "worksheet_meta-categories";

function get(spreadsheet) {
  var categories = SpreadsheetMetadata.get_object(spreadsheet, metadata_key);
  if (categories == null)
    return {};
  return categories;
}

function set(spreadsheet, categories) {
  var categories_object = {};
  for (var key in categories) {
    var category = categories[key];
    var category_object = categories_object[key] = {};
    category_object.color = HSL.to_hsl(category.color);
    category_object.name = category.name;
    category_object.filename = category.filename;
  }
  SpreadsheetMetadata.set_object(spreadsheet, metadata_key, categories_object);
}

return {get: get, set: set};
}(); // end Categories namespace

function format_category_css_(categories) {
  const category_list = Object.keys(categories).sort();
  var css_pieces = ["<style>"];
  for (let c of category_list) {
    let colour = categories[c].color;
    if (colour == null)
      continue;
    css_pieces.push(
      ".coloured.category-" + c + " { " +
        "background-color: " + HSL.to_css(colour) + "; " +
        "border-color: " + HSL.to_css(HSL.deepen(colour, 1.50)) + "; "+
      "}"
    )
  }
  css_pieces.push("</style>");
  return css_pieces.join("\n");
}

function get_category_css() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var categories = Categories.get(spreadsheet);
  return format_category_css_(categories);
}
