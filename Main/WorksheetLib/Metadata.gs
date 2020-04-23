/* ColorSchemes
 *   .get(ss) → schemes
 *   .set(ss, schemes)
 *   .get_default() → scheme
 *   .copy(scheme) → scheme
 * schemes = {`name` : scheme for each `name`}
 * scheme = {marks: hsl, rating_mid: hsl, rating_top: hsl}
 */

/* Categories
 *   .get() → categories
 *   .set(categories)
 * categories = {`name` : category for each `name`}
 * category = {color: [h,s,l], name: like "алгебра"}
 */


var ColorSchemes = function() { // namespace

const metadata_key = "worksheet_meta-color_schemes";
const colors = ["marks", "rating_mid", "rating_top"];

function get(spreadsheet) {
  var color_schemes = SpreadsheetMetadata.get_object(spreadsheet, metadata_key);
  if (color_schemes == null)
    return {};
  return multicopy(color_schemes);
}

function set(spreadsheet, schemes) {
  SpreadsheetMetadata.set_object( spreadsheet, metadata_key,
    multicopy(schemes) );
}

function copy(scheme) {
  return {
    marks:      HSL.copy(scheme.marks),
    rating_mid: HSL.copy(scheme.rating_mid),
    rating_top: HSL.copy(scheme.rating_top),
  }
}

function multicopy(schemes) {
  return Object.fromEntries( Object.entries(schemes)
    .map(([key, scheme]) => [key, copy(scheme)])
  );
}

const default_scheme = {
  marks:      {h:   0, s: 0.00, l: 0.90},
  rating_mid: {h: 300, s: 0.50, l: 0.85},
  rating_top: {h: 180, s: 0.60, l: 0.70},
};

function get_default() {
  return {
    marks:      HSL.copy(default_scheme.marks),
    rating_mid: HSL.copy(default_scheme.rating_mid),
    rating_top: HSL.copy(default_scheme.rating_top),
  };
}

return {
  get: get, set: set,
  get_default: get_default, copy: copy };
}(); // end ColorSchemes namespace


var Categories = function() { // namespace

const metadata_key = "worksheet_meta-categories";

function get(spreadsheet) {
  var categories = SpreadsheetMetadata.get_object(spreadsheet, metadata_key);
  if (categories == null)
    return {};
  return multicopy(categories);
}

function set(spreadsheet, categories) {
  SpreadsheetMetadata.set_object( spreadsheet, metadata_key,
    multicopy(categories) );
}

function copy(category) {
  return {
    name: category.name,
    filename: category.filename,
    color: category.color != null ? HSL.copy(category.color) : null
  };
}

function multicopy(categories) {
  return Object.fromEntries( Object.entries(categories)
    .map(([key, category]) => [key, copy(category)])
  );
}

return {get: get, set: set};
}(); // end Categories namespace

function format_category_css_(categories) {
  const category_list = Object.keys(categories).sort();
  var css_pieces = ["<style>"];
  for (let c of category_list) {
    let color = categories[c].color;
    if (color == null)
      continue;
    css_pieces.push(
      ".coloured.category-" + c + " { " +
        "background-color: " + HSL.to_css(color) + "; " +
        "border-color: " + HSL.to_css(HSL.deepen(color, 1.50)) + "; "+
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
