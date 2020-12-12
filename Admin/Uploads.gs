// copied from Main
function decode_hyperlink_formula_(formula) {
  var hyperlink_filter_match = /^=(?:hyperlink|HYPERLINK)\((?:filter|FILTER)\((?:uploads|'uploads')!R\d+C\d+:C\d+[,;](?:uploads|'uploads')!R\d+C\d+:C\d+="([^"]*)"\)[,;]"([^"]*)"\)$/
    .exec(formula);
  if (hyperlink_filter_match != null) {
    return [{filter: hyperlink_filter_match[1]}, hyperlink_filter_match[2]];
  }
  var hyperlink_match = /^=\s*hyperlink\s*\(\s*"([^"]*)"\s*[,;]\s*"([^"]*)"\s*\)\s*$/i
    .exec(formula);
  if (hyperlink_match != null) {
    return [{url: hyperlink_match[1]}, hyperlink_match[2]];
  }
  return null;
}

// adapted from Main
var UploadRecord = function() { // begin namespace

const required_keys = [
  "group", "category", "title", "date", "uploader", "author",
  "id",
  "pdf", "src", "initial_pdf", "initial_src",
  "stable_pdf", "stable_src", "filename",
  "archive_pdf", "archive_src", "archive_target", "archive_mtime",
  "request", "argument", "status" ];

const sheet_name = "uploads";

function exists() {
  var uploads_sheet = MainSpreadsheet.get().getSheetByName(sheet_name);
  return (uploads_sheet != null);
}

function get(mode = "full") {
  var uploads_sheet = MainSpreadsheet.get().getSheetByName(sheet_name);
  if (uploads_sheet == null)
    return null;
  return new DataTable(uploads_sheet, {required_keys: required_keys, name: sheet_name, mode: mode});
}

return {
  exists: exists, get: get };
}(); // end UploadRecord namespace

function uploads_mark_dead() {
  var spreadsheet = MainSpreadsheet.get();
  var ids = new Set();
  var errors = [];
  var record = UploadRecord.get();
  for (let group of StudyGroup.list(spreadsheet)) {
    let start = Worksheet.find_start_col(group);
    let formulas = group.sheetbuf.slice_formulas("title_row", start, group.sheetbuf.dim.sheet_width);
    for (let [i, formula] of formulas.entries()) {
      try {
        if (formula == "")
          continue;
        let formula_decode = decode_hyperlink_formula_(formula);
        if (formula_decode == null)
          throw new Error("Unrecognized formula in '" + group.name + "'!" + ACodec.encode(i + start));
        let [{filter = null}] = formula_decode;
        if (filter == null)
          throw new Error("Unrecognized hyperlink formula in '" + group.name + "'!" + ACodec.encode(i + start));
        ids.add(filter);
      } catch (error) {
        console.error(error);
        errors.push(error);
      }
    }
  }
  var backgrounds = record.get_range(null, "id").getBackgrounds().map(([c]) => c);
  for (let datum of record) {
    if (ids.has(datum.get("id"))) {
      if (backgrounds[datum.index] != "#ffffff") {
        record.get_range(datum.index, "id").setBackground(null);
      }
      continue;
    }
    record.get_range(datum.index, "id").setBackground("#dddddd");
  }
  if (errors.length > 0) {
    throw errors[0];
  }
}

