function uploads_mark_dead_single_(spreadsheet, seen_groups = []) {
  var ids = new Set();
  var errors = [];
  var record = UploadRecord.get(spreadsheet);
  for (let group of StudyGroup.list(spreadsheet)) {
    if (seen_groups.includes(group.name)) {
      continue;
    }
    seen_groups.push(group.name);
    console.log("Scanning group " + group.name);
    let start = Worksheet.find_start_col(group);
    if (start == null)
      continue; // no worksheets, no op
    let formulas = group.sheetbuf.slice_formulas( "title_row",
      start, group.sheetbuf.dim.sheet_width );
    for (let [i, formula] of formulas.entries()) {
      try {
        if (formula == "")
          continue;
        let formula_decode = decode_hyperlink_formula_(formula);
        if (formula_decode == null)
          throw new Error(
            "Unrecognized formula in " +
            "'" + group.name + "'!" + ACodec.encode(i + start) );
        let [{filter = null}] = formula_decode;
        if (filter == null)
          throw new Error(
            "Unrecognized hyperlink formula in " +
            "'" + group.name + "'!" + ACodec.encode(i + start) );
        ids.add(filter);
      } catch (error) {
        console.error(error);
        errors.push(error);
      }
    }
  }
  var backgrounds = record.get_range(null, "id")
    .getBackgrounds().map(([c]) => c);
  for (let datum of record) {
    if (datum.get("id") == "")
      continue;
    if (ids.has(datum.get("id"))) {
      if (backgrounds[datum.index] != "#ffffff") {
        record.get_range(datum.index, "id")
          .setBackground(null);
      }
      continue;
    }
    record.get_range(datum.index, "id")
      .setBackground("#dddddd");
  }
  if (errors.length > 0) {
    throw errors[0];
  }
}

function uploads_mark_dead() {
  let seen_groups = [];
  const hidden_spreadsheet = HiddenSpreadsheet.get();
  if (hidden_spreadsheet !== null) {
    uploads_mark_dead_single_(hidden_spreadsheet, seen_groups);
  }
  const main_spreadsheet = MainSpreadsheet.get();
  uploads_mark_dead_single_(main_spreadsheet, seen_groups);
}

