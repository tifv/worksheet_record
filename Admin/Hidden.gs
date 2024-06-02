function hidden_sync_now() {
  const main_spreadsheet = MainSpreadsheet.get();
  const hidden_spreadsheet = HiddenSpreadsheet.get();
  for (let group of StudyGroup.list(hidden_spreadsheet)) {
    hidden_sync_sheet_(
      hidden_spreadsheet,
      main_spreadsheet,
      group.name );
  }
}

var hidden_sync = new Scheduler(
  function hidden_sync() {
    hidden_sync_now();
  },
  "hidden_sync.schedule",
  function generate_schedule() {
    return [
      [09, 15],
      [11, 45],
      [14, 15],
      [16, 45],
      [19, 15],
      [23, 45],
    ].map(([h, m]) => {
      let date = new Date();
      date.setHours(h);
      date.setMinutes(m);
      date.setSeconds(0);
      return {date, args: []};
    });
  },
);

function hidden_sync_forever() {
  hidden_sync.never();
  ScriptApp.newTrigger("hidden_sync.today")
    .timeBased()
      .everyDays(1)
      .atHour(1)
      .nearMinute(45)
    .create();
}


var HiddenSpreadsheet = function() {

const property_key = "hidden_spreadsheet";

var spreadsheet = null;

function get() {
  if (spreadsheet != null)
    return spreadsheet;
  var id = PropertiesService.getDocumentProperties()
    .getProperty(property_key);
  if (id == null)
    return null;
  try {
    spreadsheet = SpreadsheetApp.openById(id);
    return spreadsheet;
  } catch (error) {
    return null;
  }
}

function set(spreadsheet) {
  PropertiesService.getDocumentProperties()
    .setProperty(property_key, spreadsheet.getId());
}

function is_set() {
  if (spreadsheet != null)
    return true;
  var id = PropertiesService.getDocumentProperties()
    .getProperty(property_key);
  if (id == null)
    return false;
  return true;
}

return {get: get, set: set, is_set: is_set};
}();

function connect_hidden() {
  const ui = SpreadsheetApp.getUi();
  let response = ui.prompt( "Подключить скрытую ведомость",
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
  HiddenSpreadsheet.set(spreadsheet);
}


function hidden_sync_sheet_(src_spreadsheet, dst_spreadsheet, name) {
  let src_uploads = hidden_load_uploads_(src_spreadsheet);
  let src_group = StudyGroup.find_by_name(src_spreadsheet, name);
  let src_codes = hidden_load_codes_(src_group);
  let data = hidden_load_group_data_(src_group, src_codes, src_uploads);
  let dst_group = StudyGroup.find_by_name(dst_spreadsheet, name);
  let dst_codes = hidden_load_codes_(dst_group);
  hidden_sync_group_data_(dst_group, dst_codes, data);
}

function hidden_load_group_data_(group, codes, uploads) {
  let sheetbuf = new SheetBuffer(group.sheet,
    Object.fromEntries(
      StudyGroupDim.row_names
        .map(k => [k, group.dim[k]])
        .filter(([k, v]) => v != null)
        .concat(
          codes
            .map((code, i) => ([code, i]))
            .filter(([code]) => typeof code == 'number')
            .map(([code, i]) => (['code_' + code, group.dim.data_row + i]))
        )
    ),
    group.dim );
  Object.defineProperty( group, 'sheetbuf',
    {configurable: true, value: sheetbuf} );
  let worksheet_data = [];
  for (let worksheet of Worksheet.list(group)) {
    if (worksheet.is_unused()) {
      continue;
    }
    let worksheet_datum = {};
    worksheet_datum.id = worksheet.get_title_metadata_id({validate: true});
    if (worksheet_datum.id == null) {
      continue; // XXX error?
    }
    worksheet_datum.date = worksheet.get_title_note_data().get("date");
    worksheet_datum.category = worksheet.get_category();
    let sections_data = worksheet_datum.sections = [];
    for (let section of worksheet.list_sections()) {
      let section_datum = {};
      section_datum.id = section.get_title_metadata_id({
        validate: section.dim.offset > 0 });
      if (section_datum.id == null) {
        continue;
      }
      section_datum.title = section.get_title();
      section_datum.date = section.get_title_note_data().get("date");
      section_datum.labels = {
        values: sheetbuf.slice_values( "label_row",
          section.dim.data_start, section.dim.data_end ),
        colors: sheetbuf.slice_colors( "label_row",
          section.dim.data_start, section.dim.data_end ),
      };
      section_datum.category = worksheet_datum.category;
      parse_formula: {
        let title_formula = section.get_title_formula();
        if (title_formula == "")
          break parse_formula;
        section_datum.title_formula = title_formula;
        let title_formula_decode = decode_hyperlink_formula_(title_formula);
        if (title_formula_decode == null)
          break parse_formula;
        let [{filter = null, url = null}, ] = title_formula_decode;
        if (filter != null) {
          section_datum.title_link = uploads[filter];
        } else if (url != null) {
          section_datum.title_link = url;
        }
      }
      function get_code_data(code) {
        return {
          values: sheetbuf.slice_values(
            'code_' + code,
            section.dim.data_start,
            section.dim.data_end,
          ),
          colors: sheetbuf.slice_colors(
            'code_' + code,
            section.dim.data_start,
            section.dim.data_end,
          ),
        };
      }
      section_datum.data = Object.fromEntries(
        codes
          .filter(code => typeof code == 'number')
          .map(code => [code, get_code_data(code)])
      );
      sections_data.push(section_datum);
    }
    worksheet_data.push(worksheet_datum);
  }
  return {worksheets: worksheet_data};
}

function hidden_load_uploads_(spreadsheet) {
  let uploads = {};
  let upload_record = UploadRecord.get(spreadsheet, "full");
  if (upload_record == null) {
    return uploads;
  }
  for (let datum of upload_record) {
    uploads[datum.get('id')] = datum.get('pdf');
  }
  return uploads;
}

function hidden_load_codes_(group) {
  let code_col = group.sheetbuf.find_value("label_row", "код", 1, 100);
  if (code_col == null) {
    throw new Error("A group does not contain codes");
  }
  let codes = group.sheet.getRange(
    group.dim.data_row, code_col, group.dim.data_height
  ).getValues().map(([a]) => a);
  return codes;
}

function hidden_sync_group_data_(group, codes, data) {
  let sheetbuf = new SheetBuffer(group.sheet,
    Object.fromEntries(
      StudyGroupDim.row_names
        .map(k => [k, group.dim[k]])
        .filter(([k, v]) => v != null)
        .concat(
          codes
            .map((code, i) => ([code, i]))
            .filter(([code]) => typeof code == 'number')
            .map(([code, i]) => (['code_' + code, group.dim.data_row + i]))
        )
    ),
    group.dim );
  Object.defineProperty( group, 'sheetbuf',
    {configurable: true, value: sheetbuf} );
  let worksheet_data = data.worksheets;
  let worksheets = Array.from(Worksheet.list(group)).reverse();
  for (let worksheet of worksheets) {
    if (worksheet.is_unused()) {
      sheetbuf.delete_columns(worksheet.dim.start, worksheet.dim.width + 1);
      continue;
    }
    let id = parseInt(worksheet.get_title_note_data().get('import-id'), 10);
    let [src_worksheet_index] = worksheet_data.map((datum, index) => {
      if (datum.id == id) { return index } else { return null; }
    }).filter(id => id != null);
    if (src_worksheet_index == null) {
      sheetbuf.delete_columns(worksheet.dim.start, worksheet.dim.width + 1);
      continue;
    }
    let [worksheet_datum] = worksheet_data.splice(src_worksheet_index, 1);
    if (worksheet.get_category() != worksheet_datum.category) {
      worksheet.set_category(worksheet_datum.category);
    }
    let section_data = worksheet_datum.sections;
    hidden_sync_worksheet_data_(worksheet, codes, section_data);
  }
  for (let worksheet_datum of worksheet_data) {
    let worksheet = WorksheetBuilder.build( group,
      group.sheet.getRange(1, group.sheet.getLastColumn() + 1),
      Object.assign({
        data_width: worksheet_datum.sections[0].labels.values.length,
        date: worksheet_datum.date,
        category: worksheet_datum.category,
      }),
    );
    let title_note_data = worksheet.get_title_note_data();
    title_note_data.set('import-id', worksheet_datum.id);
    worksheet.set_title_note_data(title_note_data);
    hidden_sync_worksheet_data_( worksheet,
      codes, worksheet_datum.sections,
      {overwrite: true},
    );
  }
}

function hidden_sync_worksheet_data_(worksheet, codes, section_data, options = {}) {
  let group = worksheet.group;
  let sheetbuf = group.sheetbuf;
  let {overwrite = false} = options;
  let sections = Array.from(worksheet.list_sections()).reverse();
  for (let section of sections) {
    let id = parseInt(section.get_title_note_data().get('import-id'), 10);
    let [src_section_index] = section_data.map((datum, index) => {
      if (datum.id == id) { return index } else { return null; }
    }).filter(id => id != null);
    if (src_section_index == null) {
      section.set_title("{" + section.get_title() + "}");
      continue;
    };
    let [section_datum] = section_data.splice(src_section_index, 1);
    if (section_datum.title_link != null) {
      let title_formula = (
        '=hyperlink("' + section_datum.title_link + '", ' +
        '"' + section_datum.title.replace('"', '""') + '")'
      );
      if (title_formula != section.get_title_formula()) {
        section.set_title_formula(title_formula);
      }
    } else if (section_datum.title_formula != null) {
      let title_formula = section_datum.title_formula;
      if (title_formula != section.get_title_formula()) {
        section.set_title_formula(title_formula);
      }
    } else {
      if (section_datum.title != section.get_title()) {
        section.set_title(section_datum.title);
      }
    }
    let labels = sheetbuf.slice_values( "label_row",
      section.dim.data_start, section.dim.data_end );
    let width_delta = section_datum.labels.values.length - labels.length;
    if (overwrite && width_delta >= 0) {
      if (width_delta > 0) {
        section = section.add_columns(1, width_delta);
        worksheet = section;
      }
      hidden_write_section_data_(section, codes, section_datum);
    } else {
      for (let edit_item of hidden_edit_sequence_(labels, section_datum.labels.values)) {
        if (edit_item.add_start != null) {
          let add_width = edit_item.dest_end - edit_item.dest_start;
          section = section.add_columns(edit_item.add_start, add_width);
          worksheet = section.worksheet;
          sheetbuf.set_values( "label_row",
            section.dim.data_start + edit_item.add_start,
            section.dim.data_start + edit_item.add_start + add_width - 1,
            section_datum.labels.values.slice(edit_item.dest_start, edit_item.dest_end) );
        } else if (edit_item.remove_start != null) {
          section.remove_columns((index) => (
            edit_item.remove_start <= index &&
            index < edit_item.remove_end
          ));
          section = Worksheet.Section.surrounding( group, null,
            group.sheet.getRange(1, section.dim.start) );
          worksheet = section.worksheet;
        } else {
          throw new Error("unreachable");
        }
      }
      hidden_sync_section_data_(section, codes, section_datum);
    }
  }
  for (let section_datum of section_data) {
    let last_section = Worksheet.Section.surrounding( group, null,
      group.sheet.getRange(1, worksheet.dim.data_end) );
    worksheet = last_section.worksheet;
    let new_section = worksheet.add_section_after(last_section, {
      title: section_datum.title,
      data_width: section_datum.labels.values.length,
      title_note_data: new Worksheet.NoteData([['import-id', section_datum.id]])
    });
    worksheet = new_section.worksheet;
    hidden_write_section_data_(new_section, codes, section_datum);
  }
}

function hidden_sync_section_data_(section, codes, section_datum) {
  let sheetbuf = section.worksheet.group.sheetbuf;
  let label_colors = sheetbuf.slice_colors( "label_row",
    section.dim.data_start, section.dim.data_end );
  let src_label_colors = section_datum.labels.colors;
  for (let i = 0; i < label_colors.length; ++i) {
    if (label_colors[i] != src_label_colors[i]) {
      let src_color = src_label_colors[i];
      sheetbuf.set_color( "label_row",
        section.dim.data_start + i,
        src_color != "#ffffff" ? src_color : null );
    }
  }
  let src_data = section_datum.data;
  for (let code of codes) {
    if (src_data[code] == null) {
      continue;
    }
    let values = sheetbuf.slice_values( "code_" + code,
      section.dim.data_start, section.dim.data_end );
    let colors = sheetbuf.slice_colors( "code_" + code,
      section.dim.data_start, section.dim.data_end );
    let {values: src_values, colors: src_colors} = src_data[code];
    if (values.length != src_values.length) {
      throw new Error("unreachable");
    }
    for (let i = 0; i < values.length; ++i) {
      if (values[i] != src_values[i]) {
        let src_value = src_values[i];
        sheetbuf.set_value( "code_" + code,
          section.dim.data_start + i,
          src_value != "" ? src_value : null );
      }
      if (colors[i] != src_colors[i]) {
        let src_color = src_colors[i];
        sheetbuf.set_color( "code_" + code,
          section.dim.data_start + i,
          src_color != "#ffffff" ? src_color : null );
      }
    }
  }
}

function hidden_write_section_data_(section, codes, section_datum) {
  let sheetbuf = section.worksheet.group.sheetbuf;
  sheetbuf.set_values( "label_row",
    section.dim.data_start, section.dim.data_end,
    section_datum.labels.values );
  sheetbuf.set_colors( "label_row",
    section.dim.data_start, section.dim.data_end,
    section_datum.labels.colors
      .map(src_color => src_color != "#ffffff" ? src_color : null) );
  let src_data = section_datum.data;
  for (let code of codes) {
    if (src_data[code] == null) {
      continue;
    }
    let {values: src_values, colors: src_colors} = src_data[code];
    if (section.dim.data_width != src_values.length) {
      throw new Error("unreachable");
    }
    sheetbuf.set_values( "code_" + code,
      section.dim.data_start, section.dim.data_end,
      src_values );
    sheetbuf.set_colors( "code_" + code,
      section.dim.data_start, section.dim.data_end,
      src_colors
        .map(src_color => ((src_color != "#ffffff") ? src_color : null))
    );
  }
}

/**
 * @param a {any[]}
 * @param b {any[]}
 * @return {(
 *   {add_start: number, dest_start: number, dest_end: number} |
 *   {remove_start: number, remove_end: number}
 * )[]}
 */
function hidden_edit_sequence_(a, b) {
  let distances = [[0].concat(b.map((_,i) => i + 1))];
  for (let i = 1; i <= a.length; ++i) {
    let prev_d = distances[i-1];
    let d = [i];
    for (let j = 1; j <= b.length; ++j) {
      let dd = [d[j-1] + 1, prev_d[j] + 1];
      if (a[i-1] == b[j-1]) {
        dd.push(prev_d[j-1]);
      }
      d.push(Math.min(...dd));
    }
    distances.push(d);
  }
  let edit_sequence = [];
  let add_series = 0;
  let remove_series = 0;
  function push_add() {
    if (add_series == 0) { return; }
    edit_sequence.push({add_start: i, dest_start: j, dest_end: j + add_series});
    add_series = 0;
  }
  function push_remove() {
    if (remove_series == 0) { return; }
    edit_sequence.push({remove_start: i, remove_end: i + remove_series});
    remove_series = 0;
  }
  let i = a.length, j = b.length;
  while (i > 0 || j > 0) {
    if (i > 0 && j > 0 && a[i-1] == b[j-1]) {
      push_add();
      push_remove();
      --i;
      --j;
    } else if (j > 0 && distances[i][j] == distances[i][j-1] + 1) {
      push_remove();
      ++add_series;
      --j;
    } else if (i > 0 && distances[i][j] == distances[i-1][j] + 1) {
      push_add();
      ++remove_series;
      --i;
    } else {
      throw new Error("unreachable");
    }
  }
  push_add();
  push_remove();
  return edit_sequence;
}

