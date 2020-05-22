function sidebar_show() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const categories = Categories.get(spreadsheet);
  var template = HtmlService.createTemplateFromFile("Sidebar/Front");
  template.categories = categories;
  template.category_css = format_category_css_(categories);
  template.upload_enabled = upload_enabled_();
  var output = template.evaluate().setTitle("Ведомость");
  SpreadsheetApp.getUi().showSidebar(output);
}

function sidebar_load_group_list() {
  return Array.from(StudyGroup.list_names(SpreadsheetApp.getActiveSpreadsheet()));
}

function sidebar_load_contents(group_name, {continuation = null, cached = []} = {}) {
  // XXX avoid modifying the spreadsheet at all (when getting location)
  // if it is unaviodable, return whatever contents is already scanned,
  // and set special parameter to the continuation token that will trigger
  // lock acquisition on the next iteration.
  var start_time = (new Date()).getTime();
  function execution_time() {
    return (new Date()).getTime() - start_time;
  }
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var contents = [];
  var group;
  if (group_name == null) {
    group = StudyGroup.get_active(SpreadsheetApp.getActiveSpreadsheet());
    if (group == null)
      return null;
    group_name = group.name;
    contents.push({group_name: group_name});
  } else {
    group = StudyGroup.find_by_name(spreadsheet, group_name);
  }
  if (cached.includes(group_name)) {
    if (contents.length == 0)
      throw new Error("internal error");
    contents[0].cached = true;
    return contents;
  }
  var start_column = 1;
  if (continuation != null) {
    ({start_column: start_column} = continuation);
  }
  var contents_count = 0;
  for (let worksheet of Worksheet.list(group, start_column)) {
    for (let section of worksheet.list_sections()) {
      contents.push(sidebar_load_contents_section_(section, false));
    }
    if (contents.length >= 5 && execution_time() > 2500) {
      contents.push({continuation: {start_column: worksheet.dim.end + 1}});
      break;
    }
  };
  return contents;
}

function sidebar_load_contents_validate(contents_item) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var group = StudyGroup.find_by_name(spreadsheet, contents_item.group_name);
  var title_note = contents_item.title_note;
  var title_id = contents_item.id;
  var column = contents_item.column;
  var width = contents_item.width;
  var checked_title_columns = new Set();
  function was_checked(column) {
    if (checked_title_columns.has(column))
      return true;
    checked_title_columns.add(column);
    return false;
  }
  function get_checked_id(entity) {
    // entity is section or worksheet
    return entity.get_title_metadata_id({check: !was_checked(entity.dim.title)});
  }
  function check_id(entity) {
    // entity is section or worksheet
   if (!was_checked(entity.dim.title)) {
     entity.get_title_metadata_id({check: true});
   }
  }
  var lock = ActionHelpers.acquire_lock();
  var section_by_note = null, section_by_id = null, sections = [];
  group.sheetbuf.ensure_loaded(column, column + width - 1);
  var column_by_note = group.sheetbuf.find_closest("notes", "title_row", title_note, column);
  if (column_by_note != null) {
    section_by_note = Worksheet.surrounding_section( group, null,
      group.sheet.getRange(1, column_by_note) );
    sections.push(section_by_note);
  }
  if ( section_by_note != null &&
    title_id == get_checked_id(section_by_note)
  ) {
    section_by_id = section_by_note;
    section_by_note = null;
  } else {
    let column_by_id = Worksheet.find_title_column_by_id(group, title_id);
    if (column_by_id != null) {
      var section_by_id = Worksheet.surrounding_section( group, null,
        group.sheet.getRange(1, column_by_id) );
      if (title_id != get_checked_id(section_by_id)) {
        // better just to remove this id then
        group.sheet.createDeveloperMetadataFinder()
          .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.COLUMN)
          .withId(title_id)
          .find().forEach(function(metadatum) {
            metadatum.remove(); });
      }
    }
  }
  var contents = [];
  for (let section of sections) {
    check_id(section.worksheet);
    contents.push(sidebar_load_contents_section_(section, true));
  }
  lock.releaseLock();
  return contents;
}

function sidebar_load_contents_section_(section, validated = false) {
  var worksheet = section.worksheet;
  var group = worksheet.group;
  var worksheet_location = worksheet.get_location({check_id: false});
  var section_location = section.get_location({check_id: false});
  var title_id = section.get_title_metadata_id({check: false});
  var title_note = section.get_title_note();
  var title_note_info = Worksheet.parse_title_note(title_note);
  var date = title_note_info.date;
  var contents_item = {
    id: title_id,
    validated: validated,
    group_name: group.name,
    worksheet_location: worksheet_location,
    section_location: section_location,
    column: section.dim.title,
    width: section.dim.width,
    is_unused: worksheet.get_title().startsWith("{"),
      // XXX hide such worksheets in the sidebar
    is_subsection: section.dim.offset > 0,
    title: section.get_title(),
    qualified_title: section_get_qualified_title(),
    title_note: title_note,
    category: worksheet.get_category(),
    labels: group.sheetbuf.slice_values( "label_row",
      section.dim.data_start, section.dim.data_end ),
    date: date != null ? date.to_object() : null,
    date_s: date != null ? date.format() : null,
    date_filename: date != null ? date.format({filename: true}) : null,
  };
  parse_formula: {
    let title_formula = section.get_title_formula();
    if (title_formula == "")
      break parse_formula;
    let title_formula_decode = decode_hyperlink_formula_(title_formula);
    if (title_formula_decode == null)
      break parse_formula
    let [{filter = null, url = null}, ] = title_formula_decode;
    if (filter != null) {
      contents_item.title_link = {filter: filter};
    } else if (url != null) {
      contents_item.title_link = {url: url};
    }
  }
  // XXX also add some info about hyperlink
  return contents_item;
}

function sidebar_goto(group_name, column) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var group = StudyGroup.find_by_name(spreadsheet, group_name);
  group.sheet.getRange(group.dim.title_row, column).activate();
}

function sidebar_load_uploads() {
  var response = [];
  for (let datum of UploadRecord.get("full")) {
    response.push(sidebar_load_uploads_encode_datum_(datum));
  }
  return response;
}

function sidebar_load_uploads_search(search_text) {
  console.log(search_text);
  var response = [];
  for (let datum of UploadRecord.get("full").find("initial_pdf", search_text)) {
    response.push(sidebar_load_uploads_encode_datum_(datum));
  }
  console.log(JSON.stringify(response));
  return response;
}

function sidebar_load_uploads_encode_datum_(datum) {
  return {entries: Array.from(datum.entries()), index: datum.index};
}

function sidebar_upload_get_dialog() {
  var template = HtmlService.createTemplateFromFile(
    "Upload/Dialog" );
  template.standalone = false;
  template.sidebar = true;
  template.partial = "style";
  var style = template.evaluate().getContent();
  template.partial = "script";
  var script = template.evaluate().getContent();
  template.partial = "html";
  var html = template.evaluate().getContent();
  return {
    style: style,
    script: script,
    html: html,
    author: UploadAuthor.get(),
  };
}

function sidebar_upload_get_group_filename(group_name) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(group_name);
  if (sheet == null)
    throw new Error("no sheet for the group " + group_name);
  return (new StudyGroup(sheet, group_name)).get_filename();
}

function sidebar_collapse_expand(group_name, column_actions) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(group_name);
  if (sheet == null)
    throw new Error("no sheet for the group " + group_name);

  for (let {column, action} of column_actions) {
    try {
      var col_group = sheet.getColumnGroup(column, 1);
    } catch (error) {
      console.error( "No column group at column " +
        sheet.getRange(1, collapse_col).getA1Notation().replace(/\d+/,"") );
      console.error(error);
    }
    if (col_group == null)
      continue;
    if (action == "collapse")
      col_group.collapse();
    else if (action == "expand")
      col_group.expand();
    else
      throw new Error("server error");
  }
}
