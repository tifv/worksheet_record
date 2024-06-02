// General description of upload process (initial design; not necessarily up-to-date)

// Phase server-init
// • detect presence of upload configuration
// • detect presence of upload record
// • detect worksheet section
// • collect information about worksheet and worksheet section, including
//   • group name
//   • worksheet or title category
//   • title id
//   • title
//   • title date
//   • uploader e-mail
//   • last used author name (from user properties)
// • decide filename
// • open dialog

// Phase client-form
// • present information about upload to the human, including
//   • group name
//   • worksheet category
//   • (editable) title
//   • title date
//   • (editable) author name
//   • uploader e-mail
// • collect edits
// • collect PDF file and source files

// Phase client-prepare
// • zip source files if necessary (https://stuk.github.io/jszip/)
// • determine file lengths and hashes
// • request the server for authorized upload headers, sending
//   • file names
//   • file sizes
//   • hashes

// Phase server-authorize
// • compute upload signatures
// • return get URLs, authorized upload URLs and upload headers to the client

// Phase client-upload
// • upload files to the storage using authorized upload headers
// • send a confirmation to the server that includes
//   • uploaded file urls
//   • filename (base)
//   • group name
//   • worksheet category
//   • title id
//   • title
//   • title date
//   • author name

// Phase server-finish
// • add a line to the upload record
// • locate the title range and get its value
// • replace title value with hyperlink
// • hyperlink target should be filtered from uploads record using PDF URL (as quarantine URL)


function upload_enabled_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return UploadConfig.is_configured() && UploadRecord.exists(spreadsheet);
}

function upload_start_dialog_(section, options = {}) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const categories = Categories.get(spreadsheet);
  var worksheet = section.worksheet;
  var group = worksheet.group;
  var group_name = group.name;
  var category = worksheet.get_category();
  var category_info = categories[category || "mixture"] || {};
  var category_name = category_info.name ||
    (category != null ? "категория " + category : "mixture");
  var title = options.title || section.get_qualified_title();
  var title_id = section.get_title_metadata_id({validate: true});
  var date = section.get_title_note_data().get("date");
  var author = UploadAuthor.get();
  var filename_base; {
    let filename_pieces = [];
    filename_pieces.push(group.get_filename());
    filename_pieces.push( category_info.filename ||
      (category == null ? "mixture" : "whatever") );
    if (options.filename_date != null) {
      filename_pieces.push(options.filename_date.format({filename: true}));
    } else if (date != null) {
      filename_pieces.push(date.format({filename: true}));
    }
    if (options.filename_suffix != null)
      filename_pieces.push(options.filename_suffix);
    filename_base = filename_pieces.join('-');
  }
  var upload_info = {
    group_name: group_name,
    category: category,
    category_name: category_name,
    title_id: title_id,
    title: title,
    date: date != null ? date.to_object() : null,
    date_s: date != null ? date.format() : null,
    author: author,
    filename_base: filename_base,
  };
  var template = HtmlService.createTemplateFromFile(
    "Upload/Dialog" );
  template.standalone = true;
  template.upload_info = upload_info;
  template.category_css = format_category_css_(categories);
  var output = template.evaluate();
  output.setWidth(500).setHeight(475);
  SpreadsheetApp.getUi().showModelessDialog(output, "Публикация листочка");
}

const known_file_extensions = {
  ".pdf"  : "application/pdf",
  ".tex"  : "text/x-tex",
  ".txt"  : "text/plain",
  ".zip"  : "application/zip",
  ".docx" : "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  ".doc"  : "application/msword",
  ".odt"  : "application/vnd.oasis.opendocument.text",
};

function upload_authorize(files_meta) {
  const signer = UploadConfig.get_signer();
  var result = [];
  for (let {filename, size, hash_hex, is_text} of files_meta) {
    let [, file_ext] = split_filename_(filename);
    let file_type = known_file_extensions[file_ext];
    if (file_type == null)
      throw "upload_worksheet: unknown file extension " + file_ext;
    if (is_text && file_type.startsWith('text/')) {
      file_type += '; charset=utf-8';
    }
    let [upload_url, upload_headers] = signer.sign( "PUT",
      filename, "", [
        ["Content-Length", size.toString()],
        ["Content-Type", file_type],
        ["x-amz-content-sha256", hash_hex],
      ] );
    result.push({
      get_url: upload_url,
      upload_url: upload_url,
      upload_headers: upload_headers });
  }
  return result;
}

function split_filename_(filename) {
  var ext_index = filename.lastIndexOf(".");
  if (ext_index < 0)
    return [filename, ""];
  return [
    filename.substring(0, ext_index),
    filename.substring(ext_index) ];
}

function upload_finish({
    pdf_url, src_url, filename_base,
    group_name, category, title_id, title, date, author,
}) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  UploadAuthor.set(author);
  const uploads = UploadRecord.get(spreadsheet, "minimal");
  var id = Utilities.getUuid();
  uploads.append(new Map([
    ["group", "'" + group_name],
    ["category", (category != null) ?
      "'" + category : null ],
    ["title", "'" + title],
    ["date", (date != null) ?
      "'" + WorksheetDate.from_object(date).format() :
      null ],
    ["uploader", Session.getActiveUser().getEmail()],
    ["author", author],
    ["id", id],
    ["pdf", pdf_url],
    ["src", src_url],
    ["initial_pdf", pdf_url],
    ["initial_src", src_url],
    ["status", "unstable"],
    ["filename", filename_base],
  ]));
  const group = StudyGroup.find_by_name(spreadsheet, group_name);
  const sheet = group.sheet;
  ActionLock.with_lock(() => {
    var metadata = sheet.createDeveloperMetadataFinder()
      .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.COLUMN)
      .withId(title_id)
      .find();
    if (metadata.length < 1)
      throw new Error("upload_worksheet: cannot find title to update");
    var column = metadata[0].getLocation().getColumn().getColumn();
    var cell = sheet.getRange(group.dim.title_row, column);
    function col_R1C1(key) {
      return (
        "R" + uploads.first_row +
        "C" + uploads.key_columns.get(key) +
        ":C" + uploads.key_columns.get(key)
      );
    }
    var value = cell.getValue().trim();
    cell
      .setShowHyperlink(true)
      .setRichTextValue(
        SpreadsheetApp.newRichTextValue()
          .setText(value)
          .setLinkUrl(pdf_url)
          .build() )
      .setFontColor(null)
      .setFormulaR1C1("=hyperlink(" +
        "filter(" +
          "'" + uploads.name + "'!" + col_R1C1("pdf") + ";" +
          "'" + uploads.name + "'!" + col_R1C1("id") + "=\"" + id + "\"" +
        ");" +
        "\"" + value + "\"" +
      ")");
  });
}

function upload_fake_finish_(section, options = {}) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const categories = Categories.get(spreadsheet);
  var worksheet = section.worksheet;
  var group = worksheet.group;
  var group_name = group.name;
  var category = worksheet.get_category();
  var category_info = categories[category || "mixture"] || {};
  var title = options.title || section.get_qualified_title();
  var date = section.get_title_note_data().get("date");
  var author = UploadAuthor.get();
  var filename_base; {
    let filename_pieces = [];
    filename_pieces.push(group.get_filename());
    filename_pieces.push( category_info.filename ||
      (category == null ? "mixture" : "whatever") );
    if (options.filename_date != null) {
      filename_pieces.push(options.filename_date.format({filename: true}));
    } else if (date != null) {
      filename_pieces.push(date.format({filename: true}));
    }
    if (options.filename_suffix != null)
      filename_pieces.push(options.filename_suffix);
    filename_base = filename_pieces.join('-');
  }
  date = (date != null ? date.to_object() : null);
  const uploads = UploadRecord.get(spreadsheet, "minimal");
  var id = Utilities.getUuid();
  uploads.append(new Map([
    ["group", "'" + group_name],
    ["category", (category != null) ?
      "'" + category : null ],
    ["title", "'" + title],
    ["date", (date != null) ?
      "'" + WorksheetDate.from_object(date).format() :
      null ],
    ["uploader", Session.getActiveUser().getEmail()],
    ["author", author],
    ["id", id],
    ["status", "void"],
    ["filename", filename_base],
  ]));
  var cell = section.title_range.getCell(1, 1);
  function col_R1C1(key) {
    return (
      "R" + uploads.first_row +
      "C" + uploads.key_columns.get(key) +
      ":C" + uploads.key_columns.get(key)
    );
  }
  var value = cell.getValue().trim();
  cell
    .setFormulaR1C1("=hyperlink(" +
      "filter(" +
        "'" + uploads.name + "'!" + col_R1C1("pdf") + ";" +
        "'" + uploads.name + "'!" + col_R1C1("id") + "=\"" + id + "\"" +
      ");" +
      "\"" + value + "\"" +
    ")");
}


var UploadConfig = function() { // begin namespace

const document_key = "upload_config";

var config = null;

function load() {
  if (config == null) {
    let config_json = PropertiesService.getDocumentProperties().getProperty(document_key);
    if (config_json != null)
      config = JSON.parse(config_json);
    else
      config = {configured: false};
  }
  return config;
}

function save(new_config) {
  config = new_config;
  PropertiesService.getDocumentProperties().setProperty( document_key,
    JSON.stringify(config) );
}

function is_configured() {
  return load().configured;
}

function get_signer() {
  return new S3Signer(load());
}

function set({
  region, bucket_url, access_key, secret_key,
}, preserve_secret_key = false) {
  if (preserve_secret_key) {
    secret_key = load().secret_key;
  }
  save({
    configured: true,
    region: region, bucket_url: bucket_url,
    access_key: access_key, secret_key: secret_key,
  });
}

return {
  is_configured: is_configured,
  get_signer: get_signer,
  get: load, set: set,
};
}(); // end UploadConfig namespace


var UploadAuthor = function() { // begin namespace

const user_key = "upload_author";

function get() {
  return PropertiesService.getUserProperties().getProperty(user_key);
}

function set(author) {
  PropertiesService.getUserProperties().setProperty(user_key, author);
}

return {get: get, set: set};
}(); // end UploadAuthor namespace

