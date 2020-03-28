class WorksheetError extends SpreadsheetError {};
class WorksheetInitError   extends WorksheetError {};
class WorksheetDetectError extends WorksheetError {};
class WorksheetCheckError  extends WorksheetError {};

class WorksheetSectionError extends WorksheetError {};
class WorksheetSectionInitError   extends WorksheetSectionError {};
class WorksheetSectionDetectError extends WorksheetSectionError {};
class WorksheetSectionCheckError  extends WorksheetSectionError {};


var Worksheet = function() { // namespace

const metadata_keys = {
  title: "worksheet-title",
};

const data_offset = {start: 3, end: 1};
data_offset.width = data_offset.start + data_offset.end;

const marker = {start: "‹", end: "›"};

const initial = {
  data_width: 15,
  title: "{Бланк}",
  section_data_width: 3,
  section_title: "Добавка",
};

function Worksheet(group, full_range) {
  if (group == null) {
    group = new StudyGroup(full_range.getSheet());
    group.check();
  } else if (!(group instanceof StudyGroup)) {
    throw new Error("Worksheet.constructor: type error (group)");
  }
  this.group = group;
  this.full_range = full_range;
}

define_lazy_properties_(Worksheet.prototype, {
  sheet: function() {
    return this.group.sheet; },
  dim: function() {
    var dim = {};
    dim.start = this.full_range.getColumn();
    dim.end = this.full_range.getLastColumn();
    dim.data_start = dim.start + data_offset.start;
    dim.data_end   = dim.end   - data_offset.end;
    dim.marker_start = dim.data_start - 1;
    dim.marker_end   = dim.data_end   + 1;
    dim.width = dim.end - dim.start + 1;
    dim.data_width = dim.data_end - dim.data_start + 1;
    dim.title = dim.start;
    dim.rating = dim.start;
    dim.sum = dim.start + 1;
    return dim;
  },
  title_range: function() {
    return this.sheet.getRange(
      this.group.dim.title_row, this.dim.start,
      1, this.dim.width );
  },
  data_range: function() {
    return this.sheet.getRange(
      this.group.dim.data_row, this.dim.data_start,
      this.group.dim.data_height, this.dim.data_width );
  },
  label_range: function() {
    return this.sheet.getRange(
      this.group.dim.label_row, this.dim.data_start,
      1, this.dim.data_width );
  },
  max_range: function() {
    return this.sheet.getRange(
      this.group.dim.max_row, this.dim.data_start,
      1, this.dim.data_width );
  },
  weight_range: function() {
    return this.sheet.getRange(
      this.group.dim.weight_row, this.dim.data_start,
      1, this.dim.data_width );
  },
  rating_range: function() {
    return this.sheet.getRange(
      this.group.dim.data_row, this.dim.rating,
      this.group.dim.data_height, 1 );
  },
  sum_range: function() {
    return this.sheet.getRange(
      this.group.dim.data_row, this.dim.sum,
      this.group.dim.data_height, 1 );
  },
  mirror_range: function() {
    return this.sheet.getRange(
      this.group.dim.mirror_row, this.dim.start,
      1, this.dim.width );
  },
  title_column_range: function() {
    return get_column_range_(this.sheet.getRange(
      1, this.dim.title ));
  },
  metaweight_cell: function() {
    return this.sheet.getRange(
      this.group.dim.weight_row, this.dim.rating );
  },
});

Worksheet.prototype.check = function(options = {}) {
  ({
    dimensions: options.dimensions = true,
    markers: options.markers = true,
  } = options);
  if (options.dimensions) {
    if (
      this.full_range.isStartRowBounded() ||
      this.full_range.isEndRowBounded() ||
      this.dim.width < data_offset.width + 1
    ) {
      throw new WorksheetCheckError(
        "full_range is incorrect " +
        "(must be row-unbounded and of width at least " +
          (data_offset.width + 1) + ")",
        this.full_range );
    }
  }
  if (options.markers) {
    this.group.sheetbuf.ensure_loaded(this.dim.start, this.dim.end);
    if (
        this.dim.marker_start !=
        this.group.sheetbuf.find_last_value(
          "label_row", marker.start, this.dim.end, this.dim.start ) ||
        this.dim.marker_end !=
        this.group.sheetbuf.find_value(
          "label_row", marker.end, this.dim.start, this.dim.end )
    ) {
      throw new WorksheetCheckError(
        "markers are missing or interwine",
        this.sheet.getRange(this.sheet.dim.title_row, this.dim.start, 1, this.dim.width) );
    }
  }
}

// XXX add alternative initializer that creates standalone title
// like worksheet with markers, but without formatting datarange etc.

// XXX this initializer breaks group.values structure.
// So, it must not return value.
Worksheet.add = function(group, range, options) {
  /* options:
   *   data_width (number)
   *     default is initial.data_width
   *   title
   *     default is initial.title
   *   date XXX
   *   color_scheme XXX
   *     code (string), must match spreadsheet metadata
   *     default will use group color scheme
   *   category
   *     default is not to set category
   */
  if (group == null) {
    group = new StudyGroup(range.getSheet());
    group.check();
  } else if (!(group instanceof StudyGroup)) {
    throw new Error("Worksheet.add: type error (group)");
  }
  if (options == null)
    options = {};
  if (options.data_width == null)
    options.data_width = initial.data_width;
  var sheet = group.sheet;
  if (range.getNumColumns() > 1)
    throw new WorksheetInitError(
      "only one-column ranges are accepted for initialization",
      range );
    // XXX actually, we should accept wider ranges as long as they are empty?
  var fullwidth = options.data_width + data_offset.width;
  sheet.insertColumnsAfter(range.getColumn(), fullwidth + 1);
  var fullrange = get_column_range_(sheet.getRange(
    1, range.getColumn() + 1, 1, fullwidth ));
  var worksheet = new Worksheet(group, fullrange);
  initialize.call(worksheet, options);
  return worksheet;
}

// XXX refactor this like with StudyGroup, into a separate «class»

function initialize(options) { // applied to Worksheet instance
  if (options.title == null)
    options.title = initial.title;

  const group = this.group;
  const sheet = this.sheet;
  var dim = this.dim;
  var color_scheme = group.get_color_scheme(); // XXX options.color_scheme

  title_range: { // {{{
    add_title_metadata.call(this);
    let metadata = this.get_title_metadata();
    let note_info = {lines: []};
    if (options.date != null) {
      note_info.date = options.date;
    }
    note_info.title_id = metadata.getId();
    let note = Worksheet.format_title_note(note_info);
    this.title_range.getCell(1, 1)
      .setValue(options.title)
      .setNote(note);
    this.title_range
      .merge()
      .setFontSize(12)
      .setFontWeight('bold')
      .setFontFamily("Times New Roman,serif");
  } // }}}

  // data_range {{{
    this.data_range
      .setNumberFormat('0.#;−0.#');
  // }}}

  // markers {{{
    set_marker_guard_( this.label_range.offset(0, -1, 1, 1),
      marker.start );
    set_marker_guard_( this.label_range.offset(0, dim.data_width, 1, 1),
      marker.end );
  // }}}

  // label_range {{{
    var labels = [];
    for (var i = 1; i <= dim.data_width; ++i) {
      if (i <= 3)
        labels[i-1] = i.toString();
      else // if (i < dim.data_width)
        labels[i-1] = null;
      // else
      //   labels[i-1] = null;
    }
    this.label_range
      .setValues([labels]);
  // }}}

  // max_range {{{
    var data_column_R1C1 = 'R' + group.dim.data_row + 'C[0]:C[0]';
    var max_formula_R1C1 = '=max(0;' + data_column_R1C1 + ')';
    this.max_range.offset(0, -1, 1, dim.data_width + 2)
      .setNumberFormat('0.#;−0.#')
      .setFontSize(8);
    this.max_range
      .setFormulaR1C1(max_formula_R1C1);
  // }}}

  // weight_range {{{
    var max_R1C1 = 'R' + group.dim.max_row + 'C[0]';
    var weight_formula_R1C1 = ''.concat(
      '= if(',
        max_R1C1, '; ',
        '1 / ',
        'sqrt( ',
          'max(1;sumif(', data_column_R1C1, ';">0"))',
          ' * ',
          'max(1;', max_R1C1, ')',
        ' ); ',
      '0)'
    );
    this.weight_range.offset(0, -1, 1, dim.data_width + 2)
      .setNumberFormat('#.0#;−#.0#;0')
      .setFontSize(8);
    this.weight_range
      .setFormulaR1C1(weight_formula_R1C1);
  // }}}

  // sum_range {{{
    var max_row_sum_R1C1 =
      'R' + group.dim.max_row + 'C[' + (dim.data_start - 1 - dim.sum) + ']:' +
      'R' + group.dim.max_row + 'C[' + (dim.data_end   + 1 - dim.sum) + ']';
    var data_row_sum_R1C1 =
      'R[0]C[' + (dim.data_start - 1 - dim.sum) + ']:' +
      'R[0]C[' + (dim.data_end   + 1 - dim.sum) + ']';
    var sum_formula_R1C1 = ''.concat(
      '=countifs(',
        max_row_sum_R1C1,  ';">0";',
        data_row_sum_R1C1, ';">0"',
      ')'
    );
    this.sum_range
      .setFormulaR1C1(sum_formula_R1C1)
      .setNumberFormat('0')
      .setFontSize(8);
    this.sum_range.offset(group.dim.max_row - group.dim.data_row, 0, 1, 1)
      .setFormulaR1C1(sum_formula_R1C1)
      .setNumberFormat('0')
      .setFontSize(8);
    this.sum_range.offset(group.dim.label_row - group.dim.data_row, 0, 1, 1)
      .setValue('S');
  // }}}

  // rating_range {{{
    var weight_row_rating_R1C1 =
      'R' + group.dim.weight_row + 'C[' + (dim.data_start - 1 - dim.rating) + ']:' +
      'R' + group.dim.weight_row + 'C[' + (dim.data_end   + 1 - dim.rating) + ']';
    var data_row_rating_R1C1 =
      'R[0]C[' + (dim.data_start - 1 - dim.rating) + ']:' +
      'R[0]C[' + (dim.data_end   + 1 - dim.rating) + ']';
    var rating_formula_R1C1 = ''.concat(
      '=sumproduct( ',
        weight_row_rating_R1C1, ';',
        data_row_rating_R1C1,
      ')'
    );
    this.rating_range
      .setFormulaR1C1(rating_formula_R1C1)
      .setNumberFormat('0.00;−0.00')
      .setFontSize(8);
    this.rating_range.offset(group.dim.max_row - group.dim.data_row, 0, 1, 1)
      .setFormulaR1C1(rating_formula_R1C1)
      .setNumberFormat('0.00;−0.00')
      .setFontSize(8);
    this.rating_range.offset(group.dim.label_row - group.dim.data_row, 0, 1, 1)
      .setValue('Σ');
  // }}}

  // metaweightsell {{{
    this.metaweight_cell
      .setValue(1)
      .setNote('вес листочка в рейтинге')
      .setNumberFormat('0.0;−0.0')
      .setFontSize(8);
  // }}}

  // mirror_range {{{
    var left_marker_mirror_R1C1  = 'R' + group.dim.label_row + 'C[' + (dim.data_start - 1 - dim.start) + ']';
    var right_marker_mirror_R1C1 = 'R' + group.dim.label_row + 'C[' + (dim.data_end   + 1 - dim.start) + ']';
    var title_mirror_R1C1 = 'R' + group.dim.title_row + 'C[0]';
    var label_row_mirror_R1C1 =
      'R' + group.dim.label_row + 'C[0]' + ':' +
      'R' + group.dim.label_row + 'C[' + (dim.end - dim.start) + ']';
    var mirror_formula_R1C1 = ''.concat(
      '=iferror( ',
        'if(',
          'or(',
            left_marker_mirror_R1C1,  '<>"' + marker.start + '";',
            right_marker_mirror_R1C1, '<>"' + marker.end + '"); ',
          'na(); ',
        'if(or(isblank(', title_mirror_R1C1, ');left(', title_mirror_R1C1, ')="{"); ',
          'iferror(na()); ',
        'arrayformula(', label_row_mirror_R1C1, '))); ',
      'split(rept("#N/A ";columns(', label_row_mirror_R1C1, '));" "))'
    );
    this.mirror_range.getCell(1, 1)
      .setFormulaR1C1(mirror_formula_R1C1);
  // }}}

  if (options.category)
    this.set_category(options.category, {ignore_sections: true});

  this.reset_column_widths();

  // borders {{{
    sheet.getRange(
      group.dim.data_row, dim.start - 1,
      group.dim.data_height, dim.width + 2
    )
      .setBorder(true, null, true, null, null, null)
      .setBorder( null, null, null, null, null, true,
        "black", SpreadsheetApp.BorderStyle.DOTTED );
    sheet.getRange(
      group.dim.data_row - 1, dim.start - 1,
      1, dim.width + 2
    )
      .setBorder(null, null, true, null, null, null);
    this.reset_data_borders(
      dim.data_start, dim.data_end,
      {open_left: false, open_right: false} );
    var rating_sum_range = sheet.getRange(
      group.dim.data_row, dim.start,
      group.dim.data_height, 2 );
    rating_sum_range
      .setBorder(true, true, true, true, null, null)
      .setBorder( null, null, null, null, true, null,
        "black", SpreadsheetApp.BorderStyle.DOTTED );
    rating_sum_range.offset(
        group.dim.label_row - group.dim.data_row, 0, 1 )
      .setBorder(true, true, true, true, null, null)
      .setBorder( null, null, null, null, true, null,
        "black", SpreadsheetApp.BorderStyle.DOTTED );
    rating_sum_range.offset(
        group.dim.max_row - group.dim.data_row, 0, 1 )
      .setBorder(true, true, true, true, null, null)
      .setBorder( null, null, null, null, true, null,
        "black", SpreadsheetApp.BorderStyle.DOTTED );
    this.metaweight_cell
      .setBorder(true, true, true, true, null, null);
    sheet.getRange(
      group.dim.title_row, dim.start,
      group.dim.sheet_height - group.dim.title_row + 1, dim.width
    )
      .setBorder(true, true, true, true, null, null);
  // }}}

  set_blank_guard_(this.data_range.offset(
    0, -1, group.dim.data_height, 1 ));
  set_blank_guard_(this.data_range.offset(
    0, dim.data_width, group.dim.data_height, 1));

  CFormatting.merge(sheet, [
    this.new_cf_rule_data(HSL.to_hex(color_scheme.marks)),
    this.new_cf_rule_weight(
      HSL.to_hex(HSL.deepen(color_scheme.marks, 0.35)),
      HSL.to_hex(HSL.deepen(color_scheme.marks, 4.35)) ),
    this.new_cf_rule_rating(HSL.to_hex(color_scheme.rating_mid), HSL.to_hex(color_scheme.rating_top))
  ]);

  this.add_column_group();
}

Worksheet.prototype.reset_column_widths = function() {
  this.sheet.setColumnWidth(this.dim.rating, 30);
  this.sheet.setColumnWidth(this.dim.sum,    30);
  this.sheet.setColumnWidth(this.dim.data_start - 1, 5);
  this.sheet.setColumnWidths(this.dim.data_start, this.dim.data_width, 21);
  this.sheet.setColumnWidth(this.dim.data_end   + 1, 5);
  this.sheet.setColumnWidth(this.dim.end + 1, 13);
}

// Worksheet.prototype.reset_data_borders {{{
// options = {open_left: bool, open_right: bool, skip_weight: bool}
//   open_left   — no solid border on the left
//   open_right  — no solid border on the left
//   skip_weight — do not draw borders around weightrange
Worksheet.prototype.reset_data_borders = function(col_begin, col_end, options = {}) {
  const group = this.group;
  var col_num = col_end - col_begin + 1;
  var ranges = [
    this.sheet.getRange(group.dim.label_row, col_begin, 1, col_num),
    this.sheet.getRange(group.dim.max_row, col_begin, 1, col_num),
    this.sheet.getRange(
      group.dim.data_row, col_begin, group.dim.data_height, col_num )
  ];
  if (!options.skip_weight) {
    ranges.push(
      this.sheet.getRange(group.dim.weight_row, col_begin, 1, col_num) );
  }
  // horizontal {{{
    for (let range of ranges) {
      range.setBorder(true, null, true, null, null, null);
    };
    // horizontal between weight_range and sum_range
    if ( !options.skip_weight &&
        group.dim.weight_row - group.dim.max_row == -1 )
    {
      this.sheet.getRange(
          group.dim.weight_row, col_begin,
          group.dim.max_row - group.dim.weight_row + 1, col_num )
        .setBorder( null, null, null, null, null, true,
          'black', SpreadsheetApp.BorderStyle.DOTTED );
    }
  // }}}
  // vertical {{{
    var open_left = options.open_left, open_right = options.open_right;
    if (open_left !== true) { // vertical left
      // from outside
      for (let range of ranges) {
        range.offset(0, -1, range.getNumRows(), 1)
          .setBorder( null, null, null, true, null, null );
      };
      // from inside
      for (let range of ranges) {
        range.setBorder( null, true, null, null, null, null );
      };
      open_left = null;
    } // else open_left === true
    if (open_right !== true) { // vertical right
      // from outside
      for (let range of ranges) {
        range.offset(0, col_num, range.getNumRows(), 1)
          .setBorder( null, true, null, null, null, null );
      };
      // from inside
      for (let range of ranges) {
        range.setBorder( null, null, null, true, null, null );
      };
      open_right = null;
    } // else open_right === true
    // vertical internal
    for (let range of ranges) {
      range.setBorder( null, open_left, null, open_right, true, null,
        'black', SpreadsheetApp.BorderStyle.DOTTED );
    };
  // }}}
}; // }}}

Worksheet.prototype.add_column_group = function() {
  this.title_range.shiftColumnGroupDepth(+1);
}

Worksheet.prototype.new_cf_rule_data = function(colour) {
  return SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground(colour)
    .setRanges([
      this.sheet.getRange(
        this.group.dim.data_row, this.dim.data_start - 1,
        this.group.dim.data_height, this.dim.data_width + 2 ),
      this.sheet.getRange(
        this.group.dim.max_row, this.dim.data_start - 1,
        1, this.dim.data_width + 2 )
    ])
    .build();
}

Worksheet.prototype.new_cf_rule_weight = function(colour_min, colour_max) {
  var cell_A1 = this.sheet.getRange(this.group.dim.weight_row, this.dim.data_start - 1)
    .getA1Notation();
  var weight_A1 = this.sheet.getRange(this.group.dim.weight_row, this.dim.data_start - 1)
    .getA1Notation().replace(/([A-Z]+)([0-9]+)/, "$1$$$2");
  var max_A1 = this.sheet.getRange(this.group.dim.max_row, this.dim.data_start - 1)
    .getA1Notation().replace(/([A-Z]+)([0-9]+)/, "$1$$$2");
  return SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue( colour_min,
      SpreadsheetApp.InterpolationType.NUMBER,
      "=" + cell_A1 + " - 1/power(" + weight_A1 + "*max(" + max_A1 + ";1);2) + 1" )
    .setGradientMaxpointWithValue( colour_max,
      SpreadsheetApp.InterpolationType.NUMBER,
      "=" + cell_A1 + " - 1/power(" + weight_A1 + "*max(" + max_A1 + ";1);2) + 21" )
    .setRanges([
      this.sheet.getRange(
        this.group.dim.weight_row, this.dim.data_start - 1,
        1, this.dim.data_width + 2 )
    ])
    .build();
}

Worksheet.prototype.new_cf_rule_data_limit = function(colour) {
  var limit_A1 = this.sheet.getRange(this.group.dim.weight_row, this.dim.sum)
    .getA1Notation().replace(/([A-Z]+)([0-9]+)/, "$$$1$$$2");
  return SpreadsheetApp.newConditionalFormatRule()
    .withCriteria(SpreadsheetApp.BooleanCriteria.NUMBER_GREATER_THAN_OR_EQUAL_TO, ["=" + limit_A1])
    .setBackground(colour)
    .setRanges([
      this.sheet.getRange(
        this.group.dim.data_row, this.dim.data_start - 1,
        this.group.dim.data_height, this.dim.data_width + 2 ),
      this.sheet.getRange(
        this.group.dim.max_row, this.dim.data_start - 1,
        1, this.dim.data_width + 2 )
    ])
    .build();
}

Worksheet.prototype.new_cf_rule_rating = function(colour_mid, colour_top) {
  return this.group.new_cf_rule_rating([
    this.sheet.getRange(
      this.group.dim.data_row, this.dim.rating,
      this.group.dim.data_height, 2 ),
    this.sheet.getRange(this.group.dim.max_row, this.dim.rating, 1, 2)
  ], colour_mid, colour_top);
}

Worksheet.prototype.get_category = function() {
  let category = this.group.sheetbuf.get_value("category_row", this.dim.start);
  if (category === "")
    return null;
  return category;
}

// XXX this method should be able to function in initialization,
// when group.values is broken.
Worksheet.prototype.set_category = function(code, options = {}) {
  ({
    ignore_sections: options.ignore_sections = true,
    // make the call faster by not checking for worksheet sections
  } = options);
  this.group.sheetbuf.set_value("category_row", this.dim.rating, code);
  this.group.sheetbuf.set_value("category_row", this.dim.sum, code);
  if (options.ignore_sections)
    return;
  for (let section in this.list_sections()) {
    if (section.dim.title == this.dim.rating)
      continue;
    section.set_category(code);
  }
}

Worksheet.prototype.get_title = function() {
  // also applies to WorksheetSection
  return this.group.sheetbuf.get_value("title_row", this.dim.title);
}

Worksheet.prototype.set_title = function(value) {
  // also applies to WorksheetSection
  this.group.sheetbuf.set_value("title_row", this.dim.title, value);
}

Worksheet.prototype.get_title_formula = function() {
  // also applies to WorksheetSection
  return this.group.sheetbuf.get_formula("title_row", this.dim.title);
}

Worksheet.prototype.set_title_formula = function(formula, value_replace = "") {
  // also applies to WorksheetSection
  this.group.sheetbuf.set_formula("title_row", this.dim.title, formula, value_replace);
}

Worksheet.prototype.get_title_note = function() {
  // also applies to WorksheetSection
  return this.group.sheetbuf.get_note("title_row", this.dim.title);
}

Worksheet.prototype.set_title_note = function(note) {
  // also applies to WorksheetSection
  this.group.sheetbuf.set_note("title_row", this.dim.title, note);
}

Worksheet.parse_title_note = function(note) {
  var lines = note.split("\n");
  while (lines.length > 0 && lines[lines.length-1] == "")
    --lines.length;
  var note_info = {lines: lines};
  for (let i = 0; i < lines.length; ++i) {
    let line = lines[i];
    if (note_info.date == null) {
      let date = WorksheetDate.parse(line);
      if (date != null) {
        note_info.date = date;
        note_info.date_line = i;
        lines[i] = ""; continue;
      }
    }
    if (note_info.title_id == null) {
      let title_match = /^id=(\d+)$/.exec(line);
      if (title_match != null) {
        note_info.title_id = parseInt(title_match[1]);
        note_info.title_id_line = i;
        lines[i] = ""; continue;
      }
    }
  }
  return note_info;
}

Worksheet.format_title_note = function(note_info) {
  var note_lines = Array.from(note_info.lines);
  if (note_info.date != null) {
    if (note_info.date_line == null) {
      note_info.date_line = 0;
      note_lines.unshift("");
    }
    note_lines[note_info.date_line] = note_info.date.format();
  }
  if (note_info.title_id != null) {
    if (note_info.title_id_line == null) {
      note_info.title_id_line = note_lines.length;
    }
    note_lines[note_info.title_id_line] = "id=" + note_info.title_id;
  }
  return note_lines.join("\n");
}

Worksheet.prototype.get_title_metadata_id = function(options = {}) {
  // also applies to WorksheetSection
  ({
    check: options.check = false,
      // check value from the title note agains actual metadata
  } = options);
  var note_info = Worksheet.parse_title_note(this.get_title_note());
  if (note_info.title_id != null && !options.check) {
    return note_info.title_id;
  }
  var title_id = this.get_title_metadata().getId();
  if (note_info.title_id == null || note_info.title_id != title_id) {
    note_info.title_id = title_id;
    this.set_title_note(
      Worksheet.format_title_note(note_info) );
  }
  return title_id;
}

Worksheet.prototype.get_title_metadata = function(_recursed) {
  // also applies to WorksheetSection
  var metadata = this.title_column_range.createDeveloperMetadataFinder()
    .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.COLUMN)
    .withVisibility(SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
    .withKey(metadata_keys.title)
    .find();
  if (metadata.length > 0)
    return metadata[0];
  if (_recursed)
    throw new Error("Worksheet().get_title_metadata: internal error");
  add_title_metadata.call(this);
  return this.get_title_metadata(true);
}

function add_title_metadata() {
  // applies to Worksheet or WorksheetSection
  this.title_column_range.addDeveloperMetadata(
    metadata_keys.title,
    SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT );
}

Worksheet.prototype.has_weight_row = function() {
  // also applies to WorksheetSection
  return this.group.sheetbuf.slice_values( "weight_row",
    this.dim.data_start, this.dim.data_end
  ).some(x => (x !== ""));
}

Worksheet.find_title_column_by_id = function(group, title_id) {
  if (!(group instanceof StudyGroup)) {
    throw new Error("Worksheet.find_title_column_by_id: type error (group)");
  }
  var metadata = group.sheet.createDeveloperMetadataFinder()
    .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.COLUMN)
    .withId(title_id)
    .find();
  if (metadata.length < 1)
    return null;
  return metadata[0].getLocation().getColumn().getColumn();
}

Worksheet.list = function*(group, start = 1, end) {
  if (!(group instanceof StudyGroup)) {
    throw new Error("Worksheet.list: type error (group)");
  }
  var last_end = start - 1;
  while (true) {
    let marker_start = group.sheetbuf.find_value("label_row", marker.start, last_end + 1, end);
    if (marker_start == null)
      break;
    let marker_end = group.sheetbuf.find_value("label_row", marker.end, marker_start + 2, end);
    if (marker_end == null)
      break;
    let rogue_start = group.sheetbuf.find_value("label_row", marker.start, marker_start + 1, marker_end);
    if (rogue_start != null) {
      if (last_end >= marker_start)
        throw new Error("Worksheet.list: internal error");
      last_end = marker_start;
      continue;
    }
    yield new Worksheet( group,
      get_column_range_(group.sheet.getRange(
        1, marker_start - data_offset.start + 1,
        1, marker_end - marker_start + 1 + data_offset.width - 2 ))
    );
    if (last_end >= marker_end)
      throw new Error("Worksheet.list: internal error");
    last_end = marker_end;
  }
}

Worksheet.surrounding = function(group, range) {
  if (group == null) {
    group = new StudyGroup(range.getSheet());
    group.check();
  } else if (!(group instanceof StudyGroup)) {
    throw new Error("Worksheet.surrounding: type error (group)");
  }
  var range_start = range.getColumn(), range_end = range.getLastColumn();
  group.sheetbuf.ensure_loaded(range_start, range_end);
  var start = group.sheetbuf.find_last_value("label_row", marker.start, range_end + data_offset.start);
  var end = group.sheetbuf.find_value("label_row", marker.end, range_start - data_offset.end);
  if (start == null || end == null || end - start + 1 <= data_offset.width) {
    throw new WorksheetDetectError(
      "unable to locate surrounding worksheet",
      range );
  }
  var worksheet = new Worksheet( group,
    get_column_range_(group.sheet.getRange(
      1, start - data_offset.start + 1,
      1, end - start + 1 + data_offset.width - 2 ))
  );
  try {
    worksheet.check();
  } catch (error) {
    if (error instanceof WorksheetCheckError) {
      throw new WorksheetDetectError(error.message, error.range);
    } else {
      throw error;
    }
  }
  return worksheet;
}

Worksheet.prototype.get_location = function(options = {}) {
  ({
    check_id: options.check_id = true,
      // check value from the title note agains actual metadata
  } = options);
  var title_id = this.get_title_metadata_id({check: options.check_id});
  return {
    title_id: title_id,
    column: this.dim.start,
    width: this.dim.width,
  };
}

Worksheet.find_by_location = function(group, location) {
  if (!(group instanceof StudyGroup)) {
    throw new Error("Worksheet.find_by_location: type error (group)");
  }
  var {title_id, column = null, width = null} = location;
  var sheet = group.sheet;
  find_column: {
    if (column != null) {
      var title_column_range = get_column_range_(sheet.getRange(1, column));
      var title_metadata = title_column_range.createDeveloperMetadataFinder()
        .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.COLUMN)
        .withId(title_id)
        .find();
      if (title_metadata.length > 0)
        break find_column;
    }
    title_metadata = sheet.createDeveloperMetadataFinder()
      .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.COLUMN)
      .withId(title_id)
      .find();
    if (title_metadata.length > 0) {
      column = title_metadata[0].getLocation().getColumn().getColumn();
      break find_column;
    }
    throw new WorksheetDetectError(
      "unable to locate worksheet starting column" );
  }
  if (width != null)
    group.sheetbuf.ensure_loaded(column, column + width - 1);
  let end_column = group.sheetbuf.find_value("label_row", marker.end, column);
  var worksheet = new Worksheet( group,
    get_column_range_(sheet.getRange(1, column, 1, end_column + data_offset.end - column))
  );
  worksheet.check();
  return worksheet;
}

Worksheet.prototype.alloy_subproblems = function() {
  var options = {
    skip_weight: ! this.has_weight_row()
  };
  for (let section of this.list_sections()) {
    section.alloy_subproblems(Object.assign({}, options));
  }
}

function WorksheetSection(worksheet, full_range) {
  this.worksheet = worksheet;
  this.full_range = full_range;
}

define_lazy_properties_(WorksheetSection.prototype, {
  group: function() { return this.worksheet.group; },
  sheet: function() { return this.worksheet.sheet; },
  dim: function() {
    var dim = {};
    dim.start = this.full_range.getColumn();
    dim.end = this.full_range.getLastColumn();
    dim.offset = dim.start - this.worksheet.dim.start;
    dim.data_start = dim.start > this.worksheet.dim.data_start ?
      dim.start : this.worksheet.dim.data_start;
    dim.data_end = dim.end < this.worksheet.dim.data_end ?
      dim.end : this.worksheet.dim.data_end;
    dim.data_offset = dim.data_start - this.worksheet.dim.data_start;
    dim.width = dim.end - dim.start + 1;
    dim.data_width = dim.data_end - dim.data_start + 1;
    dim.title = dim.start;
    return dim;
  },
  title_range: function() {
    return this.sheet.getRange(
      this.group.dim.title_row, this.dim.start,
      1, this.dim.width );
  },
  data_range: function() {
    return this.sheet.getRange(
      this.group.dim.data_row, this.dim.data_start,
      this.group.dim.data_height, this.dim.data_width );
  },
  label_range: function() {
    return this.sheet.getRange(
      this.group.dim.label_row, this.dim.data_start,
      1, this.dim.data_width );
  },
  max_range: function() {
    return this.sheet.getRange(
      this.group.dim.max_row, this.dim.data_start,
      1, this.dim.data_width );
  },
  weight_range: function() {
    return this.sheet.getRange(
      this.group.dim.weight_row, this.dim.data_start,
      1, this.dim.data_width );
  },
  mirror_range: function() {
    return this.sheet.getRange(
      this.group.dim.mirror_row, this.dim.start,
      1, this.dim.width );
  },
  title_column_range: function() {
    return get_column_range_(this.sheet.getRange(
      1, this.dim.title ));
  },
});

WorksheetSection.prototype.check = function(options = {}) {
  ({
    dimensions: options.dimensions = true,
    title: options.title = true,
  } = options);
  if (options.dimensions) {
    if (
      this.full_range.isStartRowBounded() ||
      this.full_range.isEndRowBounded() ||
      this.dim.start < this.worksheet.dim.start ||
      this.dim.start > this.worksheet.dim.data_end ||
      this.dim.end < this.worksheet.dim.data_start ||
      this.dim.end > this.worksheet.dim.end
    ) {
      throw new WorksheetSectionCheckError(
        "range is incorrect " +
        "(must be row-unbounded and contained in the worksheet range " +
          this.worksheet.full_range.getA1Notation() + ")",
        this.full_range );
    }
  }
  if (options.title) {
    let title_cols = this.group.sheetbuf.find_merge("title_row", this.dim.start, this.dim.end);
    if (title_cols != null) {
      let [, title_end] = title_cols;
      if (this.group.sheetbuf.find_merge("title_row", title_end + 1, this.dim.end) != null) {
        throw new WorksheetSectionCheckError(
          "misaligned title detected",
          this.title_range );
      }
    }
  }
}

WorksheetSection.prototype.set_category = function(
  code = this.worksheet.get_category()
) {
  this.group.sheetbuf.set_value("category_row", this.dim.title, code);
}

WorksheetSection.prototype.get_title =
  Worksheet.prototype.get_title ;
WorksheetSection.prototype.set_title =
  Worksheet.prototype.set_title ;
WorksheetSection.prototype.get_title_formula =
  Worksheet.prototype.get_title_formula ;
WorksheetSection.prototype.set_title_formula =
  Worksheet.prototype.set_title_formula ;
WorksheetSection.prototype.get_title_note =
  Worksheet.prototype.get_title_note ;
WorksheetSection.prototype.set_title_note =
  Worksheet.prototype.set_title_note ;
WorksheetSection.prototype.get_title_metadata_id =
  Worksheet.prototype.get_title_metadata_id ;
WorksheetSection.prototype.get_title_metadata =
  Worksheet.prototype.get_title_metadata ;
WorksheetSection.prototype.has_weight_row =
  Worksheet.prototype.has_weight_row ;

function* list_titles(group, start, end) {
  var title_start = start;
  var current_start = title_start;
  while (true) {
    let title_cols;
    title_cols = group.sheetbuf.find_merge("title_row", current_start, end);
    if (title_cols == null) {
      yield [title_start, end];
      return;
    }
    current_start = title_cols[1] + 1;
    if (title_cols[0] == title_start) {
      continue;
    }
    yield [title_start, title_cols[0] - 1];
    title_start = title_cols[0];
  }
}

Worksheet.prototype.list_sections = function*() {
  try {
    for ( let [section_start, section_end]
      of list_titles(this.group, this.dim.start, this.dim.end)
    ) {
      if (section_end < this.data_start || section_start > this.data_end) {
        throw new WorksheetSectionDetectError(
          "worksheet section titles are malformed " +
          "(each section must contain at least one data column)",
          this.title_range );
      }
      let section = new WorksheetSection( this,
        get_column_range_(this.sheet.getRange(
          1, section_start, 1, section_end - section_start + 1 ))
      );
      //section.check();
      yield section;
    }
  } catch (error) {
    if (error instanceof SheetBufferMergeOverlap) {
      throw new WorksheetSectionDetectError(
        "merged ranges overlap worksheet title range",
        this.title_range );
    } else {
      throw error;
    }
  }
}

Worksheet.surrounding_section = function(group, worksheet, range) {
  if (group != null && !(group instanceof StudyGroup)) {
    throw new Error("Worksheet.surrounding_section: type error (group)");
  }
  if (worksheet == null) {
    worksheet = Worksheet.surrounding(group, range);
  } else if (!(worksheet instanceof Worksheet)) {
    throw new Error("Worksheet.surrounding_section: type error (worksheet)");
  }
  group = worksheet.group; // may have been null
  var range_start = range.getColumn(), range_end = range.getLastColumn();
  var section_start, section_end;
  let title_cols = group.sheetbuf.find_last_merge( "title_row",
    range_end, worksheet.dim.start,
    {allow_overlap_start: true} );
  if (title_cols == null) {
    section_start = worksheet.dim.start;
  } else {
    section_start = title_cols[0];
  }
  let next_title_cols = group.sheetbuf.find_merge( "title_row",
    title_cols != null ? title_cols[1] + 1 : section_start + 1,
    worksheet.dim.end,
    {allow_overlap_start: true} );
  if (next_title_cols == null) {
    section_end = worksheet.dim.end;
  } else {
    section_end = next_title_cols[0] - 1;
  }
  if (section_start > range_start || section_end < range_end) {
    throw new WorksheetSectionDetectError(
      "unable to locate surrounding worksheet section",
      range );
  }
  var section = new WorksheetSection( worksheet,
    get_column_range_(worksheet.sheet.getRange(
      1, section_start, 1, section_end - section_start + 1 ))
  );
  section.check();
  return section;
}

WorksheetSection.prototype.get_location = function(options = {}) {
  ({
    check_id: options.check_id = true,
      // check value from the title note agains actual metadata
  } = options);
  var title_id = this.get_title_metadata_id({check: options.check_id});
  return {
    title_id: title_id,
    offset: this.dim.offset,
    width: this.dim.width,
  };
}

Worksheet.prototype.find_section_by_location = function(location) {
  var {title_id, offset = null} = location;
  var column = offset != null ? this.dim.start + offset : null;
  var sheet = this.group.sheet;
  find_column: {
    if (column != null) {
      let title_column_range = get_column_range_(sheet.getRange(1, column));
      var title_metadata = title_column_range.createDeveloperMetadataFinder()
        .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.COLUMN)
        .withId(title_id)
        .find();
      if (title_metadata.length > 0)
        break find_column;
    }
    console.log(this.full_range.getA1Notation());
    console.log(title_id);
    title_metadata = this.full_range.createDeveloperMetadataFinder()
      .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.COLUMN)
      .onIntersectingLocations() // XXX shouldn't be really necessary
      .withId(title_id)
      .find();
    if (title_metadata.length > 0) {
      column = title_metadata[0].getLocation().getColumn().getColumn();
      break find_column;
    }
    throw new WorksheetSectionDetectError(
      "unable to locate worksheet section starting column" );
  }
  var end_column;
  {
    let title_cols = this.group.sheetbuf.find_merge( "title_row",
      column, this.dim.end );
    if (title_cols != null && title_cols[0] == column) {
      title_cols = this.group.sheetbuf.find_merge( "title_row",
        title_cols[1] + 1, this.dim.end );
    }
    if (title_cols == null) {
      end_column = this.dim.end;
    } else {
      end_column = title_cols[0] - 1;
    }
  }
  var section = new WorksheetSection(this,
    get_column_range_(sheet.getRange(1, column, 1, end_column - column + 1)) );
  section.check();
  return section;
}

Worksheet.prototype.add_section_after = function(section, options = {}) {
  // after this function is applied, all worksheet structures
  // (worksheets, sections) should be discarded
  ({
    data_width: options.data_width = initial.section_data_width,
    title: options.title = initial.section_title,
    date: options.date = null,
  } = options);
  if (section.worksheet !== this)
    throw new Error( "Worksheet().add_section_after: " +
      "the section does not belong to this worksheet" );
  var dim = {prev_end: section.dim.data_end};
  var max_formula_R1C1 = this.group.sheetbuf.get_formula("max_row", dim.prev_end);
  var weight_formula_R1C1 = this.group.sheetbuf.get_formula("weight_row", dim.prev_end);
  var category = this.get_category();
  if (weight_formula_R1C1 === "")
    weight_formula_R1C1 = null;
  if (section.dim.end > section.dim.data_end) {
    section.title_range.breakApart();
  }
  // XXX use sheetbuf's set and insert methods
  this.sheet.insertColumnsAfter(dim.prev_end, options.data_width);
  if (section.dim.end > section.dim.data_end) {
    section.title_range.offset(0,0,1,section.dim.data_end - section.dim.start + 1).merge();
  }
  dim.start = dim.title = dim.data_start = dim.prev_end + 1;
  dim.data_end = dim.prev_end + options.data_width;
  dim.end = section.dim.data_end < this.dim.data_end ? dim.data_end : this.dim.end + options.data_width;
  dim.width = dim.end - dim.start + 1;
  dim.data_width = dim.data_end - dim.data_start + 1;

  // XXX reset backgrounds
  // XXX and maybe in worksheet creation too
  // XXX and in add_columns

  this.sheet.getRange(this.group.dim.max_row, dim.data_start, 1, dim.data_width)
    .setFormulaR1C1(max_formula_R1C1);
  if (weight_formula_R1C1 != null)
    this.sheet.getRange(this.group.dim.weight_row, dim.data_start, 1, dim.data_width)
      .setFormulaR1C1(weight_formula_R1C1);
  if (category != null)
    this.sheet.getRange(this.group.dim.category_row, dim.start)
      .setValue(category);

  this.sheet.setColumnWidths(this.dim.data_start, this.dim.data_width, 21);

  this.reset_data_borders(
    dim.data_start, dim.data_end,
    {
      open_left: false,
      open_right: false,
      skip_weight: (weight_formula_R1C1 == null) }
  );
  var title_range = this.sheet.getRange(this.group.dim.title_row, dim.start, 1, dim.width);
  title_range.merge();
  title_range
    .setBorder(true, true, null, true, null, null);
  var title_id;
  {
    let quasi_section = {
      title_column_range: get_column_range_(this.sheet.getRange(1, dim.title)),
    };
    add_title_metadata.call(quasi_section);
    title_id = Worksheet.prototype.get_title_metadata.call(quasi_section).getId();
  }
  var title_note = Worksheet.format_title_note({ lines: [],
    date: options.date, title_id: title_id });
  title_range.getCell(1,1)
    .setValue(options.title)
    .setNote(title_note);
}

WorksheetSection.prototype.add_columns = function(data_index, data_width) {
  if (data_index < 0 || data_index > this.dim.data_width) {
    throw new Error( "WorksheetSection().add_columns: " +
      "data_index is invalid" );
  }
  var dim = {};
  dim.data_start = this.dim.data_start + data_index;
  dim.data_end = dim.data_start + data_width - 1;
  dim.data_width = data_width;
  var insert_column =  data_index > 0 ?
    this.dim.data_start + data_index - 1 :
    this.dim.data_start;
  var max_formula_R1C1 = this.group.sheetbuf.get_formula("max_row", insert_column);
  var weight_formula_R1C1 = this.group.sheetbuf.get_formula("weight_row", insert_column);
  if (weight_formula_R1C1 === "")
    weight_formula_R1C1 = null;
  // XXX use sheetbuf's set and insert methods
  if (data_index > 0) {
    this.sheet.insertColumnsAfter(insert_column, dim.data_width);
    if (insert_column == this.dim.end) {
      this.title_range.offset(0, 0, 1, this.dim.width + data_width).merge();
    }
  } else {
    if (dim.data_start == this.dim.start) {
      var category = this.worksheet.get_category();
      var metadata = this.get_title_metadata();
      var metadata_range = this.title_column_range;
      this.set_category(null);
    }
    this.sheet.insertColumnsBefore(insert_column, dim.data_width);
    if (dim.data_start == this.dim.start) {
      let title_range = this.title_range.offset(0, 0, 1, this.dim.width + dim.data_width);
      title_range.merge();
      title_range.offset(this.group.dim.category_row - this.group.dim.title_row, 0, 1, 1)
        .setValue(category);
      metadata.moveToColumn(metadata_range);
    }
  }
  this.sheet.getRange(this.group.dim.max_row, dim.data_start, 1, dim.data_width)
    .setFormulaR1C1(max_formula_R1C1);
  if (weight_formula_R1C1 != null)
    this.sheet.getRange(this.group.dim.weight_row, dim.data_start, 1, dim.data_width)
      .setFormulaR1C1(weight_formula_R1C1);

  this.sheet.setColumnWidths(this.dim.data_start, this.dim.data_width, 21);

  this.worksheet.reset_data_borders(
    dim.data_start, dim.data_end,
    {
      open_left: data_index > 0,
      open_right: data_index < this.dim.data_width,
      skip_weight: (weight_formula_R1C1 == null) }
  );
}

WorksheetSection.prototype.remove_excess_columns = function() { // {{{
  var label_values = this.group.sheetbuf.slice_values("label_row", this.dim.data_start, this.dim.data_end);
  var data_values = this.data_range.getValues();

  // XXX use sheetbuf's set and insert methods
  var removed_count = 0;
  var removing_series = 0;
  for (let i = this.dim.data_width - 1; i >= -1; --i) {
    let col_is_blank = true;
    if (i < 0) {
      col_is_blank = false;
    } else if (label_values[i] != "") {
      col_is_blank = false;
    } else {
      for (var data_row of data_values) {
        if (data_row[i] != "") {
          col_is_blank = false;
          break;
        }
      }
    }
    if (col_is_blank) {
      ++removing_series;
    } else if (removing_series > 0) {
      removed_count += removing_series;
      if (i < 0) {
        if (removing_series == this.dim.data_width) {
          return -1;
        }
        if (this.dim.start == this.dim.data_start) {
          var title = this.get_title();
          var title_note = this.get_title_note();
          var category = this.worksheet.get_category();
          var metadata = this.get_title_metadata();
          metadata.moveToColumn(get_column_range_(this.sheet.getRange(1, this.dim.title + removing_series)));
        }
      }
      this.sheet.deleteColumns(this.dim.data_start + i + 1, removing_series);
      if (i < 0 && this.dim.start == this.dim.data_start) {
        this.title_range.getCell(1,1)
          .setValue(title)
          .setNote(title_note)
        .offset(this.group.dim.category_row - this.group.dim.title_row, 0)
          .setValue(category);
      }
      removing_series = 0;
    }
  }
  return removed_count;
} // }}}

WorksheetSection.prototype.alloy_subproblems = function(options = {}) {
  var labels = this.group.sheetbuf.slice_values( "label_row",
    this.dim.data_start, this.dim.data_end );
  ({
    skip_weight: options.skip_weight = ! this.has_weight_row()
  } = options);
  var ranges_heights = [
    [this.label_range, 1],
    [this.data_range, this.group.dim.data_height],
    [this.max_range, 1] ];
  if (!options.skip_weight) {
    ranges_heights.push([this.weight_range, 1]);
  }
  let l = null, lb = null;
  let alloying = 0;
  labels.push(null);
  for (let i = 0; i < labels.length; ++i) {
    let label = labels[i];
    if (label != null)
      label = label.toString();
    let labelbase;
    if (label == null) {
      labelbase = null;
    } else {
      let match = /[a-zа-яё'*]*$/.exec(label);
      if (match == null)
        throw new Error("WorksheetSection().alloy_subproblems: internal error");
      labelbase = label.substring(0, label.length - match[0].length);
    }
    if (l != null && label != null && labelbase != "" && lb == labelbase && l != label) {
      l = label;
      ++alloying;
    } else {
      if (alloying > 1) {
        for (let [range, height] of ranges_heights) {
          range.offset(0, i - alloying, height, alloying)
            .setBorder( null, null, null, null, true, null,
              "#cccccc", SpreadsheetApp.BorderStyle.DOTTED );
        }
        SpreadsheetFlusher.reset();
      }
      l = label;
      lb = labelbase;
      alloying = 1;
    }
  }
}

return Worksheet;
}(); // end Worksheet namespace

// vim: set fdm=marker sw=2 :
