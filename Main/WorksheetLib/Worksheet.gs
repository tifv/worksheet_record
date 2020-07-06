class WorksheetError extends SpreadsheetError {};
class WorksheetInitError extends WorksheetError {};
class WorksheetDetectionError extends WorksheetError {};

class WorksheetSectionError extends WorksheetError {};
class WorksheetSectionInitError extends WorksheetSectionError {};
class WorksheetSectionDetectionError extends WorksheetSectionError {};


var Worksheet = function() { // namespace {{{1

class WorksheetBase {}; // {{{2

WorksheetBase.metadata_keys = { // {{{
    title: "worksheet-title",
}; // }}}

define_lazy_properties_(WorksheetBase.prototype, { // {{{
    group: function() {
        throw new Error("not implemented"); },
    sheet: function() {
        return this.group.sheet; },
    dim: function() {
        throw new Error("not implemented"); },
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
        if (this.group.dim.max_row == null)
            return null;
        return this.sheet.getRange(
            this.group.dim.max_row, this.dim.data_start,
            1, this.dim.data_width );
    },
    weight_range: function() {
        if (this.group.dim.weight_row == null)
            return null;
        return this.sheet.getRange(
            this.group.dim.weight_row, this.dim.data_start,
            1, this.dim.data_width );
    },
    title_column_range: function() {
        return get_column_range_(this.sheet, this.dim.title);
    },
}); // }}}

// WorksheetBase().get_title {{{
WorksheetBase.prototype.get_title = function() {
    // also applies to WorksheetSection
    return this.group.sheetbuf.get_value( "title_row",
        this.dim.title ).toString();
} // }}}

// WorksheetBase().set_title (value) {{{
WorksheetBase.prototype.set_title = function(value) {
    // also applies to WorksheetSection
    this.group.sheetbuf.set_value("title_row", this.dim.title, value);
} // }}}

// WorksheetBase().get_title_formula {{{
WorksheetBase.prototype.get_title_formula = function() {
    // also applies to WorksheetSection
    return this.group.sheetbuf.get_formula("title_row", this.dim.title);
} // }}}

// WorksheetBase().set_title_formula (formula, value_replace?) {{{
WorksheetBase.prototype.set_title_formula = function(formula, value_replace = "") {
    // also applies to WorksheetSection
    this.group.sheetbuf.set_formula( "title_row",
        this.dim.title, formula, value_replace );
} // }}}

// WorksheetBase().get_title_note {{{
WorksheetBase.prototype.get_title_note = function() {
    // also applies to WorksheetSection
    return this.group.sheetbuf.get_note("title_row", this.dim.title);
} // }}}

// WorksheetBase().set_title_note (note) {{{
WorksheetBase.prototype.set_title_note = function(note) {
    // also applies to WorksheetSection
    this.group.sheetbuf.set_note("title_row", this.dim.title, note);
} // }}}

WorksheetBase.NoteData = class NoteData extends Map { // {{{
    constructor(entries = [], lines = []) {
        super(entries);
        this.lines = lines;
    }
} // }}}

// WorksheetBase.NoteData.parse (note) {{{
WorksheetBase.NoteData.parse = function(note) {
    var data = new this();
    for (let line of note.split("\n")) {
        line_date: {
            let date = WorksheetDate.parse(line);
            if (date == null || data.has("date"))
                break line_date;
            data.set("date", date);
            data.lines.push({key: "date"});
            continue;
        }
        line_key: {
            let key_match = /^([0-9a-zA-Z_\-]+)=(.*)$/.exec(line);
            if (key_match == null)
                break line_key;
            let [, key, value] = key_match;
            if (data.has(key))
                break line_key;
            if (key == "id") {
                if (/^[0-9]+$/.exec(value) == null) {
                    break line_key;
                } else {
                    value = parseInt(value, 10);
                }
            } else if (key == "date") {
                break line_key;
            }
            data.set(key, value)
            data.lines.push({key: key});
            continue;
        }
        data.lines.push(line);
    }
    return data;
} // }}}

// WorksheetBase.NoteData().format {{{
WorksheetBase.NoteData.prototype.format = function() {
    var keys = new Set(this.keys());
    var lines = [];
    var push_key = (key) => {
        let value = this.get(key);
        if (key == "date") {
            lines.push(value.format());
        } else {
            lines.push(key + "=" + value);
        }
    }
    for (let line of (this.lines || [])) {
        if (typeof line == "string") {
            lines.push(line);
            continue
        }
        let key = line.key;
        if (key == null)
            continue;
        if (!keys.delete(key))
            continue;
        push_key(key);
    }
    for (let key of keys) {
        push_key(key);
    }
    return lines.join("\n");
} // }}}

// WorksheetBase().get_title_note_data {{{
WorksheetBase.prototype.get_title_note_data = function() {
    // also applies to WorksheetSection
    return this.constructor.NoteData.parse(this.get_title_note());
} // }}}

// WorksheetBase().set_title_note_data (data) {{{
WorksheetBase.prototype.set_title_note_data = function(data) {
    // also applies to WorksheetSection
    this.set_title_note(data.format());
} // }}}

// WorksheetBase().get_title_metadata_id (options?) {{{
WorksheetBase.prototype.get_title_metadata_id = function(options = {}) {
    // also applies to WorksheetSection
    ({
        check: options.check = true,
            // check value from the title note agains actual metadata
        write: options.write = true,
            // write value to the title note if necessary
            // if false, the method may return null
    } = options);
    if (!options.check || options.write) {
        var note_data = this.get_title_note_data();
        var note_id = note_data.get("id");
        if (!options.check && note_id != null)
            return note_id;
    }
    var metadatum = this.get_title_metadata({
        create: options.write ? null : false });
    if (!options.write && metadatum == null)
        return null;
    var metadatum_id = metadatum.getId();
    if (options.write && note_id != metadatum_id) {
        note_data.set("id", metadatum_id);
        this.set_title_note_data(note_data);
    }
    return metadatum_id;
} // }}}

// WorksheetBase().get_title_metadata {{{
WorksheetBase.prototype.get_title_metadata = function(options = {}) {
    // also applies to WorksheetSection
    ({
        create: options.create = null,
        // true: create metadatum, find it and return;
        // false: find metadatum and return it; otherwise return null;
        // null: find metadatum and return it; otherwise create, etc.
    } = options);
    if (options.create) {
        this.title_column_range.addDeveloperMetadata(
            this.constructor.metadata_keys.title,
            SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT );
    }
    var metadata = this.title_column_range.createDeveloperMetadataFinder()
        .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.COLUMN)
        .withVisibility(SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT)
        .withKey(this.constructor.metadata_keys.title)
        .find();
    if (metadata.length > 0)
        return metadata[0];
    if (options.create)
        throw new Error("WorksheetBase().get_title_metadata: internal error");
    if (options.create == null)
        return this.get_title_metadata({create: true});
    if (options.create == false)
        return null;
} // }}}

// WorksheetBase().get_weight_formula {{{
WorksheetBase.prototype.get_weight_formula = function() {
    // also applies to WorksheetSection
    for (let i = 0; i < this.dim.data_width; ++i) {
        let formula = this.group.sheetbuf.get_formula( "weight_row",
            this.dim.data_start + i )
        if (formula !== "")
            return formula;
        let value = this.group.sheetbuf.get_value( "weight_row",
            this.dim.data_start + i )
        if (value !== "")
            return value;
    }
    return null;
} // }}}

// WorksheetBase().has_weight_row {{{
WorksheetBase.prototype.has_weight_row = function() {
    // also applies to WorksheetSection
    return this.get_weight_formula() != null;
} // }}}

// WorksheetBase().get_max_formula {{{
WorksheetBase.prototype.get_max_formula = function() {
    // also applies to WorksheetSection
    for (let i = 0; i < this.dim.data_width; ++i) {
        let formula = this.group.sheetbuf.get_formula( "max_row",
            this.dim.data_start + i )
        if (formula !== "")
            return formula;
        let value = this.group.sheetbuf.get_value( "max_row",
            this.dim.data_start + i )
        if (value !== "")
            return value;
    }
    return null;
} // }}}

// WorksheetBase().has_max_row {{{
WorksheetBase.prototype.has_max_row = function() {
    // also applies to WorksheetSection
    return this.get_max_formula() != null;
} // }}}

// end WorksheetBase definition }}}2

// Worksheet constructor (group, full_range) {{{
class Worksheet extends WorksheetBase {
    constructor(group, full_range) {
        super();
        if (group == null) {
            group = new StudyGroup(full_range.getSheet());
            group.check();
        } else if (!(group instanceof StudyGroup)) {
            throw new Error("Worksheet.constructor: type error (group)");
        }
        Object.defineProperty(this, "group", { value: group,
            configurable: true });
        this.full_range = full_range;
    }
} // }}}

Worksheet.marker = {start: "‹", end: "›"};

Worksheet.data_offset = {start: 3, end: 1, width: 4};

Worksheet.initial = { // {{{
    data_width: 15,
    title: "{Бланк}",
}; // }}}

define_lazy_properties_(Worksheet.prototype, { // {{{
    dim: function() {
        var dim = {};
        dim.start = this.full_range.getColumn();
        dim.end = this.full_range.getLastColumn();
        dim.data_start = dim.start + this.constructor.data_offset.start;
        dim.data_end   = dim.end   - this.constructor.data_offset.end;
        dim.marker_start = dim.data_start - 1;
        dim.marker_end   = dim.data_end   + 1;
        dim.width = dim.end - dim.start + 1;
        dim.data_width = dim.data_end - dim.data_start + 1;
        dim.title = dim.start;
        dim.rating = dim.start;
        dim.sum = dim.start + 1;
        return dim;
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
        if (this.group.dim.mirror_row == null)
            return null;
        return this.sheet.getRange(
            this.group.dim.mirror_row, this.dim.start,
            1, this.dim.width );
    },
    metaweight_cell: function() {
        return this.sheet.getRange(
            this.group.dim.weight_row, this.dim.rating );
    },
}); // }}}

// Worksheet().check (options?) {{{
Worksheet.prototype.check = function(options = {}) {
    ({
        dimensions: options.dimensions = true,
        markers: options.markers = true,
    } = options);
    if (options.dimensions) {
        if (
            this.full_range.isStartRowBounded() ||
            this.full_range.isEndRowBounded() ||
            this.dim.width < this.constructor.data_offset.width + 1
        ) {
            throw new WorksheetDetectionError(
                "full_range is incorrect " +
                "(must be row-unbounded and of width at least " +
                    (this.constructor.data_offset.width + 1) + ")",
                this.full_range );
        }
    }
    if (options.markers) {
        this.group.sheetbuf.ensure_loaded(this.dim.start, this.dim.end);
        if (
            this.dim.marker_start !=
            this.group.sheetbuf.find_last_value( "label_row",
                this.constructor.marker.start, this.dim.end, this.dim.start ) ||
            this.dim.marker_end !=
            this.group.sheetbuf.find_value( "label_row",
                this.constructor.marker.end, this.dim.start, this.dim.end )
        ) {
            throw new WorksheetDetectionError(
                "markers are missing or interwine",
                this.sheet.getRange(
                    this.sheet.dim.title_row, this.dim.start,
                    1, this.dim.width )
            );
        }
    }
} // }}}

// Worksheet().reset_column_widths {{{
Worksheet.prototype.reset_column_widths = function() {
    this.sheet.setColumnWidth(this.dim.rating, 30);
    this.sheet.setColumnWidth(this.dim.sum,    30);
    this.sheet.setColumnWidth(this.dim.data_start - 1, 5);
    this.sheet.setColumnWidths(this.dim.data_start, this.dim.data_width, 21);
    this.sheet.setColumnWidth(this.dim.data_end   + 1, 5);
    this.sheet.setColumnWidth(this.dim.end + 1, 13);
} // }}}

// Worksheet().reset_data_borders (col_begin, col_end, options) {{{
// options = { open_left: bool, open_right: bool,
//     skip_max: bool, skip_weight: bool }
//   open_left   --- no solid border on the left
//   open_right  --- no solid border on the left
//   skip_max --- do not draw borders around max_range
//   skip_weight --- do not draw borders around weight_range
Worksheet.prototype.reset_data_borders = function( col_begin, col_end,
    options = {}
) {
    const group = this.group;
    if (group.dim.max_row == null)
        options.skip_max = true;
    if (group.dim.weight_row == null)
        options.skip_weight = true;
    var col_num = col_end - col_begin + 1;
    var ranges = [
        this.sheet.getRange(group.dim.label_row, col_begin, 1, col_num),
        this.sheet.getRange(
            group.dim.data_row, col_begin, group.dim.data_height, col_num )
    ];
    if (!options.skip_max) {
        ranges.push(
            this.sheet.getRange(group.dim.max_row, col_begin, 1, col_num) );
    }
    if (!options.skip_weight) {
        ranges.push(
            this.sheet.getRange(group.dim.weight_row, col_begin, 1, col_num) );
    }
    // horizontal {{{
        for (let range of ranges) {
            range.setBorder(true, null, true, null, null, null);
        };
        // horizontal between weight_range and sum_range
        if ( !options.skip_max && !options.skip_weight &&
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

// Worksheet().add_column_group {{{
Worksheet.prototype.add_column_group = function() {
    this.title_range.shiftColumnGroupDepth(+1);
} // }}}

// Worksheet.recolor_cf_rules (group, color_scheme, cfrules, start_col) {{{
Worksheet.recolor_cf_rules = function( group, color_scheme,
    ext_cfrules = null, start_col = this.find_start_col(group),
) {
    if (start_col == null)
      return;
    var cfrules = ext_cfrules || ConditionalFormatting.RuleList.load(group.sheet);
    var location_width = group.sheetbuf.dim.sheet_width - start_col + 1;
    var location_data = [ group.dim.data_row, start_col,
        group.dim.data_height, location_width ];
    var location_max = [group.dim.max_row, start_col,1, location_width];
    var location_weight = [group.dim.weight_row, start_col,1, location_width];
    cfrules.replace({ type: "boolean",
        condition: this.get_cfcondition_data(),
        locations: [location_data, location_max],
    }, this.get_cfeffect_data(color_scheme));
    var data_limit_filter = ConditionalFormatting.RuleFilter.from_object({
        type: "boolean",
        condition: {
            type: SpreadsheetApp.BooleanCriteria
                .NUMBER_GREATER_THAN_OR_EQUAL_TO,
            values: [null] },
        locations: [location_data, location_max],
    });
    var data_limit_formula_regex = new RegExp(
      "=R" + group.dim.weight_row + "C\d+" );
    data_limit_filter.condition.match = (cfcondition) => {
        return (
            cfcondition instanceof ConditionalFormatting.BooleanCondition &&
            cfcondition.type ==
                SpreadsheetApp.BooleanCriteria.NUMBER_GREATER_THAN_OR_EQUAL_TO &&
            cfcondition.values.length == 1 &&
            typeof cfcondition.values[0] == "string" &&
            data_limit_formula_regex.exec(cfcondition.values[0])
        );
    }
    cfrules.replace( data_limit_filter,
      this.get_cfeffect_data_limit(color_scheme) );
    cfrules.replace({ type: "gradient",
      condition: this.get_cfcondition_weight(group),
      locations: [location_weight],
    }, this.get_cfeffect_weight(color_scheme));
    cfrules.replace({ type: "gradient",
      condition: group.get_cfcondition_rating(),
      locations: [location_data, location_max],
    }, group.get_cfeffect_rating(color_scheme));
    if (ext_cfrules == null)
        cfrules.save(this.sheet);
} // }}}

// Worksheet().recolor_cf_rules (color_scheme) {{{
Worksheet.prototype.recolor_cf_rules = function(color_scheme) {
    var cfrules = ConditionalFormatting.RuleList.load(this.sheet);
    var cfrule_data_obj = this.new_cfrule_data(color_scheme);
    cfrules.remove(Object.assign({}, cfrule_data_obj, {effect: null}));
    cfrules.insert(cfrule_data_obj);
    var cfrule_data_limit_obj = this.new_cfrule_data_limit(color_scheme);
    cfrules.replace(
        Object.assign({}, cfrule_data_limit_obj, {effect: null}),
        cfrule_data_limit_obj.effect );
    var cfrule_weight_obj = this.new_cfrule_weight(color_scheme);
    // XXX weight coloring may be missing
    // (e.g. when weights are reset by hand)
    cfrules.remove(Object.assign({}, cfrule_weight_obj, {effect: null}));
    cfrules.insert(cfrule_weight_obj);
    var cfrule_rating_obj = this.new_cfrule_rating(color_scheme);
    cfrules.remove(Object.assign({}, cfrule_rating_obj, {effect: null}));
    cfrules.insert(cfrule_rating_obj);
    cfrules.save(this.sheet);
} // }}}

// Worksheet.get_cfcondition_data {{{
Worksheet.get_cfcondition_data = function() {
    return new ConditionalFormatting.BooleanCondition({
        type: SpreadsheetApp.BooleanCriteria.NUMBER_GREATER_THAN,
        values: [0] });
} // }}}

// Worksheet.get_cfeffect_data (color_scheme) {{{
Worksheet.get_cfeffect_data = function(color_scheme) {
    return new ConditionalFormatting.BooleanEffect(
      {background: HSL.to_hex(color_scheme.mark)} );
} // }}}

// Worksheet().new_cfrule_data (color_scheme) {{{
Worksheet.prototype.new_cfrule_data = function(color_scheme) {
    return { type: "boolean",
        condition: this.constructor.get_cfcondition_data(this.group),
        ranges: [
            [
                this.group.dim.data_row, this.dim.data_start - 1,
                this.group.dim.data_height, this.dim.data_width + 2 ],
            [
                this.group.dim.max_row, this.dim.data_start - 1,
                1, this.dim.data_width + 2 ]
        ],
        effect: this.constructor.get_cfeffect_data(color_scheme),
    };
} // }}}

// Worksheet.get_cfeffect_data_limit (color_scheme) {{{
Worksheet.get_cfeffect_data_limit = function(color_scheme) {
    return new ConditionalFormatting.BooleanEffect(
        {background: HSL.to_hex(HSL.deepen(color_scheme.mark, 2))} );
} // }}}

// Worksheet().new_cfrule_data_limit {{{
Worksheet.prototype.new_cfrule_data_limit = function(color_scheme) {
    return { type: "boolean",
        condition: {
            type: SpreadsheetApp.BooleanCriteria
                .NUMBER_GREATER_THAN_OR_EQUAL_TO ,
            values: ["=R" + this.group.dim.weight_row + "C" + this.dim.sum] },
        ranges: [
            [
                this.group.dim.data_row, this.dim.data_start - 1,
                this.group.dim.data_height, this.dim.data_width + 2 ],
            [
                this.group.dim.max_row, this.dim.data_start - 1,
                1, this.dim.data_width + 2 ]
        ],
        effect: this.constructor.get_cfeffect_data_limit(color_scheme),
    };
} // }}}

// Worksheet.get_cfcondition_weight (group) {{{
Worksheet.get_cfcondition_weight = function(group) {
    var weight_R1C1 = "R" + group.dim.weight_row + "C[0]";
    var max_R1C1 = "R" + group.dim.max_row + "C[0]";
    var student_count_R1C1 = group.student_count_cell != null ? (
        "R" + group.student_count_cell.getRow() +
        "C" + group.student_count_cell.getColumn()
    ) : null;
    var formula_base = ( "=R[0]C[0]" +
        " - 1/power(" + weight_R1C1 + "*max(" + max_R1C1 + ",1),2)" );
    return new ConditionalFormatting.GradientCondition({
        min_type: SpreadsheetApp.InterpolationType.NUMBER,
        min_value: formula_base + " + 1",
        max_type: SpreadsheetApp.InterpolationType.NUMBER,
        max_value: formula_base + " + " +
            ( student_count_R1C1 != null ?
                "max(" + student_count_R1C1 + ",7)" : "7" ),
    });
} // }}}

// Worksheet.get_cfeffect_weight {{{
Worksheet.get_cfeffect_weight = function(color_scheme) {
    return new ConditionalFormatting.GradientEffect({
        min_color: HSL.to_hex(HSL.deepen(color_scheme.mark, 0.35)),
        max_color: HSL.to_hex(HSL.deepen(color_scheme.mark, 7.00)),
    });
} // }}}

// Worksheet().new_cfrule_weight {{{
Worksheet.prototype.new_cfrule_weight = function(color_scheme) {
    return { type: "gradient",
        condition: this.constructor.get_cfcondition_weight(this.group),
        ranges: [
            [
                this.group.dim.weight_row, this.dim.data_start - 1,
                1, this.dim.data_width + 2 ]
        ],
        effect: this.constructor.get_cfeffect_weight(color_scheme),
    };
} // }}}

// Worksheet().new_cfrule_rating {{{
Worksheet.prototype.new_cfrule_rating = function(color_scheme) {
    return this.group.new_cfrule_rating([
        [
            this.group.dim.data_row, this.dim.rating,
            this.group.dim.data_height, 2 ],
        [this.group.dim.max_row, this.dim.rating, 1, 2],
    ], color_scheme);
} // }}}

// Worksheet().get_category {{{
Worksheet.prototype.get_category = function() {
    var category = this.group.sheetbuf.get_value( "category_row",
        this.dim.start );
    if (category === "")
        return null;
    return category;
} // }}}

// Worksheet().set_category (code, options?) {{{
Worksheet.prototype.set_category = function(code, options = {}) {
    ({
        ignore_sections: options.ignore_sections = false,
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
} // }}}

// Worksheet().get_metaweight {{{
Worksheet.prototype.get_metaweight = function() {
    if (this.dim.rating == null || this.group.dim.weight_row == null)
        return null;
    var metaweight = this.group.sheetbuf.get_value( "weight_row",
        this.dim.rating );
    if (typeof metaweight != "number")
        return null;
    return metaweight;
} // }}}

// Worksheet().set_metaweight (value, options?) {{{
Worksheet.prototype.set_metaweight = function(value, options = {}) {
    ({
        add: options.add = false,
        // add the supplied value to the metaweight instead of replacing
    } = options);
    if (this.dim.rating == null || this.group.dim.weight_row == null)
        return;
    var metaweight;
    if (options.add) {
        var metaweight = this.get_metaweight();
        if (metaweight == null)
            return;
        metaweight += value;
    } else {
        metaweight = value;
    }
    this.group.sheetbuf.set_value( "weight_row",
        this.dim.rating, metaweight );
} // }}}

// Worksheet.find_title_column_by_id (group, title_id) {{{
Worksheet.find_title_column_by_id = function(group, title_id) {
    if (!(group instanceof StudyGroup)) {
        throw new Error( "Worksheet.find_title_column_by_id: " +
            "type error (group)" );
    }
    var metadata = group.sheet.createDeveloperMetadataFinder()
        .withLocationType(SpreadsheetApp.DeveloperMetadataLocationType.COLUMN)
        .withId(title_id)
        .find();
    if (metadata.length < 1)
        return null;
    return metadata[0].getLocation().getColumn().getColumn();
} // }}}

// Worksheet.find_start (group) {{{
Worksheet.find_start_col = function(group) {
    if (!(group instanceof StudyGroup)) {
        throw new Error("Worksheet.find_start: type error (group)");
    }
    var marker_start = group.sheetbuf.find_value( "label_row",
        this.marker.start, 1 );
    if (marker_start == null)
        return null;
    return marker_start - this.data_offset.start;
} // }}}

// Worksheet.list (group, start?, end?) {{{
Worksheet.list = function*(group, start = 1, end) {
    if (!(group instanceof StudyGroup)) {
        throw new Error("Worksheet.list: type error (group)");
    }
    var last_end = start - 1;
    while (true) {
        let marker_start = group.sheetbuf.find_value( "label_row",
            this.marker.start, last_end + 1, end );
        if (marker_start == null)
            break;
        let marker_end = group.sheetbuf.find_value( "label_row",
            this.marker.end, marker_start + 2, end );
        if (marker_end == null)
            break;
        let rogue_start = group.sheetbuf.find_value( "label_row",
            this.marker.start, marker_start + 1, marker_end );
        if (rogue_start != null) {
            if (last_end >= marker_start)
                throw new Error("Worksheet.list: internal error");
            last_end = marker_start;
            continue;
        }
        yield new this( group,
            get_column_range_( group.sheet,
                marker_start - this.data_offset.start + 1,
                marker_end - marker_start + 1 + this.data_offset.width - 2 )
        );
        if (last_end >= marker_end)
            throw new Error("Worksheet.list: internal error");
        last_end = marker_end;
    }
} // }}}

// Worksheet.surrounding (group, range) {{{
Worksheet.surrounding = function(group, range) {
    if (group == null) {
        group = new StudyGroup(range.getSheet());
        group.check();
    } else if (!(group instanceof StudyGroup)) {
        throw new Error("Worksheet.surrounding: type error (group)");
    }
    var range_start = range.getColumn(), range_end = range.getLastColumn();
    group.sheetbuf.ensure_loaded(range_start, range_end);
    var marker_start = group.sheetbuf.find_last_value( "label_row",
        this.marker.start, range_end + this.data_offset.start - 1 );
    var marker_end = group.sheetbuf.find_value( "label_row",
        this.marker.end, range_start - this.data_offset.end + 1);
    if ( marker_start == null || marker_end == null ||
        marker_end - marker_start <= 1 )
    {
        throw new WorksheetDetectionError(
            "unable to locate surrounding worksheet",
            range );
    }
    var worksheet = new this( group,
        get_column_range_( group.sheet,
            marker_start - this.data_offset.start + 1,
            marker_end - marker_start + 1 + this.data_offset.width - 2 )
    );
    worksheet.check();
    return worksheet;
} // }}}

// Worksheet().get_location (options?) {{{
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
} // }}}

// Worksheet.find_by_location (group, location) {{{
Worksheet.find_by_location = function(group, location) {
    if (!(group instanceof StudyGroup)) {
        throw new Error("Worksheet.find_by_location: type error (group)");
    }
    var {title_id, column = null, width = null} = location;
    var sheet = group.sheet;
    find_column: {
        if (column != null) {
            var title_column_range = get_column_range_(sheet, column);
            var title_metadata = title_column_range
                .createDeveloperMetadataFinder()
                .withLocationType(
                    SpreadsheetApp.DeveloperMetadataLocationType.COLUMN )
                .withId(title_id)
                .find();
            if (title_metadata.length > 0)
                break find_column;
        }
        title_metadata = sheet
            .createDeveloperMetadataFinder()
            .withLocationType(
                SpreadsheetApp.DeveloperMetadataLocationType.COLUMN )
            .withId(title_id)
            .find();
        if (title_metadata.length > 0) {
            column = title_metadata[0].getLocation().getColumn().getColumn();
            break find_column;
        }
        throw new WorksheetDetectionError(
            "unable to locate worksheet starting column" );
    }
    if (width != null)
        group.sheetbuf.ensure_loaded(column, column + width - 1);
    let end_column = group.sheetbuf.find_value( "label_row",
        this.marker.end, column );
    var worksheet = new this( group,
        get_column_range_( sheet,
            column, end_column + this.data_offset.end - column )
    );
    worksheet.check();
    return worksheet;
} // }}}

// Worksheet().alloy_subproblems {{{
Worksheet.prototype.alloy_subproblems = function() {
    for (let section of this.list_sections()) {
        section.alloy_subproblems();
    }
} // }}}

// WorksheetSection constructor (worksheet, full_range) {{{
class WorksheetSection extends WorksheetBase {
    constructor(worksheet, full_range) {
        super();
        this.worksheet = worksheet;
        this.full_range = full_range;
    }
} // }}}

WorksheetSection.initial = { // {{{
    data_width: 3,
    title: "Добавка",
}; // }}}

Worksheet.Section = WorksheetSection;
WorksheetSection.Worksheet = Worksheet;

define_lazy_properties_(WorksheetSection.prototype, { // {{{
    group: function() { return this.worksheet.group; },
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
}); // }}}

// WorksheetSection().check (options?) {{{
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
            throw new WorksheetSectionDetectionError(
                "range is incorrect " +
                "(must be row-unbounded and " +
                    "contained in the worksheet range " +
                    this.worksheet.full_range.getA1Notation() + ")",
                this.full_range );
        }
    }
    if (options.title) {
        let title_cols = this.group.sheetbuf.find_merge( "title_row",
            this.dim.start, this.dim.end );
        if (title_cols != null) {
            let [, title_end] = title_cols;
            if ( this.group.sheetbuf.find_merge( "title_row",
                    title_end + 1, this.dim.end ) != null
            ) {
                throw new WorksheetSectionDetectionError(
                    "misaligned title detected",
                    this.title_range );
            }
        }
    }
} // }}}

// WorksheetSection().set_category (code?) {{{
WorksheetSection.prototype.set_category = function(
    code = this.worksheet.get_category()
) {
    this.group.sheetbuf.set_value("category_row", this.dim.title, code);
} // }}}

// WorksheetSection().get_qualified_title {{{
WorksheetSection.prototype.get_qualified_title = function() {
    function capitalize(s) {
        if (s.length < 1)
            return s;
        return s[0].toUpperCase() + s.substring(1);
    }
    if (!this.is_addendum()) {
        return this.worksheet.get_title() +
            (this.dim.offset > 0 ? ". " + this.get_title() : "");
    } else {
        return this.get_original().get_qualified_title() + ". " +
            capitalize(this.get_title());
    }
} // }}}

// list_titles* (group, start, end) {{{
function* list_titles(group, start, end) {
    var title_start = start;
    var current_start = title_start;
    while (true) {
        let title_cols;
        title_cols = group.sheetbuf.find_merge( "title_row",
            current_start, end );
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
} // }}}

// Worksheet().list_sections* {{{
Worksheet.prototype.list_sections = function*(start_offset = 0) {
    var start = this.dim.start + start_offset
    try {
        for ( let [section_start, section_end]
            of list_titles(this.group, start, this.dim.end)
        ) {
            if ( section_end < this.data_start ||
                section_start > this.data_end
            ) {
                throw new WorksheetSectionDetectionError(
                    "worksheet section titles are malformed " +
                    "(each section must contain at least one data column)",
                    this.title_range );
            }
            let section = new this.constructor.Section( this,
                get_column_range_( this.sheet,
                    section_start, section_end - section_start + 1 )
            );
            //section.check();
            yield section;
        }
    } catch (error) {
        if (error instanceof SheetBufferMergeOverlap) {
            throw new WorksheetSectionDetectionError(
                "merged ranges overlap worksheet title range",
                (start_offset == 0) ?
                    this.title_range :
                    this.title_range.offset(
                        0, start_offset, 1, this.dim.width - offset )
            );
        } else {
            throw error;
        }
    }
} // }}}

// WorksheetSection.surrounding (group?, worksheet?, range) {{{
WorksheetSection.surrounding = function(group, worksheet, range) {
    if (group != null && !(group instanceof StudyGroup)) {
        throw new Error( "WorksheetSection.surrounding: " +
            "type error (group)" );
    }
    if (worksheet == null) {
        worksheet = Worksheet.surrounding(group, range);
    } else if (!(worksheet instanceof Worksheet)) {
        throw new Error( "WorksheetSection.surrounding: " +
            "type error (worksheet)" );
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
        throw new WorksheetSectionDetectionError(
            "unable to locate surrounding worksheet section",
            range );
    }
    var section = new this( worksheet,
        get_column_range_( worksheet.sheet,
            section_start, section_end - section_start + 1 )
    );
    section.check({dimensions: true, title: false});
    return section;
} // }}}

// WorksheetSection().get_location (options?) {{{
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
} // }}}

// Worksheet().find_section_by_location (location) {{{
Worksheet.prototype.find_section_by_location = function(location) {
    var {title_id, offset = null} = location;
    var column = offset != null ? this.dim.start + offset : null;
    var sheet = this.group.sheet;
    find_column: {
        if (column != null) {
            let title_column_range = get_column_range_(sheet, column);
            var title_metadata = title_column_range
                .createDeveloperMetadataFinder()
                .withLocationType(
                    SpreadsheetApp.DeveloperMetadataLocationType.COLUMN )
                .withId(title_id)
                .find();
            if (title_metadata.length > 0)
                break find_column;
        }
        title_metadata = this.full_range
            .createDeveloperMetadataFinder()
            .withLocationType(
                SpreadsheetApp.DeveloperMetadataLocationType.COLUMN )
            .onIntersectingLocations() // XXX shouldn't be really necessary
            .withId(title_id)
            .find();
        if (title_metadata.length > 0) {
            column = title_metadata[0].getLocation().getColumn().getColumn();
            break find_column;
        }
        throw new WorksheetSectionDetectionError(
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
    var section = new this.constructor.Section( this,
        get_column_range_(sheet, column, end_column - column + 1)
    );
    section.check();
    return section;
} // }}}

// WorksheetSection().is_addendum {{{
WorksheetSection.prototype.is_addendum = function() {
    return this.get_addendum_type() != null;
} // }}}

// WorksheetSection().get_addendum_type {{{
WorksheetSection.prototype.get_addendum_type = function() {
    return this.get_title_note_data().get("addendum-type");
} // }}}

// WorksheetSection().get_addendum (options?) {{{
WorksheetSection.prototype.get_addendum = function(options = {}) {
    ({
        type: options.type = "addendum",
        title: options.title = "addendum",
        label: options.label = "◦",
        column_width: options.column_width = 80,
    } = options);
    if (this.is_addendum())
        throw new WorksheetSectionDetectionError(
            "Cannot have an addendum section for an addendum section." );
    var after_section = this;
    for ( let section of
        this.worksheet.list_sections(this.dim.offset + this.dim.width) )
    {
        let section_type = section.get_addendum_type();
        if (section_type == null)
            break;
        if (section_type == options.type)
            return section;
        after_section = section;
    }
    const group = this.group;
    const sheet = this.sheet;
    // XXX remove weight and max in the section
    var section = this.worksheet.add_section_after(after_section, {
        title: options.title, data_width: 1,
        title_note_data: new this.constructor.NoteData(
            [["addendum-type", options.type]],
            [{key: "id"}, {key: "addendum-type"}] )
    });
    group.sheetbuf.set_value( "label_row", section.dim.data_start,
        options.label );
    sheet.getRange(group.dim.label_row, section.dim.data_start)
        .setFontWeight("normal");
    sheet.getRange(
        group.dim.data_row, section.dim.data_start,
        group.dim.data_height, 1
    ).setValue("◦");
    sheet.setColumnWidth(section.dim.data_start, options.column_width);
    return section;
} // }}}

// WorksheetSection().get_original () {{{
WorksheetSection.prototype.get_original = function() {
    var group = this.group;
    var original_section = null;
    for (var section of this.worksheet.list_sections()) {
        if (section.dim.offset >= this.dim.offset)
            break;
        if (!section.is_addendum())
            original_section = section;
    }
    if (original_section == null)
        throw new WorksheetSectionError(
            "An addendum section cannot be the first in its worksheet" );
    return original_section;
} // }}}

// Worksheet().add_section_after (section, options) {{{
Worksheet.prototype.add_section_after = function(section, options = {}) {
    // after this function is applied, all worksheet structures
    // (worksheets, sections) should be discarded
    ({
        data_width: options.data_width =
            this.constructor.Section.initial.data_width,
        title: options.title =
            this.constructor.Section.initial.title,
        title_note_data: options.title_note_data =
            new this.constructor.NoteData(),
    } = options);
    if (section.worksheet !== this)
        throw new Error( "Worksheet().add_section_after: " +
            "the section does not belong to this worksheet" );
    var dim = {prev_end: section.dim.data_end};
    var max_formula = this.group.dim.max_row != null ?
        this.get_max_formula() : null;
    var weight_formula = this.group.dim.weight_row != null ?
        this.get_weight_formula() : null;
    var category = this.get_category();
    if (section.dim.end > section.dim.data_end) {
        this.group.sheetbuf.unmerge( "title_row",
            section.dim.start, section.dim.end );
        this.group.sheetbuf.merge( "title_row",
            section.dim.start, section.dim.data_end );
    }
    this.group.sheetbuf.insert_columns_after(dim.prev_end, options.data_width)
    dim.start = dim.title = dim.data_start = dim.prev_end + 1;
    dim.data_end = dim.prev_end + options.data_width;
    dim.end = (section.dim.data_end < this.dim.data_end) ?
        dim.data_end : this.dim.end + options.data_width;
    dim.width = dim.end - dim.start + 1;
    dim.data_width = dim.data_end - dim.data_start + 1;

    // XXX reset backgrounds
    // XXX and maybe in worksheet creation too
    // XXX and also in add_columns

    // XXX refactor this as set_max_formula and set_weight_formula
    // XXX set numbers as values and formulas as formulas
    if (max_formula != null)
        this.group.sheetbuf.set_formulas( "max_row",
            dim.data_start, dim.data_end, max_formula, 0 );
    if (weight_formula != null)
        this.group.sheetbuf.set_formulas( "weight_row",
            dim.data_start, dim.data_end, weight_formula, 0 );
    if (category != null)
        this.group.sheetbuf.set_value("category_row", dim.start, category);

    this.sheet.setColumnWidths(dim.data_start, dim.data_width, 21);

    this.reset_data_borders(
        dim.data_start, dim.data_end,
        {
            open_left: false,
            open_right: false,
            skip_max: (max_formula == null),
            skip_weight: (weight_formula == null)
        },
    );
    this.group.sheetbuf.merge("title_row", dim.start, dim.end);
    this.sheet.getRange(this.group.dim.title_row, dim.start, 1, dim.width)
        .setBorder(true, true, null, true, null, null);
    var new_worksheet = new this.constructor( this.group,
        get_column_range_( this.sheet,
            this.dim.start, this.dim.width + options.data_width )
    );
    var new_section = new this.constructor.Section( new_worksheet,
        get_column_range_(this.sheet, dim.start, dim.width) );
    var title_id = new_section.get_title_metadata({create: true}).getId();
    options.title_note_data.set("id", title_id);
    new_section.set_title_note_data(options.title_note_data);
    this.group.sheetbuf.set_value("title_row", dim.start, options.title);
    return new_section;
} // }}}

// XXX return new section object
// WorksheetSection().add_columns (data_index, data_width) {{{
WorksheetSection.prototype.add_columns = function(data_index, data_width) {
    // after this function is applied, all worksheet structures
    // (worksheets, sections) should be discarded
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
    var max_formula = this.group.dim.max_row != null ?
        this.get_max_formula() : null;
    var weight_formula = this.group.dim.weight_row != null ?
        this.get_weight_formula() : null;
    if (data_index > 0) {
        this.group.sheetbuf.insert_columns_after(
            insert_column, dim.data_width );
        if (insert_column == this.dim.end)
            this.group.sheetbuf.merge( "title_row",
                this.dim.start, this.dim.end + dim.data_width );
    } else {
        if (dim.data_start == this.dim.start) {
            var category = this.worksheet.get_category();
            var metadata = this.get_title_metadata();
            var metadata_range = this.title_column_range;
            this.set_category(null);
        }
        this.group.sheetbuf.insert_columns_before(
            insert_column, dim.data_width );
        if (dim.data_start == this.dim.start) {
            this.group.sheetbuf.merge( "title_row",
                this.dim.start, this.dim.end + dim.data_width );
            this.group.sheetbug.set_value( "category_row",
                this.dim.title, category );
            metadata.moveToColumn(metadata_range);
        }
    }
    if (max_formula != null)
        this.group.sheetbuf.set_formulas( "max_row",
            dim.data_start, dim.data_end, max_formula );
    if (weight_formula != null)
        this.group.sheetbuf.set_formulas( "weight_row",
            dim.data_start, dim.data_end, weight_formula );

    this.sheet.setColumnWidths(dim.data_start, dim.data_width, 21);

    this.worksheet.reset_data_borders(
        dim.data_start, dim.data_end,
        {
            open_left: data_index > 0,
            open_right: data_index < this.dim.data_width,
            skip_max: (max_formula == null),
            skip_weight: (weight_formula == null)
        },
    );
} // }}}

// XXX return new section object
// WorksheetSection().remove_excess_columns {{{
WorksheetSection.prototype.remove_excess_columns = function() {
    // after this function is applied, all worksheet structures
    // (worksheets, sections) should be discarded
    var label_values = this.group.sheetbuf.slice_values( "label_row",
        this.dim.data_start, this.dim.data_end );
    var data_values = this.data_range.getValues();

    var removed_count = 0;
    var removing_series = 0;
    for (let i = this.dim.data_width - 1; i >= -1; --i) {
        let col_is_blank = true;
        if (i < 0) {
            col_is_blank = false;
        } else if (label_values[i] != "") {
            col_is_blank = false;
        } else {
            for (let data_row of data_values) {
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
                    metadata.moveToColumn(get_column_range_( this.sheet,
                        this.dim.title + removing_series ));
                }
            }
            this.group.sheetbuf.delete_columns(
                this.dim.data_start + i + 1, removing_series );
            if (i < 0 && this.dim.start == this.dim.data_start) {
                this.group.sheetbuf.set_value( "title_row",
                    this.dim.title, title );
                this.group.sheetbuf.set_note ( "title_row",
                    this.dim.title, title_note );
                if (category != null)
                    this.group.sheetbuf.set_value( "category_row",
                        this.dim.title, category );
            }
            removing_series = 0;
        }
    }
    return removed_count;
} // }}}

// WorksheetSection().alloy_subproblems (options) {{{
WorksheetSection.prototype.alloy_subproblems = function() {
    var labels = this.group.sheetbuf.slice_values( "label_row",
        this.dim.data_start, this.dim.data_end );
    var ranges_heights = [
        [this.label_range, 1],
        [this.data_range, this.group.dim.data_height],
    ];
    if (this.has_weight_row()) {
        ranges_heights.push([this.weight_range, 1]);
    }
    if (this.has_max_row()) {
        ranges_heights.push([this.max_range, 1]);
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
                throw new Error( "WorksheetSection().alloy_subproblems: " +
                    "internal error" );
            labelbase = label.substring(0, label.length - match[0].length);
        }
        if (
            l != null && label != null &&
            labelbase != "" && lb == labelbase && l != label
        ) {
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
} // }}}

return Worksheet;
}(); // end Worksheet namespace }}}1

var WorksheetBuilder = function() { // begin namespace {{{1

// XXX add alternative builder that creates standalone title
// like worksheet with markers, but without formatting datarange etc.

// XXX add alternative builder that creates a series of worksheets
// and returns an array

WorksheetBuilder.Worksheet = Worksheet;
Object.defineProperty( WorksheetBuilder, "initial",
    {enumerable: false, get: function() {
        return this.Worksheet.initial;
    }} );

// WorksheetBuilder.build (group, range, options) {{{
WorksheetBuilder.build = function(group, range, options) {
    /* options:
     *   data_width (number)
     *     default is initial.data_width
     *   and other options listed below
     */
    if (group == null) {
        group = new StudyGroup(range.getSheet());
        group.check();
    } else if (!(group instanceof StudyGroup)) {
        throw new Error("Worksheet build: type error (group)");
    }
    if (options == null)
        options = {};
    var {data_width = this.initial.data_width} = options;
    var sheet = group.sheet;
    // XXX refactor data_offset so that it may depend on options
    var full_width = data_width + this.Worksheet.data_offset.width;
    var range_width = range.getNumColumns();
    if (range_width < full_width + 2) {
        group.sheetbuf.insert_columns_after(
            range.getColumn(), full_width + 2 - range_width );
    }
    var full_range = get_column_range_( sheet,
        range.getColumn() + 1, full_width );
    return (new this(group, full_range, options)).worksheet;
} // }}}

/* options: {{{
 *     title (string)
 *         default is initial.title
 *     date (WorksheetDate)
 *         default is no date
 *     color_scheme (object)
 *         default will use group color scheme
 *     category (string)
 *         default is not to set category
 * }}} */

function WorksheetBuilder(group, full_range, options) { // {{{
    this.group = group;
    this.sheet = group.sheet;
    this.options = this.rectify_options(options);
    this.worksheet = new this.constructor.Worksheet(group, full_range);
    this.dim = this.worksheet.dim;
    if (this.options.color_scheme == null) {
        this.color_scheme = this.group.get_color_scheme();
    } else {
        this.color_scheme = ColorSchemes.copy(this.options.color_scheme);
    }
    this.worksheet.reset_column_widths();
    this.init_markers();
    this.title_id = this.add_title_metadata();
    this.init_title_range();
    this.init_data_range();
    this.init_label_range();
    this.init_max_range();
    this.init_weight_range();
    this.init_rating_range();
    this.init_sum_range();
    this.init_metaweight_cell();
    this.init_mirror_range();
    if (options.category)
        this.worksheet.set_category( options.category,
            {ignore_sections: true} );
    set_fixed_value_validation_(this.sheet.getRange(
        this.group.dim.data_row, this.dim.data_start - 1,
        this.group.dim.data_height, 1
    ), "");
    set_fixed_value_validation_(this.sheet.getRange(
        this.group.dim.data_row, this.dim.data_end + 1,
        this.group.dim.data_height, 1
    ), "");
    this.init_borders();
    this.add_cf_rules();
    this.worksheet.add_column_group();
} // }}}

// WorksheetBuilder().rectify_options (options) => (options) {{{
WorksheetBuilder.prototype.rectify_options = function(options) {
    if (options == null)
        options = {};
    ({
        title: options.title = this.constructor.initial.title,
        date: options.date = null,
        color_scheme: options.color_scheme = null,
        category: options.category = null,
    } = options);
    return options
} // }}}

// WorksheetBuilder().add_title_metadata () => (title_id) {{{
WorksheetBuilder.prototype.add_title_metadata = function() {
    return this.worksheet.get_title_metadata({create: true}).getId();
} // }}}

// WorksheetBuilder().init_markers {{{
WorksheetBuilder.prototype.init_markers = function() {
    this.group.sheetbuf.set_value( "label_row",
        this.dim.data_start - 1, this.constructor.Worksheet.marker.start );
    set_fixed_value_validation_(
        this.sheet.getRange(this.group.dim.label_row, this.dim.data_start - 1),
        this.constructor.Worksheet.marker.start );
    this.group.sheetbuf.set_value( "label_row",
        this.dim.data_end + 1, this.constructor.Worksheet.marker.end );
    set_fixed_value_validation_(
        this.sheet.getRange(this.group.dim.label_row, this.dim.data_end + 1),
        this.constructor.Worksheet.marker.end );
} // }}}

// WorksheetBuilder().init_title_range {{{
WorksheetBuilder.prototype.init_title_range = function() {
    this.group.sheetbuf.set_value( "title_row",
        this.dim.title, this.options.title );
    var note_data = new this.constructor.Worksheet.NoteData();
    if (this.options.date != null) {
        note_data.set("date", this.options.date);
    }
    note_data.set("id", this.title_id);
    this.worksheet.set_title_note_data(note_data);
    this.group.sheetbuf.merge("title_row", this.dim.start, this.dim.end);
    this.worksheet.title_range
        .setFontSize(12)
        .setFontWeight('bold')
        .setFontFamily("Times New Roman,serif");
} // }}}

// WorksheetBuilder().init_data_range {{{
WorksheetBuilder.prototype.init_data_range = function() {
    this.worksheet.data_range
        .setNumberFormat('0.#;−0.#');
} // }}}

// WorksheetBuilder().init_label_range {{{
WorksheetBuilder.prototype.init_label_range = function() {
    var labels = [];
    for (let i = 1; i <= this.dim.data_width; ++i) {
        if (i <= 3)
            labels[i-1] = i.toString();
        else
            labels[i-1] = null;
    }
    this.group.sheetbuf.set_values( "label_row",
        this.dim.data_start, this.dim.data_end, labels );
} // }}}

// WorksheetBuilder().init_max_range {{{
WorksheetBuilder.prototype.init_max_range = function() {
    var data_column_R1C1 = 'R' + this.group.dim.data_row + 'C[0]:C[0]';
    var max_formula_R1C1 = '=max(0;' + data_column_R1C1 + ')';
    this.group.sheetbuf.set_formulas( "max_row",
        this.dim.data_start, this.dim.data_end, max_formula_R1C1, 0 );
    this.worksheet.max_range
        .setNumberFormat('0.#;−0.#')
        .setFontSize(8);
} // }}}

// WorksheetBuilder().init_weight_range {{{
WorksheetBuilder.prototype.init_weight_range = function() {
    var data_column_R1C1 = 'R' + this.group.dim.data_row + 'C[0]:C[0]';
    var max_R1C1 = 'R' + this.group.dim.max_row + 'C[0]';
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
    this.group.sheetbuf.set_formulas( "weight_row",
        this.dim.data_start, this.dim.data_end, weight_formula_R1C1, 0 );
    this.worksheet.weight_range
        .setNumberFormat('#.0#;−#.0#;0')
        .setFontSize(8);
} // }}}

// WorksheetBuilder().init_rating_range {{{
WorksheetBuilder.prototype.init_rating_range = function() {
    var weight_row_rating_R1C1 =
        'R' + this.group.dim.weight_row +
            'C[' + (this.dim.data_start - 1 - this.dim.rating) + ']:' +
        'R' + this.group.dim.weight_row +
            'C[' + (this.dim.data_end   + 1 - this.dim.rating) + ']';
    var data_row_rating_R1C1 =
        'R[0]C[' + (this.dim.data_start - 1 - this.dim.rating) + ']:' +
        'R[0]C[' + (this.dim.data_end   + 1 - this.dim.rating) + ']';
    var rating_formula_R1C1 = ''.concat(
        '=sumproduct(',
            weight_row_rating_R1C1, ';',
            data_row_rating_R1C1,
        ')'
    );
    this.worksheet.rating_range
        .setFormulaR1C1(rating_formula_R1C1)
        .setNumberFormat('0.00;−0.00')
        .setFontSize(8);
    this.group.sheetbuf.set_formula( "max_row",
        this.dim.rating, rating_formula_R1C1 );
    this.sheet.getRange(this.group.dim.max_row, this.dim.rating)
        .setNumberFormat('0.00;−0.00')
        .setFontSize(8);
    this.group.sheetbuf.set_value( "label_row",
        this.dim.rating, "Σ" );
} // }}}

// WorksheetBuilder().init_sum_range {{{
WorksheetBuilder.prototype.init_sum_range = function() {
    var max_row_sum_R1C1 =
        'R' + this.group.dim.max_row +
            'C[' + (this.dim.data_start - 1 - this.dim.sum) + ']:' +
        'R' + this.group.dim.max_row +
            'C[' + (this.dim.data_end   + 1 - this.dim.sum) + ']';
    var data_row_sum_R1C1 =
        'R[0]C[' + (this.dim.data_start - 1 - this.dim.sum) + ']:' +
        'R[0]C[' + (this.dim.data_end   + 1 - this.dim.sum) + ']';
    var sum_formula_R1C1 = ''.concat(
        '=countifs(',
            max_row_sum_R1C1,  ';">0";',
            data_row_sum_R1C1, ';">0"',
        ')'
    );
    this.worksheet.sum_range
        .setFormulaR1C1(sum_formula_R1C1)
        .setNumberFormat('0')
        .setFontSize(8);
    this.group.sheetbuf.set_formula( "max_row",
        this.dim.sum, sum_formula_R1C1 );
    this.sheet.getRange(this.group.dim.max_row, this.dim.sum)
        .setNumberFormat('0')
        .setFontSize(8);
    this.group.sheetbuf.set_value( "label_row",
        this.dim.sum, "S" );
} // }}}

// WorksheetBuilder().init_metaweight_cell {{{
WorksheetBuilder.prototype.init_metaweight_cell = function() {
    this.group.sheetbuf.set_value( "weight_row",
        this.dim.rating, 1 );
    this.group.sheetbuf.set_note( "weight_row",
        this.dim.rating, "вес листочка в рейтинге" );
    this.worksheet.metaweight_cell
        .setNumberFormat('0.0;−0.0')
        .setFontSize(8);
} // }}}

// WorksheetBuilder().init_mirror_range {{{
WorksheetBuilder.prototype.init_mirror_range = function() {
    if (this.group.dim.mirror_row == null)
        return;
    var left_marker_mirror_R1C1  = 'R' + this.group.dim.label_row +
        'C[' + (this.dim.data_start - 1 - this.dim.start) + ']';
    var right_marker_mirror_R1C1 = 'R' + this.group.dim.label_row +
        'C[' + (this.dim.data_end   + 1 - this.dim.start) + ']';
    var title_mirror_R1C1 = 'R' + this.group.dim.title_row + 'C[0]';
    var label_row_mirror_R1C1 =
        'R' + this.group.dim.label_row + 'C[0]' + ':' +
        'R' + this.group.dim.label_row +
            'C[' + (this.dim.end - this.dim.start) + ']';
    var mirror_formula_R1C1 = ''.concat(
        '=iferror( ',
            'if(',
                'or(',
                    left_marker_mirror_R1C1,
                        '<>"' + this.constructor.Worksheet.marker.start + '";',
                    right_marker_mirror_R1C1,
                        '<>"' + this.constructor.Worksheet.marker.end + '"',
                '); ',
                'na(); ',
                'if(',
                    'or(',
                        'isblank(', title_mirror_R1C1, ');',
                        'left(', title_mirror_R1C1, ')="{"',
                    '); ',
                    'iferror(na()); ',
                    'arrayformula(', label_row_mirror_R1C1, ')',
                ')',
            '); ',
            'split(',
                'rept("#N/A ";columns(', label_row_mirror_R1C1, '));',
            '" ")',
        ')'
    );
    this.group.sheetbuf.set_formula( "mirror_row",
        this.dim.start, mirror_formula_R1C1 );
} // }}}

// WorksheetBuilder().init_borders {{{
WorksheetBuilder.prototype.init_borders = function() {
    this.sheet.getRange(
        this.group.dim.data_row, this.dim.start - 1,
        this.group.dim.data_height, this.dim.width + 2
    )
        .setBorder(true, null, true, null, null, null)
        .setBorder( null, null, null, null, null, true,
            "black", SpreadsheetApp.BorderStyle.DOTTED );
    this.sheet.getRange(
        this.group.dim.data_row - 1, this.dim.start - 1,
        1, this.dim.width + 2
    )
        .setBorder(null, null, true, null, null, null);
    this.worksheet.reset_data_borders(
        this.dim.data_start, this.dim.data_end,
        {open_left: false, open_right: false} );
    var rating_sum_range = this.sheet.getRange(
        this.group.dim.data_row, this.dim.start,
        this.group.dim.data_height, 2 );
    rating_sum_range
        .setBorder(true, true, true, true, null, null)
        .setBorder( null, null, null, null, true, null,
            "black", SpreadsheetApp.BorderStyle.DOTTED );
    rating_sum_range.offset(
            this.group.dim.label_row - this.group.dim.data_row, 0, 1 )
        .setBorder(true, true, true, true, null, null)
        .setBorder( null, null, null, null, true, null,
            "black", SpreadsheetApp.BorderStyle.DOTTED );
    rating_sum_range.offset(
            this.group.dim.max_row - this.group.dim.data_row, 0, 1 )
        .setBorder(true, true, true, true, null, null)
        .setBorder( null, null, null, null, true, null,
            "black", SpreadsheetApp.BorderStyle.DOTTED );
    this.worksheet.metaweight_cell
        .setBorder(true, true, true, true, null, null);
    this.sheet.getRange(
        this.group.dim.title_row,
        this.dim.start,
        this.group.dim.sheet_height - this.group.dim.title_row + 1,
        this.dim.width
    )
        .setBorder(true, true, true, true, null, null);
} // }}}

// WorksheetBuilder().add_cf_rules {{{
WorksheetBuilder.prototype.add_cf_rules = function() {
    ConditionalFormatting.merge(this.sheet,
        this.worksheet.new_cfrule_data(this.color_scheme),
        this.worksheet.new_cfrule_weight(this.color_scheme),
        this.worksheet.new_cfrule_rating(this.color_scheme),
    );
} // }}}

return WorksheetBuilder;
}(); // end WorksheetBuilder namespace }}}1

// vim: set fdm=marker sw=4 :
