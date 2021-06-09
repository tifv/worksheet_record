class WorksheetError extends SpreadsheetError {};
class WorksheetInitError extends WorksheetError {};
class WorksheetDetectionError extends WorksheetError {};

class WorksheetSectionError extends WorksheetError {};
class WorksheetSectionInitError extends WorksheetSectionError {};
class WorksheetSectionDetectionError extends WorksheetSectionError {};


function worksheet_blank_namer_(date) {
    if (date == null)
        return "{Бланк}";
    let date_format = WorksheetDate.from_object(
        Object.assign({}, date, {period: null})
    ).format();
    if (date.period != null)
        date_format += ", пара " + date.period;
    return "{" + date_format + "}";
}

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
    full_range: function() {
        return get_column_range_(this.sheet, this.dim.start, this.dim.width);
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
    full_column_range: function() {
        return get_column_range_(this.sheet, this.dim.start, this.dim.width);
    },
    //data_column_range: function() {
    //    return get_column_range_(this.sheet, this.dim.data_start, this.dim.data_width);
    //},
}); // }}}

// WorksheetBase().get_title {{{
WorksheetBase.prototype.get_title = function() {
    return this.group.sheetbuf.get_value( "title_row",
        this.dim.title ).toString();
} // }}}

// WorksheetBase().set_title (value) {{{
WorksheetBase.prototype.set_title = function(value) {
    this.group.sheetbuf.set_value("title_row", this.dim.title, value);
} // }}}

// WorksheetBase().get_title_formula {{{
WorksheetBase.prototype.get_title_formula = function() {
    return this.group.sheetbuf.get_formula("title_row", this.dim.title);
} // }}}

// WorksheetBase().set_title_formula (formula, value_replace?) {{{
WorksheetBase.prototype.set_title_formula = function(formula, value_replace = "") {
    this.group.sheetbuf.set_formula( "title_row",
        this.dim.title, formula, value_replace );
} // }}}

// WorksheetBase().get_title_note {{{
WorksheetBase.prototype.get_title_note = function() {
    return this.group.sheetbuf.get_note("title_row", this.dim.title);
} // }}}

// WorksheetBase().set_title_note (note) {{{
WorksheetBase.prototype.set_title_note = function(note) {
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
    if (note == "")
        return data;
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
    return this.constructor.NoteData.parse(this.get_title_note());
} // }}}

// WorksheetBase().set_title_note_data (data) {{{
WorksheetBase.prototype.set_title_note_data = function(data) {
    this.set_title_note(data.format());
} // }}}

// WorksheetBase().get_title_metadata_id (options?) {{{
WorksheetBase.prototype.get_title_metadata_id = function(options = {}) {
    ({
        validate: options.validate = true,
            // true
            //   check value from the title note agains actual metadata
            //   creating metadata if necessary
            // false
            //   metadata will never be checked
            //   and method may return null
    } = options);
    var note_data = this.get_title_note_data();
    var note_id = note_data.get("id");
    if (!options.validate)
        return note_id;
    var metadatum = this.get_title_metadata();
    var metadatum_id = metadatum.getId();
    if (note_id != metadatum_id) {
        note_data.set("id", metadatum_id);
        this.set_title_note_data(note_data);
    }
    return metadatum_id;
} // }}}

// WorksheetBase().get_title_metadata {{{
WorksheetBase.prototype.get_title_metadata = function(options = {}) {
    ({
        create: options.create = null,
        // true: create metadatum, find it and return;
        // false: find metadatum and return it; otherwise return null;
        // null: find metadatum and return it; otherwise create, etc.
    } = options);
    if (options.create === true) {
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
    if (options.create === true)
        throw new Error("WorksheetBase().get_title_metadata: internal error");
    if (options.create === null)
        return this.get_title_metadata({create: true});
    if (options.create === false)
        return null;
} // }}}

// WorksheetBase().get_weight_formula {{{
WorksheetBase.prototype.get_weight_formula = function() {
    if (this.group.dim.weight_row == null)
        return null;
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
    return this.get_weight_formula() != null;
} // }}}

// WorksheetBase().get_max_formula {{{
WorksheetBase.prototype.get_max_formula = function() {
    if (this.group.dim.max_row == null)
        return null;
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
    return this.get_max_formula() != null;
} // }}}

// WorksheetBase().set_data_borders (start, end, options) {{{
WorksheetBase.prototype.set_data_borders = function( start, end,
    options = {}
) {
    ({
        horizontal: options.horizontal = true,
        max_row: options.max_row = null,
        weight_row: options.weight_row = null,
        left: options.left = {},
        right: options.right = {},
        alloy_columns: options.alloy_columns = null,
            // if alloy_subproblems is non-null, all other borders are ignored
    } = options);
    if (!options.left.open) {
        ({
            open: options.left.open = false,
            max_row: options.left.max_row = false,
            weight_row: options.left.weight_row = false,
            outer: options.left.outer = false,
        } = options.left);
    }
    if (!options.right.open) {
        ({
            open: options.right.open = false,
            max_row: options.right.max_row = false,
            weight_row: options.right.weight_row = false,
            outer: options.right.outer = false,
        } = options.right);
    }
    if (this.group.dim.max_row == null) {
        options.max_row = null;
        options.left.max_row = false;
        options.right.max_row = false;
    }
    if (this.group.dim.weight_row == null) {
        options.weight_row = null;
        options.left.weight_row = false;
        options.right.weight_row = false;
    }
    function get_R1C1( start_row, end_row = start_row,
        start_col = start, end_col = end,
    ) {
        return ( "R" + start_row + "C" + start_col +
            ":R" + end_row + "C" + end_col );
    }
    var row_dim = {
        title_row: this.group.dim.title_row,
        label_row: this.group.dim.label_row,
        max_row: this.group.dim.max_row,
        weight_row: this.group.dim.weight_row,
        data_start_row: this.group.dim.data_row,
        data_end_row: this.group.dim.data_row + this.group.dim.data_height - 1,
    };
    var solid_rows = [
        [row_dim.data_start_row, row_dim.data_end_row],
        [row_dim.label_row, row_dim.label_row],
    ];
    if ( options.max_row && options.weight_row &&
        row_dim.max_row - row_dim.weight_row == 1
    ) {
        solid_rows.push([row_dim.weight_row, row_dim.max_row]);
    } else if ( options.max_row && options.weight_row &&
        row_dim.max_row - row_dim.weight_row == -1
    ) {
        solid_rows.push([row_dim.max_row, row_dim.weight_row]);
    } else {
        if (options.max_row) {
            solid_rows.push([row_dim.max_row, row_dim.max_row]);
        }
        if (options.weight_row) {
            solid_rows.push([row_dim.weight_row, row_dim.weight_row]);
        }
    }
    const sheet = this.group.sheet;
    if (options.alloy_columns != null) { // {{{
        let ranges_R1C1 = [];
        for (let [a_row, b_row] of solid_rows) {
            for (let [a_col, b_col] of options.alloy_columns)
                ranges_R1C1.push(get_R1C1(a_row, b_row, a_col, b_col));
        }
        if (ranges_R1C1.length > 0)
            sheet.getRangeList(ranges_R1C1)
                .setBorder( null, null, null, null, true, null,
                    "#cccccc", SpreadsheetApp.BorderStyle.DOTTED );
        return;
    } // }}}
    if (options.horizontal) { // remove borders that should not be there {{{
        let ranges_R1C1 = [];
        if (options.max_row === false)
            ranges_R1C1.push(get_R1C1(row_dim.max_row))
        if (options.weight_row === false)
            ranges_R1C1.push(get_R1C1(row_dim.weight_row))
        if (ranges_R1C1.length > 0)
            sheet.getRangeList(ranges_R1C1)
                .setBorder(false, false, false, false, false, false);
    } // }}}
    if (options.horizontal) { // {{{
        let ranges_R1C1 = [];
        for (let [a_row, b_row] of solid_rows)
            ranges_R1C1.push(get_R1C1(a_row, b_row));
        sheet.getRangeList(ranges_R1C1)
            .setBorder(true, null, true, null, null, null)
            .setBorder(null, null, null, null, null, true,
                'black', SpreadsheetApp.BorderStyle.DOTTED );
        sheet.getRange(row_dim.title_row, start, 1, end - start + 1)
            .setBorder(true, null, null, null, null, null)
    } // }}}
    { // vertical internal {{{
        let ranges_R1C1 = [];
        for (let [a_row, b_row] of solid_rows)
            ranges_R1C1.push(get_R1C1(a_row, b_row));
        sheet.getRangeList(ranges_R1C1)
            .setBorder(
                null, options.left.open ? true : null,
                null, options.right.open ? true : null,
                true, null,
                'black', SpreadsheetApp.BorderStyle.DOTTED );
    } // }}}
    if (!options.left.open) { // vertical left {{{
        let ranges_R1C1 = [
            get_R1C1(row_dim.data_start_row, row_dim.data_end_row),
            get_R1C1(row_dim.label_row, row_dim.label_row),
        ];
        if (options.max_row || options.left.max_row) {
            ranges_R1C1.push(get_R1C1(row_dim.max_row));
        }
        if (options.weight_row || options.left.weight_row) {
            ranges_R1C1.push(get_R1C1(row_dim.weight_row));
        }
        sheet.getRangeList(ranges_R1C1)
            .setBorder(null, true, null, null, null, null);
        if (options.left.outer) { // vertical left outer {{{
            let colcol = [start - 1, start - 1];
            let ranges_R1C1 = [
                get_R1C1(row_dim.data_start_row, row_dim.data_end_row, ...colcol),
                get_R1C1(row_dim.label_row, row_dim.label_row, ...colcol),
            ];
            if (options.max_row) {
                ranges_R1C1.push(get_R1C1(row_dim.max_row, row_dim.max_row, ...colcol));
            }
            if (options.weight_row) {
                ranges_R1C1.push(get_R1C1(row_dim.weight_row, row_dim.weight_row, ...colcol));
            }
            sheet.getRangeList(ranges_R1C1)
                .setBorder(null, null, null, true, null, null);
        } // }}}
    } // }}}
    if (!options.right.open) { // vertical right {{{
        let ranges_R1C1 = [
            get_R1C1(row_dim.data_start_row, row_dim.data_end_row),
            get_R1C1(row_dim.label_row, row_dim.label_row),
        ];
        if (options.max_row || options.right.max_row) {
            ranges_R1C1.push(get_R1C1(row_dim.max_row));
        }
        if (options.weight_row || options.right.weight_row) {
            ranges_R1C1.push(get_R1C1(row_dim.weight_row));
        }
        sheet.getRangeList(ranges_R1C1)
            .setBorder(null, null, null, true, null, null);
        if (options.right.outer) { // vertical right outer {{{
            let colcol = [end + 1, end + 1];
            let ranges_R1C1 = [
                get_R1C1(row_dim.data_start_row, row_dim.data_end_row, ...colcol),
                get_R1C1(row_dim.label_row, row_dim.label_row, ...colcol),
            ];
            if (options.max_row || options.right.max_row) {
                ranges_R1C1.push(get_R1C1(row_dim.max_row, row_dim.max_row, ...colcol));
            }
            if (options.weight_row || options.right.weight_row) {
                ranges_R1C1.push(get_R1C1(row_dim.weight_row, row_dim.weight_row, ...colcol));
            }
            sheet.getRangeList(ranges_R1C1)
                .setBorder(null, true, null, null, null, null);
        } // }}}
    } // }}}
    SpreadsheetFlusher.reset();
} // }}}

// end WorksheetBase definition }}}2

// Worksheet constructor (group, start, end, data_start, data_end) {{{
class Worksheet extends WorksheetBase {
    constructor(group, start, end, data_start, data_end) {
        super();
        if (!(group instanceof StudyGroup)) {
            throw new Error("Worksheet.constructor: type error (group)");
        }
        Object.defineProperty(this, "group", { value: group,
            configurable: true });
        if (
            typeof start != "number"      || isNaN(start) ||
            typeof end != "number"        || isNaN(end)   ||
            typeof data_start != "number" || isNaN(data_start) ||
            typeof data_end != "number"   || isNaN(data_end)
        ) {
            throw new Error("Worksheet.constructor: type error (columns)");
        }
        Object.defineProperty(this, "dim", { value: {
            start: start, end: end,
            data_start: data_start, data_end: data_end,
            marker_start: data_start - 1, marker_end: data_end + 1,
            width: end - start + 1,
            data_width: data_end - data_start + 1,
            title: start,
        }, configurable: true });
    }
} // }}}

Worksheet.marker = {start: "‹", end: "›"};

Worksheet.initial = { // {{{
    data_width: 15,
    sum_column: +2,
    rating_column: +1,
    title: worksheet_blank_namer_(),
}; // }}}

define_lazy_properties_(Worksheet.prototype, { // {{{
    mirror_range: function() {
        if (this.group.dim.mirror_row == null)
            return null;
        return this.sheet.getRange(
            this.group.dim.mirror_row, this.dim.start,
            1, this.dim.width );
    },
    sum_column: function() {
        return (
            this.group.sheetbuf.find_value( "label_row", "S",
                this.dim.start, this.dim.marker_start - 1 ) ||
            this.group.sheetbuf.find_value( "label_row", "S",
                this.dim.marker_end + 1, this.dim.end)
        );
    },
    rating_column: function() {
        return (
            this.group.sheetbuf.find_value( "label_row", "Σ",
                this.dim.start, this.dim.marker_start - 1 ) ||
            this.group.sheetbuf.find_value( "label_row", "Σ",
                this.dim.marker_end + 1, this.dim.end)
        );
    },
    separator_column_range: function() {
        return get_column_range_(this.sheet, this.dim.end + 1);
    },
}); // }}}

// Worksheet().check (options?) {{{
Worksheet.prototype.check = function(options = {}) {
    ({
        dimensions: options.dimensions = true,
        markers:    options.markers    = true,
        titles:     options.titles     = true,
    } = options);
    if (options.dimensions) {
        if (
            this.dim.start > this.dim.marker_start ||
            this.dim.end < this.dim.marker_start ||
            this.dim.data_start > this.dim.data_end ||
            this.dim.start < 1
        ) {
            throw new WorksheetDetectionError(
                "worksheet dimensions are invalid (" +
                    "start=" + ACodec.debug(this.dim.start) + ", " +
                    "end="   + ACodec.debug(this.dim.end)   + ", " +
                    "data_start=" + ACodec.debug(this.dim.data_start) + ", " +
                    "data_end="   + ACodec.debug(this.dim.data_end)   +
                ")" );
        }
    }
    if (options.markers) {
        this.group.sheetbuf.ensure_loaded(this.dim.start, this.dim.end);
        if (
            this.dim.marker_start !=
            this.group.sheetbuf.find_value( "label_row",
                this.constructor.marker.start, this.dim.start, this.dim.end ) ||
            this.dim.marker_start !=
            this.group.sheetbuf.find_last_value( "label_row",
                this.constructor.marker.start, this.dim.end, this.dim.start ) ||
            this.dim.marker_end !=
            this.group.sheetbuf.find_value( "label_row",
                this.constructor.marker.end, this.dim.start, this.dim.end ) ||
            this.dim.marker_end !=
            this.group.sheetbuf.find_last_value( "label_row",
                this.constructor.marker.end, this.dim.end, this.dim.start )
        ) {
            throw new WorksheetDetectionError(
                "markers are missing or interwine",
                this.sheet.getRange(
                    this.group.dim.label_row, this.dim.marker_start,
                    1, this.dim.marker_end - this.dim.marker_start + 1 )
            );
        }
    }
    if (options.titles) {
        try {
            let start_title_cols = this.group.sheetbuf.find_merge( "title_row",
                this.dim.start, this.dim.end );
            if (start_title_cols == null)
                throw new WorksheetDetectionError(
                    "no title at the start of the title range",
                    this.title_range );
            let [title_start, title_end] = start_title_cols;
            if (
                title_start != this.dim.start ||
                title_end < this.dim.data_start &&
                this.group.sheetbuf.find_merge( "title_row",
                    title_end + 1, this.dim.data_start,
                    {allow_overlap_end: true} ) != null
            ) {
                throw new WorksheetDetectionError(
                    "misaligned title at the start of the title range",
                    this.title_range );
            }
            let end_title_cols = this.group.sheetbuf.find_merge( "title_row",
                this.dim.marker_end, this.dim.marker_end,
                {allow_overlap_start: true, allow_overlap_end: true} );
            if (end_title_cols != null && (
                end_title_cols[0] > this.dim.data_end ||
                end_title_cols[1] != this.dim.end
            ) || end_title_cols == null && (
                this.dim.marker_end < this.dim.end
            )) {
                throw new WorksheetDetectionError(
                    "misaligned title at the end of the title range",
                    this.title_range );
            }
        } catch (error) {
            if (error instanceof SheetBufferMergeOverlap) {
                throw new WorksheetDetectionError(
                    "merged ranges overlap worksheet title range",
                    this.title_range );
            } else {
                throw error;
            }
        }
    }
    return this;
} // }}}

// Worksheet().add_column_group {{{
Worksheet.prototype.add_column_group = function() {
    this.title_range.shiftColumnGroupDepth(+1);
} // }}}

// Worksheet.recolor_cf_rules (group, color_scheme, cfrules, start, end) {{{
Worksheet.recolor_cf_rules = function( group, color_scheme,
    ext_cfrules = null,
    start_col = this.find_start_col(group),
    end_col = group.sheetbuf.dim.sheet_width,
) {
    if (start_col == null)
      return;
    var cfrules = ext_cfrules || ConditionalFormatting.RuleList.load(group.sheet);
    var location_width = end_col - start_col + 1;
    var location_data = [ group.dim.data_row, start_col,
        group.dim.data_height, location_width ];
    var location_max = group.dim.max_row == null ? null :
        [group.dim.max_row, start_col,1, location_width];
    var location_weight = group.dim.weight_row == null ? null :
        [group.dim.weight_row, start_col,1, location_width];
    cfrules.replace({ type: "boolean",
        condition: this.get_cfcondition_data(),
        locations: [location_data, location_max].filter(l => l != null),
    }, this.get_cfeffect_data(color_scheme));
    var data_limit_filter = ConditionalFormatting.RuleFilter.from_object({
        type: "boolean",
        condition: {
            type: SpreadsheetApp.BooleanCriteria
                .NUMBER_GREATER_THAN_OR_EQUAL_TO,
            values: [null] },
        locations: [location_data, location_max].filter(l => l != null),
    });
    var data_limit_formula_regex = new RegExp(
        "=R" + group.dim.weight_row + "C\\d+" );
    data_limit_filter.condition.match = (cfcondition) => {
        return (
            cfcondition instanceof ConditionalFormatting.BooleanCondition &&
            cfcondition.type ==
                SpreadsheetApp.BooleanCriteria.NUMBER_GREATER_THAN_OR_EQUAL_TO &&
            cfcondition.values.length == 1 && (
                typeof cfcondition.values[0] == "string" &&
                data_limit_formula_regex.exec(cfcondition.values[0])
            ||
                typeof cfcondition.values[0] == "number" &&
                cfcondition.values[0] > 1
            )
        );
    }
    cfrules.replace( data_limit_filter,
        this.get_cfeffect_data_limit(color_scheme) );
    if (group.dim.weight_row != null) {
        cfrules.replace({ type: "gradient",
            condition: this.get_cfcondition_weight(group),
            locations: [location_weight],
        }, this.get_cfeffect_weight(color_scheme));
    }
    cfrules.replace({ type: "gradient",
        condition: group.get_cfcondition_rating(),
        locations: [location_data, location_max].filter(l => l != null),
    }, group.get_cfeffect_rating(color_scheme));
    if (ext_cfrules == null)
        cfrules.save(group.sheet);
} // }}}

// Worksheet().recolor_cf_rules (color_scheme) {{{
Worksheet.prototype.recolor_cf_rules = function( color_scheme,
    ext_cfrules = null,
) {
    this.constructor.recolor_cf_rules( this.group, color_scheme,
        ext_cfrules, this.dim.start, this.dim.end );
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
            [ this.group.dim.data_row, this.dim.data_start - 1,
                this.group.dim.data_height, this.dim.data_width + 2 ],
            [ this.group.dim.max_row, this.dim.data_start - 1,
                1, this.dim.data_width + 2 ],
        ].filter(([r,]) => r != null),
        effect: this.constructor.get_cfeffect_data(color_scheme),
    };
} // }}}

// Worksheet.get_cfeffect_data_limit (color_scheme) {{{
Worksheet.get_cfeffect_data_limit = function(color_scheme) {
    return new ConditionalFormatting.BooleanEffect(
        {background: HSL.to_hex(HSL.deepen(color_scheme.mark, 2))} );
} // }}}

// Worksheet().new_cfrule_data_limit {{{
Worksheet.prototype.new_cfrule_data_limit = function(color_scheme, limit = null) {
    if (limit == null && this.group.dim.weight_row == null)
        throw new WorksheetError( "Worksheet().new_cfrule_data_limit: " +
            "impossible without weight_row" );
    if (limit == null && this.sum_column == null)
        throw new WorksheetError( "Worksheet().new_cfrule_data_limit: " +
            "impossible without sum_column" );
    return { type: "boolean",
        condition: {
            type: SpreadsheetApp.BooleanCriteria
                .NUMBER_GREATER_THAN_OR_EQUAL_TO ,
            values: [ limit != null ? limit :
                ("=R" + this.group.dim.weight_row + "C" + this.sum_column) ]
        },
        ranges: [
            [ this.group.dim.data_row, this.dim.data_start - 1,
                this.group.dim.data_height, this.dim.data_width + 2 ],
            [ this.group.dim.max_row, this.dim.data_start - 1,
                1, this.dim.data_width + 2 ],
        ].filter(([r,]) => r != null),
        effect: this.constructor.get_cfeffect_data_limit(color_scheme),
    };
} // }}}

// Worksheet.get_cfcondition_weight (group) {{{
Worksheet.get_cfcondition_weight = function(group) {
    if (group.dim.weight_row == null)
        throw new WorksheetError( "Worksheet().get_cfcondition_weight: " +
            "impossible without weight_row" );
    var weight_R1C1 = "R" + group.dim.weight_row + "C[0]";
    var max_R1C1 = group.dim.max_row != null ?
        ("R" + group.dim.max_row + "C[0]") :
        ("R" + group.dim.data_row + "C[0]:C[0]");
    var student_count_R1C1 = group.student_count_cell == null ? null : (
        "R" + group.student_count_cell.getRow() +
        "C" + group.student_count_cell.getColumn() );
    var formula_base = ( "=R[0]C[0]" +
        " - 1/power(" + weight_R1C1 + "*max(1," + max_R1C1 + "),2)" );
    return new ConditionalFormatting.GradientCondition({
        min_type: SpreadsheetApp.InterpolationType.NUMBER,
        min_value: formula_base + " + 1",
        max_type: SpreadsheetApp.InterpolationType.NUMBER,
        max_value: formula_base + " + " +
            ( student_count_R1C1 == null ? "7" :
                "max(7," + student_count_R1C1 + ")" ),
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
    if (this.group.dim.weight_row == null)
        throw new WorksheetError( "Worksheet().new_cfrule_weight: " +
            "impossible without weight_row" );
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
    let cfranges = [];
    if (this.sum_column != null)
        cfranges.push([ this.group.dim.data_row, this.sum_column,
            this.group.dim.data_height, 2 ]);
    if (this.rating_column != null)
        cfranges.push([ this.group.dim.data_row, this.rating_column,
            this.group.dim.data_height, 2 ]);
    if (cfranges.length == 0)
        throw new WorksheetError( "Worksheet().new_cfrule_rating: " +
            "impossible without sum_column or rating_column" );
    if (this.group.dim.max_row != null)
        cfranges.push(...cfranges.map( ([, c, , w]) =>
            [this.group.dim.max_row, c, 1, w] ));
    return this.group.new_cfrule_rating(cfranges, color_scheme);
} // }}}

// Worksheet().get_category {{{
Worksheet.prototype.get_category = function() {
    if (this.group.dim.category_row == null)
        return null;
    var category = this.group.sheetbuf.get_value( "category_row",
        this.dim.title );
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
    if (this.group.dim.category_row == null)
        throw new WorksheetError( "Worksheet().get_category: " +
            "impossible without category_row" );
    var columns = new Set([this.dim.title]);
    if (this.rating_column != null)
        columns.add(this.rating_column);
    if (this.sum_column != null)
        columns.add(this.sum_column);
    for (let column of columns) {
        this.group.sheetbuf.set_value("category_row", column, code);
    }
    if (options.ignore_sections)
        return;
    for (let section in this.list_sections()) {
        if (columns.has(section.dim.title))
            continue;
        section.set_category(code);
    }
} // }}}

// Worksheet().get_metaweight {{{
Worksheet.prototype.get_metaweight = function() {
    if (this.rating_column == null || this.group.dim.weight_row == null)
        return null;
    var metaweight = this.group.sheetbuf.get_value( "weight_row",
        this.rating_column );
    if (typeof metaweight != "number")
        return null;
    return metaweight;
} // }}}

// Worksheet().set_metaweight (value, options?) {{{
Worksheet.prototype.set_metaweight = function(value, options = {}) {
    ({
        add: options.add = false,
        // add the value to the metaweight instead of replacing
    } = options);
    if (this.group.dim.weight_row == null)
        throw new WorksheetError( "Worksheet().set_metaweight: " +
            "impossible without weight_row" );
    if (this.rating_column == null)
        throw new WorksheetError( "Worksheet().set_metaweight: " +
            "impossible without rating_column" );
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
        this.rating_column, metaweight );
} // }}}

// Worksheet().is_unused () {{{
Worksheet.prototype.is_unused = function() {
    let title = this.get_title();
    return (typeof title == "string" && (
        title.startsWith("{") && title.endsWith("}") ||
        title == "" ));
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
    var first_title = group.sheetbuf.find_last_merge( "title_row",
        marker_start, null, {allow_overlap_start: true} );
    if (first_title == null)
        return null;
    return first_title[0] - 1;
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
        if (last_end >= marker_start)
            throw new Error("Worksheet.list: internal error");
        let marker_end = group.sheetbuf.find_value( "label_row",
            this.marker.end, marker_start + 2, end );
        if (marker_end == null)
            break;
        let rogue_start = group.sheetbuf.find_value( "label_row",
            this.marker.start, marker_start + 1, marker_end );
        if (rogue_start != null) {
            last_end = marker_start;
            continue;
        }
        let start_title_cols = group.sheetbuf.find_last_merge( "title_row",
            marker_start, last_end + 1,
            {allow_overlap_start: true, allow_overlap_end: true} )
        if (start_title_cols == null ||
            start_title_cols[0] <= last_end
        ) {
            last_end = marker_start;
            continue;
        }
        let end_title_cols = group.sheetbuf.find_merge( "title_row",
            marker_end, marker_end,
            {allow_overlap_start: true, allow_overlap_end: true} )
        if (end_title_cols != null &&
            end_title_cols[0] >= marker_end
        ) {
            last_end = marker_start;
            continue;
        }
        let worksheet_end = end_title_cols == null ? marker_end : end_title_cols[1];
        yield (new this( group,
            start_title_cols[0], worksheet_end,
            marker_start + 1, marker_end - 1 )).check();
        if (last_end >= worksheet_end)
            throw new Error("Worksheet.list: internal error");
        last_end = worksheet_end;
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
    var start_title_cols = group.sheetbuf.find_merge( "title_row",
        range_start, range_start,
        {allow_overlap_start: true, allow_overlap_end: true} );
    var marker_start = group.sheetbuf.find_last_value( "label_row",
        this.marker.start,
        start_title_cols != null ?
            start_title_cols[1] : range_start - 2 );
    var end_title_cols = group.sheetbuf.find_merge( "title_row",
        range_end, range_end,
        {allow_overlap_start: true, allow_overlap_end: true} );
    var marker_end = group.sheetbuf.find_value( "label_row",
        this.marker.end,
        end_title_cols != null ?
            end_title_cols[0] : range_end + 2 );
    if ( marker_start == null || marker_end == null ||
        marker_end - marker_start <= 1 )
    {
        throw new WorksheetDetectionError(
            "unable to locate surrounding worksheet",
            range );
    }
    var start_title_cols = group.sheetbuf.find_last_merge( "title_row",
        marker_start, null,
        {allow_overlap_start: true} );
    if (start_title_cols == null ||
        start_title_cols[0] > range_start
    ) {
        throw new WorksheetDetectionError(
            "unable to locate start title of the surrounding worksheet",
            range );
    }
    var end_title_cols = group.sheetbuf.find_merge( "title_row",
        marker_end, marker_end,
        {allow_overlap_start: true, allow_overlap_end: true} )
    if (end_title_cols != null && (
        end_title_cols[0] >= marker_end ||
        end_title_cols[1] < range_end
    ) || end_title_cols == null && (
        marker_end < range_end
    )) {
        throw new WorksheetDetectionError(
            "unable to locate end title of the surrounding worksheet",
            range );
    }
    return (new this( group,
        start_title_cols[0],
        end_title_cols == null ? marker_end : end_title_cols[1],
        marker_start + 1, marker_end - 1 )).check();
} // }}}

// Worksheet().get_location (options?) {{{
Worksheet.prototype.get_location = function(options = {}) {
    ({
        validate: options.validate = true,
            // check value from the title note agains actual metadata
    } = options);
    var title_id = this.get_title_metadata_id({validate: options.validate});
    if (!options.validate && title_id == null)
        return null;
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
    var marker_start = group.sheetbuf.find_value( "label_row",
        this.marker.start, column );
    if (marker_start == null) {
        throw new WorksheetDetectionError(
            "unable to locate worksheet starting marker" );
    }
    var marker_end = group.sheetbuf.find_value( "label_row",
        this.marker.end, marker_start );
    if (marker_start == null) {
        throw new WorksheetDetectionError(
            "unable to locate worksheet ending marker" );
    }
    var end_title_cols = group.sheetbuf.find_merge( "title_row",
        marker_end, marker_end,
        {allow_overlap_start: true, allow_overlap_end: true} )
    if (end_title_cols != null && end_title_cols[0] >= marker_end) {
        throw new WorksheetDetectionError(
            "unable to locate worksheet end title" );
    }
    return (new this( group,
        column,
        end_title_cols == null ? marker_end : end_title_cols[1],
        marker_start + 1, marker_end - 1 )).check();
} // }}}

// Worksheet().alloy_subproblems {{{
Worksheet.prototype.alloy_subproblems = function() {
    for (let section of this.list_sections()) {
        section.alloy_subproblems();
    }
} // }}}

// WorksheetSection constructor (worksheet, start, end) {{{
class WorksheetSection extends WorksheetBase {
    constructor(worksheet, start, end) {
        super();
        if (!(worksheet instanceof this.constructor.Worksheet)) {
            throw new Error("WorksheetSection.constructor: type error (worksheet)");
        }
        Object.defineProperty(this, "worksheet", { value: worksheet,
            configurable: true });
        if (
            typeof start != "number"      || isNaN(start) ||
            typeof end != "number"        || isNaN(end)
        ) {
            throw new Error("WorksheetSection.constructor: type error (columns)");
        }
        var data_start = start > worksheet.dim.data_start ?
                start : worksheet.dim.data_start,
            data_end = end < worksheet.dim.data_end ?
                end : worksheet.dim.data_end;
        Object.defineProperty(this, "dim", { value: {
            start: start, end: end,
            data_start: data_start, data_end: data_end,
            width: end - start + 1,
            data_width: data_end - data_start + 1,
            offset: start - worksheet.dim.start,
            data_offset: data_start - worksheet.dim.data_start,
            title: start,
        }, configurable: true });
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
}); // }}}

// WorksheetSection().check (options?) {{{
WorksheetSection.prototype.check = function(options = {}) {
    ({
        dimensions: options.dimensions = true,
        title: options.title = true,
    } = options);
    if (options.dimensions) {
        if (
            this.dim.start > this.dim.end ||
            this.dim.start < this.worksheet.dim.start ||
            this.dim.start > this.worksheet.dim.data_end ||
            this.dim.end < this.worksheet.dim.data_start ||
            this.dim.end > this.worksheet.dim.end
        ) {
            throw new WorksheetSectionDetectionError(
                "worksheet section dimensions are invalid (" +
                    "start=" + ACodec.debug(this.dim.start) + ", " +
                    "end="   + ACodec.debug(this.dim.end)   + ", " +
                    "data_start=" + ACodec.debug(this.dim.data_start) + ", " +
                    "data_end="   + ACodec.debug(this.dim.data_end)   +
                ")" );
        }
    }
    check_title:
    if (options.title) {
        try {
            let title_cols = this.group.sheetbuf.find_merge( "title_row",
                this.dim.start, this.dim.end );
            if (title_cols == null)
                break check_title;
            let [title_start, title_end] = title_cols;
            if (
                title_start != this.dim.start ||
                this.group.sheetbuf.find_merge( "title_row",
                    title_end + 1, this.dim.end ) != null
            ) {
                throw new WorksheetSectionDetectionError(
                    "misaligned title detected",
                    this.title_range );
            }
        } catch (error) {
            if (error instanceof SheetBufferMergeOverlap) {
                throw new WorksheetDetectionError(
                    "merged ranges overlap worksheet section title range",
                    this.title_range );
            } else {
                throw error;
            }
        }
    }
    return this;
} // }}}

// WorksheetSection().set_category (code?) {{{
WorksheetSection.prototype.set_category = function(
    code = this.worksheet.get_category()
) {
    if (this.group.dim.category_row == null)
        throw new WorksheetSectionError( "WorksheetSection().set_category: " +
            "impossible without category_row" );
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
    if (start == end + 1) {
        return;
    } else if (start > end) {
        throw new Error("internal error");
    }
    var title_start = start;
    var current_start = title_start;
    while (true) {
        let title_cols = group.sheetbuf.find_merge( "title_row",
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
            yield (new this.constructor.Section( this,
                section_start, section_end )).check();
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
    if (group == null)
        group = worksheet.group;
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
    return (new this( worksheet,
        section_start, section_end )).check();
} // }}}

// WorksheetSection().get_location (options?) {{{
WorksheetSection.prototype.get_location = function(options = {}) {
    ({
        validate: options.validate = true,
            // check value from the title note agains actual metadata
    } = options);
    var title_id = this.get_title_metadata_id({validate: options.validate});
    if (!options.validate && title_id == null)
        return null;
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
    return (new this.constructor.Section( this,
        column, end_column )).check();
} // }}}

// WorksheetSection().get_previous {{{
WorksheetSection.prototype.get_previous = function() {
    if (this.dim.start <= this.worksheet.dim.data_start)
        return null;
    return this.constructor.surrounding( this.group, this.worksheet,
        this.sheet.getRange(1, this.dim.start - 1) );
} // }}}

// WorksheetSection().get_next {{{
WorksheetSection.prototype.get_next = function() {
    if (this.dim.end >= this.worksheet.dim.data_end)
        return null;
    return this.constructor.surrounding( this.group, this.worksheet,
        this.sheet.getRange(1, this.dim.end + 1) );
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
    var section = this.worksheet.add_section_after(after_section, {
        title: options.title, data_width: 1,
        title_note_data: new this.constructor.NoteData(
            [["addendum-type", options.type]],
            [{key: "id"}, {key: "addendum-type"}] ),
        max_row: false,
        weight_row: false,
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
        max_row: options.max_row = true,
        weight_row: options.weight_row = true,
    } = options);
    if (section.worksheet !== this)
        throw new Error( "Worksheet().add_section_after: " +
            "the section does not belong to this worksheet" );
    var dim = {prev_end: section.dim.data_end};
    var max_formula =
        (options.max_row && this.group.dim.max_row != null) ?
            this.get_max_formula() : null;
    var weight_formula =
        (options.weight_row && this.group.dim.weight_row != null) ?
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

    this.group.sheetbuf.merge("title_row", dim.start, dim.end);
    this.sheet.getRange(this.group.dim.title_row, dim.start, 1, dim.width)
        .setBorder(true, true, null, true, null, null);
    var new_worksheet = new this.constructor( this.group,
        this.dim.start, this.dim.end + dim.data_width,
        this.dim.data_start, this.dim.data_end + dim.data_width,
    ).check();
    var new_section = new this.constructor.Section( new_worksheet,
        dim.start, dim.end ).check();
    var title_id = new_section.get_title_metadata({create: true}).getId();
    options.title_note_data.set("id", title_id);
    new_section.set_title_note_data(options.title_note_data);
    this.group.sheetbuf.set_value("title_row", dim.start, options.title);
    var prev_section = new_section.get_previous();
    var next_section = new_section.get_next();
    new_section.set_data_borders(
        new_section.dim.data_start, new_section.dim.data_end,
        {
            max_row: (max_formula != null),
            weight_row: (weight_formula != null),
            left: { open: false,
                max_row: (prev_section == null) ? false :
                    prev_section.has_max_row(),
                weight_row: (prev_section == null) ? false :
                    prev_section.has_weight_row(),
            },
            right: { open: false,
                max_row: (next_section == null) ? false :
                    next_section.has_max_row(),
                weight_row: (next_section == null) ? false :
                    next_section.has_weight_row(),
            },
        },
    );
    return new_section;
} // }}}

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
            if (this.group.dim.category_row != null)
                this.set_category(null); // XXX replace this with direct set
        }
        this.group.sheetbuf.insert_columns_before(
            insert_column, dim.data_width );
        if (dim.data_start == this.dim.start) {
            this.group.sheetbuf.merge( "title_row",
                this.dim.start, this.dim.end + dim.data_width );
            if (this.group.dim.category_row != null)
                this.group.sheetbuf.set_value( "category_row",
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

    var new_worksheet = new this.constructor.Worksheet( this.group,
        this.worksheet.dim.start, this.worksheet.dim.end + dim.data_width,
        this.worksheet.dim.data_start, this.worksheet.dim.data_end + dim.data_width,
    ).check();
    var new_section = new this.constructor( new_worksheet,
        this.dim.start, this.dim.end + dim.data_width ).check();
    let left_border_opt = {};
    if (data_index > 0) {
        left_border_opt.open = true;
    } else {
        let prev_section = new_section.get_previous();
        left_border_opt.max_row = (prev_section == null) ? false :
            prev_section.has_max_row();
        left_border_opt.weight_row = (prev_section == null) ? false :
            prev_section.has_weight_row();
    }
    let right_border_opt = {};
    if (data_index < this.dim.data_width) {
        right_border_opt.open = true;
    } else {
        let next_section = new_section.get_next();
        right_border_opt.max_row = (next_section == null) ? false :
            next_section.has_max_row();
        right_border_opt.weight_row = (next_section == null) ? false :
            next_section.has_weight_row();
    }
    this.worksheet.set_data_borders(
        dim.data_start, dim.data_end,
        {
            max_row: (max_formula != null),
            weight_row: (weight_formula != null),
            left: left_border_opt,
            right: right_border_opt,
        },
    );
    return new_section;
} // }}}

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
    var alloy_columns = [];
    labels.push(null);
    for (let i = 0, l = null, lb = null, s = 0; i < labels.length; ++i) {
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
            ++s;
        } else {
            if (s > 1) {
                alloy_columns.push([i - s, i - 1]);
            }
            l = label;
            lb = labelbase;
            s = 1;
        }
    }
    if (alloy_columns.length == 0)
        return;
    var data_start = this.dim.data_start;
    this.set_data_borders(this.dim.data_start, this.dim.data_end, {
        horizontal: false,
        max_row: this.has_max_row() ? true : null,
        weight_row: this.has_weight_row() ? true : null,
        alloy_columns: alloy_columns.map( ([x,y]) =>
            [data_start + x, data_start + y] ),
    });
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
     *   sum_column (integer)
     *     +n means that the sum column will be at the n'th column from
     *     the start of the worksheet
     *     -n means that it will be at the end of the worksheet
     *     0 means no sum column will be created
     *     default is +2
     *   rating_column (integer)
     *     default is +1
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
    var {
        data_width = this.initial.data_width,
        sum_column = this.initial.sum_column,
        rating_column = this.initial.rating_column,
    } = options;
    var sheet = group.sheet;
    var full_width = data_width + 2 +
      Math.max(0, sum_column, rating_column) +
      Math.max(0, -sum_column, -rating_column);
    var range_width = range.getNumColumns();
    if (range_width < full_width + 2) {
        group.sheetbuf.insert_columns_after(
            range.getColumn(), full_width + 2 - range_width );
    }
    var start = range.getColumn() + 1, end = start + full_width - 1;
    return (new this(group, start, end, options)).worksheet;
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

function WorksheetBuilder(group, start, end, options) { // {{{
    this.group = group;
    this.sheet = group.sheet;
    this.options = this.rectify_options(options);
    if (
        options.rating_column != 0 &&
        options.sum_column != 0 &&
        options.sum_column == options.rating_column
    ) {
        throw new Error("sum and rating columns cannot coincide");
    }
    var
        data_start = start + 1 + Math.max(0, +options.rating_column, +options.sum_column),
        data_end   = end   - 1 - Math.max(0, -options.rating_column, -options.sum_column);
    this.worksheet = new this.constructor.Worksheet( group,
        start, end, data_start, data_end );
    this.dim = this.worksheet.dim;
    if (options.rating_column != 0) {
        this.rating_column = (options.rating_column > 0) ?
            (this.dim.start + options.rating_column - 1) :
            (this.dim.end   + options.rating_column + 1);
    } else {
        this.rating_column = null;
    }
    if (options.sum_column != 0) {
        this.sum_column = (options.sum_column > 0) ?
            (this.dim.start + options.sum_column - 1) :
            (this.dim.end   + options.sum_column + 1);
    } else {
      this.sum_column = null;
    }
    if (this.options.color_scheme == null) {
        this.color_scheme = this.group.get_color_scheme();
    } else {
        this.color_scheme = ColorSchemes.copy(this.options.color_scheme);
    }
    this.set_column_widths();
    this.init_markers();
    this.title_id = this.add_title_metadata();
    this.init_title_range();
    this.init_data_range();
    this.init_label_range();
    if (this.group.dim.max_row != null)
        this.init_max_range();
    if (this.group.dim.weight_row != null)
        this.init_weight_range();
    if (this.rating_column != null)
        this.init_rating_range();
    if (this.sum_column != null)
        this.init_sum_range();
    if (this.group.dim.weight_row != null && this.rating_column != null)
        this.init_metaweight_cell();
    if (this.group.dim.mirror_row != null)
        this.init_mirror_range();
    if (options.category)
        this.worksheet.set_category( options.category,
            {ignore_sections: true} );
    // set_fixed_value_validation_(this.sheet.getRange(
    //     this.group.dim.data_row, this.dim.data_start - 1,
    //     this.group.dim.data_height, 1
    // ), "");
    // set_fixed_value_validation_(this.sheet.getRange(
    //     this.group.dim.data_row, this.dim.data_end + 1,
    //     this.group.dim.data_height, 1
    // ), "");
    this.init_borders();
    this.add_cf_rules();
    this.worksheet.add_column_group();
} // }}}

// WorksheetBuilder().rectify_options (options) => (options) {{{
WorksheetBuilder.prototype.rectify_options = function(options) {
    if (options == null)
        options = {};
    ({
        sum_column: options.sum_column = this.constructor.initial.sum_column,
        rating_column: options.rating_column = this.constructor.initial.rating_column,
        title: options.title = this.constructor.initial.title,
        date: options.date = null,
        color_scheme: options.color_scheme = null,
        category: options.category = null,
    } = options);
    if (this.group.dim.category_row == null)
        options.category = null;
    return options;
} // }}}

// WorksheetBuilder().set_column_widths {{{
WorksheetBuilder.prototype.set_column_widths = function() {
    if (this.dim.start < this.dim.marker_start) {
        this.sheet.setColumnWidths(
            this.dim.start, this.dim.marker_start - this.dim.start, 30 );
    }
    this.sheet.setColumnWidth(this.dim.marker_start, 5);
    this.sheet.setColumnWidths(this.dim.data_start, this.dim.data_width, 21);
    this.sheet.setColumnWidth(this.dim.marker_end, 5);
    if (this.dim.end > this.dim.marker_end) {
        this.sheet.setColumnWidths(
            this.dim.marker_end + 1, this.dim.end - this.dim.marker_end, 30 );
    }
    this.sheet.setColumnWidth(this.dim.end + 1, 13);
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
    if (this.group.dim.max_row == null)
        throw new Error("internal error");
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
    if (this.group.dim.weight_row == null)
        throw new Error("internal error");
    var data_column_R1C1 = 'R' + this.group.dim.data_row + 'C[0]:C[0]';
    var max_R1C1, max_formula;
    if (this.group.dim.max_row != null) {
        max_R1C1 = 'R' + this.group.dim.max_row + 'C[0]';
        max_formula = max_R1C1;
    } else {
        max_R1C1 = 'R' + this.group.dim.data_row + 'C[0]:C[0]';
        max_formula = 'max(0;' + max_R1C1 + ')';
    }
    var weight_formula_R1C1 = ''.concat(
        '= if(',
            max_formula, '; ',
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
    var data_row_rating_R1C1 =
        'R[0]C[' + (this.dim.data_start - 1 - this.rating_column) + ']:' +
        'R[0]C[' + (this.dim.data_end   + 1 - this.rating_column) + ']';
    var rating_formula_R1C1;
    if (this.group.dim.weight_row != null) {
        var weight_row_rating_R1C1 =
            'R' + this.group.dim.weight_row +
                'C[' + (this.dim.data_start - 1 - this.rating_column) + ']:' +
            'R' + this.group.dim.weight_row +
                'C[' + (this.dim.data_end   + 1 - this.rating_column) + ']';
        rating_formula_R1C1 = ''.concat(
            '=sumproduct(',
                weight_row_rating_R1C1, ';',
                data_row_rating_R1C1,
            ')'
        );
    } else {
        rating_formula_R1C1 = ''.concat(
            '=sum(',
                data_row_rating_R1C1,
            ')'
        );
    }
    var number_format = (this.group.dim.weight_row != null) ?
        "0.00;−0.00" : "0.#;−0.#";
    this.sheet.getRange(this.group.dim.data_row, this.rating_column, this.group.dim.data_height, 1)
        .setFormulaR1C1(rating_formula_R1C1)
        .setNumberFormat(number_format)
        .setFontSize(8);
    if (this.group.dim.max_row != null) {
        this.group.sheetbuf.set_formula( "max_row",
            this.rating_column, rating_formula_R1C1 );
        this.sheet.getRange(this.group.dim.max_row, this.rating_column)
            .setNumberFormat(number_format)
            .setFontSize(8);
    }
    this.group.sheetbuf.set_value( "label_row",
        this.rating_column, "Σ" );
} // }}}

// WorksheetBuilder().init_sum_range {{{
WorksheetBuilder.prototype.init_sum_range = function() {
    var data_row_sum_R1C1 =
        'R[0]C[' + (this.dim.data_start - 1 - this.sum_column) + ']:' +
        'R[0]C[' + (this.dim.data_end   + 1 - this.sum_column) + ']';
    var sum_formula_R1C1;
    if (this.group.dim.max_row != null) {
        var max_row_sum_R1C1 =
            'R' + this.group.dim.max_row +
                'C[' + (this.dim.data_start - 1 - this.sum_column) + ']:' +
            'R' + this.group.dim.max_row +
                'C[' + (this.dim.data_end   + 1 - this.sum_column) + ']';
        sum_formula_R1C1 = ''.concat(
            '=countifs(',
                max_row_sum_R1C1,  ';">0";',
                data_row_sum_R1C1, ';">0"',
            ')'
        );
    } else {
        sum_formula_R1C1 = ''.concat(
            '=countif(',
                data_row_sum_R1C1, ';">0"',
            ')'
        );
    }
    this.sheet.getRange(this.group.dim.data_row, this.sum_column, this.group.dim.data_height, 1)
        .setFormulaR1C1(sum_formula_R1C1)
        .setNumberFormat('0')
        .setFontSize(8);
    if (this.group.dim.max_row != null) {
        this.group.sheetbuf.set_formula( "max_row",
            this.sum_column, sum_formula_R1C1 );
        this.sheet.getRange(this.group.dim.max_row, this.sum_column)
            .setNumberFormat('0')
            .setFontSize(8);
    }
    this.group.sheetbuf.set_value( "label_row",
        this.sum_column, "S" );
} // }}}

// WorksheetBuilder().init_metaweight_cell {{{
WorksheetBuilder.prototype.init_metaweight_cell = function() {
    if (this.group.dim.weight_row == null)
        throw new Error("internal error");
    this.group.sheetbuf.set_value( "weight_row",
        this.rating_column, 1 );
    this.group.sheetbuf.set_note( "weight_row",
        this.rating_column, "вес листочка в рейтинге" );
    this.sheet.getRange(this.group.dim.weight_row, this.rating_column)
        .setNumberFormat('0.0;−0.0')
        .setFontSize(8);
} // }}}

// WorksheetBuilder().init_mirror_range {{{
WorksheetBuilder.prototype.init_mirror_range = function() {
    if (this.group.dim.mirror_row == null)
        throw new Error("internal error");
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
    this.worksheet.set_data_borders(
        this.dim.data_start, this.dim.data_end,
        {
            max_row: this.group.dim.max_row != null ? true : null,
            weight_row: this.group.dim.weight_row != null ? true : null,
            left:  {open: false, outer: true},
            right: {open: false, outer: true},
        } );
    var rating_sum_ranges = [];
    if (
        this.rating_column != null &&
        this.sum_column != null && (
            this.rating_column - this.sum_column == -1 ||
            this.rating_column - this.sum_column == +1
        )
    ) {
        if (this.rating_column - this.sum_column == -1) {
            rating_sum_ranges.push(this.sheet.getRange(
                this.group.dim.data_row, this.rating_column,
                this.group.dim.data_height, 2 ));
        } else {
            rating_sum_ranges.push(this.sheet.getRange(
                this.group.dim.data_row, this.sum_column,
                this.group.dim.data_height, 2 ));
        }
    } else {
        if (this.rating_column != null) {
            rating_sum_ranges.push(this.sheet.getRange(
                this.group.dim.data_row, this.rating_column,
                this.group.dim.data_height, 1 ));
        }
        if (this.sum_column != null) {
            rating_sum_ranges.push(this.sheet.getRange(
                this.group.dim.data_row, this.sum_column,
                this.group.dim.data_height, 1 ));
        }
    }
    for (let rating_sum_range of rating_sum_ranges) {
        rating_sum_range
            .setBorder(true, true, true, true, null, null)
            .setBorder( null, null, null, null, true, null,
                "black", SpreadsheetApp.BorderStyle.DOTTED );
        rating_sum_range.offset(
                this.group.dim.label_row - this.group.dim.data_row, 0, 1 )
            .setBorder(true, true, true, true, null, null)
            .setBorder( null, null, null, null, true, null,
                "black", SpreadsheetApp.BorderStyle.DOTTED );
        if (this.group.dim.max_row != null) {
            rating_sum_range.offset(
                    this.group.dim.max_row - this.group.dim.data_row, 0, 1 )
                .setBorder(true, true, true, true, null, null)
                .setBorder( null, null, null, null, true, null,
                    "black", SpreadsheetApp.BorderStyle.DOTTED );
        }
    }
    if (this.group.dim.weight_row != null && this.rating_column != null) {
        // metaweight cell
        this.sheet.getRange(this.group.dim.weight_row, this.rating_column)
            .setBorder(true, true, true, true, null, null);
    }
    this.sheet.getRange(
        this.group.dim.title_row,
        this.dim.start,
        this.group.dim.sheet_height - this.group.dim.title_row + 1,
        this.dim.width
    ).setBorder(true, true, true, true, null, null);
} // }}}

// WorksheetBuilder().add_cf_rules {{{
WorksheetBuilder.prototype.add_cf_rules = function() {
    var cfrules = [];
    cfrules.push(this.worksheet.new_cfrule_data(this.color_scheme));
    if (this.group.dim.weight_row != null) {
        cfrules.push(this.worksheet.new_cfrule_weight(this.color_scheme));
    }
    if (this.rating_column != null || this.sum_column != null) {
        cfrules.push(this.worksheet.new_cfrule_rating(this.color_scheme));
    }
    ConditionalFormatting.merge(this.sheet, ...cfrules);
} // }}}

return WorksheetBuilder;
}(); // end WorksheetBuilder namespace }}}1

// vim: set fdm=marker sw=4 :
