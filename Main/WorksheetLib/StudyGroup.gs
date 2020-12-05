class StudyGroupError extends SpreadsheetError {};
class StudyGroupInitError extends StudyGroupError {};
class StudyGroupDetectionError extends StudyGroupError {};

// StudyGroupDim class
// * load() constructor loads from sheet
// * init(dim) constructor inits sheet, takes values from dim

var StudyGroupDim = function() { // namespace {{{1

const row_markers = { // {{{
    title_row:    "α",
    label_row:    "β",
    mirror_row:   "κ",
    category_row: "τ",
    max_row:      "μ",
    weight_row:   "σ",
}; // }}}

const row_names = Object.keys(row_markers);

const required_rows = new Set(["title_row", "label_row"]);

function assert_number_type(value) {
    if (typeof value != "number" || isNaN(value))
        throw new Error("StudyGroupDim: value is not a number");
    return value;
}

// load (sheet) {{{
function load(sheet) {
    var dim = {};
    define_lazy_properties_(dim, {
        sheet_id: function get_sheet_id() {
            return assert_number_type(
                sheet.getSheetId() );
        },
        sheet_height: function get_sheet_height() {
            return assert_number_type(
                sheet.getMaxRows() );
        },
        frozen_height: function get_frozen_height() {
            return assert_number_type(
                sheet.getFrozenRows() );
        },
        data_row: function get_data_row() {
            return assert_number_type(
                this.frozen_height + 1 );
        },
        data_height: function get_data_height() {
            return assert_number_type(
                this.sheet_height - this.data_row + 1 );
        },
    });
    setup_set_row.call(dim);
    if (dim.data_row <= 1) {
        throw new StudyGroupDetectionError(
            "no frozen rows in the sheet" );
    }
    SpreadsheetFlusher.add_dimensions( true,
        dim.sheet_id, dim.sheet_height, null,
        1, 1, dim.data_row - 1, 1 );
    var marker_values = sheet.getRange(1, 1, dim.data_row - 1, 1)
        .getValues()
        .map(v => v[0]);
    for (let name of row_names) {
        let row = marker_values.indexOf(row_markers[name]) + 1;
        if (row < 1) {
            set_row.call(dim, name, null);
        } else {
            set_row.call(dim, name, row);
        }
    }
    return dim;
} // }}}

// init (sheet, init_dim) {{{
function init(sheet, init_dim) {
    var dim = {};
    dim.data_row = init_dim.data_row;
    define_lazy_properties_(dim, {
        sheet_id: function() {
            return init_dim.sheet_id || sheet.getSheetId(); },
        data_row: function() {
            return init_dim.data_row; },
        sheet_height: function() {
            return init_dim.sheet_height || sheet.getMaxRows(); },
        frozen_height: function() {
            return init_dim.frozen_height || this.data_row - 1; },
        data_height: function() {
            return init_dim.data_height ||
                this.sheet_height - this.data_row + 1; },
    });
    setup_set_row.call(dim);
    if (dim.data_row == null) {
        throw new StudyGroupInitError(
            "data_row not defined" );
    }
    for (let name of row_names) {
        let row = init_dim[name];
        try {
            set_row.call(dim, name, row);
        } catch (error) {
            if (error instanceof StudyGroupDetectionError) {
                throw new StudyGroupInitError(error.message);
            }
            throw error;
        }
        if (row != null) {
            sheet.getRange(row, 1)
                .setValue(row_markers[name]);
        }
    }
    return dim;
} // }}}

// [dim].setup_set_row {{{
function setup_set_row() {
    Object.defineProperty(this, "inverse", {value: {}, enumerable: false});
} // }}}

// [dim].set_row (name, row) {{{
function set_row(name, row) {
    if (row == null && required_rows.has(name)) {
        throw new StudyGroupDetectionError(
            "missing required header row: " + name );
    }
    if (!(row == null || typeof row == "number" && !isNaN(row))) {
        throw new StudyGroupDetectionError(
            "invalid row number (" + name + "=" + row + ")" );
    }
    if (row != null && row >= this.data_row) {
        throw new StudyGroupDetectionError(
            "all header rows must lie before data_row, " +
            "i.e. in the frozen area (" + name + "=" + row + ")" );
    }
    if (name in this && this[name] != row) {
        throw new StudyGroupDetectionError(
            "row conflict: " +
            name + "=" + this[name] + " and " +
            name + "=" + row );
    }
    this[name] = row;
    if (row != null) {
        if (row in this.inverse && this.inverse[row] != name) {
            throw new StudyGroupDetectionError(
                "row conflict: " +
                this.inverse[row] + "=" + row + " and " +
                name + "=" + row );
        }
        this.inverse[row] = name;
    }
} // }}}

return {
    load: load, init: init,
    row_names: row_names,
    required_rows: required_rows,
};
}(); // end StudyGroupDim namespace }}}1

var StudyGroup = function() { // namespace {{{1

StudyGroup.metadata_keys = { // {{{
    main:           "worksheet_group",
    filename:       "worksheet_group-filename",
    color_scheme:   "worksheet_group-color_scheme",
    timetable:      "worksheet_group-timetable",
    worksheet_plan: "worksheet_group-worksheet_plan",
    student_count:  "worksheet_group-student_count",
}; // }}}

// StudyGroup constructor (sheet, name) {{{
function StudyGroup(sheet, name = null) {
    if (sheet == null) {
        throw new Error("StudyGroup: sheet must not be null");
    }
    this.sheet = sheet;
    if (name != null) {
        Object.defineProperty(this, "name", {value: name});
    }
} // }}}

// StudyGroup properties: name, sheetbuf, dim, student_count_cell {{{
define_lazy_properties_(StudyGroup.prototype, {
    name: function() {
        return this.sheet.getName(); },
    sheetbuf: function() {
        return new SheetBuffer( this.sheet,
            Object.fromEntries(
                StudyGroupDim.row_names
                    .map(k => [k, this.dim[k]])
                    .filter(([k, v]) => v != null) ),
            this.dim );
    },
    dim: function() {
        return StudyGroupDim.load(this.sheet);
    },
    student_count_cell: function() {
        return SheetMetacell.get( this.sheet,
            this.constructor.metadata_keys.student_count );
    }
}); // }}}

// StudyGroup().get_cfcondition_rating {{{
StudyGroup.prototype.get_cfcondition_rating = function() {
    var max_R1C1;
    if (this.dim.max_row != null) {
        max_R1C1 = "R" + this.dim.max_row + "C[0]";
    } else {
        max_R1C1 = "R" + this.dim.data_row + "C[0]:C[0]";
    }
    return new ConditionalFormatting.GradientCondition({
        min_type: SpreadsheetApp.InterpolationType.NUMBER,
        min_value: "=0.1*max(0.01," + max_R1C1 + ")",
        mid_type: SpreadsheetApp.InterpolationType.NUMBER,
        mid_value: "=0.5*max(0.01," + max_R1C1 + ")",
        max_type: SpreadsheetApp.InterpolationType.NUMBER,
        max_value: "=0.9*max(0.01," + max_R1C1 + ")",
    });
} // }}}

// StudyGroup().get_cfeffect_rating (color_scheme) {{{
StudyGroup.prototype.get_cfeffect_rating = function(color_scheme) {
    return new ConditionalFormatting.GradientEffect({
        min_color: "#ffffff",
        mid_color: HSL.to_hex(color_scheme.rating_mid),
        max_color: HSL.to_hex(color_scheme.rating_top),
    });
} // }}}

// StudyGroup().new_cfrule_rating (ranges, color_scheme) {{{
StudyGroup.prototype.new_cfrule_rating = function(
    cfranges, color_scheme
) {
    return { type: "gradient",
        condition: this.get_cfcondition_rating(),
        ranges: cfranges,
        effect: this.get_cfeffect_rating(color_scheme),
    };
} // }}}

// StudyGroup().add_metadatum (options) {{{
StudyGroup.prototype.add_metadatum = function(options = {}) {
    ({
        skip_remove: options.skip_remove = false,
            // make the call faster by not trying to remove old metadata
    } = options);
    if (!options || !options.skip_remove) {
        var metadata = this.sheet.createDeveloperMetadataFinder()
            .withLocationType(
                SpreadsheetApp.DeveloperMetadataLocationType.SHEET )
            .withVisibility(
                SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT )
            .withKey(this.constructor.metadata_keys.main)
            .find()
            .forEach(function(metadatum) {
                metadatum.remove();
            });
    }
    this.sheet.addDeveloperMetadata(
        this.constructor.metadata_keys.main,
        SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT );
} // }}}

// StudyGroup().set_student_count_cell (range) {{{
StudyGroup.prototype.set_student_count_cell = function(range) {
    Object.defineProperty( this, "student_count_cell",
      {configurable: true, value: range} );
    if (range == null) {
        SheetMetacell.unset( this.sheet,
            this.constructor.metadata_keys.student_count );
        return;
    }
    SheetMetacell.set( this.sheet,
        this.constructor.metadata_keys.student_count, range );
} // }}}

// StudyGroup().check (options) {{{
StudyGroup.prototype.check = function(options = {}) {
    ({
        metadata: options.metadata = true,
    } = options);
    if (options.metadata) {
        var metadata = this.sheet.createDeveloperMetadataFinder()
            .withLocationType(
                SpreadsheetApp.DeveloperMetadataLocationType.SHEET )
            .withVisibility(
                SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT )
            .withKey(this.constructor.metadata_keys.main)
            .find();
        if (metadata.length < 1)
            throw new StudyGroupDetectionError(
                "sheet " + this.name + " " +
                "is not marked as group by metadata" );
    }
} // }}}

// StudyGroup.find_by_name (spreadsheet, group_name) {{{
StudyGroup.find_by_name = function(spreadsheet, group_name) {
    var sheet = spreadsheet.getSheetByName(group_name);
    if (sheet == null)
        return null;
    var group = new StudyGroup(sheet, group_name);
    try {
        group.check();
    } catch (error) {
        if (error instanceof StudyGroupDetectionError)
            return null;
        throw error;
    }
    return group;
} // }}}

// StudyGroup.get_active (spreadsheet) {{{
StudyGroup.get_active = function(spreadsheet) {
    var sheet = spreadsheet.getActiveSheet();
    var group = new StudyGroup(sheet);
    try {
        group.check();
    } catch (error) {
        if (error instanceof StudyGroupDetectionError)
            return null;
        throw error;
    }
    return group;
} // }}}

// StudyGroup.list* (spreadsheet) {{{
StudyGroup.list = function*(spreadsheet) {
    for ( let metadatum of
        spreadsheet.createDeveloperMetadataFinder()
            .withLocationType(
                SpreadsheetApp.DeveloperMetadataLocationType.SHEET )
            .withVisibility(
                SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT )
            .withKey(this.metadata_keys.main)
            .find()
    ) {
        yield new StudyGroup(metadatum.getLocation().getSheet());
    }
} // }}}

// StudyGroup.list_names* (spreadsheet) {{{
StudyGroup.list_names = function*(spreadsheet) {
    for (let workgroup of StudyGroup.list(spreadsheet)) {
        yield workgroup.name;
    }
} // }}}

// StudyGroup().get_filename {{{
StudyGroup.prototype.get_filename = function() {
    var filename = SheetMetadata.get( this.sheet,
        this.constructor.metadata_keys.filename );
    if (filename == null)
        return this.name;
} // }}}

// StudyGroup().set_filename (filename) {{{
StudyGroup.prototype.set_filename = function(filename) {
    if (filename == null) {
        SheetMetadata.unset( this.sheet,
            this.constructor.metadata_keys.filename );
        return;
    }
    SheetMetadata.set( this.sheet,
        this.constructor.metadata_keys.filename, filename );
} // }}}

// StudyGroup().get_color_scheme {{{
StudyGroup.prototype.get_color_scheme = function() {
    // returned color_scheme may include one additional field, 'name'
    var color_scheme = SheetMetadata.get_object( this.sheet,
        this.constructor.metadata_keys.color_scheme );
    if (color_scheme == null)
        return ColorSchemes.get_default();
    color_scheme = ColorSchemes.copy(color_scheme, ["name"]);
    return color_scheme;
} // }}}

// StudyGroup().set_color_scheme (color_scheme) {{{
StudyGroup.prototype.set_color_scheme = function(color_scheme) {
    // color_scheme may include one additional field, 'name'
    if (color_scheme == null) {
        SheetMetadata.unset( this.sheet,
            this.constructor.metadata_keys.color_scheme );
        return;
    }
    color_scheme = ColorSchemes.copy(color_scheme, ["name"]);
    SheetMetadata.set_object( this.sheet,
        this.constructor.metadata_keys.color_scheme, color_scheme );
} // }}}

// StudyGroup().get_timetable {{{
StudyGroup.prototype.get_timetable = function() {
    var timetable = SheetMetadata.get_object( this.sheet,
        this.constructor.metadata_keys.timetable );
    if (timetable == null)
        return null;
    return timetable;
} // }}}

// StudyGroup().get_today_timetable (today?) {{{
StudyGroup.prototype.get_today_timetable = function(
    today = WorksheetDate.today()
) {
    var timetable = this.get_timetable();
    if (timetable == null)
        return null;
    return timetable[today.format()];
} // }}}

// StudyGroup().get_current_period {{{
StudyGroup.prototype.get_current_period = function(offset = 0) {
    var today_timetable = this.get_today_timetable();
    if (today_timetable == null)
      return null;
    let now = new Date();
    let current_time = now.getHours() * 60 + now.getMinutes() + offset;
    let current_period = null;
    for ( let [period, {time: period_time}]
        of Object.entries(today_timetable)
    ) {
        period = parseInt(period);
        if (isNaN(period))
            continue;
        if (period_time > current_time)
            continue;
        if (current_period == null || period > current_period)
            current_period = period;
    }
    return current_period;
} // }}}

// StudyGroup().set_timetable (timetable) {{{
StudyGroup.prototype.set_timetable = function(timetable) {
    if (timetable == null) {
        SheetMetadata.unset( this.sheet,
            this.constructor.metadata_keys.timetable );
        return null;
    }
    function objgenmap(obj, gen) {
      return Object.fromEntries(gen(Object.entries(obj)));
    }
    function copy_date_value(date_value) {
        return objgenmap(date_value || {}, function*(entries) {
            for (let [period_key, period_value] of entries) {
                if (!/^\d$/.exec(period_key))
                    continue;
                let {time = null, duration = null} = period_value;
                if (time != null && typeof time != "number")
                    time = null;
                if ( duration != null &&
                    (time == null || typeof duration != "number") )
                    duration = null;
                yield [period_key, Object.assign( {},
                    time != null ? {time: time} : null,
                    duration != null ? {duration: duration} : null,
                )];
            }
        })
    }
    SheetMetadata.set_object( this.sheet,
        this.constructor.metadata_keys.timetable,
        objgenmap(timetable, function*(entries) {
            for (let [date_key, date_value] of entries) {
                if (WorksheetDate.parse_date(date_key) == null)
                    continue;
                yield [ date_key, copy_date_value(date_value) ];
            }
        })
    );
} // }}}

// StudyGroup().get_worksheet_plan {{{
StudyGroup.prototype.get_worksheet_plan = function() {
    var worksheet_plan = SheetMetadata.get_object( this.sheet,
        this.constructor.metadata_keys.worksheet_plan );
    if (worksheet_plan == null)
        return null;
    return worksheet_plan;
} // }}}

// StudyGroup().get_today_worksheet_plan {{{
StudyGroup.prototype.get_today_worksheet_plan = function(
    today = WorksheetDate.today()
) {
    var worksheet_plan = SheetMetadata.get_object( this.sheet,
        this.constructor.metadata_keys.worksheet_plan );
    if (worksheet_plan == null)
        return null;
    return worksheet_plan[today.format()];
} // }}}

// StudyGroup().set_worksheet_plan (worksheet_plan) {{{
StudyGroup.prototype.set_worksheet_plan = function(worksheet_plan) {
    if (worksheet_plan == null) {
        SheetMetadata.unset( this.sheet,
            this.constructor.metadata_keys.worksheet_plan );
        return null;
    }
    function listgenmap(list, gen) {
      return Array.from(gen(list));
    }
    function objgenmap(obj, gen) {
      return Object.fromEntries(gen(Object.entries(obj)));
    }
    function copy_date_value(date_value) {
        return listgenmap(date_value || [], function*(items) {
            for (let worksheet_item of items) {
                let {period, category} = worksheet_item;
                if (period != null && !/^\d$/.exec(period))
                    period = null;
                if (category != null && typeof category != "string")
                    category = null;
                yield Object.assign( {},
                    period != null ? {period: period} : {},
                    category != null ? {category: category} : {}, );
            }
        })
    }
    SheetMetadata.set_object( this.sheet,
        this.constructor.metadata_keys.worksheet_plan,
        objgenmap(worksheet_plan, function*(entries) {
            for (let [date_key, date_value] of entries) {
                if (WorksheetDate.parse_date(date_key) == null)
                    continue;
                yield [date_key, copy_date_value(date_value)];
            }
        })
    );
} // }}}

return StudyGroup;
}(); // end StudyGroup namespace }}}1

var StudyGroupBuilder = function() { // namespace {{{1

// StudyGroupBuilder.build (spreadsheet, name, options) {{{
StudyGroupBuilder.build = function(spreadsheet, name, options) {
    try {
        var sheet = spreadsheet.insertSheet(name);
    } catch (error) {
        console.error(error);
        throw new StudyGroupInitError(
            "Unable to create a new sheet with the name “" + name + "” " +
            "(probably because it already exists)." );
    }
    return (new this(sheet, name, options)).group;
} // }}}

// XXX make this a class property of StudyGroup
const initial = { // {{{
    data_height: 20,
    attendance: {
        columns: 20,
        colors: {
            mark:    {h:  60, s: 0.55, l: 0.80},
            past:    {h:   0, s: 0.00, l: 0.90},
            present: {h:   0, s: 0.00, l: 0.95}
        }
    },
    separator_col_width: 13,
    rows: {
        data_row:     8,
        title_row:    6,
        label_row:    7,
        mirror_row:   1,
        category_row: 5,
        max_row:      3,
        weight_row:   2,
    },
}; // }}}

/* options: {{{
 *   rows (object)
 *     map of header row names to row numbers
 *     should include 'data_row' and all row names
 *   data_height (number)
 *     number of data rows, excluding the first (hidden) data row
 *     default is initial.data_height
 *   filename (string)
 *     filename prefix for uploads
 *     default will use group name (even if it will change)
 *   color_scheme (string)
 *     code, must match spreadsheet metadata
 *     default will use default color scheme
 *   timetable (object)
 *     saved in sheet metadata, no direct effect
 *   timetable["YYYY-MM-DD"] (object)
 *   timetable["YYYY-MM-DD"]["1"…"9"] (object)
 *     object that describes the period
 *   timetable["YYYY-MM-DD"]["1"…"9"].time
 *     minutes since 00:00
 *   worksheet_plan (object)
 *     saved in sheet metadata, no direct effect
 *   worksheet_plan["YYYY-MM-DD"] (list)
 *   worksheet_plan["YYYY-MM-DD"][*].period (string)
 *     integer from 1…9; optional
 *   worksheet_plan["YYYY-MM-DD"][*].category (string)
 *     category code; optional
 *   worksheet_plan.initial (list)
 *     XXX not implemented
 *     works same as dated worksheet lists, except is not saved in metadata
 *     and used immediately to add worksheets
 *   rating (boolean)
 *     whether to create a total rating
 *     default is yes
 *   sum (boolean)
 *     whether to create a total sum
 *     default is yes
 *   categories (array)
 *     default is no categories
 *   categories[*].code
 *   categories[*].rating (false or object)
 *     whether to create a rating for this category
 *     default is yes
 *   categories[*].rating.integrate (boolean)
 *     whether to include this category in bases for the total rating
 *     effective if total rating is created at all
 *     default is yes
 *   categories[*].sum (false or object)
 *     whether to create a sum for this category
 *     default is yes
 *   categories[*].sum.integrate (boolean)
 *     whether to include this category in bases for the total sum
 *     effective if total sum is created at all
 *     default is yes
 *   category_musthave (boolean)
 *     whether empty category (or category that was not listed) in worksheet
 *       will be highlighted as an error
 *     default is yes, unless no categories were specified
 *   attendance (object)
 *     parameters for creating attendance
 *     default is to not create attendance
 *   attendance.sum (boolean)
 *     whether to create attendance sum
 *     default is yes
 *   attendance.columns (number or object)
 *     number of columns in attendance
 *     default is initial.attendance.columns
 *   attendance.columns.date_list (array or object)
 *     date list (array of Date objects) or
 *     parameters for date list
 *   attendance.columns.date_list.start (Date)
 *     first date, included
 *   attendance.columns.date_list.end (Date)
 *     last date, not included
 *   attendance.columns.date_list.weekdays (array)
 *     list of seven boolean
 *     whether to include corresponding weekdays (starting with Monday)
 *     default is to include every day
 *   attendance.columns.date_lists (array)
 *     list of date lists
 *     each list is specified in the same way as above
 *   attendance.columns.date_lists[*] (array or object)
 *   attendance.columns.date_lists[*].start (Date)
 *   attendance.columns.date_lists[*].end (Date)
 *   attendance.columns.date_lists[*].weekdays (array)
 *   attendance.columns.date_lists[*].title (string)
 * }}} */

function StudyGroupBuilder(sheet, name, options) { // {{{
    this.sheet = sheet;
    this.options = this.rectify_options(options);
    this.dim = {};
    for (let row_name in initial.rows) {
        this.dim[row_name] = this.options.rows[row_name];
    }
    this.attendance_total_row = this.dim.weight_row;
    this.cfrule_objs = [];
    this.cfrule_error_objs = [];
    this.dim.data_height = this.options.data_height + 1;
    this.dim.sheet_height = this.dim.data_row - 1 + this.dim.data_height;
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
        .setVerticalAlignment("middle")
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    sheet.setHiddenGridlines(true);
    this.init_rows();
    var {student_count_cell} = this.init_columns();

    this.group = new StudyGroup(sheet, name);
    this.group.add_metadatum({skip_remove: true});
    if (student_count_cell != null) {
        this.group.set_student_count_cell(student_count_cell);
    }
    if (this.options.filename != null) {
        this.group.set_filename(this.options.filename);
    }
    if (this.options.color_scheme != null) {
        this.color_scheme = ColorSchemes.copy(this.options.color_scheme, ["name"]);
        this.group.set_color_scheme(this.options.color_scheme);
    } else {
        this.color_scheme = ColorSchemes.get_default();
    }
    if (this.options.timetable != null)
        this.group.set_timetable(this.options.timetable);
    if (this.options.worksheet_plan != null)
        this.group.set_worksheet_plan(this.options.worksheet_plan);
    Object.defineProperty(this.group, "dim", {value:
        StudyGroupDim.init(this.sheet, this.dim) });

    this.frozen_columns = sheet.getFrozenColumns();
    this.current_max_columns = sheet.getMaxColumns();
    this.next_column = this.frozen_columns + 1;
    this.separator_required = true;
    if (this.options.attendance && this.options.attendance.sum)
        this.attendance_sum_range = this.add_attendance_sum();
    if (this.options.rating)
        this.rating_range = this.add_rating_range();
    if (this.options.sum)
        this.sum_range = this.add_sum_range();
    this.category_info = this.prepare_category_info();
    if (this.category_info.rating.codes.length > 0)
        this.category_rating_range = this.add_category_rating_range();
    if (this.category_info.sum.codes.length > 0)
        this.category_sum_range = this.add_category_sum_range();
    this.separator_required = true;
    if (this.options.attendance)
        this.attendance_range = this.add_attendance();
    this.sheet.setColumnWidth(this.next_column, initial.separator_col_width);
    this.worksheet_start_column = this.next_column;
    if (this.current_max_columns > this.worksheet_start_column) {
        sheet.deleteColumns(
            this.worksheet_start_column + 1,
            this.current_max_columns - this.worksheet_start_column );
        this.current_max_columns = this.worksheet_start_column;
    }
    this.set_ratinglike_formulas();
    sheet.getRange(this.dim.label_row, 1, 1, this.current_max_columns)
        .setBorder(null, null, true, null, null, null);
    sheet.getRange(
        this.dim.data_row, 1,
        this.dim.data_height, this.current_max_columns
    )
        .setBorder(true, true, true, null, null, null)
        .setBorder( null, null, null, null, null, true,
            "black", SpreadsheetApp.BorderStyle.DOTTED );

    this.push_worksheet_cfrules();
    this.push_category_cfrules();
    this.push_rating_cfrules();
    var cfrules = new ConditionalFormatting.RuleList();
    this.cfrule_objs.unshift(...this.cfrule_error_objs);
    this.cfrule_objs.forEach(cfrule_obj => cfrules.push(
        ConditionalFormatting.Rule.from_object(cfrule_obj)
    ));
    cfrules.save(sheet);
} // }}}

// StudyGroupBuilder().rectify_options [options -> options] {{{
StudyGroupBuilder.prototype.rectify_options = function(options) {
    if (options == null)
        options = {};
    ({
        rows: options.rows = initial.rows,
        data_height: options.data_height = initial.data_height,
        filename: options.filename = null,
        color_scheme: options.color_scheme = null,
    } = options);
    for (let row_name in initial.rows) {
        ({
            [row_name] : options.rows[row_name] =
                initial.rows[row_name]
        } = options.rows);
        if (options.rows[row_name] == null && StudyGroupDim.required_rows.has(row_name)) {
            throw new StudyGroupInitError(
                "rectify_options: missing required row " + row_name );
        }
    }
    ({
        rating: options.rating = true,
        sum: options.sum = true,
        categories: options.categories = [],
        category_musthave: options.category_musthave =
            (options.categories.length > 0)
    } = options);
    if (options.rows.category_row == null) {
        options.categories = [];
    }
    for (let item of options.categories) {
        if (item.code == null)
            throw new StudyGroupInitError(
                "rectify_options: category must have a code" );
        if (item.rating == null)
            item.rating = {};
        if (item.rating && item.rating.integrate == null)
            item.rating.integrate = true;
        if (item.sum == null)
            item.sum = {};
        if (item.sum && item.sum.integrate == null)
            item.sum.integrate = true;
    }
    if (options.attendance == null)
        options.attendance = false;
    if (options.attendance) {
        var attendance = options.attendance;
        if (attendance.sum == null)
            attendance.sum = true;
        if (attendance.columns == null)
            attendance.columns = initial.attendance.columns;
    }
    return options;
} // }}}

// StudyGroupBuilder().init_rows {{{
StudyGroupBuilder.prototype.init_rows = function() {
    const sheet = this.sheet;
    const sheet_height = this.dim.sheet_height;
    const max_columns = sheet.getMaxColumns();
    var current_height = sheet.getMaxRows();
    if (current_height > sheet_height) {
        sheet.deleteRows(sheet_height + 1, current_height - sheet_height);
    } else if (current_height < sheet_height) {
        sheet.insertRowsAfter(current_height, sheet_height - current_height);
    }
    sheet.setFrozenRows(this.dim.data_row - 1);
    sheet.hideRows(this.dim.data_row);
    sheet.setRowHeights(1, this.dim.data_row - 1, 5);
    sheet.setRowHeights(this.dim.data_row, this.dim.data_height, 21);
    sheet.setRowHeight(this.dim.title_row, 24);
    sheet.getRange(this.dim.title_row, 1, 1, max_columns)
        .setNumberFormat('@STRING@')
        .setFontWeight('bold')
        .setFontFamily("Times New Roman,serif")
        .setFontSize(12);
    sheet.setRowHeight(this.dim.label_row, 21);
    sheet.getRange(this.dim.label_row, 1, 1, max_columns)
        .setNumberFormat('@STRING@')
        .setFontWeight('bold')
        .setFontFamily("Times New Roman,serif")
        .setVerticalAlignment("bottom");
    if (this.dim.mirror_row != null) {
        sheet.setRowHeight(this.dim.mirror_row, 21);
        sheet.hideRows(this.dim.mirror_row);
    }
    if (this.dim.category_row != null) {
        sheet.setRowHeight(this.dim.category_row, 17);
        sheet.getRange(this.dim.category_row, 1, 1, max_columns)
            .setNumberFormat('@STRING@')
            .setFontSize(8);
    }
    if (this.dim.max_row != null) {
        sheet.setRowHeight(this.dim.max_row, 17);
        sheet.getRange(this.dim.max_row, 1, 1, max_columns)
            .setFontSize(8);
    }
    if (this.dim.weight_row != null) {
        sheet.setRowHeight(this.dim.weight_row, 17);
        sheet.getRange(this.dim.weight_row, 1, 1, max_columns)
            .setFontSize(8);
    }
} // }}}

// StudyGroupBuilder().init_columns {{{
StudyGroupBuilder.prototype.init_columns = function() {
    const sheet = this.sheet;
    const frozen_columns = 4;
    const header_columns = frozen_columns - 1;
    const max_rows = sheet.getMaxRows();
    const max_columns = sheet.getMaxColumns()
    sheet.setFrozenColumns(frozen_columns);
    sheet.setColumnWidth(1, initial.separator_col_width);
    sheet.hideColumns(1);
    sheet.getRange(1, 1, this.dim.data_row - 1, 1)
        .setFontSize(8)
        .setFontFamily("Arial")
        .setFontWeight("normal")
        .setFontColor("#d0d0d0");
    sheet.setColumnWidths(2, header_columns, 75);
    sheet.getRange(1, 1, max_rows, frozen_columns)
        .setHorizontalAlignment("left");
    var group_title_range, student_count_cell = null;
    if (this.dim.mirror_row == 1 && this.dim.label_row >= 5) {
        group_title_range = sheet.getRange(2, 2, 2, header_columns);
        student_count_cell = sheet.getRange(this.dim.label_row - 1, 2);
    } else if (this.dim.mirror_row == null && this.dim.label_row >= 4) {
        group_title_range = sheet.getRange(1, 2, 2, header_columns);
        student_count_cell = sheet.getRange(this.dim.label_row - 1, 2);
    } else {
        group_title_range = sheet.getRange(
            this.dim.title_row, 2, 1, header_columns );
        if (this.dim.max_row != null)
            student_count_cell = sheet.getRange(this.dim.max_row, 2);
    }
    group_title_range.getCell(1, 1)
        .setValue("Название группы");
    group_title_range.merge()
        .setNumberFormat('@STRING@')
        .setFontSize(14)
        .setFontWeight("bold")
        .setFontFamily("Verdana")
        .setVerticalAlignment("top");
    if (student_count_cell != null) {
        const data_column_R1C1 = "R" + this.dim.data_row + "C[0]:C[0]";
        student_count_cell
            .setFormulaR1C1(
                '= counta(' + data_column_R1C1 + ') - ' +
                'sum(arrayformula(' +
                    'iferror(N(find("(→)",' + data_column_R1C1 + ')>0))' +
                '))' )
            .setNumberFormat("0 чел\\.")
            .setFontSize(8)
            .setVerticalAlignment("bottom");
        if (frozen_columns >= 3) {
            student_count_cell.offset(0, 1)
                .setValue("каб. ?")
                .setNumberFormat('@STRING@')
                .setFontSize(8)
                .setVerticalAlignment("bottom");
        }
    }
    sheet.getRange(this.dim.label_row, 2)
        .setValue("Фамилия Имя");
    sheet.getRange(this.dim.label_row, 2, 1, header_columns).merge();
    sheet.getRange(this.dim.data_row, 2, this.dim.data_height, header_columns)
        .mergeAcross();
    var nospace_cfrule_obj = { type: "boolean",
        condition: {
            type: SpreadsheetApp.BooleanCriteria.TEXT_STARTS_WITH,
            values: [" "] },
        ranges: [[
            this.dim.data_row, 2, this.dim.data_height, header_columns ]],
        effect: {background: "#ff0000"}
    };
    this.cfrule_error_objs.push(
        nospace_cfrule_obj,
        Object.assign({}, nospace_cfrule_obj, {condition: {
            type: SpreadsheetApp.BooleanCriteria.TEXT_CONTAINS,
            values: ["  "],
        }}),
        Object.assign({}, nospace_cfrule_obj, {condition: {
            type: SpreadsheetApp.BooleanCriteria.TEXT_ENDS_WITH,
            values: [" "],
        }})
    );
    sheet.getRange(
        1, frozen_columns + 1, max_rows, max_columns - frozen_columns
    )
        .setHorizontalAlignment("center");
    sheet.getRange(this.dim.label_row, 2, 1, header_columns)
        .setBorder(true, true, true, true, null, null);
    sheet.getRange(this.dim.data_row, 2, this.dim.data_height, header_columns)
        .setBorder(true, true, true, true, null, null);
    sheet.setColumnWidths(
        frozen_columns + 1, max_columns - frozen_columns, 21 );
    return {student_count_cell: student_count_cell};
} // }}}

// StudyGroupBuilder().allocate_columns (num_columns) => (start_column) {{{
StudyGroupBuilder.prototype.allocate_columns = function(num_columns) {
    if (typeof num_columns != "number" || isNaN(num_columns)) {
        throw new StudyGroupInitError(
            "allocate_columns: invalid argument num_columns" );
    }
    var next_next_column = this.next_column + num_columns +
        this.separator_required;
    if (this.current_max_columns < next_next_column) {
        this.sheet.insertColumnsAfter( this.current_max_columns,
            next_next_column - this.current_max_columns );
        this.current_max_columns = next_next_column;
    }
    if (this.separator_required) {
        this.separator_required = false;
        this.sheet.setColumnWidth( this.next_column,
            initial.separator_col_width );
        ++this.next_column;
    }
    var start_column = this.next_column;
    this.next_column = next_next_column;
    return start_column;
} // }}}

// StudyGroupBuilder().add_attendance_sum () => (range) {{{
StudyGroupBuilder.prototype.add_attendance_sum = function() {
    const sheet = this.sheet;
    var column = this.allocate_columns(1);
    sheet.setColumnWidth(column, 40);
    var range = sheet.getRange(
        this.dim.data_row, column, this.dim.data_height, 1 );
    var max_range = this.dim.max_row == null ? null :
        range.offset(this.dim.max_row - this.dim.data_row, 0, 1);
    range
        .setFontFamily("Georgia");
    range.offset(this.dim.label_row - this.dim.data_row, 0, 1)
        .setValue("At")
        .setNote("посещаемость");
    range.offset(
            this.dim.label_row - this.dim.data_row, 0,
            this.dim.data_height + this.dim.data_row - this.dim.label_row )
        .setBorder(true, true, true, true, null, null);
    if (max_range != null)
        max_range.setFontFamily("Georgia");
    var max_R1C1;
    if (this.dim.max_row != null) {
        max_R1C1 = "R" + this.dim.max_row + "C[0]";
    } else {
        max_R1C1 = "R" + this.dim.data_row + "C[0]:C[0]";
    }
    var cf_ranges = [[this.dim.data_row, column, this.dim.data_height, 1]];
    if (this.dim.max_row != null) {
        cf_ranges.push([this.dim.max_row, column, 1, 1]);
    }
    this.cfrule_objs.push({ type: "gradient",
        condition: {
            min_type: SpreadsheetApp.InterpolationType.NUMBER,
            min_value: "0",
            mid_type: SpreadsheetApp.InterpolationType.NUMBER,
            mid_value: "= min(0.5 * max(7," + max_R1C1 + "), 4)",
            max_type: SpreadsheetApp.InterpolationType.NUMBER,
            max_value: "= 0.9 * max(7," + max_R1C1 + ")",
        },
        ranges: cf_ranges,
        effect: {
            min_color: "#ffffff",
            mid_color: HSL.to_hex(initial.attendance.colors.mark),
            max_color: HSL.to_hex(
                HSL.deepen(initial.attendance.colors.mark, 4) ),
        },
    });
    return range;
} // }}}

// StudyGroupBuilder().add_ratinglike_range (label, note) => (range) {{{
StudyGroupBuilder.prototype.add_ratinglike_range = function(label, note) {
    const sheet = this.sheet;
    var column = this.allocate_columns(1);
    this.sheet.setColumnWidth(column, 40);
    var range = sheet.getRange(
        this.dim.data_row, column, this.dim.data_height, 1);
    range.offset(this.dim.label_row - this.dim.data_row, 0, 1)
        .setValue(label)
        .setNote(note);
    range.offset(
            this.dim.label_row - this.dim.data_row, 0,
            this.dim.data_height + this.dim.data_row - this.dim.label_row )
        .setBorder(true, true, true, true, null, null);
    return range;
    //add_ratinglike_range.call(this, rating_range);
} // }}}

// StudyGroupBuilder().add_rating_range () => (range) {{{
StudyGroupBuilder.prototype.add_rating_range = function() {
    return this.add_ratinglike_range("ΣΣ", "рейтинг");
} // }}}

// StudyGroupBuilder().add_sum_range () => (range) {{{
StudyGroupBuilder.prototype.add_sum_range = function() {
    return this.add_ratinglike_range("SS", "кол-во задач");
} // }}}

// StudyGroupBuilder().prepare_category_info () => (object) {{{
StudyGroupBuilder.prototype.prepare_category_info = function() {
    function prepare_rating_info(rating_type) {
        var codes = [], weights = [], integrate = false;
        this.options.categories.forEach(function(category) {
            if (!category[rating_type])
                return;
            codes.push(category.code);
            if (category[rating_type].integrate) {
                weights.push(1);
                integrate = true;
            } else {
                weights.push(0);
            }
        });
        if (!this.options[rating_type])
            integrate = false;
        return { type: rating_type,
            codes: codes, weights: weights, integrate: integrate
        };
    }
    return {
        rating: prepare_rating_info.call(this, "rating"),
        sum:    prepare_rating_info.call(this, "sum")
    }
} // }}}

// StudyGroupBuilder().add_category_ratinglike_range (info, label) => (range) {{{
StudyGroupBuilder.prototype.add_category_ratinglike_range = function(info, label) {
    if (this.dim.category_row == null)
        throw new StudyGroupInitError(
            "add_category_ratinglike_range: internal error" );
    const sheet = this.sheet;
    var width = info.codes.length;
    var column = this.allocate_columns(width);
    sheet.setColumnWidths(column, width, 40);
    var range = sheet.getRange(
        this.dim.data_row, column, this.dim.data_height, width );
    range.offset(this.dim.label_row - this.dim.data_row, 0, 1)
        .setValue(label);
    if (info.integrate && this.dim.weight_row != null) {
        range.offset(this.dim.weight_row - this.dim.data_row, 0, 1)
            .setValues([info.weights])
            .setFontSize(8);
    }
    range.offset(this.dim.category_row - this.dim.data_row, 0, 1)
        .setValues([info.codes]);
    range.offset(
            this.dim.label_row - this.dim.data_row, 0,
            this.dim.data_height + this.dim.data_row - this.dim.label_row )
        .setBorder(true, true, true, true, null, null)
        .setBorder( null, null, null, null, true, null,
            "black", SpreadsheetApp.BorderStyle.DOTTED );
    return range;
} // }}}

// StudyGroupBuilder().add_category_rating_range () => (range) {{{
StudyGroupBuilder.prototype.add_category_rating_range = function() {
    return this.add_category_ratinglike_range(this.category_info.rating, "Σ");
} // }}}

// StudyGroupBuilder().add_category_sum_range () => (range) {{{
StudyGroupBuilder.prototype.add_category_sum_range = function() {
    return this.add_category_ratinglike_range(this.category_info.sum, "S");
} // }}}

// StudyGroupBuilder().generate_date_list (start, end, weekdays) => (array) {{{
StudyGroupBuilder.prototype.generate_date_list = function(start, end, weekdays) {
    var dates = [];
    for (
        var i = new Date(start.getTime());
        i < end;
        i.setDate(i.getDate() + 1)
    ) {
        if (!weekdays || weekdays[(i.getDay()+6)%7])
            dates.push(new Date(i.getTime()));
    }
    return dates;
} // }}}

// StudyGroupBuilder().add_attendance () => (range) {{{
StudyGroupBuilder.prototype.add_attendance = function() {
    const total_row = this.attendance_total_row;
    var dim = {};
    var date_list, date_lists;
    const columns_option = this.options.attendance.columns;
    if (typeof columns_option == "number") {
        dim.data_width = columns_option;
    } else if (columns_option.date_list) {
        date_list = columns_option.date_list
        if (date_list.start != null && date_list.end != null) {
            date_list = this.generate_date_list(
                date_list.start,
                date_list.end,
                date_list.weekdays );
        }
        dim.data_width = date_list.length;
    } else if (columns_option.date_lists) {
        date_lists = columns_option.date_lists.map((date_list) => {
            if (date_list.start != null && date_list.end != null) {
                var title = date_list.title;
                date_list = this.generate_date_list(
                    date_list.start,
                    date_list.end,
                    date_list.weekdays );
                date_list.title = title;
            }
            return date_list;
        });
        dim.data_width = date_lists
            .reduce(function(sum, list) {
                return sum + list.length + 1;
            }, 0);
    } else {
        throw new StudyGroupInitError(
            "add_attendance: invalid option attendance.columns" );
    }
    dim.width = dim.data_width + 2;
    dim.start = this.allocate_columns(dim.width);
    dim.data_start = dim.start + 1;
    dim.end = dim.start + dim.width - 1;
    dim.data_end = dim.data_start + dim.data_width - 1;
    var title_range = this.sheet.getRange(
        this.dim.title_row, dim.start, 1, dim.width );
    var data_ext_range = this.sheet.getRange(
        this.dim.data_row, dim.start,
        this.dim.data_height, dim.width );
    this.sheet.setColumnWidth(dim.start, 5);
    this.sheet.hideColumns(dim.start);
    var data_range = this.sheet.getRange(
        this.dim.data_row, dim.data_start,
        this.dim.data_height, dim.data_width );
    this.sheet.setColumnWidth(dim.end, 5);
    this.sheet.hideColumns(dim.end);
    this.sheet.setColumnWidths(dim.data_start, dim.data_width, 21);
    title_range.offset(
        0, 0,
        this.dim.data_height + this.dim.data_row - this.dim.title_row
    )
        .setBorder(true, true, true, true, null, null);
    this.sheet.getRange(this.dim.label_row, dim.end)
        .setValue(".") // force column non-empty
    if (date_list != null) {
        this.format_attendance_data(data_range);
        data_range.offset(this.dim.label_row - this.dim.data_row, 0, 1)
            .setNumberFormat("dd")
            .setValues([date_list]);
        title_range
            .setNumberFormat('yyyy"-"mm');
        this.monthize_attendance(
            data_range.offset(this.dim.title_row - this.dim.data_row, 0, 1),
            date_list );
    } else if (date_lists != null) {
        let substart = data_range.getColumn();
        for (let date_list of date_lists) {
            let metatitle_range = this.sheet.getRange(
                this.dim.title_row, substart, 1, 1 );
            this.sheet.setColumnWidth(substart, initial.separator_col_width);
            ++substart;
            metatitle_range
                .setNumberFormat("@STRING@")
                .setValue(date_list.title)
                .setBorder(true, true, true, true, null, null)
                .setBackground(HSL.to_hex(
                    HSL.deepen(initial.attendance.colors.mark, 0.75) ));
            let subrange = ( (row, height) =>
                this.sheet.getRange(row, substart, height, date_list.length) );
            this.format_attendance_data(subrange(
                this.dim.data_row, this.dim.data_height ));
            subrange(this.dim.label_row, 1)
                .setNumberFormat("dd")
                .setValues([date_list])
                .shiftColumnGroupDepth(1);
            let title_subrange = subrange(this.dim.title_row, 1);
            title_subrange
                .setNumberFormat('yyyy"-"mm');
            this.monthize_attendance(title_subrange, date_list);
            substart += date_list.length;
        };
    } else {
        this.format_attendance_data(data_range);
        title_range.getCell(1, 1).setValue("Посещаемость");
        title_range.merge();
    }
    var label_R1C1 = "R" + this.dim.label_row + "C[0]";
    var noshadow_formula = "or(isdate(R[0]C[0]),R[0]C[0]=0)";
    var today_cfranges = [
        [this.dim.data_row,   dim.start, this.dim.data_height, dim.width],
        [this.dim.label_row,  dim.start, 1, dim.width],
        ...[
            [this.dim.max_row,  dim.start, 1, dim.width],
            [this.dim.weight_row,  dim.start, 1, dim.width],
        ].filter(([r,]) => r != null)
    ]
    this.cfrule_objs.push({ type: "boolean",
        condition: {
            type: SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA,
            values: ["".concat(
                "=and(",
                    noshadow_formula, ",",
                    "isdate(", label_R1C1, "),",
                    label_R1C1, "<today()",
                ")" )],
        },
        ranges: today_cfranges,
        effect: {
            background: HSL.to_hex(initial.attendance.colors.past) },
    });
    this.cfrule_objs.push({ type: "boolean",
        condition: {
            type: SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA,
            values: ["".concat(
                "=and(",
                    noshadow_formula, ",",
                    "isdate(", label_R1C1, "),",
                    label_R1C1, "<=today()",
                ")" )],
        },
        ranges: today_cfranges,
        effect: {
            background: HSL.to_hex(initial.attendance.colors.present) },
    });
    var title_R1C1 = "R" + this.dim.title_row + "C[0]";
    this.cfrule_objs.push({ type: "boolean",
        condition: {
            type: SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA,
            values: ["".concat(
                "=and(",
                    "isdate(", title_R1C1, "),",
                    "eomonth(", title_R1C1, ",0)<today()",
                ")" )],
        },
        ranges: [[this.dim.title_row, dim.start, 1, dim.width]],
        effect: {
            background: HSL.to_hex(initial.attendance.colors.past) },
    });
    this.cfrule_objs.push({ type: "boolean",
        condition: {
            type: SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA,
            values: ["".concat(
                "=and(",
                    "isdate(", title_R1C1, "),",
                    "date(year(", title_R1C1, "),month(", title_R1C1, "),1)",
                    "<=today()",
                ")" )],
        },
        ranges: [[this.dim.title_row, dim.start, 1, dim.width]],
        effect: {
            background: HSL.to_hex(initial.attendance.colors.present) },
    });
    var ones_cfranges = [
        [this.dim.data_row, dim.start, this.dim.data_height, dim.width],
        ...[
            [this.dim.max_row, dim.start, 1, dim.width],
        ].filter(([r,]) => r != null)
    ];
    this.cfrule_objs.push({ type: "boolean",
        condition: {
            type: SpreadsheetApp.BooleanCriteria.NUMBER_GREATER_THAN,
            values: [0] },
        ranges: ones_cfranges,
        effect: {
            background: HSL.to_hex(initial.attendance.colors.mark) },
    });
    if (total_row != null) {
        this.cfrule_objs.push({ type: "gradient",
            condition: {
                min_type: SpreadsheetApp.InterpolationType.NUMBER,
                min_value: "0.5",
                mid_type: SpreadsheetApp.InterpolationType.NUMBER,
                mid_value: "10",
                max_type: SpreadsheetApp.InterpolationType.NUMBER,
                max_value: "25" }, // XXX make it depend on student count? or maybe just maximum
            ranges: [[total_row, dim.start, 1, dim.width]],
            effect: {
                min_color: "#ffffff",
                mid_color: HSL.to_hex(initial.attendance.colors.mark),
                max_color: HSL.to_hex(
                    HSL.deepen(initial.attendance.colors.mark, 4) ),
            },
        });
    }
    return data_ext_range;
} // }}}

// StudyGroupBuilder().format_attendance_data (range) {{{
StudyGroupBuilder.prototype.format_attendance_data = function(data_range) {
    const total_row = this.attendance_total_row;
    data_range.setFontFamily("Georgia");
    const extra_rows = [total_row, this.dim.max_row].filter(r => r != null);
    for (let row of extra_rows) {
        data_range.offset(row - this.dim.data_row, 0, 1)
            .setFontFamily("Georgia");
    }
    const data_column_R1C1 = "R" + this.dim.data_row + "C[0]:C[0]";
    var total_formula = 'sum(' + data_column_R1C1 + ')';
    if (total_row != null) {
        data_range.offset(total_row - this.dim.data_row, 0, 1)
            .setFormulaR1C1('=' + total_formula);
        total_formula = 'R' + total_row + 'C[0]';
    }
    if (this.dim.max_row != null) {
        data_range.offset(this.dim.max_row - this.dim.data_row, 0, 1)
            .setFormulaR1C1('=N(' + total_formula + '>0)');
    }
    for (let row of [null, this.dim.label_row, ...extra_rows]) {
        ( row == null ? data_range :
                data_range.offset(row - this.dim.data_row, 0, 1) )
            .setBorder(true, true, true, true, null, null)
            .setBorder( null, null, null, null, true, null,
                "black", SpreadsheetApp.BorderStyle.DOTTED );
    };
    if (total_row != null && this.dim.max_row != null && total_row - this.dim.max_row == -1)
        data_range.offset(total_row - this.dim.data_row, 0, 2)
            .setBorder( null, null, null, null, null, true,
                "black", SpreadsheetApp.BorderStyle.DOTTED );
} // }}}

// StudyGroupBuilder().monthize_attendance (range, dates) {{{
StudyGroupBuilder.prototype.monthize_attendance = function(title_range, dates) {
    const first_column = title_range.getColumn();
    const extra_rows = [this.attendance_total_row, this.dim.max_row].filter(r => r != null);
    for (let {start, end, length, value} of
        group_by_(dates, function(date) {
            var month = new Date(date.getTime());
            month.setDate(1);
            return month.getTime();
        }) )
    {
        var month_title_range = this.sheet.getRange(
            this.dim.title_row, first_column + start,
            1, length );
        month_title_range.getCell(1, 1)
            .setValue(new Date(value));
        month_title_range.merge();
        month_title_range.offset(
            0, 0,
            this.dim.data_height + this.dim.data_row - this.dim.title_row
        )
            .setBorder(true, true, true, true, null, null);
        for (let row of extra_rows) {
            month_title_range.offset(
                row - this.dim.title_row, 0, 1
            )
                .setBorder(null, true, null, true, null, null);
        }
    };
} // }}}

// StudyGroupBuilder().get_worksheet_row_R1C1 (row_R1) => (R1C1) {{{
StudyGroupBuilder.prototype.get_worksheet_row_R1C1 = function(row_R1) {
    if (typeof row_R1 == "number")
        row_R1 = 'R' + row_R1;
    return row_R1 + 'C' + this.worksheet_start_column + ':' + row_R1;
} // }}}

// StudyGroupBuilder().set_ratinglike_formulas {{{
StudyGroupBuilder.prototype.set_ratinglike_formulas = function() {
    if (this.options.attendance && this.options.attendance.sum)
        this.set_attendance_sum_formulas();
    if (this.options.rating) {
        if (this.category_info.rating.integrate)
            this.set_rating_category_formulas();
        else
            this.set_rating_direct_formulas();
    }
    if (this.options.sum) {
        if (this.category_info.sum.integrate)
            this.set_sum_category_formulas();
        else
            this.set_sum_direct_formulas();
    }
    if (this.category_info.rating.codes.length > 0)
        this.set_category_rating_formulas();
    if (this.category_info.sum.codes.length > 0)
        this.set_category_sum_formulas();
} // }}}

// StudyGroupBuilder().set_attendance_sum_formulas {{{
StudyGroupBuilder.prototype.set_attendance_sum_formulas = function() {
    var attendance_R1C1 =
        'R[0]C' + this.attendance_range.getColumn() + ':' +
        'R[0]C' + this.attendance_range.getLastColumn();
    var attendance_sum_formula_R1C1 = '=sum(' + attendance_R1C1 + ')';
    this.attendance_sum_range
        .setFormulaR1C1(attendance_sum_formula_R1C1);
    if (this.dim.max_row != null) {
        this.attendance_sum_range.offset(
            this.dim.max_row - this.dim.data_row, 0, 1
        )
            .setFormulaR1C1(attendance_sum_formula_R1C1);
    }
} // }}}

// StudyGroupBuilder().set_rating_direct_formulas {{{
StudyGroupBuilder.prototype.set_rating_direct_formulas = function() {
    // XXX create intermediate column
    // so that rating will be in percent
    const label_row = this.dim.mirror_row != null ? this.dim.mirror_row : this.dim.label_row;
    var rating_formula_R1C1 = ''.concat('=sum(iferror(filter( ',
        'arrayformula(',
            this.get_worksheet_row_R1C1('R[0]'),
            this.dim.weight_row == null ? '' : ('*' + this.get_worksheet_row_R1C1(this.dim.weight_row)),
            this.dim.max_row    == null ? '' : ('/' + this.get_worksheet_row_R1C1(this.dim.max_row)),
        '), ',
        this.get_worksheet_row_R1C1(label_row), '="Σ"',
    ' )))');
    var number_format = (this.dim.max_row != null || this.dim.weight_row != null) ?
        "0.00;−0.00" : "0.#;−0.#"
    this.rating_range
        .setFormulaR1C1(rating_formula_R1C1)
        .setNumberFormat(number_format);
    if (this.dim.max_row != null) {
        this.rating_range.offset(this.dim.max_row - this.dim.data_row, 0, 1)
            .setFormulaR1C1(rating_formula_R1C1)
            .setNumberFormat(number_format);
    }
} // }}}

// StudyGroupBuilder().set_rating_category_formulas {{{
StudyGroupBuilder.prototype.set_rating_category_formulas = function() {
    const category_rating_range = this.category_rating_range;
    function get_category_rating_row(row_R1) {
        if (typeof row_R1 == "number")
            row_R1 = 'R' + row_R1;
        return row_R1 + 'C' + category_rating_range.getColumn() +
            ':' + row_R1 + 'C' + category_rating_range.getLastColumn();
    }
    var rating_formula_R1C1, number_format;
    if (this.dim.max_row != null) {
        var rating_arrayformula_R1C1 = ''.concat(
            'arrayformula(iferror(',
                get_category_rating_row('R[0]'), "/",
                get_category_rating_row(this.dim.max_row),
            ',0))' );
        if (this.dim.weight_row != null) {
            rating_formula_R1C1 = ''.concat(
                '=average.weighted(',
                    rating_arrayformula_R1C1, ',',
                    get_category_rating_row(this.dim.weight_row),
                ')' );
        } else {
            rating_formula_R1C1 = ''.concat(
                '=average(',
                    rating_arrayformula_R1C1,
                ')' );
        }
        number_format = "00%;−00%";
    } else {
        if (this.dim.weight_row != null) {
            rating_formula_R1C1 = ''.concat(
                '=sumproduct(',
                    get_category_rating_row('R[0]'), ',',
                    get_category_rating_row(this.dim.weight_row),
                ')' );
            number_format = "0.00;−0.00";
        } else {
            rating_formula_R1C1 = ''.concat(
                '=sum(',
                    get_category_rating_row('R[0]'),
                ')' );
            number_format = "0.#;−0.#";
        }
    }
    this.rating_range
        .setFormulaR1C1(rating_formula_R1C1)
        .setNumberFormat(number_format);
    if (this.dim.max_row != null) {
        this.rating_range.offset(this.dim.max_row - this.dim.data_row, 0, 1)
            .setFormulaR1C1(rating_formula_R1C1)
            .setNumberFormat(number_format);
    }
} // }}}

// StudyGroupBuilder().set_sum_direct_formulas {{{
StudyGroupBuilder.prototype.set_sum_direct_formulas = function() {
    const label_row = this.dim.mirror_row != null ? this.dim.mirror_row : this.dim.label_row;
    var sum_formula_R1C1 = ''.concat('=sum(iferror(filter( ',
        this.get_worksheet_row_R1C1('R[0]'), ', ',
        this.get_worksheet_row_R1C1(label_row), '="S"',
    ' )))');
    this.sum_range
        .setFormulaR1C1(sum_formula_R1C1)
        .setNumberFormat("0");
    if (this.dim.max_row != null) {
        this.sum_range.offset(this.dim.max_row - this.dim.data_row, 0, 1)
            .setFormulaR1C1(sum_formula_R1C1)
            .setNumberFormat("0");
    }
} // }}}

// StudyGroupBuilder().set_sum_category_formulas {{{
StudyGroupBuilder.prototype.set_sum_category_formulas = function() {
    const category_sum_range = this.category_sum_range;
    function get_category_sum_row(row_R1) {
        if (typeof row_R1 == "number")
            row_R1 = 'R' + row_R1;
        return row_R1 + 'C' + category_sum_range.getColumn() +
            ':' + row_R1 + 'C' + category_sum_range.getLastColumn();
    }
    var sum_formula_R1C1;
    if (this.dim.weight_row != null) {
        sum_formula_R1C1 = ''.concat(
            '=sumproduct(',
                get_category_sum_row('R[0]'), ',',
                get_category_sum_row(this.dim.weight_row),
            ')' );
    } else {
        sum_formula_R1C1 = ''.concat(
            '=sum(',
                get_category_sum_row('R[0]'),
            ')' );
    }
    this.sum_range
        .setFormulaR1C1(sum_formula_R1C1)
        .setNumberFormat("0");
    if (this.dim.max_row != null) {
        this.sum_range.offset(this.dim.max_row - this.dim.data_row, 0, 1)
            .setFormulaR1C1(sum_formula_R1C1)
            .setNumberFormat("0");
    }
} // }}}

// StudyGroupBuilder().set_category_rating_formulas {{{
StudyGroupBuilder.prototype.set_category_rating_formulas = function() {
    if (this.dim.category_row == null)
        throw new StudyGroupInitError(
            "add_category_ratinglike_range: internal error" );
    const label_row = this.dim.mirror_row != null ? this.dim.mirror_row : this.dim.label_row;
    var category_rating_formula_R1C1 = ''.concat('=sum(iferror(filter( ',
        'arrayformula(',
            this.get_worksheet_row_R1C1('R[0]'),
            this.dim.weight_row == null ? '' : ('*' + this.get_worksheet_row_R1C1(this.dim.weight_row)),
            this.dim.max_row    == null ? '' : ('/' + this.get_worksheet_row_R1C1(this.dim.max_row)),
        '), ',
        this.get_worksheet_row_R1C1(label_row), '="Σ", ',
        'exact(',
            this.get_worksheet_row_R1C1(this.dim.category_row), ',',
            'R', this.dim.category_row, 'C[0])',
    ' )))');
    var number_format = (this.dim.max_row != null || this.dim.weight_row != null) ?
        "0.00;−0.00" : "0.#;−0.#"
    this.category_rating_range
        .setFormulaR1C1(category_rating_formula_R1C1)
        .setNumberFormat(number_format);
    if (this.dim.max_row != null) {
        this.category_rating_range.offset(
            this.dim.max_row - this.dim.data_row, 0, 1
        )
            .setFormulaR1C1(category_rating_formula_R1C1)
            .setNumberFormat(number_format);
    }
} // }}}

// StudyGroupBuilder().set_category_sum_formulas {{{
StudyGroupBuilder.prototype.set_category_sum_formulas = function() {
    if (this.dim.category_row == null)
        throw new StudyGroupInitError(
            "add_category_ratinglike_range: internal error" );
    const label_row = this.dim.mirror_row != null ? this.dim.mirror_row : this.dim.label_row;
    var category_sum_formula_R1C1 = ''.concat('=sum(iferror(filter( ',
        this.get_worksheet_row_R1C1('R[0]'), ', ',
        this.get_worksheet_row_R1C1(label_row), '="S", ',
        'exact(',
            this.get_worksheet_row_R1C1(this.dim.category_row), ',',
            'R', this.dim.category_row, 'C[0])',
    ' )))');
    this.category_sum_range
        .setFormulaR1C1(category_sum_formula_R1C1)
        .setNumberFormat("0");
    if (this.dim.max_row != null) {
        this.category_sum_range.offset(this.dim.max_row - this.dim.data_row, 0, 1)
            .setFormulaR1C1(category_sum_formula_R1C1)
            .setNumberFormat("0");
    }
} // }}}

// StudyGroupBuilder().push_worksheet_cfrules {{{
StudyGroupBuilder.prototype.push_worksheet_cfrules = function() {
    const sheet = this.sheet;
    if (this.dim.mirror_row != null) {
        var mirror_R1C1 = "R" + this.dim.mirror_row + "C[0]";
        this.cfrule_error_objs.push({ type: "boolean",
            condition: {
                type: SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA,
                values: ["".concat(
                    "=iserror(", mirror_R1C1, ")" )],
            },
            ranges: [[1, this.worksheet_start_column, this.dim.sheet_height, 1]],
            effect: {background: "#ff0000"},
        });
        this.cfrule_error_objs.push({ type: "boolean",
            condition: {
                type: SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA,
                values: ["".concat(
                    "=isblank(", mirror_R1C1, ")" )],
            },
            ranges: [
                [ this.dim.data_row, this.worksheet_start_column,
                    this.dim.data_height ],
                [this.dim.label_row,  this.worksheet_start_column],
                ...[
                    [this.dim.weight_row, this.worksheet_start_column],
                    [this.dim.max_row,    this.worksheet_start_column],
               ].filter(([r,]) => r != null)
            ],
            effect: {font_color: "#808080"},
        });
    }
} // }}}

// StudyGroupBuilder().push_category_cfrules {{{
StudyGroupBuilder.prototype.push_category_cfrules = function() {
    if (this.dim.category_row == null)
        return;
    var ranges = [this.category_rating_range, this.category_sum_range];
    var cfranges = [];
    for (let range of ranges) {
        if (range == null)
            continue;
        cfranges.push(
            [this.dim.label_row,    range.getColumn(), 1, range.getWidth()],
            [this.dim.category_row, range.getColumn(), 1, range.getWidth()],
        );
    }
    var worksheet_category_row =
        [this.dim.category_row, this.worksheet_start_column];
    cfranges.push(
        worksheet_category_row,
        [this.dim.title_row,    this.worksheet_start_column],
    );
    var categories = Categories.get(this.sheet.getParent());
    var category_R1C1 = "R" + this.dim.category_row + "C[0]";
    for (let category_option of this.options.categories) {
        let code = category_option.code;
        let category = categories[code];
        if (category == null || category.color == null)
            continue;
        this.cfrule_objs.push({ type: "boolean",
            condition: {
                type: SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA,
                values: ["".concat(
                    "=exact(",
                        "\"", code.replace('"', '""'), "\",",
                        category_R1C1,
                    ")" )],
            },
            ranges: cfranges,
            effect: {
                background: HSL.to_hex(category.color) },
        });
    }
    this.cfrule_objs.push({ type: "boolean",
        condition: {
            type: SpreadsheetApp.BooleanCriteria.CELL_NOT_EMPTY,
            values: [], },
        ranges: [worksheet_category_row],
        effect: {background: "#ff0000"},
    });
    if (this.options.category_musthave) {
        const label_row = this.dim.mirror_row != null ? this.dim.mirror_row : this.dim.label_row;
        var label_R1C1 = "R" + label_row + "C[0]";
        this.cfrule_objs.push({ type: "boolean",
            condition: {
                type: SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA,
                values: ["".concat(
                    "=or(",
                        "exact(\"Σ\"," + label_R1C1 + ")", ",",
                        "exact(\"S\"," + label_R1C1 + ")",
                    ")" )],
            },
            ranges: [worksheet_category_row],
            effect: {background: "#ff0000"},
        });
    }
} // }}}

// StudyGroupBuilder().push_rating_cfrules {{{
StudyGroupBuilder.prototype.push_rating_cfrules = function() {
    var ranges = [
        this.rating_range, this.sum_range,
        this.category_rating_range, this.category_sum_range ];
    var cfranges = [];
    for (let range of ranges) {
        if (range == null)
            continue;
        cfranges.push([
            this.dim.data_row, range.getColumn(),
            this.dim.data_height, range.getWidth() ]);
        if (this.dim.max_row) {
            cfranges.push([
                this.dim.max_row,  range.getColumn(),
                1, range.getWidth() ]);
        }
    }
    if (cfranges.length < 1)
      return;
    this.cfrule_objs.push(this.group.new_cfrule_rating(
        cfranges, this.color_scheme,
    ));
} // }}}

return StudyGroupBuilder;
}(); // end StudyGroupBuilder namespace }}}1

// vim: set fdm=marker sw=4 :
