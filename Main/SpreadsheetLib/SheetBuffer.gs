/* Sheet value getter-setter
 * Optimized for access to values, notes and formulas of the same range.
 * Optimized for random access (on first get call).
 */

class SheetBufferError extends SpreadsheetError {};
class SheetBufferMergeOverlap extends SheetBufferError {};

var SheetBuffer = function() { // begin namespace

function MergedRangesBuffer(sheet, row_map, dim) {
    this.sheet = sheet;
    this._row_map = row_map;
    this.dim = dim;
    this._loaded = false;
    this._indices = {};
    for (let row_name in row_map) {
        this._indices[row_name] = [];
    }
}

MergedRangesBuffer.prototype.ensure_loaded = function() {
    if (this._loaded)
        return;
    for (let row_name in this._row_map) {
        let indices = this._indices[row_name];
        indices.length = this.dim.sheet_width;
        indices.fill(0);
    }
    var merged_ranges = this.sheet.getRange(
        1, 1, this.dim.frozen_height, this.dim.sheet_width
    ).getMergedRanges();
    var rev_row_map = Object.fromEntries( Object.entries(this._row_map)
      .map(([k,v]) => [v,k])
    );
    for (let range of merged_ranges) {
        if (range.getNumRows() != 1)
            continue;
        let row_name = rev_row_map[range.getRow()];
        if (row_name == null)
            continue;
        let indices = this._indices[row_name];
        let x = range.getColumn(), y = range.getLastColumn();
        let width = y - x + 1;
        if (width < 2)
            throw new Error(
                "MergedRangesBuffer.ensure_loaded: internal error" );
        indices[x-1] = y - x;
        for (let i = 1; i < width; ++i) {
            indices[x-1+i] = -i;
        }
    }
    this._loaded = true;
}

MergedRangesBuffer.prototype.get_merge = function(row_name, column) {
    // return [start, end] or null
    this.ensure_loaded();
    if (this._row_map[row_name] == null)
        throw new SheetBufferError("unknown row: " + row_name)
    if (typeof column != "number" || isNaN(column))
        throw new Error( "MergedRangesBuffer().get_merge: " +
            "internal type error (column)" );
    let indices = this._indices[row_name];
    let index = indices[column-1];
    if (index == 0) {
        return null;
    } else if (index > 0) {
        return [column, column + index];
    } else {
        let start = column + index;
        return [start, start + indices[start-1]];
    }
}

MergedRangesBuffer.prototype.is_overlapped = function(row_name, column) {
    // return true if the cell is part of a merge and not a start of it
    if (typeof column !== "number")
        throw new Error("MergedRangesBuffer().is_overlapped: internal error");
    if (this._row_map[row_name] == null)
        throw new SheetBufferError("unknown row: " + row_name)
    this.ensure_loaded();
    let indices = this._indices[row_name];
    return indices[column-1] < 0;
}

MergedRangesBuffer.prototype.merge = function(row_name, start, end) {
    if (!this._loaded)
        return;
    if (this._row_map[row_name] == null)
        throw new SheetBufferError("unknown row: " + row_name)
    if (this.is_overlapped(row_name, start))
        throw new SheetBufferMergeOverlap(
            "an existing merge overlaps merge area boundary",
            this.sheet.getRange(this._row_map[row_name], start) );
    if (end < this.dim.sheet_width && this.is_overlapped(row_name, end + 1))
        throw new SheetBufferMergeOverlap(
            "an existing merge overlaps merge area boundary",
            this.sheet.getRange(this._row_map[row_name], end) );
    let indices = this._indices[row_name];
    let width = end - start + 1;
    if (width >= 2) {
        indices[start-1] = end - start;
        for (let i = 1; i < width; ++i) {
            indices[start-1+i] = -i;
        }
    }
}

MergedRangesBuffer.prototype.unmerge = function(row_name, start, end) {
    if (!this._loaded)
        return;
    if (this._row_map[row_name] == null)
        throw new SheetBufferError("unknown row: " + row_name)
    if (this.is_overlapped(row_name, start))
        throw new SheetBufferMergeOverlap(
            "an existing merge overlaps unmerge area boundary",
            this.sheet.getRange(this._row_map[row_name], start) );
    if (end < this.dim.sheet_width && this.is_overlapped(row_name, end + 1))
        throw new SheetBufferMergeOverlap(
            "an existing merge overlaps unmerge area boundary",
            this.sheet.getRange(this._row_map[row_name], end) );
    let indices = this._indices[row_name];
    indices.fill(0, start - 1, end);
}

MergedRangesBuffer.prototype.insert_columns = function(column, num_columns) {
    if (!this._loaded)
        return;
    if (typeof column != "number" || isNaN(column))
        throw new Error( "MergedRangesBuffer().insert_columns: " +
            "internal type error (column)" );
    if (typeof num_columns != "number" || isNaN(num_columns))
        throw new Error( "MergedRangesBuffer().insert_columns: " +
            "internal type error (num_columns)" );
    for (let row_name in this._row_map) {
        let indices = this._indices[row_name];
        if (column < 1 || column > indices.length + 1)
            throw new Error( "MergedRangesBuffer().insert_columns: " +
                "internal error (index out of bounds)" );
        let index = column <= indices.length ?
            indices[column-1] : 0;
        if (index >= 0) {
            let zeroes = new Array(num_columns);
            zeroes.fill(0);
            indices.splice(column-1, 0, ...zeroes);
        } else {
            let start = column + index;
            let old_width = indices[start-1] + 1;
            let old_end = start + old_width - 1;
            let new_width = old_width + num_columns;
            let new_end = start + new_width - 1;
            let new_indices = new Array(new_width);
            new_indices[0] = new_end - start;
            for (let i = 1; i < new_width; ++i) {
                new_indices[i] = -i;
            }
            indices.splice(start - 1, old_width, ...new_indices);
        }
        if (indices.length != this.dim.sheet_width)
            throw new Error( "MergedRangesBuffer().insert_columns: " +
                "internal error (incorrect array length)" );
    }
}

MergedRangesBuffer.prototype.delete_columns = function(column, num_columns) {
    if (!this._loaded)
        return;
    var start = column, end = column + num_columns - 1;
    for (let row_name in this._row_map) {
        let indices = this._indices[row_name];
        let survived_merges = [];
        let inside_merge = false;
        if (indices[start-1] < 0) {
            let merge_start = start + indices[start-1];
            let merge_end = merge_start + indices[merge_start-1];
            if (merge_end > end) {
                merge_end -= num_columns;
                inside_merge = true;
            } else {
                merge_end = start - 1;
            }
            survived_merges.push([merge_start, merge_end]);
        }
        if (!inside_merge && indices[end] < 0) {
            let merge_start = (end + 1) + indices[end];
            let merge_end = merge_start + indices[merge_start-1];
            if (merge_start < start || merge_end <= end)
                throw new Error( "MergedRangesBuffer.delete_columns: " +
                    "internal error" );
            merge_start = start;
            merge_end -= num_columns;
            survived_merges.push([merge_start, merge_end]);
        }
        indices.splice(column-1, num_columns);
        for (let [merge_start, merge_end] of survived_merges) {
            let merge_width = merge_end - merge_start + 1;
            if (merge_width == 1) {
                indices[merge_start-1] = 0;
                continue;
            }
            indices[merge_start-1] = merge_end - merge_start;
            for (let i = 1; i < merge_width; ++i) {
                indices[merge_start-1+i] = -i;
            }
        }
    }
}





const chunk_size = 100;

function SheetBuffer(sheet, row_map, super_dim = {}) {
    this.sheet = sheet;
    this.dim = {};
    define_lazy_properties_(this.dim, {
        sheet_id: function() {
            return super_dim.sheet_id || sheet.getSheetId();
        },
        sheet_height: function() {
            return super_dim.sheet_height || sheet.getMaxRows();
        },
        frozen_height: function() {
            return super_dim.frozen_height || sheet.getFrozenRows();
        },
        sheet_width: function() {
            return sheet.getMaxColumns();
        },
    });
    this._row_map = row_map;
    this._load_height = Math.max(...Object.entries(row_map).map(([,v]) => v));
    this._loaded_start = null;
    this._loaded_end = null;
    this._data = {
        values:   new_loaded_values(row_map),
        formulas: new_loaded_values(row_map),
        notes:    new_loaded_values(row_map),
    };
    this._merges = new MergedRangesBuffer(sheet, row_map, this.dim);
}

function new_loaded_values(row_map) {
    var loaded = {};
    for (let row_name in row_map) {
        loaded[row_name] = [];
    }
    return loaded;
}

SheetBuffer.prototype.ensure_loaded = function(start, end) {
    const dim = this.dim;
    const sheet_width = dim.sheet_width;
    if (start == null) {
        throw new SheetBufferError("ensure_loaded: invalid index");
    }
    if (start < 1) {
        start = 1;
    }
    if (end == null)
        end = start;
    if (end > sheet_width) {
        end = sheet_width;
    }
    return SheetBuffer_ensure_loaded.call(this, start, end);
}

function SheetBuffer_ensure_loaded(start, end) {
    if (end == null)
        end = start;
    const dim = this.dim;
    const sheet_width = dim.sheet_width;
    if (start < 1 || end > sheet_width || isNaN(start) || isNaN(end))
        throw new SheetBufferError(
            "ensure_loaded (internal): index out of bounds " +
            "(1 < " + start + " < " + end + " < " + sheet_width + ")");
    if (this._loaded_start == null) {
        this._loaded_start = start +
            Math.ceil((end - start + 1 - chunk_size) / 2);
        if (this._loaded_start > start)
            this._loaded_start = start;
        else if (this._loaded_start < 1)
            this._loaded_start = 1;
        this._loaded_end = this._loaded_start - 1;
    }
    while (start < this._loaded_start) {
        let load_start = this._loaded_start - chunk_size;
        if (load_start < 1)
            load_start = 1;
        let [added_values, added_formulas, added_notes] =
            SheetBuffer_load_chunk.call( this,
                load_start, this._loaded_start - 1 );
        for (let [loaded, added] of [
            [this._data.values,   added_values  ],
            [this._data.formulas, added_formulas],
            [this._data.notes,    added_notes   ],
        ]) {
            for (let row_name in this._row_map) {
                Array.prototype.unshift.apply( loaded[row_name],
                    added[this._row_map[row_name] - 1] );
            }
        }
        this._loaded_start = load_start;
    }
    while (end > this._loaded_end) {
        let load_end = this._loaded_end + chunk_size;
        if (load_end > sheet_width)
            load_end = sheet_width;
        let [added_values, added_formulas, added_notes] = 
            SheetBuffer_load_chunk.call( this,
                this._loaded_end + 1, load_end );
        for (let [loaded, added] of [
            [this._data.values,   added_values  ],
            [this._data.formulas, added_formulas],
            [this._data.notes,    added_notes   ],
        ]) {
            for (let row_name in this._row_map) {
                Array.prototype.push.apply( loaded[row_name],
                    added[this._row_map[row_name] - 1] );
            }
        }
        this._loaded_end = load_end;
    }
}

function SheetBuffer_load_chunk(start, end) {
    SpreadsheetFlusher.add_dimensions( true,
        this.dim.sheet_id, this.dim.sheet_height, this.dim.sheet_width,
        1, start, this._load_height, end - start + 1 );
    var range = this.sheet.getRange(
        1, start, this._load_height, end - start + 1 );
    return [range.getValues(), range.getFormulasR1C1(), range.getNotes()];
}

function SheetBuffer_slice(value_type, row_name, start, end) {
    // start and end are inclusive
    if ( start < 1 || start > this.dim.sheet_width ||
            end < 1 || end > this.dim.sheet_width )
        throw new SheetBufferError("slice: index out of bounds");
    if (end < start)
        throw new SheetBufferError("slice: invalid indices");
    SheetBuffer_ensure_loaded.call(this, start, end);
    if (this._row_map[row_name] == null)
        throw new SheetBufferError("unknown row: " + row_name)
    return this._data[value_type][row_name].slice(
        start - this._loaded_start, end - this._loaded_start + 1 );
}

SheetBuffer.prototype.slice_values = function(row_name, start, end) {
    return SheetBuffer_slice.call(this, "values", row_name, start, end);
}

SheetBuffer.prototype.slice_formulas = function(row_name, start, end) {
    return SheetBuffer_slice.call(this, "formulas", row_name, start, end);
}

function SheetBuffer_get(value_type, row_name, column) {
    if (column < 1 || column > this.dim.sheet_width)
        throw new SheetBufferError("get: index out of bounds");
    SheetBuffer_ensure_loaded.call(this, column);
    if (this._row_map[row_name] == null)
        throw new SheetBufferError("unknown row: " + row_name)
    return this._data[value_type][row_name][column - this._loaded_start];
}

SheetBuffer.prototype.get_value = function(row_name, column) {
    return SheetBuffer_get.call(this, "values", row_name, column);
}

SheetBuffer.prototype.get_formula = function(row_name, column) {
    return SheetBuffer_get.call(this, "formulas", row_name, column);
}

SheetBuffer.prototype.get_note = function(row_name, column) {
    return SheetBuffer_get.call(this, "notes", row_name, column);
}

SheetBuffer.prototype.set_value = function(row_name, column, value) {
    if (column < 1 || column > this.dim.sheet_width)
        throw new SheetBufferError("set_value: index out of bounds");
    if (typeof value === "string" && value.startsWith("=")) {
        throw new SheetBufferError( "set_value: " +
            "cannot set a formula value with this method".
            this.sheet.getRange(this._row_map[row_name], column) );
    }
    if (this._row_map[row_name] == null)
        throw new SheetBufferError("unknown row: " + row_name)
    if ( this._loaded_start != null &&
        column >= this._loaded_start && column <= this._loaded_end
    ) {
        let i = column - this._loaded_start;
        this._data.values  [row_name][i] = value != null ? value : "";
        this._data.formulas[row_name][i] = "";
    }
    this.sheet.getRange(this._row_map[row_name], column).setValue(value);
    SpreadsheetFlusher.reset();
}

SheetBuffer.prototype.set_values = function(row_name, start, end, values) {
    if (start < 1 || end > this.dim.sheet_width)
        throw new SheetBufferError("set_values: index out of bounds");
    if (start > end)
        throw new SheetBufferError("set_values: invalid indices");
    var width = end - start + 1;
    if (!(values instanceof Array)) {
        let values_array = new Array(width);
        values_array.fill(values);
        values = values_array;
    } else if (values.length != width) {
        throw new SheetBufferError("set_values: invalid values array");
    }
    if (this._row_map[row_name] == null)
        throw new SheetBufferError("unknown row: " + row_name)
    var formulas = new Array(width);
    formulas.fill("");
    if ( this._loaded_start != null &&
        end >= this._loaded_start && start <= this._loaded_end
    ) {
        let i = start - this._loaded_start;
        this._data.values  [row_name].splice( i, width,
            ...values.map(v => v != null ? v : "") );
        this._data.formulas[row_name].splice(i, width, ...formulas);
    }
    this.sheet.getRange(this._row_map[row_name], start, 1, width)
        .setValues([values]);
    SpreadsheetFlusher.reset();
}

SheetBuffer.prototype.set_formula = function( row_name, column, formula,
    value_replace = ""
) {
    if (column < 1 || column > this.dim.sheet_width)
        throw new SheetBufferError("set_formula: index out of bounds");
    if (this._row_map[row_name] == null)
        throw new SheetBufferError("unknown row: " + row_name)
    if ( this._loaded_start != null &&
        column >= this._loaded_start && column <= this._loaded_end
    ) {
        let i = column - this._loaded_start;
        this._data.formulas[row_name][i] = formula;
        this._data.values  [row_name][i] = value_replace;
    }
    this.sheet.getRange(this._row_map[row_name], column)
        .setFormulaR1C1(formula);
    SpreadsheetFlusher.reset();
}

SheetBuffer.prototype.set_formulas = function( row_name, start, end, formulas,
    values_replace = ""
) {
    if (start < 1 || end > this.dim.sheet_width)
        throw new SheetBufferError("set_formulas: index out of bounds");
    if (start > end)
        throw new SheetBufferError("set_formulas: invalid indices");
    var width = end - start + 1;
    if (!(formulas instanceof Array)) {
        let formulas_array = new Array(width);
        formulas_array.fill(formulas);
        formulas = formulas_array;
    } else if (formulas.length != width) {
        throw new SheetBufferError("set_formulas: invalid formulas array");
    }
    if (this._row_map[row_name] == null)
        throw new SheetBufferError("unknown row: " + row_name)
    if (!(values_replace instanceof Array)) {
        let values_array = new Array(width);
        values_array.fill(values_replace);
        values_replace = values_array;
    } else if (values_replace.length != width) {
        throw new SheetBufferError( "set_formulas: " +
            "invalid values_replace array" );
    }
    if ( this._loaded_start != null &&
        end >= this._loaded_start && start <= this._loaded_end
    ) {
        let i = start - this._loaded_start;
        this._data.formulas[row_name].splice(i, width, ...formulas);
        this._data.values  [row_name].splice(i, width, ...values_replace);
    }
    this.sheet.getRange(this._row_map[row_name], start, 1, width)
        .setFormulasR1C1([formulas]);
    SpreadsheetFlusher.reset();
}

SheetBuffer.prototype.set_note = function(row_name, column, note) {
    if (column < 1 || column > this.dim.sheet_width)
        throw new SheetBufferError("set_note: index out of bounds");
    if (this._row_map[row_name] == null)
        throw new SheetBufferError("unknown row: " + row_name)
    if ( this._loaded_start != null &&
        column >= this._loaded_start && column <= this._loaded_end
    ) {
        let i = column - this._loaded_start;
        this._data.notes[row_name][i] = note != null ? note : "";
    }
    this.sheet.getRange(this._row_map[row_name], column).setNote(note);
    SpreadsheetFlusher.reset();
}

SheetBuffer.prototype.merge = function(row_name, start, end) {
    if (start < 1 || end > this.dim.sheet_width)
        throw new SheetBufferError("merge: index out of bounds");
    if (start > end)
        throw new SheetBufferError("merge: invalid indices");
    if (this._row_map[row_name] == null)
        throw new SheetBufferError("unknown row: " + row_name)
    this._merges.merge(row_name, start, end);
    if ( this._loaded_start &&
        start <= this._loaded_end && end >= this._loaded_start
    ) {
        this.ensure_loaded(start, end);
        let first_value = this.find("values", row_name, start, end);
        let first_formula = this.find("formulas", row_name, start, end);
        let first_cell = (first_formula == null || first_value == null) ?
            (first_value || first_formula || start) :
            ((first_value <= first_formula) ? first_value : first_formula);
        let i = start - this._loaded_start;
        let j = first_cell - this._loaded_start;
        let e = end + 1 - this._loaded_start;
        for (let values of [
            this._data.values,
            this._data.formulas, // XXX fix formulas (they are A1-preserved)
            this._data.notes,
        ]) {
            let row_values = values[row_name];
            row_values[i] = row_values[j];
            row_values.fill("", i + 1, e);
        }
    }
    this.sheet.getRange(
        this._row_map[row_name], start, 1, end - start + 1
    ).merge();
    SpreadsheetFlusher.reset();
}

SheetBuffer.prototype.unmerge = function(row_name, start, end) {
    if (start < 1 || end > this.dim.sheet_width)
        throw new SheetBufferError("unmerge: index out of bounds");
    if (start > end)
        throw new SheetBufferError("unmerge: invalid indices");
    if (this._row_map[row_name] == null)
        throw new SheetBufferError("unknown row: " + row_name)
    this._merges.unmerge(row_name, start, end);
    this.sheet.getRange(
        this._row_map[row_name], start, 1, end - start + 1
    ).breakApart();
    SpreadsheetFlusher.reset();
}

function SheetBuffer_insert_columns(column, num_columns) {
    Object.defineProperty( this.dim, "sheet_width",
        {configurable: true, value: this.dim.sheet_width + num_columns} );
    if (this._loaded_start == null || column > this._loaded_end) {
        // no-op
    } else if (column <= this._loaded_start) {
        this._loaded_start += num_columns;
        this._loaded_end += num_columns;
    } else {
        this._loaded_end += num_columns;
        let splice_index = column - this._loaded_start;
        let blanks = new Array(num_columns);
        blanks.fill("");
        for (let loaded of [
            this._data.values, this._data.formulas, this._data.notes
        ]) {
            for (let row_name in this._row_map) {
                loaded[row_name].splice(splice_index, 0, ...blanks);
            }
        }
    }
    this._merges.insert_columns(column, num_columns);
}

SheetBuffer.prototype.insert_columns_after = function(column, num_columns) {
    if (column < 1 || column > this.dim.sheet_width)
        throw new SheetBufferError( "insert_columns_after: " +
            "index out of bounds" );
    SheetBuffer_insert_columns.call(this, column + 1, num_columns);
    this.sheet.insertColumnsAfter(column, num_columns);
    SpreadsheetFlusher.reset();
}

SheetBuffer.prototype.insert_columns_before = function(column, num_columns) {
    if (column < 1 || column > this.dim.sheet_width)
        throw new SheetBufferError( "insert_columns_before: " +
            "index out of bounds" );
    SheetBuffer_insert_columns.call(this, column, num_columns);
    this.sheet.insertColumnsBefore(column, num_columns);
    SpreadsheetFlusher.reset();
}

function SheetBuffer_delete_columns(column, num_columns) {
    var start = column, end = column + num_columns - 1;
    Object.defineProperty( this.dim, "sheet_width",
        {configurable: true, value: this.dim.sheet_width - num_columns} );
    if (this._loaded_start != null) {
        let delete_start = start > this._loaded_start ?
            start : this._loaded_start;
        let delete_end = end < this._loaded_end ?
            end : this._loaded_end ;
        if (delete_start > this._loaded_end) {
            // no-op
        } else if (delete_end < this._loaded_start) {
            this._loaded_start -= num_columns;
            this._loaded_end -= num_columns;
        } else if (delete_end < delete_start) {
            throw new Error("SheetBuffer.delete_columns: internal error");
        } else if (start <= this._loaded_start && end >= this._loaded_end) {
            this._loaded_start = this._loaded_end = null;
            for (let loaded of [
                this._data.values, this._data.formulas, this._data.notes
            ]) {
                for (let row_name in this._row_map) {
                    loaded[row_name].length = 0;
                }
            }
        } else {
            let splice_index = delete_start - this._loaded_start;
            let delete_count = delete_end - delete_start + 1;
            if (delete_start > start)
                this._loaded_start -= (delete_start - start)
            this._loaded_end -= delete_end - start + 1;
            for (let loaded of [
                this._data.values, this._data.formulas, this._data.notes
            ]) {
                for (let row_name in this._row_map) {
                    loaded[row_name].splice(splice_index, delete_count);
                }
            }
        }
    }
    this._merges.delete_columns(column, num_columns);
}

SheetBuffer.prototype.delete_columns = function(column, num_columns) {
    if (typeof num_columns !== "number" || num_columns < 1)
        throw new SheetBufferError( "delete_columns: " +
            "invalid argument" );
    if (column < 1 || column + num_columns - 1 > this.dim.sheet_width)
        throw new SheetBufferError( "delete_columns: " +
            "index out of bounds" );
    SheetBuffer_delete_columns.call(this, column, num_columns);
    this.sheet.deleteColumns(column, num_columns);
    SpreadsheetFlusher.reset();
}

SheetBuffer.prototype.find = function(
    value_type, row_name, value, start, end
) {
    if (start == null || start < 1)
        start = 1;
    if (end == null || end > this.dim.sheet_width)
        end = this.dim.sheet_width;
    if (end < start)
        return null;
    if (this._row_map[row_name] == null)
        throw new SheetBufferError("unknown row: " + row_name)
    if (this._loaded_start == null) {
        if (end - start + 1 <= chunk_size) {
            SheetBuffer_ensure_loaded.call(this, start, end);
        } else {
            SheetBuffer_ensure_loaded.call( this,
                start, start + chunk_size - 1 );
        }
    } else {
        SheetBuffer_ensure_loaded.call(this, start);
    }
    var current_start = start;
    while (true) {
        let current_end = (end > this._loaded_end) ? this._loaded_end : end;
        let i_end = current_end - this._loaded_start;
        let values = this._data[value_type][row_name];
        for (let i = current_start - this._loaded_start; i <= i_end; ++i) {
            if (values[i] === value)
                return i + this._loaded_start;
        }
        if (current_end < end) {
            current_start = current_end + 1;
            SheetBuffer_ensure_loaded.call(this, current_start);
        } else {
            return null;
        }
    }
}

SheetBuffer.prototype.find_last = function(
    value_type, row_name, value, start, end
) {
    if (start == null || start > this.dim.sheet_width)
        start = this.dim.sheet_width;
    if (end == null || end < 1)
        end = 1;
    if (end > start)
        return null;
    if (this._row_map[row_name] == null)
        throw new SheetBufferError("unknown row: " + row_name)
    if (this._loaded_start == null) {
        if (start - end + 1 <= chunk_size) {
            SheetBuffer_ensure_loaded.call(this, end, start);
        } else {
            SheetBuffer_ensure_loaded.call( this,
                start - chunk_size + 1, start );
        }
    } else {
        SheetBuffer_ensure_loaded.call(this, start);
    }
    var current_start = start;
    while (true) {
        let current_end = (end < this._loaded_start) ?
            this._loaded_start : end ;
        let i_end = current_end - this._loaded_start;
        let values = this._data[value_type][row_name];
        for (let i = current_start - this._loaded_start; i >= i_end; --i) {
            if (values[i] === value)
                return i + this._loaded_start;
        }
        if (current_end > end) {
            current_start = current_end - 1;
            SheetBuffer_ensure_loaded.call(this, current_start);
        } else {
            return null;
        }
    }
}

SheetBuffer.prototype.find_closest = function(
    value_type, row_name, value, center
) {
    if (center == null)
        throw new SheetBufferError("find_closest: invalid argument");
    if (center <= 1)
        return this.find(value_type, row_name, value);
    if (center >= this.dim.sheet_width)
        return this.find_last(value_type, row_name, value);
    if (this._row_map[row_name] == null)
        throw new SheetBufferError("unknown row: " + row_name)
    SheetBuffer_ensure_loaded.call(this, center);
    var current_start = center;
    while (true) {
        let alt_end = 2 * center - this._loaded_start;
        let current_end = (alt_end < this._loaded_end) ?
            alt_end : this._loaded_end;
        if (current_end < current_start)
            throw new Error("SheetBuffer.find_closest: internal error");
        let values = this._data[value_type][row_name];
        for (let i = current_start; i <= current_end; ++i) {
            if (values[i - this._loaded_start] === value)
                return i;
            if ( i > center &&
                values[2 * center - i - this._loaded_start] === value
            ) {
                return 2 * center - i;
            }
        }
        current_start = current_end + 1;
        if (current_start > this.dim.sheet_width) {
            return this.find_last( value_type, row_name,
                value, 2 * center - current_start );
        } else if (2 * center - current_start < 1) {
            return this.find(value_type, row_name, value, current_start);
        }
        SheetBuffer_ensure_loaded.call(this, current_start);
        SheetBuffer_ensure_loaded.call(this, 2 * center - current_start);
    }
}

SheetBuffer.prototype.find_value = function(
    row_name, value, start, end
) {
    return this.find("values", row_name, value, start, end);
}

SheetBuffer.prototype.find_last_value = function(
    row_name, value, start, end
) {
    return this.find_last("values", row_name, value, start, end);
}

SheetBuffer.prototype.find_merge = function(
    row_name, start, end, options = {}
) {
    /* return
     *   * either [s, e] where s and e are boundary columns of the next merge
     *     (from start);
     *   * or [x, x] if cell in column x has nonempty value, formula or note;
     *   * or throw an error if a merge overlaps start column
     *     (and not starts at it);
     *   * or throw an error if the next merge overlaps end column
     *     (and not ends at it);
     *   * or null.
     * if options.allow_start_overlap is true, then returned merge may
     * overlap start column
     * if options.allow_end_overlap is true, then returned merge may
     * overlap end column
     */   
    ({
        allow_overlap_start: options.allow_overlap_start = false,
        allow_overlap_end:   options.allow_overlap_end   = false,
    } = options);
    if (start != null && typeof start != "number")
        throw new Error("SheetBuffer().find_merge: internal type error");
    if (end != null && typeof end != "number")
        throw new Error("SheetBuffer().find_merge: internal type error");
    if (start == null || start < 1)
        start = 1;
    if (end == null || end > this.dim.sheet_width)
        end = this.dim.sheet_width;
    if (end < start)
        return null;
    if (this._row_map[row_name] == null)
        throw new SheetBufferError("unknown row: " + row_name)
    if (this._loaded_start == null) {
        if (end - start + 1 <= chunk_size) {
            SheetBuffer_ensure_loaded.call(this, start, end);
        } else {
            SheetBuffer_ensure_loaded.call( this,
                start, start + chunk_size - 1 );
        }
    } else {
        SheetBuffer_ensure_loaded.call(this, start);
    }
    this._merges.ensure_loaded();
    var current_start = start;
    while (true) {
        let current_end = (end > this._loaded_end) ?
            this._loaded_end : end;
        let values = this._data.values[row_name];
        let formulas = this._data.formulas[row_name];
        let notes = this._data.notes[row_name];
        for (let i = current_start; i <= current_end; ++i) {
            let merge = this._merges.get_merge(row_name, i);
            if (merge != null) {
                let [merge_start, merge_end] = merge;
                if (merge_start < i) {
                    if (i > start)
                        throw new Error(
                            "SheetBuffer.find_merge: internal error" );
                    if (!options.allow_overlap_start)
                        throw new SheetBufferMergeOverlap(
                            "a merge overlaps search area boundary",
                            this.sheet.getRange(this._row_map[row_name], start)
                        );
                } else if (merge_end > end) {
                    if (!options.allow_overlap_end)
                        throw new SheetBufferMergeOverlap(
                            "a merge overlaps search area boundary",
                            this.sheet.getRange(this._row_map[row_name], end)
                        );
                }
                return [merge_start, merge_end];
            }
            let ii = i - this._loaded_start;
            if (values[ii] != "" || formulas[ii] != "" || notes[ii] != "") {
                return [i, i];
            }
        }
        if (current_end < end) {
            current_start = current_end + 1;
            this.storage.ensure_loaded(current_start);
        } else {
            break;
        }
    }
    return null;
}

SheetBuffer.prototype.find_last_merge = function(
    row_name, start, end, options = {}
) {
    ({
        allow_overlap_start: options.allow_overlap_start = false,
        allow_overlap_end:   options.allow_overlap_end   = false,
    } = options);
    if (start != null && typeof start != "number")
        throw new Error("SheetBuffer().find_last_merge: internal type error");
    if (end != null && typeof end != "number")
        throw new Error("SheetBuffer().find_last_merge: internal type error");
    if (start == null || start > this.dim.sheet_width)
        start = this.dim.sheet_width;
    if (end == null || end < 1)
        end = 1;
    if (end > start)
        return null;
    if (this._row_map[row_name] == null)
        throw new SheetBufferError("unknown row: " + row_name)
    if (this._loaded_start == null) {
        if (start - end + 1 <= chunk_size) {
            SheetBuffer_ensure_loaded.call(this, end, start);
        } else {
            SheetBuffer_ensure_loaded.call( this,
                start - chunk_size + 1, start );
        }
    } else {
        SheetBuffer_ensure_loaded.call(this, start);
    }
    this._merges.ensure_loaded();
    var current_start = start;
    while (true) {
        let current_end = (end < this._loaded_start) ?
            this._loaded_start : end;
        let values = this._data.values[row_name];
        let formulas = this._data.formulas[row_name];
        let notes = this._data.notes[row_name];
        for (let i = current_start; i >= current_end; --i) {
            let merge = this._merges.get_merge(row_name, i);
            if (merge != null) {
                let [merge_start, merge_end] = merge;
                if (merge_end > i) {
                    if (i < start)
                        throw new Error(
                            "SheetBuffer.find_merge: internal error" );
                    if (!options.allow_overlap_start)
                        throw new SheetBufferMergeOverlap(
                            "a merge overlaps search area boundary",
                            this.sheet.getRange(this._row_map[row_name], start)
                        );
                } else if (merge_start < end) {
                    if (!options.allow_overlap_end)
                        throw new SheetBufferMergeOverlap(
                            "a merge overlaps search area boundary",
                            this.sheet.getRange(this._row_map[row_name], end)
                        );
                }
                return [merge_start, merge_end];
            }
            let ii = i - this._loaded_start;
            if (values[ii] != "" || formulas[ii] != "" || notes[ii] != "") {
                return [i, i];
            }
        }
        if (current_end > end) {
            current_start = current_end - 1;
            this.storage.ensure_loaded(current_start);
        } else {
            break;
        }
    }
    return null;
}

SheetBuffer.prototype.test = function() {
    if (this._loaded_start == null)
        return;
    var loaded_width = this._loaded_end - this._loaded_start + 1;
    var range = this.sheet.getRange(
        1, this._loaded_start,
        this.dim.frozen_height, loaded_width );
    var values = range.getValues();
    var formulas = range.getFormulasR1C1();
    var notes = range.getNotes();
    for (let [value_type, loaded_values, real_values] of [
        ["values",   this._data.values,   values  ],
        ["formulas", this._data.formulas, formulas],
        ["notes",    this._data.notes,    notes   ],
    ]) {
        for (let row_name in this._row_map) {
            let loaded_row = loaded_values[row_name];
            let real_row = real_values[this._row_map[row_name]-1];
            for (let i = 0; i < loaded_width; ++i) {
                 if (loaded_row[i] != real_row[i]) {
                     console.error(
                         this.sheet.getRange(
                             this._row_map[row_name], i + this._loaded_start )
                             .getA1Notation() + ": (" + value_type + ") " +
                         loaded_row[i] + " ≠ " + real_row[i]
                     );
                 }
            }
        }
    }
}

return SheetBuffer;
}(); // end SheetBuffer namespace

// vim: set fdm=marker sw=4 :
