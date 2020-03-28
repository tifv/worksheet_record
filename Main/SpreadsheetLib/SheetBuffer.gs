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
  var merged_ranges = this.sheet.getRange(1, 1, this.dim.frozen_height, this.dim.sheet_width).getMergedRanges();
  var rev_row_map = Object.fromEntries(Object.entries(this._row_map).map(([k,v]) => [v,k]));
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
      throw new Error("MergedRangesBuffer.ensure_loaded: internal error");
    indices[x-1] = y - x;
    for (let i = 1; i < width; ++i) {
      indices[x-1+i] = -i;
    }
  }
  this._loaded = true;
}

MergedRangesBuffer.prototype.get_merge = function(row_name, column) {
  // return [start, end] or null
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

MergedRangesBuffer.prototype.insert_columns = function(column, num_columns) {
  if (!this.loaded)
    return;
  for (let row_name of this._row_map) {
    let indices = this.indices[row_name];
    let index = indices[column-1];
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
      indices.splice(start, old_width, new_indices);
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
    throw new Error("SheetBuffer.ensure_loaded: invalid index");
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
    throw new Error("SheetBuffer.ensure_loaded: index out of bounds");
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
      SheetBuffer_load_chunk.call(this, load_start, this._loaded_start - 1);
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
      SheetBuffer_load_chunk.call(this, this._loaded_end + 1, load_end);
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
    1, start, this.dim.frozen_height, end - start + 1 );
  var range = this.sheet.getRange(1, start, this.dim.frozen_height, end - start + 1);
  return [range.getValues(), range.getFormulasR1C1(), range.getNotes()];
}

function SheetBuffer_slice(value_type, row_name, start, end) {
  // start and end are inclusive
  if ( start < 1 || start > this.dim.sheet_width ||
      end < 1 || end > this.dim.sheet_width )
    throw new Error("SheetBuffer: index out of bounds");
  if (end < start)
    throw new Error("SheetBuffer: invalid indices");
  SheetBuffer_ensure_loaded.call(this, start, end);
  var values_map = this._data[value_type];
  return values_map[row_name].slice(
    start - this._loaded_start, end - this._loaded_start + 1 );
}

SheetBuffer.prototype.slice_values = function(row_name, start, end) {
  return SheetBuffer_slice.call(this, "values", row_name, start, end);
}

function SheetBuffer_get(value_type, row_name, column) {
  if (column < 1 || column > this.dim.sheet_width)
    throw new Error("SheetBuffer: index out of bounds");
  SheetBuffer_ensure_loaded.call(this, column);
  var values_map = this._data[value_type];
  return values_map[row_name][column - this._loaded_start];
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
    throw new Error("SheetBuffer: index out of bounds");
  if (value != null && value.startsWith("=")) {
    throw new Error("SheetBuffer: cannot set a formula value with this method");
  }
  if (column >= this._loaded_start && column <= this._loaded_end) {
    this._data.values  [row_name][column - this._loaded_start] = value != null ? value : "";
    this._data.formulas[row_name][column - this._loaded_start] = "";
  }
  this.sheet.getRange(this._row_map[row_name], column).setValue(value);
  SpreadsheetFlusher.reset();
}

SheetBuffer.prototype.set_formula = function(row_name, column, formula, value_replace = "") {
  if (column < 1 || column > this.dim.sheet_width)
    throw new Error("SheetBuffer: index out of bounds");
  if (column >= this._loaded_start && column <= this._loaded_end) {
    this._data.formulas[row_name][column - this._loaded_start] = formula;
    this._data.values  [row_name][column - this._loaded_start] = value_replace;
  }
  this.sheet.getRange(this._row_map[row_name], column).setFormulaR1C1(formula);
  SpreadsheetFlusher.reset();
}

SheetBuffer.prototype.set_note = function(row_name, column, note) {
  if (column < 1 || column > this.dim.sheet_width)
    throw new Error("SheetBuffer: index out of bounds");
  if (column >= this._loaded_start && column <= this._loaded_end) {
    this._data.notes[row_name][column - this._loaded_start] = note != null ? note : "";
  }
  this.sheet.getRange(this._row_map[row_name], column).setNote(note);
  SpreadsheetFlusher.reset();
}

function SheetBuffer_insert_columns(column, num_columns) {
  if (column <= this._loaded_start) {
    this._loaded_start += num_columns;
    this._loaded_end += num_columns;
    Object.defineProperty(this.dim, "sheet_width",
      {configurable: true, value: this.dim.sheet_width + num_columns} );
  } else if (column > this._loaded_end) {
    Object.defineProperty(this.dim, "sheet_width",
      {configurable: true, value: this.dim.sheet_width + num_columns} );
  } else {
    this._loaded_end += num_columns;
    Object.defineProperty(this.dim, "sheet_width",
      {configurable: true, value: this.dim.sheet_width + num_columns} );
    let splice_index = column - this._loaded_start;
    let blanks = new Array(num_columns);
    blanks.fill("");
    for (let loaded of [this._data.values, this._data.formulas, this._data.notes]) {
      for (let row_name in this._row_map) {
        loaded[row_name].splice(splice_index, 0, ...blanks);
      }
    }
  }
  this._merges.insert_columns(column, num_columns);
}

SheetBuffer.prototype.insert_columns_after = function(column, num_columns) {
  SheetBuffer_insert_columns.call(this, column + 1, num_columns);
  this.sheet.insertColumnsAfter(column, num_columns);
  SpreadsheetFlusher.reset();
}

SheetBuffer.prototype.insert_columns_before = function(column, num_columns) {
  SheetBuffer_insert_columns.call(this, column, num_columns);
  this.sheet.insertColumnsBefore(column, num_columns);
  SpreadsheetFlusher.reset();
}

SheetBuffer.prototype.delete_columns = function(column, num_columns) {
  throw "XXX not implemented";
  this.sheet.deleteColumns(column, num_columns);
  SpreadsheetFlusher.reset();
}

SheetBuffer.prototype.find = function(value_type, row_name, value, start, end) {
  if (start == null || start < 1)
    start = 1;
  if (end == null || end > this.dim.sheet_width)
    end = this.dim.sheet_width;
  if (end < start)
    return null;
  if (this._loaded_start == null) {
    if (end - start + 1 <= chunk_size) {
      SheetBuffer_ensure_loaded.call(this, start, end);
    } else {
      SheetBuffer_ensure_loaded.call(this, start, start + chunk_size - 1);
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

SheetBuffer.prototype.find_last = function(value_type, row_name, value, start, end) {
  if (start == null || start > this.dim.sheet_width)
    start = this.dim.sheet_width;
  if (end == null || end < 1)
    end = 1;
  if (end > start)
    return null;
  if (this._loaded_start == null) {
    if (start - end + 1 <= chunk_size) {
      SheetBuffer_ensure_loaded.call(this, end, start);
    } else {
      SheetBuffer_ensure_loaded.call(this, start - chunk_size + 1, start);
    }
  } else {
    SheetBuffer_ensure_loaded.call(this, start);
  }
  var current_start = start;
  while (true) {
    let current_end = (end < this._loaded_start) ? this._loaded_start : end;
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

SheetBuffer.prototype.find_closest = function(value_type, row_name, value, center) {
  if (center == null)
    throw new Error("SheetBuffer.find_closest: invalid argument");
  if (center <= 1)
    return this.find(value_type, row_name, value);
  if (center >= this.dim.sheet_width)
    return this.find_last(value_type, row_name, value);
  SheetBuffer_ensure_loaded.call(this, center);
  var current_start = center;
  while (true) {
    let alt_end = 2 * center - this._loaded_start;
    let current_end = (alt_end < this._loaded_end) ? alt_end : this._loaded_end;
    if (current_end < current_start)
      throw new Error("SheetBuffer.find_closest: internal error");
    let values = this._data[value_type][row_name];
    for (let i = current_start; i <= current_end; ++i) {
      if (values[i - this._loaded_start] === value)
        return i;
      if (i > center && values[2 * center - i - this._loaded_start] === value)
        return 2 * center - i;
    }
    current_start = current_end + 1;
    if (current_start > this.dim.sheet_width) {
      return this.find_last(value_type, row_name, value, 2 * center - current_start);
    } else if (2 * center - current_start < 1) {
      return this.find(value_type, row_name, value, current_start);
    }
    SheetBuffer_ensure_loaded.call(this, current_start);
    SheetBuffer_ensure_loaded.call(this, 2 * center - current_start);
  }
}

SheetBuffer.prototype.find_value = function(row_name, value, start, end) {
  return this.find("values", row_name, value, start, end);
}

SheetBuffer.prototype.find_last_value = function(row_name, value, start, end) {
  return this.find_last("values", row_name, value, start, end);
}

SheetBuffer.prototype.find_merge = function(row_name, start, end, options = {}) {
  // return
  //   either [s, e] where s and e are boundary columns of the next (from start) merge
  //   or [x, x] if cell in column x has nonempty value, formula or note
  //   or throw an error if a merge overlaps (and not starts at) start column
  //   or throw an error if the next merge overlaps (and not ends at) end column
  // if options.allow_start_overlap is true, then returned merge may overlap start column
  // if options.allow_end_overlap is true, then returned merge may overlap start column
  //   
  ({
    allow_overlap_start: options.allow_overlap_start = false,
    allow_overlap_end:   options.allow_overlap_end   = false,
  } = options);
  if (start == null || start < 1)
    start = 1;
  if (end == null || end > this.dim.sheet_width)
    end = this.dim.sheet_width;
  if (end < start)
    return null;
  if (this._loaded_start == null) {
    if (end - start + 1 <= chunk_size) {
      SheetBuffer_ensure_loaded.call(this, start, end);
    } else {
      SheetBuffer_ensure_loaded.call(this, start, start + chunk_size - 1);
    }
  } else {
    SheetBuffer_ensure_loaded.call(this, start);
  }
  this._merges.ensure_loaded();
  var current_start = start;
  while (true) {
    let current_end = (end > this._loaded_end) ? this._loaded_end : end;
    let values = this._data.values[row_name];
    let formulas = this._data.formulas[row_name];
    let notes = this._data.notes[row_name];
    for (let i = current_start; i <= current_end; ++i) {
      let merge = this._merges.get_merge(row_name, i);
      if (merge != null) {
        let [merge_start, merge_end] = merge;
        if (merge_start < i) {
          if (i > start)
            throw new Error("SheetBuffer.find_merge: internal error");
          if (!options.allow_overlap_start)
            throw new SheetBufferMergeOverlap(
              "a merge overlaps search area boundary",
              this.sheet.getRange(this._row_map[row_name], start) );
        } else if (merge_end > end) {
          if (!options.allow_overlap_end)
            throw new SheetBufferMergeOverlap(
              "a merge overlaps search area boundary",
              this.sheet.getRange(this._row_map[row_name], end) );
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

SheetBuffer.prototype.find_last_merge = function(row_name, start, end, options = {}) {
  ({
    allow_overlap_start: options.allow_overlap_start = false,
    allow_overlap_end:   options.allow_overlap_end   = false,
  } = options);
  if (start == null || start > this.dim.sheet_width)
    start = this.dim.sheet_width;
  if (end == null || end < 1)
    end = this.dim.sheet_width;
  if (end > start)
    return null;
  if (this._loaded_start == null) {
    if (start - end + 1 <= chunk_size) {
      SheetBuffer_ensure_loaded.call(this, end, start);
    } else {
      SheetBuffer_ensure_loaded.call(this, start - chunk_size + 1, start);
    }
  } else {
    SheetBuffer_ensure_loaded.call(this, start);
  }
  this._merges.ensure_loaded();
  var current_start = start;
  while (true) {
    let current_end = (end < this._loaded_start) ? this._loaded_start : end;
    let values = this._data.values[row_name];
    let formulas = this._data.formulas[row_name];
    let notes = this._data.notes[row_name];
    for (let i = current_start; i >= current_end; --i) {
      let merge = this._merges.get_merge(row_name, i);
      if (merge != null) {
        let [merge_start, merge_end] = merge;
        if (merge_end > i) {
          if (i < start)
            throw new Error("SheetBuffer.find_merge: internal error");
          if (!options.allow_overlap_start)
            throw new SheetBufferMergeOverlap(
              "a merge overlaps search area boundary",
              this.sheet.getRange(this._row_map[row_name], start) );
        } else if (merge_start < end) {
          if (!options.allow_overlap_end)
            throw new SheetBufferMergeOverlap(
              "a merge overlaps search area boundary",
              this.sheet.getRange(this._row_map[row_name], end) );
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

SheetBuffer.prototype.merge = function(row_name, start, end) {
  throw "XXX not implemented";
  // remove deleted values
  // and alter merges
}

SheetBuffer.prototype.unmerge = function(row_name, start, end) {
  throw "XXX not implemented";
  // alter merges
}

return SheetBuffer;
}(); // end SheetBuffer namespace

