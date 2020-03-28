var CFormatting = function() { // namespace

// Known pitfalls:
//   for localization purposes, it is better to never use decimal fractions at all

// XXX change this function; always search for range
function find(sheet, rule, options) {
  if (options == null)
    options = {ignore_effect: false, includes_range: null};
  if (options.includes_range != null) {
    var target_dim = {
      t: options.includes_range.getRow(),
      l: options.includes_range.getColumn(),
      b: options.includes_range.getLastRow(),
      r: options.includes_range.getLastColumn()
    };
  }

  var rules = sheet.getConditionalFormatRules();
  var count = rules.length;

  var target_code = encode(rule, {R1C1_mustdo: true, ignore_effect: options.ignore_effect});
  var codes = [];
  for (var i = 0; i < count; ++i) {
    codes.push(encode(rules[i], {R1C1_mustdo: false, ignore_effect: options.ignore_effect}));
  }

  for (var i = 0; i < codes.length; ++i) {
    if (!code_eq(codes[i], target_code))
      continue;
    if (options.includes_range != null) {
      var found_it = false;
      var ranges = rules[i].getRanges();
      for (var j = 0; j < ranges.length; ++j) {
        var range = ranges[j];
        if (
          range.getRow()        != target_dim.t ||
          range.getColumn()     != target_dim.l ||
          range.getLastRow()    != target_dim.b ||
          range.getLastColumn() != target_dim.r
        )
          continue;
        found_it = true;
        break;
      }
      if (!found_it)
        continue;
    }
    return i;
  }
  return -1;
}

function merge(sheet, new_rules) {
  var new_count = new_rules.length;
  var rules = sheet.getConditionalFormatRules();
  var old_count = rules.length;

  var new_codes = [];
  for (var ni = 0; ni < new_count; ++ni) {
    new_codes.push(encode(new_rules[ni], {R1C1_mustdo: true}));
  }
  var old_codes = [];
  for (var i = 0; i < old_count; ++i) {
    old_codes.push(encode(rules[i], {R1C1_mustdo: false}));
  }

  var merge_targets = [], rev_merge_targets = {};
  for (var ni = 0; ni < new_count; ++ni) {
    for (var i = 0; i < old_count; ++i) {
      if (code_eq(new_codes[ni], old_codes[i])) {
        merge_targets[ni] = i;
        if (rev_merge_targets[i] != null)
          throw "no new rules can be the same: " + rev_merge_targets[i] + " " + ni;
        rev_merge_targets[i] = ni;
        break;
      }
    }
  }

  for (var ni = 0; ni < new_count; ++ni) {
    var new_rule = new_rules[ni];
    var i = merge_targets[ni];
    if (i == null) {
      rules.push(new_rule);
    } else {
      var old_rule = rules[i];
      var ranges = old_rule.getRanges();
      Array.prototype.push.apply(ranges, new_rule.getRanges());
      var rule = old_rule.copy().setRanges(ranges).build();
      rules[i] = rule;
    }
  }

  sheet.setConditionalFormatRules(rules);
}

function encode(rule, options) {
  if (options == null)
    options = {R1C1_mustdo: false, ignore_effect: false};
  var R1C1_refcell = rule.getRanges()[0];
  var code = [];
  var boolean_condition = rule.getBooleanCondition();
  if (boolean_condition !== null) {
    code.push(true);
    var criteria_type = boolean_condition.getCriteriaType();
    code.push(criteria_type);
    var criteria_values = boolean_condition.getCriteriaValues();
    for (var i = 0; i < criteria_values.length; ++i) {
      code.push(encode_formula(criteria_values[i], R1C1_refcell, options.R1C1_mustdo));
    }
    if (!options.ignore_effect) {
      code.push(normalize_colour(boolean_condition.getBackground()));
      code.push(normalize_colour(boolean_condition.getFontColor()));
      code.push(boolean_condition.getBold());
      code.push(boolean_condition.getItalic());
      code.push(boolean_condition.getStrikethrough());
      code.push(boolean_condition.getUnderline());
    }
  } else {
    code.push(false);
    var gradient_condition = rule.getGradientCondition();
    code.push(gradient_condition.getMinType());
    code.push(encode_formula(gradient_condition.getMinValue(), R1C1_refcell, options.R1C1_mustdo));
    if (!options.ignore_effect)
      code.push(normalize_colour(gradient_condition.getMinColor()));
    var mid_type;
    code.push(mid_type = gradient_condition.getMidType());
    if (mid_type == null) {
      code.push(null, null);
    } else {
      code.push(encode_formula(gradient_condition.getMidValue(), R1C1_refcell, options.R1C1_mustdo));
      if (!options.ignore_effect)
        code.push(normalize_colour(gradient_condition.getMidColor()));
    }
    code.push(gradient_condition.getMaxType());
    code.push(encode_formula(gradient_condition.getMaxValue(), R1C1_refcell, options.R1C1_mustdo));
    if (!options.ignore_effect)
      code.push(normalize_colour(gradient_condition.getMaxColor()));
  }
  return code;
}

function normalize_colour(colour) {
  if (colour == null)
    return null;
  if (/#[0-9A-Ha-h]{6}/.exec(colour) == null) {
    if (colour == "black")
      return "#000000";
    else if (colour == "white")
      return "#FFFFFF";
    else if (colour == "red")
      return "#FF0000";
    else
      throw "CFormatting.encode: colour is presented in a non-standard form: " + colour;
  }
  return colour.toUpperCase();
}

function encode_formula(formula, R1C1_refcell, R1C1_mustdo) {
  // R1C1_mustdo, if true, will error if formula is too syntactically complicated,
  // i.e., contains strings or multicell ranges.
  if (typeof formula != "string" || formula.lastIndexOf("=", 0) < 0) {
    // not really a formula, just a value
    return formula;
  }
  if (formula.indexOf('"') > 0 || formula.indexOf(':') > 0) {
    if (R1C1_mustdo)
      throw "impossible to encode " + formula;
    return "=~" + formula;
  }
  formula = formula
    .replace(/,/g, ";")
    .replace(/(\$?)([A-Z]+)(\$?)([0-9]+)/g, function(m, p1, p2, p3, p4) {
      var cell = R1C1_refcell.getSheet().getRange(p2+p4);
      var r, c;
      if (p1 == "$") {
        c = 'C' + cell.getColumn();
      } else {
        c = 'C[' + (cell.getColumn() - R1C1_refcell.getColumn()) + ']';
      }
      if (p3 == "$") {
        r = 'R' + cell.getRow();
      } else {
        r = 'R[' + (cell.getRow() - R1C1_refcell.getRow()) + ']';
      }
      return r + c;
    });
  return formula;
}

function code_eq(acode, bcode) {
  var length = acode.length;
  if (length != bcode.length)
    return false;
  for (var i = 0; i < length; ++i) {
    if (acode[i] != bcode[i])
      return false;
  }
  return true;
}

return {merge: merge};
}() // end CFormatting namespace
