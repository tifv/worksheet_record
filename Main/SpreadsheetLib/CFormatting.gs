/* Possible actions:
 * • find all rules that:
 *    (a) have given condition
 *    (b) (optional) cover given range
 *    (c) (optional) have given effect
 * • find existing rule with (ac) and extend it to given range or create a new rule that covers this range
 *   • ensure that this new rule comes before any other rule specified with (a)
 * • find all rules with (ab) or (abc) and remove them from given range
 *   • replace them
 * • find all rules with (a) and replace them with similar rule that covers the same ranges and has given effect
 */

class ConditionalFormattingError extends Error {};

var ConditionalFormatting = function() {

class CFCondition {
}

class CFBooleanCondition extends CFCondition {
  constructor({type, values}) {
    super();
    this.type = type;
    this.values = values;
  }
  static from_condition(boolean_condition, value_encode) {
    return new this({
      type: boolean_condition.getCriteriaType(),
      values: boolean_condition.getCriteriaValues()
        .map(value_encode)
    });
  }
  impose_on_rule_builder(builder, value_decode) {
    builder.withCriteria(this.type, this.values.map(value_decode));
  }
  match(other) {
    if (!(other instanceof CFBooleanCondition)) {
      return false;
    }
    var other_values = other.values;
    return (
      this.type == other.type &&
      this.values.length == other.values.length &&
      this.values.every( (element, index) =>
        (element == other_values[index]) )
    );
  }
}

class CFGradientCondition extends CFCondition {
  constructor({
    min_type, min_value = "",
    mid_type, mid_value = "",
    max_type, max_value = "",
  }) {
    super();
    this.min_type  = min_type;
    this.min_value = min_value;
    this.mid_type  = mid_type;
    this.mid_value = mid_value;
    this.max_type  = max_type;
    this.max_value = max_value;
  }
  static from_condition(gradient_condition, value_encode) {
    return new this({
      min_type:  gradient_condition.getMinType(),
      min_value: value_encode(gradient_condition.getMinValue()),
      mid_type:  gradient_condition.getMidType(),
      mid_value: value_encode(gradient_condition.getMidValue()),
      max_type:  gradient_condition.getMaxType(),
      max_value: value_encode(gradient_condition.getMaxValue()),
    });
  }
  match(other) {
    if (!(other instanceof CFGradientCondition)) {
      return false;
    }
    return (
      this.min_type == other.min_type &&
        this.min_values == other.min_values &&
      this.mid_type == other.mid_type &&
        this.mid_values == other.mid_values &&
      this.max_type == other.max_type &&
        this.max_values == other.max_values
    );
  }
}

class CFRange {
  constructor(start_row, start_col, height = 1, width = 1) {
    this.start_row = start_row;
    this.height = height;
    this.end_row = start_row + height - 1;
    this.start_col = start_col;
    this.width = width;
    this.end_col = start_col + width - 1;
  }
  static from_range(range) {
    return new this(
      range.getRow(), range.getColumn(),
      range.getHeight(), range.getWidth() );
  }
  match(cfrange) {
    return (
      this.start_row <= cfrange.start_row &&
      this.end_row >= cfrange.end_row &&
      this.start_column <= cfrange.start_column &&
      this.end_column >= cfrange.end_column );
  }
  *[Symbol.iterator]() {
    // effectively allows supplying a CFRangeList objec
    // to CFRangeList.from_dimensions()
    yield this.start_row;
    yield this.start_col;
    yield this.height;
    yield this.width;
  }
}

class CFRangeList extends Array {
  static from_ranges(ranges) {
    return this.from(ranges.map(range => CFRange.from_range(range)));
  }
  static from_dimensions(ranges) {
    return this.from(ranges.map(range_dim => new CFRange(...range_dim)));
  }
  impose_on_rule_builder(builder, sheet) {
    builder.setRanges(this.map(cfrange => sheet.getRange(...cfrange)));
  }
  match(other) {
    return other.every( other_cfrange =>
      this.some( our_cfrange =>
        our_cfrange.match(other_cfrange)
      )
    );
  }
  split_match(other) {
    throw "not implemented";
  }
}

class CFEffect {
  static normalize_color(color) {
    if (color == null || color == "")
      return color;
    if (/#[0-9A-Ha-h]{6}/.exec(color) == null) {
      throw Error("non-standard form of color: " + color);
    }
    return color.toLowerCase();
  }
}

class CFBooleanEffect extends CFEffect {
  constructor({
    background, font_color,
    bold, italic, strikethrough, underline
  }) {
    super();
    this.background = this.constructor.normalize_color(background);
    this.font_color = this.constructor.normalize_color(font_color);
    this.bold = bold;
    this.italic = italic;
    this.strikethrough = strikethrough;
    this.underline = underline;
  }
  static from_condition(boolean_condition) {
    return new this({
      background: boolean_condition.getBackground(),
      font_color: boolean_condition.getFontColor(),
      bold: boolean_condition.getBold(),
      italic: boolean_condition.getItalic(),
      strikethrough: boolean_condition.getStrikethrough(),
      underline: boolean_condition.getUnderline(),
    });
  }
  impose_on_rule_builder(builder) {
    builder
      .setBackground(this.background)
      .setFontColor(this.font_color)
      .setBold(this.bold)
      .setItalic(this.italic)
      .setStrikethrough(this.strikethrough)
      .setUnderline(this.underline);
  }
  match(other) {
    if (!(other instanceof CFBooleanEffect)) {
      return false;
    }
    return (
      this.background == other.background &&
      this.font_color == other.font_color &&
      this.bold   == other.bold &&
      this.italic == other.italic &&
      this.strikethrough == other.strikethrough &&
      this.underline == other.underline
    );
  }
}

class CFGradientEffect extends CFEffect {
  constructor({min_color = "", mid_color = "", max_color = ""}) {
    super();
    this.min_color = this.constructor.normalize_color(min_color);
    this.mid_color = this.constructor.normalize_color(mid_color);
    this.max_color = this.constructor.normalize_color(max_color);
  }
  static from_condition(gradient_condition) {
    return new this({
      min_color: gradient_condition.getMinColor(),
      mid_color: gradient_condition.getMidColor(),
      max_color: gradient_condition.getMaxColor(),
    });
  }
  match(other) {
    if (!(other instanceof CFGradientEffect)) {
      return false;
    }
    return (
      this.min_color == other.min_color &&
      this.mid_color == other.mid_color &&
      this.max_color == other.max_color
    );
  }
}

class CFRule {
  constructor(cfcondition, cfranges, cfeffect) {
    if (cfcondition instanceof CFGradientCondition) {
      this.type = "gradient";
    } else {
      this.type = "boolean";
    }
    this.condition = cfcondition;
    this.ranges = cfranges;
    this.effect = cfeffect;
  }
  static from_rule(rule) {
    var cfcondition, cfranges, cfeffect;
    cfranges = CFRangeList.from_ranges(rule.getRanges());
    if (cfranges.length < 1)
      throw new ConditionalFormattingError("A rule has no ranges");
    const base_cell = [cfranges[0].start_row, cfranges[0].start_col];
    var boolean_condition = rule.getBooleanCondition();
    if (boolean_condition !== null) {
      cfcondition = CFBooleanCondition.from_condition(
        boolean_condition, (value) => {
          if (typeof value != "string")
            return value;
          if (!value.startsWith("="))
            return value;
          return Formula.translate( value,
            {notation: "A1", base_cell},
            {locale: "en", notation: "R1C1"} );
        } );
      cfeffect = CFBooleanEffect.from_condition(boolean_condition);
    } else {
      var gradient_condition = rule.getGradientCondition();
      cfcondition = CFGradientCondition.from_condition(
        gradient_condition, value => {
          if (typeof value != "string")
            return value;
          return Formula.translate( value,
            {notation: "A1", base_cell},
            {locale: "en", notation: "R1C1"} );
        });
      cfeffect = CFGradientEffect.from_condition(gradient_condition);
    }
    return new this(cfcondition, cfranges, cfeffect);
  }
  static from_object(cfrule_obj) {
    var cfcondition, cfranges, cfeffect;
    var { type: rule_type,
      condition, ranges, effect,
    } = cfrule_obj;
    cfranges = CFRangeList.from_dimensions(ranges);
    if (rule_type == "boolean") {
      cfcondition = new CFBooleanCondition(condition);
      cfeffect = new CFBooleanEffect(effect);
    } else if (rule_type == "gradient") {
      cfcondition = new CFGradientCondition(condition);
      cfeffect = new CFGradientEffect(effect);
    }
    return new this(cfcondition, cfranges, cfeffect);
  }
  to_rule(sheet) {
    var builder = SpreadsheetApp.newConditionalFormatRule();
    if (this.ranges.length < 1)
      throw new ConditionalFormattingError("CFRule object has no ranges");
    const base_cell = [this.ranges[0].start_row, this.ranges[0].start_col];
    this.ranges.impose_on_rule_builder(builder, sheet);
    if (this.condition instanceof CFBooleanCondition) {
      this.condition.impose_on_rule_builder(builder, (value) => {
        if (typeof value != "string")
          return value;
        if (!value.startsWith("="))
          return value;
        return Formula.translate( value,
          {locale: "en", notation: "R1C1"},
          {notation: "A1", base_cell} );
      });
      this.effect.impose_on_rule_builder(builder);
    } else if (this.condition instanceof CFGradientCondition) {
      function value_decode(value) {
        if (typeof value != "string")
          return value;
        return Formula.translate( value,
          {locale: "en", notation: "R1C1"},
          {notation: "A1", base_cell} );
      }
      builder.setGradientMinpointWithValue(
        this.effect.min_color,
        this.condition.min_type,
        value_decode(this.condition.min_value) );
      builder.setGradientMidpointWithValue(
        this.effect.mid_color,
        this.condition.mid_type,
        value_decode(this.condition.mid_value) );
      builder.setGradientMaxpointWithValue(
        this.effect.max_color,
        this.condition.max_type,
        value_decode(this.condition.max_value) );
    } else {
      throw new Error("internal error");
    }
    return builder.build();
  }
  match(cfrule_filter) {
    var {
      condition: cfcondition,
      ranges: cfranges,
      effect: cfeffect
    } = cfrule_filter;
    return ( this.condition.match(cfcondition) &&
      (cfranges == null || this.ranges.match(cfranges)) &&
      (cfeffect == null || this.effect.match(cfeffect)) );
  }
  split_match(cfrule_filter) {
    // Return [this, null] if there is no match;
    // return [null, this] if all ranges match
    //   (or there is match and filter does not specify ranges);
    // return [this, split] if some ranges match (modify this);
    var {
      condition: cfcondition,
      ranges: cfranges,
      effect: cfeffect
    } = cfrule_filter;
    if (!this.condition.match(cfcondition))
      return [this, null];
    if (cfeffect != null && !this.effect.match(cfeffect))
      return [this, null];
    if (cfranges == null)
      return [null, this];
    var [our_cfranges, split_cfranges] = this.ranges.split_match(cfranges);
    if (split_cfranges == null)
      return [this, null];
    if (our_cfranges == null)
      return [null, this];
    this.ranges = our_cfranges;
    var split = new this.constructor(
      this.condition, split_ranges, this.effect );
    return [this, split];
  }
}

class CFRuleFilter {
  // Essentially a partial CFRule object, for matching against.
  // Must contain a condition; ranges and effect are optional.
  constructor(cfcondition, cfranges, cfeffect) {
    this.condition = cfcondition;
    this.ranges = cfranges;
    this.effect = cfeffect;
  }
  static from_object(filter_obj) {
    var cfcondition, cfranges = null, cfeffect = null;
    var { type: rule_type,
      condition, ranges = null, effect = null,
    } = filter_obj;
    if (ranges != null)
      cfranges = CFRangeList.from_dimensions(ranges);
    if (rule_type == "boolean") {
      cfcondition = new CFBooleanCondition(condition);
      if (effect != null)
        cfeffect = new CFBooleanEffect(effect);
    } else if (rule_type == "gradient") {
      cfcondition = new CFGradientCondition(condition);
      if (effect != null)
        cfeffect = new CFGradientEffect(effect);
    }
    return new this(cfcondition, cfranges, cfeffect);
  }
}

class CFRuleList extends Array {
  static load(sheet) {
    var rules = sheet.getConditionalFormatRules();
    var cfrules = new this();
    cfrules.push(...rules.map(rule => CFRule.from_rule(rule)))
    return cfrules;
  }
  insert(cfrule_obj, ...before_filter_objs) {
    if (before_filter_objs.length > 0)
      throw "not implemented";
    var new_cfrule = CFRule.from_object(cfrule_obj);
    var cfrule_filter = CFRuleFilter.from_object(cfrule_obj);
    cfrule_filter.ranges = null;
    for (let cfrule of this) {
      if (!cfrule.match(cfrule_filter))
        continue;
      cfrule.ranges.push(...new_cfrule.ranges);
      return;
    }
    this.push(new_cfrule);
  }
  remove(filter_obj) {
    throw "not implemented";
  }
  replace(filter_obj, effect_obj) {
    throw "not implemented";
  }
  save(sheet) {
    sheet.setConditionalFormatRules(this.map(cfrule => cfrule.to_rule(sheet)));
  }
}

function merge(sheet, ...cfrule_objs) {
  var cfrules = CFRuleList.load(sheet);
//  for (let cfrule of cfrules) { // FIXME
//    console.log(JSON.stringify(cfrule));
//  }
//  console.log("All is well…");
  for (let cfrule_obj of cfrule_objs) {
//    console.log(JSON.stringify(CFRule.from_object(cfrule_obj))); // FIXME
    cfrules.insert(cfrule_obj);
  }
  cfrules.save(sheet);
}

return {
  merge: merge,
  RuleList: CFRuleList,
  Rule: CFRule,
};
}();

