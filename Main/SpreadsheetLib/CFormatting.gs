/* Possible actions:
 * • find all rules that:
 *    (a) have given condition;
 *    (b) (optional) cover given range;
 *    (c) (optional) have given effect;
 * • find existing rule with (ac) and extend it to given range or
 *   create a new rule that covers this range;
 *   • ensure that this new rule comes before any other rule
 *     specified with (a);
 * • find all rules with (ab) or (abc) and remove them from given range;
 *   • replace them;
 * • find all rules with (a) and replace them with similar rule
 *   that covers the same ranges and has given effect;
 */

class ConditionalFormattingError extends Error {};

var ConditionalFormatting = function() {

class CFCondition {};

class CFBooleanCondition extends CFCondition {
    constructor({type, values}) {
        super();
        this.type = type;
        this.values = Array.from(values);
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
    match(cfcondition) {
        if (!(cfcondition instanceof CFBooleanCondition)) {
            return false;
        }
        var filter_values = this.values;
        return (
            cfcondition.type == this.type &&
            cfcondition.values.length == filter_values.length &&
            cfcondition.values.every( (element, index) =>
                (element == filter_values[index]) )
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
    match(cfcondition) {
        if (!(cfcondition instanceof CFGradientCondition)) {
            return false;
        }
        return (
            cfcondition.min_type == this.min_type &&
                cfcondition.min_value == this.min_value &&
            cfcondition.mid_type == this.mid_type &&
                cfcondition.mid_value == this.mid_value &&
            cfcondition.max_type == this.max_type &&
                cfcondition.max_value == this.max_value
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
    subset(cfrange) {
        return (
            this.start_row >= cfrange.start_row &&
            this.end_row <= cfrange.end_row &&
            this.start_col >= cfrange.start_col &&
            this.end_col <= cfrange.end_col );
    }
    superset(cfrange) {
        return (
            this.start_row <= cfrange.start_row &&
            this.end_row >= cfrange.end_row &&
            this.start_col <= cfrange.start_col &&
            this.end_col >= cfrange.end_col );
    }
    detract_from(cfrange) {
        if (!this.subset(cfrange))
            return null;
        var rest = [];
        if (cfrange.start_row < this.start_row) {
            rest.push(new cfrange.constructor(
                cfrange.start_row, cfrange.start_col,
                this.start_row - cfrange.start_row, cfrange.width ));
        }
        if (cfrange.end_row > this.end_row) {
            rest.push(new cfrange.constructor(
                this.end_row + 1, cfrange.start_col,
                cfrange.end_row - this.end_row, cfrange.width ));
        }
        if (cfrange.start_col < this.start_col) {
            rest.push(new cfrange.constructor(
                this.start_row, cfrange.start_col,
                this.height, this.start_col - cfrange.start_col ));
        }
        if (cfrange.end_col > this.end_col) {
            rest.push(new cfrange.constructor(
                this.start_row, this.end_col + 1,
                this.height, cfrange.end_col - this.end_col ));
        }
        return rest;
    }
    *[Symbol.iterator]() {
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
    match(cfranges) {
        return this.every( filter_range =>
            cfranges.some( cfrange =>
                filter_range.subset(cfrange)
            )
        );
    }
    split_match(cfranges) {
        // Return
        //     [copy, null ] if there are no intersections;
        //     [null, copy ] if we have all cfranges' ranges;
        //     [rest, split] if some our ranges are contained
        //         in cfranges' ranges (they go in split).
        cfranges = cfranges.constructor.from(cfranges);
        var filter_ranges = Array.from(this);
        var split_cfranges = new cfranges.constructor();
        for (let i = 0; i < cfranges.length; ++i) {
            let cfrange = cfranges[i];
            for (let j = 0; j < filter_ranges.length; ++j) {
                let filter_range = filter_ranges[j];
                let rest = filter_range.detract_from(cfrange);
                if (rest == null)
                    continue;
                cfranges.splice(i, 1, ...rest);
                --i;
                split_cfranges.push(filter_range);
                filter_ranges.splice(j, 1);
                break;
            }
        }
        if (split_cfranges.length == 0)
            return [cfranges, null];
        if (cfranges.length == 0) {
            return [null, split_cfranges];
        }
        return [cfranges, split_cfranges];
    }
}

class CFLocationList extends CFRangeList {
    match(cfranges) {
        return this.some( filter_range =>
            cfranges.some( cfrange =>
                filter_range.superset(cfrange)
            )
        );
    }
    split_match(cfranges) {
        // Return
        //     [copy, null ] if there are no intersections,
        //         where copy is cfranges' copy;
        //     [null, copy ] if we superset all cfranges' ranges;
        //     [rest, split] if some our ranges superset
        //         some cfranges' ranges.
        cfranges = cfranges.constructor.from(cfranges);
        var split_cfranges = new cfranges.constructor();
        for (let i = 0; i < cfranges.length; ++i) {
            let cfrange = cfranges[i];
            for (let j = 0; j < this.length; ++j) {
                let filter_range = this[j];
                if (!filter_range.superset(cfrange))
                    continue;
                cfranges.splice(i, 1);
                --i;
                split_cfranges.push(cfrange);
                break;
            }
        }
        if (split_cfranges.length == 0)
            return [cfranges, null];
        if (cfranges.length == 0) {
            return [null, split_cfranges];
        }
        return [cfranges, split_cfranges];
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
    match(cfeffect) {
        if (!(cfeffect instanceof CFBooleanEffect)) {
            return false;
        }
        return (
            cfeffect.background == this.background &&
            cfeffect.font_color == this.font_color &&
            cfeffect.bold   == this.bold &&
            cfeffect.italic == this.italic &&
            cfeffect.strikethrough == this.strikethrough &&
            cfeffect.underline == this.underline
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
    match(cfeffect) {
        if (!(cfeffect instanceof CFGradientEffect)) {
            return false;
        }
        return (
            cfeffect.min_color == this.min_color &&
            cfeffect.mid_color == this.mid_color &&
            cfeffect.max_color == this.max_color
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
        } else {
            throw new Error("invalid rule type");
        }
        return new this(cfcondition, cfranges, cfeffect);
    }
    to_rule(sheet) {
        var builder = SpreadsheetApp.newConditionalFormatRule();
        if (this.ranges.length < 1)
            throw new ConditionalFormattingError(
              "CFRule object has no ranges" );
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
}

class CFRuleFilter {
    // Essentially a partial CFRule object, for matching against.
    // Must contain a condition; ranges, locations, and effect are optional.
    // locations are like ranges, but to match they must be a superset
    //   of the target range, not a subset.
    constructor(cfcondition, cfranges, cfeffect, cflocations) {
        if (cfcondition instanceof CFGradientCondition) {
            this.type = "gradient";
        } else {
            this.type = "boolean";
        }
        if (cfranges != null && cflocations != null) {
            throw new Error(
                "filter can only have ranges or locations, not both" );
        }
        this.condition = cfcondition;
        this.ranges = cfranges;
        this.effect = cfeffect;
        this.locations = cflocations;
    }
    static from_object(filter_obj) {
        var cfcondition, cfranges = null, cfeffect = null, cflocations = null;
        var { type: rule_type,
            condition, ranges = null, effect = null,
            locations = null
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
        } else {
            throw new Error("invalid rule type");
        }
        if (locations != null)
            cflocations = CFLocationList.from_dimensions(locations);
        return new this(cfcondition, cfranges, cfeffect, cflocations);
    }
    static convert_object(filter_obj) {
        if (filter_obj instanceof this) {
            return filter_obj;
        } else {
            return this.from_object(filter_obj);
        }
    }
    match(cfrule) {
        var {
            condition: cfcondition,
            ranges: cfranges,
            effect: cfeffect
        } = cfrule;
        var filter_ranges = this.ranges || this.locations;
        return ( this.condition.match(cfcondition) &&
            (filter_ranges == null || filter_ranges.match(cfranges)) &&
            (this.effect == null || this.effect.match(cfeffect))
        );
    }
    split_match(cfrule) {
        // Return
        //     [copy, null ] if there is no match;
        //     [null, copy ] if all ranges match
        //         (or if there is match and we do not specify ranges);
        //     [rest, split] if some ranges match;
        var {
            condition: cfcondition,
            ranges: cfranges,
            effect: cfeffect
        } = cfrule;
        if (!this.condition.match(cfcondition))
            return [cfrule, null];
        if (this.effect != null && !this.effect.match(cfeffect))
            return [cfrule, null];
        var filter_ranges = this.ranges || this.locations;
        if (filter_ranges == null)
            return [null, cfrule];
        var [rest_cfranges, split_cfranges] = filter_ranges.split_match(cfranges);
        if (split_cfranges == null)
            return [cfrule, null];
        if (rest_cfranges == null)
            return [null, cfrule];
        var rest_cfrule = new cfrule.constructor(
            new cfcondition.constructor(cfcondition),
            rest_cfranges,
            new cfeffect.constructor(cfeffect) );
        var split_cfrule = new cfrule.constructor(
            new cfcondition.constructor(cfcondition),
            split_cfranges,
            new cfeffect.constructor(cfeffect) );
        return [rest_cfrule, split_cfrule];
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
        // Note: if the new rule is merged into existing rule,
        // before_filters are not checked.
        var new_cfrule = CFRule.from_object(cfrule_obj);
        var cfrule_filter = CFRuleFilter.from_object(new_cfrule);
        cfrule_filter.ranges = null;
        for (let cfrule of this) {
            if (!cfrule_filter.match(cfrule))
                continue;
            cfrule.ranges.push(...new_cfrule.ranges);
            return;
        }
        var splice_index = this.length;
        for (let before_filter_obj of before_filter_objs) {
            let before_filter = CFRuleFilter.convert_object(before_filter_obj);
            for (let [index, cfrule] of this.entries()) {
                if (!before_filter.match(cfrule))
                    continue;
                splice_index = splice_index > index ? index : splice_index;
            }
        }
        this.splice(splice_index, 0, new_cfrule);
    }
    remove(filter_obj) {
        var cfrule_filter = CFRuleFilter.convert_object(filter_obj);
        for (let i = 0; i < this.length; ++i) {
            let cfrule = this[i];
            [cfrule,] = cfrule_filter.split_match(cfrule);
            if (cfrule != null) {
                this[i] = cfrule;
            } else {
                this.splice(i, 1);
            }
        }
    }
    replace(filter_obj, effect_obj) {
        var cfrule_filter = CFRuleFilter.convert_object(filter_obj);
        var cfeffect;
        if (cfrule_filter.type == "boolean") {
            cfeffect = new CFBooleanEffect(effect_obj);
        } else if (cfrule_filter.type == "gradient") {
            cfeffect = new CFGradientEffect(effect_obj);
        } else {
            throw new Error("invalid rule type " + cfrule_filter.type);
        }
        var old_cfrules = Array.from(this);
        this.length = 0;
        var new_cfrules = new Array();
        for (let cfrule of old_cfrules) {
            let [rest_cfrule, split_cfrule] = cfrule_filter.split_match(cfrule);
            if (rest_cfrule != null)
                this.push(rest_cfrule);
            if (split_cfrule != null) {
                split_cfrule.effect = new cfeffect.constructor(cfeffect);
                new_cfrules.push([this.length, split_cfrule]);
            }
        }
        var splice_shift = 0;
        merge_new_cfrules:
        for (let [splice_index, new_cfrule] of new_cfrules) {
            let new_cfrule_filter = CFRuleFilter.from_object(new_cfrule);
            new_cfrule_filter.ranges = null;
            for (let cfrule of this) {
                if (!new_cfrule_filter.match(cfrule))
                    continue;
                cfrule.ranges.push(...new_cfrule.ranges);
                continue merge_new_cfrules;
            }
            this.splice(splice_index + (splice_shift++), 0, new_cfrule);
        }
    }
    save(sheet) {
        sheet.setConditionalFormatRules(
            this.map(cfrule => cfrule.to_rule(sheet)) );
    }
    save_find_error(sheet) {
        try {
            sheet.setConditionalFormatRules(
                this.map(cfrule => cfrule.to_rule(sheet)) );
        } catch (error) {
            while (this.length > 0) {
                let rogue_cfrule = this.pop();
                try {
                    sheet.setConditionalFormatRules(
                        this.map(cfrule => cfrule.to_rule(sheet)) );
                    console.error(JSON.stringify(rogue_cfrule));
                    break;
                } catch (another_error) {
                }
            }
            throw error;
        }
    }
}
//CFRuleList.prototype.save = CFRuleList.prototype.save_find_error;

function merge(sheet, ...cfrule_objs) {
    var cfrules = CFRuleList.load(sheet);
    for (let cfrule_obj of cfrule_objs) {
        cfrules.insert(cfrule_obj);
    }
    cfrules.save(sheet);
}

return {
    merge: merge,
    RuleList:   CFRuleList,
    Rule: CFRule, RuleFilter: CFRuleFilter,
    RangeList: CFRangeList, LocationList: CFLocationList,
    Range: CFRange,
    BooleanCondition: CFBooleanCondition, GradientCondition: CFGradientCondition,
    BooleanEffect: CFBooleanEffect, GradientEffect: CFGradientEffect,
};
}();

// vim: set fdm=marker sw=4 :
