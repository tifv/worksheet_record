/* Limitations:
 * * support for decimal fractions in formulas is not guaranteed;
 *   they should work as long as precision of neither JavaScript
 *   nor Google Sheets is tested.
 */

class FormulaError extends Error {};

var Formula = function() { // begin namespace

const known_locales = ["en", "ru"];
var default_locale = "en";

const known_notations = ["R1C1", "A1"];
var default_notation = "R1C1";

class Formula {
    static rectify_options(options) {
        var {
            locale = default_locale,
            notation = default_notation,
            base_cell = null,
        } = options || {};
        if (!known_locales.includes(locale)) {
            throw new Error("unknown locale " + locale);
        }
        if (!known_notations.includes(notation)) {
            throw new Error("invalid notation " + notation);
        }
        if (notation == "A1" && base_cell == null) {
            throw new Error(
                "working with A1 notation requires base_cell option" );
        }
        return {locale: locale, notation: notation, base_cell: base_cell};
    }
    constructor(formula, options) {
        this.tokens = Array.from(this.constructor.tokenize(
            formula, this.constructor.rectify_options(options) ));
    }
    static *tokenize(formula, options) {
        var index = 0, length = formula.length;
        try {
            var context = "toplevel", contexts = [context];
                // "toplevel" or "arguments" or "expression" or "array"
            while (index < length) {
                let subformula = formula.substring(index);
                let tokenmatch =
                    WhitespaceToken.match(subformula, context, options) ||
                    BraketToken    .match(subformula, context, options) ||
                    DelimiterToken .match(subformula, context, options) ||
                    OperatorToken  .match(subformula, context, options) ||
                    StringToken    .match(subformula, context, options) ||
                    NumberToken    .match(subformula, context, options) ||
                    RangeRefToken  .match(subformula, context, options) ||
                    CellRefToken   .match(subformula, context, options) ||
                    ForeignRefToken.match(subformula, context, options) ||
                    FunctionToken  .match(subformula, context, options) ||
                    IdentifierToken.match(subformula, context, options) ||
                    null;
                if (tokenmatch == null)
                    throw new FormulaError("unrecognized token");
                let [size, ...tokens] = tokenmatch;
                for (let token of tokens) {
                    if (token instanceof OpeningBraketToken) {
                        contexts.push(context = token.start_context)
                    } else if (token instanceof ClosingBraketToken) {
                        if (contexts.pop() != token.end_context) {
                            throw new FormulaError("incorrect closing token");
                        }
                        context = contexts[contexts.length-1];
                    }
                    yield token;
                }
                index += size;
            }
            if (contexts.length != 1) {
                throw new FormulaError("unbalanced parentheses");
            }
        } catch (error) {
            if (error instanceof FormulaError) {
                error.message += (
                    " (index " + index + " of formula " + formula + ")" );
            }
            throw error;
        }
    }
    toString(options) {
        options = this.constructor.rectify_options(options);
        var pieces = [];
        for (let token of this.tokens) {
            pieces.push(token.toString(options));
        }
        return pieces.join("");
    }
    static translate(formula, from_options, to_options) {
      return (new Formula(formula, from_options)).toString(to_options);
    }
}

class Token {
    static match(subformula, context, options) {
        var match = subformula.match(this.regex);
        if (match == null)
            return null;
        if (match.index > 0)
            throw new FormulaError("internal error");
        return [ match[0].length,
            this.match_constructor(match, context, options) ];
    }
    static match_constructor(match, context, options) {
        return new this();
    }
    toString(options) {
        throw new Error( "abstract method " +
          this.constructor.name + "." + "toString");
    }
}

class CommonToken extends Token {
    constructor(string) {
        super();
        this.string = string;
    }
    static match_constructor(match, context, options) {
        return new this(match[0]);
    }
    toString(options) {
        return this.string;
    }
}

class LocalizedToken extends Token {
    static match(subformula, context, options) {
        const {locale} = options;
        var match = subformula.match(this["regex_" + locale]);
        if (match == null)
            return null;
        if (match.index > 0)
            throw new FormulaError("internal error");
        return [ match[0].length,
            this["match_constructor_" + locale](match, context, options) ];
    }
    static match_constructor_en(match, context, options) {
        throw new Error( "abstract method " +
          this.constructor.name + "." + "match_constructor_en");
    }
    static match_constructor_ru(match, context, options) {
        throw new Error( "abstract method " +
          this.constructor.name + "." + "match_constructor_ru");
    }
    toString(options) {
        const {locale} = options;
        return this["toString_" + locale](options);
    }
}

class WhitespaceToken extends CommonToken {};
WhitespaceToken.regex = /^\s+/;

class BraketToken extends Token {
    static match_constructor(match, context, options) {
        var string = match[0];
        if (string == "(") {
            return new OpeningParenToken(string);
        } else if (string == ")") {
            if (context == "arguments") {
                return new ClosingArgsToken(string);
            }
            return new ClosingParenToken(string);
        } else if (string == "{") {
            return new OpeningBraceToken(string);
        } else if (string == "}") {
            return new ClosingBraceToken(string);
        } else {
            throw new FormulaError("internal error");
        }
    }
    toString(options) {
        return this.braket_s;
    }
};
BraketToken.regex = /^[(){}]/;

class DelimiterToken extends LocalizedToken {
    static match_constructor_en(match, context, options) {
        if (match[0] == ",") {
            if (context == "arguments") {
                return new ArgDelimiterToken(match[0]);
            } else if (context == "array") {
                return new ColumnDelimiterToken(match[0]);
            } else {
                throw new FormulaError( "out-of-place delimiter " +
                    "(context=" + context + ")" );
            }
        } else if (match[0] == ";") {
            if (context == "arguments") {
                return new ArgDelimiterToken(match[0]);
            } else if (context == "array") {
                return new RowDelimiterToken(match[0]);
            } else {
                throw new FormulaError( "out-of-place delimiter " +
                    "(context=" + context + ")" );
            }
        } else {
            throw new FormulaError("invalid delimiter");
        }
    }
    static match_constructor_ru(match, context, options) {
        if (match[0] == ";") {
            if (context == "arguments") {
                return new ArgDelimiterToken(match[0]);
            } else if (context == "array") {
                return new RowDelimiterToken(match[0]);
            } else {
                throw new FormulaError( "out-of-place delimiter " +
                    "(context=" + context + ")" );
            }
        } else if (match[0] == "\\") {
            if (context == "array") {
                return new ColumnDelimiterToken(match[0]);
            } else {
                throw new FormulaError( "out-of-place delimiter " +
                    "(context=" + context + ")" );
            }
        } else {
            throw new FormulaError("invalid delimiter");
        }
    }
    toString_en(options) {
        return this.delimiter_s_en;
    }
    toString_ru(options) {
        return this.delimiter_s_ru;
    }
};
DelimiterToken.regex_en = /^[,;]/;
DelimiterToken.regex_ru = /^[;\\]/;

class OperatorToken extends CommonToken {};
OperatorToken.regex = /^(?:==|<=|>=|<>|[+\-*\/^=<>&:])/;

class NotationalToken extends Token {
    static match(subformula, context, options) {
        const {notation} = options;
        var match = subformula.match(this["regex_" + notation]);
        if (match == null)
            return null;
        if (match.index > 0)
            throw new FormulaError("internal error");
        return [ match[0].length,
            this["match_constructor_" + notation](match, context, options) ];
    }
    static match_constructor_R1C1(match, context, options) {
        throw new Error( "abstract method " +
          this.constructor.name + "." + "match_constructor_R1C1");
    }
    static match_constructor_A1(match, context, options) {
        throw new Error( "abstract method " +
          this.constructor.name + "." + "match_constructor_A1");
    }
    toString(options) {
        const {notation} = options;
        return this["toString_" + notation](options);
    }
}

function decode_R1C1_index(index_s) {
    if (index_s == null)
        return null;
    let index_type = index_s[0];
    if (index_type != "R" && index_type != "C") {
        throw new FormulaError("internal error: " + index_type);
    }
    index_s = index_s.substring(1);
    if (index_s == "")
        return {relative: 0};
    var relative_match = /^\[(-?\d+)\]$/.exec(index_s);
    if (relative_match != null) {
        return {relative: parseInt(relative_match[1], 10)};
    }
    if (/^\d+$/.exec(index_s)) {
        return {absolute: parseInt(index_s, 10)};
    }
    throw new FormulaError( "internal error: " +
        index_s + " is not a valid R1C1 index" );
}

function encode_R1C1_index(prefix, index) {
    if (index == null)
        return "";
    if (index.relative != null) {
        //if (index.relative == 0)
        //    return prefix;
        return prefix + "[" + index.relative.toString() + "]";
    } else if (index.absolute != null) {
        return prefix + index.absolute.toString();
    } else {
        throw new FormulaError("internal error");
    }
}

function encode_R1C1_row_index(index) {
    return encode_R1C1_index("R", index);
}

function encode_R1C1_col_index(index) {
    return encode_R1C1_index("C", index);
}

function decode_A1_row_index(index_s, ref_row) {
    if (index_s == null)
        return null;
    if (/^\$[0-9]+$/.exec(index_s)) {
        return {absolute: parseInt(index_s.substring(1), 10)};
    }
    if (/^[0-9]+$/.exec(index_s)) {
        return {relative: parseInt(index_s, 10) - ref_row};
    }
    throw new FormulaError( "internal error: " +
        index_s + " is not a valid A1 row index");
}

function decode_A1_col_index(index_s, ref_col) {
    if (index_s == null)
        return null;
    if (/^\$[A-Za-z]+$/.exec(index_s)) {
        return {absolute: ACodec.decode(index_s.substring(1))};
    }
    if (/^[A-Za-z]+$/.exec(index_s)) {
        return {relative: ACodec.decode(index_s) - ref_col};
    }
    throw new FormulaError( "internal error: " +
        index_s + " is not a valid A1 column index");
}

function encode_A1_row_index(index, ref_row) {
    if (index == null)
        return "";
    if (index.relative != null) {
        return (ref_row + index.relative).toString();
    } else if (index.absolute != null) {
        return "$" + index.absolute.toString();
    } else {
        throw new FormulaError("internal error");
    }
}

function encode_A1_col_index(index, ref_col) {
    if (index == null)
        return "";
    if (index.relative != null) {
        return ACodec.encode(ref_col + index.relative);
    } else if (index.absolute != null) {
        return "$" + ACodec.encode(index.absolute);
    } else {
        throw new FormulaError("internal error");
    }
}

class RangeRefToken extends NotationalToken {
    constructor(start_row, start_col, end_row, end_col) {
        super();
        if (start_row == null && end_row != null) {
            throw new FormulaError( "start row cannot be unbounded " +
                "unless end row is" );
        }
        if (start_col == null && end_col != null) {
            throw new FormulaError( "start column cannot be unbounded " +
                "unless end column is" );
        }
        if (end_row == null && end_col == null) {
            throw new FormulaError("cannot have unbounded both row and column");
        }
        this.start_row = start_row;
        this.start_col = start_col;
        this.end_row = end_row;
        this.end_col = end_col;
    }
    static match_constructor_R1C1(match, context, options) {
        var start_row = decode_R1C1_index(match[1]);
        var start_col = decode_R1C1_index(match[2]);
        var end_row = decode_R1C1_index(match[3]);
        var end_col = decode_R1C1_index(match[4]);
        return new this(start_row, start_col, end_row, end_col);
    }
    static match_constructor_A1(match, context, options) {
        const {base_cell: [base_row, base_col]} = options;
        var start_col = decode_A1_col_index(match[1], base_col);
        var start_row = decode_A1_row_index(match[2], base_row);
        var end_col = decode_A1_col_index(match[3], base_col);
        var end_row = decode_A1_row_index(match[4], base_row);
        return new this(start_row, start_col, end_row, end_col);
    }
    toString_R1C1(options) {
        return (
            encode_R1C1_row_index(this.start_row) +
            encode_R1C1_col_index(this.start_col) +
            ":" +
            encode_R1C1_row_index(this.end_row) +
            encode_R1C1_col_index(this.end_col) );
    }
    toString_A1(options) {
        const {base_cell: [base_row, base_col]} = options;
        return (
            encode_A1_col_index(this.start_col, base_col) +
            encode_A1_row_index(this.start_row, base_row) +
            ":" +
            encode_A1_col_index(this.end_col, base_col) +
            encode_A1_row_index(this.end_row, base_row) );
    }
};
RangeRefToken.regex_R1C1 = new RegExp( "^" +
    "(?!:)" +
    "(R(?:\\d+|\\[-?\\d+\\]|))?(C(?:\\d+|\\[-?\\d+\\]|))?:" +
    "(R(?:\\d+|\\[-?\\d+\\]|))?(C(?:\\d+|\\[-?\\d+\\]|))?" +
    "(?<!:)" );
RangeRefToken.regex_A1 = new RegExp( "^" +
    "(?!:)" +
    "(\\$?[A-Za-z]+)?(\\$?[0-9]+)?:" +
    "(\\$?[A-Za-z]+)?(\\$?[0-9]+)?" +
    "(?<!:)" );

class CellRefToken extends NotationalToken {
    constructor(row, col) {
        super();
        this.row = row;
        this.col = col;
    }
    static match_constructor_R1C1(match, context, options) {
        var row = decode_R1C1_index(match[1]);
        var col = decode_R1C1_index(match[2]);
        if (row == null || row == null) {
            throw new FormulaError("internal error");
        }
        return new this(row, col);
    }
    static match_constructor_A1(match, context, options) {
        const {base_cell: [base_row, base_col]} = options;
        var col = decode_A1_col_index(match[1], base_col);
        var row = decode_A1_row_index(match[2], base_row);
        if (row == null || row == null) {
            throw new FormulaError("internal error");
        }
        return new this(row, col);
    }
    toString_R1C1(options) {
        return (
            encode_R1C1_row_index(this.row) +
            encode_R1C1_col_index(this.col) );
    }
    toString_A1(options) {
        const {base_cell: [base_row, base_col]} = options;
        return (
            encode_A1_col_index(this.col, base_col) +
            encode_A1_row_index(this.row, base_row) );
    }
};
CellRefToken.regex_R1C1 = new RegExp( "^" +
    "(R(?:\\d+|\\[-?\\d+\\]|))(C(?:\\d+|\\[-?\\d+\\]|))"
);
CellRefToken.regex_A1 = new RegExp( "^" +
    "(\\$?[A-Za-z]+)(\\$?[0-9]+)"
);

class ForeignRefToken extends NotationalToken {
    constructor(sheet_name, range_ref) {
        super();
        this.sheet_name = sheet_name;
        this.range_ref = range_ref;
    }
    static match_constructor_R1C1(match, context, options) {
        var [ , // match[0]
            sheet_name_s,
            range_ref_s,
            range_ref_start_row, range_ref_start_col,
            range_ref_end_row, range_ref_end_col,
            cell_ref_s,
            cell_ref_row, cell_ref_col
        ] = match;
        var sheet_name = this.parse_sheet_name(sheet_name_s);
        var range_ref;
        if (range_ref_s != null) {
            range_ref = RangeRefToken.match_constructor_R1C1([ range_ref_s,
                range_ref_start_row, range_ref_start_col,
                range_ref_end_row, range_ref_end_col ]);
        } else if (cell_ref_s != null) {
            range_ref = CellRefToken.match_constructor_R1C1([ cell_ref_s,
                cell_ref_row, cell_ref_col ]);
        }
        return new this(sheet_name, range_ref);
    }
    static match_constructor_A1(match, context, options) {
        var [ , // match[0]
            sheet_name_s,
            range_ref_s,
            range_ref_start_col, range_ref_start_row,
            range_ref_end_col, range_ref_end_row,
            cell_ref_s,
            cell_ref_col, cell_ref_row,
        ] = match;
        var sheet_name = this.parse_sheet_name(sheet_name_s);
        var range_ref;
        if (range_ref_s != null) {
            range_ref = RangeRefToken.match_constructor_A1([ range_ref_s,
                range_ref_start_col, range_ref_start_row,
                range_ref_end_col, range_ref_end_row
            ], context, options);
        } else if (cell_ref_s != null) {
            range_ref = CellRefToken.match_constructor_A1([ cell_ref_s,
                cell_ref_row, cell_ref_col
            ], context, options);
        }
        return new this(sheet_name, range_ref);
    }
    static parse_sheet_name(sheet_name_s) {
        if (sheet_name_s.startsWith("'")) {
            return sheet_name_s.substring(1, sheet_name_s.length - 1)
                .replace(/''/, "'");
        } else {
            return sheet_name_s;
        }
    }
    format_sheet_name() {
        if (/^[A-Za-z_][A-Za-z0-9_.]*$/.exec(this.sheet_name)) {
            return this.sheet_name;
        } else {
            return "'" + this.sheet_name.replace(/'/g, "''") + "'";
        }
    }
    toString(options) {
        return this.format_sheet_name() + "!" +
            this.range_ref.toString(options);
    }
};
ForeignRefToken.regex_R1C1 = new RegExp(
    "^([A-Za-z_][A-Za-z0-9_.]*|'(?:[^']|'')*'(?!'))!" +
    "(?:(" +
        "(?!:)" +
        "(R(?:\\d+|\\[-?\\d+\\]|))?(C(?:\\d+|\\[-?\\d+\\]|))?:" +
        "(R(?:\\d+|\\[-?\\d+\\]|))?(C(?:\\d+|\\[-?\\d+\\]|))?" +
        "(?<!:)" +
    ")|(" +
        "(R(?:\\d+|\\[-?\\d+\\]|))(C(?:\\d+|\\[-?\\d+\\]|))" +
    "))" );
ForeignRefToken.regex_A1 = new RegExp(
    "^([A-Za-z_][A-Za-z0-9_.]*|'(?:[^']|'')*'(?!'))!" +
    "(?:(" +
        "(?!:)" +
        "(\\$?[A-Za-z]+)?(\\$?[0-9]+)?:" +
        "(\\$?[A-Za-z]+)?(\\$?[0-9]+)?" +
        "(?<!:)" +
    ")|(" +
        "(\\$?[A-Za-z]+)(\\$?[0-9]+)" +
    "))" );

class NumberToken extends LocalizedToken {
    constructor(value) {
        super();
        this.value = value;
    }
    static match_constructor_en(match, context, options) {
        return new this(parseFloat(match[0]));
    }
    static match_constructor_ru(match, context, options) {
        return new this(parseFloat(match[0].replace(",", ".")));
    }
    toString_en() {
        return this.value.toString();
    }
    toString_ru() {
        return this.value.toString().replace(".", ",");
    }
};
NumberToken.regex_en = /^(?:\d+\.\d+|\d+\.|\.\d+|\d+)(?:[Ee][+\-]\d+)?/;
NumberToken.regex_ru = /^(?:\d+,\d+|\d+,|,\d+|\d+)(?:[Ee][+\-]\d+)?/;

class StringToken extends Token {
    constructor(value) {
        super();
        this.value = value;
    }
    static match_constructor(match) {
        var value_s = match[0];
        return new this(
            value_s.substring(1, value_s.length - 1)
                .replace(/""/g, '"') );
    }
    toString() {
        return '"' + this.value.replace(/"/g, '""') + '"';
    }
};
StringToken.regex = /^"(?:[^"]|"")*"(?!")/;

class IdentifierToken extends CommonToken {};
IdentifierToken.regex = /^[A-Za-z_][A-Za-z0-9_.]*/;

class FunctionToken extends IdentifierToken {
    static match(subformula, context, options) {
        var match = subformula.match(/^([A-Za-z_][A-Za-z0-9_.]*)(\s*)\(/);
        if (match == null)
            return null;
        if (match.index > 0)
            throw new FormulaError("internal error");
        var tokens = [new this(match[1])];
        if (match[2] != "")
            tokens.push(new WhitespaceToken(match[2]));
        tokens.push(new OpeningArgsToken());
        return [match[0].length, ...tokens];
    }
};

class OpeningBraketToken extends BraketToken {};
class ClosingBraketToken extends BraketToken {};

class OpeningParenToken extends OpeningBraketToken {};
OpeningParenToken.prototype.start_context = "expression";
OpeningParenToken.prototype.braket_s = "(";

class ClosingParenToken extends ClosingBraketToken {};
ClosingParenToken.prototype.end_context = "expression";
ClosingParenToken.prototype.braket_s = ")";

class OpeningArgsToken extends OpeningBraketToken {};
OpeningArgsToken.prototype.start_context = "arguments";
OpeningArgsToken.prototype.braket_s = "(";

class ClosingArgsToken extends ClosingBraketToken {};
ClosingArgsToken.prototype.end_context = "arguments";
ClosingArgsToken.prototype.braket_s = ")";

class OpeningBraceToken extends OpeningBraketToken {};
OpeningBraceToken.prototype.start_context = "array";
OpeningBraceToken.prototype.braket_s = "{";

class ClosingBraceToken extends ClosingBraketToken {};
ClosingBraceToken.prototype.end_context = "array";
ClosingBraceToken.prototype.braket_s = "}";

class ArgDelimiterToken extends DelimiterToken {};
ArgDelimiterToken.prototype.delimiter_s_en = ",";
ArgDelimiterToken.prototype.delimiter_s_ru = ";";

class RowDelimiterToken extends DelimiterToken {};
RowDelimiterToken.prototype.delimiter_s_en = ";";
RowDelimiterToken.prototype.delimiter_s_ru = ";";

class ColumnDelimiterToken extends DelimiterToken {};
ColumnDelimiterToken.prototype.delimiter_s_en = ",";
ColumnDelimiterToken.prototype.delimiter_s_ru = "\\";

Formula.tokens = {
    Whitespace : WhitespaceToken,
    Operator   : OperatorToken,
    RangeRef   : RangeRefToken,
    CellRef    : CellRefToken,
    Number     : NumberToken,
    String     : StringToken,
    Identifier : IdentifierToken,
    ForeignRef : ForeignRefToken,
    OpeningParen    : OpeningParenToken,
    ClosingParen    : ClosingParenToken,
    OpeningArgs     : OpeningArgsToken,
    ClosingArgs     : ClosingArgsToken,
    OpeningBrace    : OpeningBraceToken,
    ClosingBrace    : ClosingBraceToken,
    ArgDelimiter    : ArgDelimiterToken,
    RowDelimiter    : RowDelimiterToken,
    ColumnDelimiter : ColumnDelimiterToken,
};

return Formula;
}(); // end Formula namespace

// vim: set fdm=marker sw=4 :
