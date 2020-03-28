class SpreadsheetError extends Error {
  constructor(message, range) {
    super();
    this.name = this.constructor.name;
    this.message = message;
    this.range = range;
  }

  toString() {
    var prefix = this.name;
    if (this.range) {
      prefix = prefix + " (" +
        this.range.getSheet().getName() + "!" +
        this.range.getA1Notation() +
      ")";
    }
    return prefix + ": " + this.message; 
  }
}
