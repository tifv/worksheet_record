var WorksheetDate = function() { // begin namespace

function WorksheetDate(year, month, day, period) {
  this.year = year;
  this.month = month;
  this.day = day;
  this.period = period;
}
WorksheetDate.from_object = function(date_obj) {
  return new WorksheetDate(date_obj.year, date_obj.month, date_obj.day, date_obj.period);
};

const date_regex = /^(\d{4,})-(\d{2,})-(\d{2,})$/;
const period_regex = /^(\d{4,})-(\d{2,})-(\d{2,})(?:[ -]p(\d{1,}))?$/;
const period_regex_m = new RegExp(period_regex.source, 'm');

WorksheetDate.parse = function(string, regex = period_regex) {
  var match = regex.exec(string);
  if (match == null)
    return null;
  var
    year = parseInt(match[1], 10),
    month = parseInt(match[2], 10),
    day = parseInt(match[3], 10),
    period = (match[4] != null) ? parseInt(match[4], 10) : null;
  return new WorksheetDate(year, month, day, period);
};

WorksheetDate.find = function(string) {
  return WorksheetDate.parse(string, period_regex_m);
};

WorksheetDate.parse_date = function(string) {
  return WorksheetDate.parse(string, date_regex);
};

WorksheetDate.prototype.format = function({filename = false} = {}) {
  var string =
    this.year.toString().padStart(4, "0") + '-' +
    this.month.toString().padStart(2, "0") + '-' +
    this.day.toString().padStart(2, "0");
  if (this.period != null) {
    if (filename)
      string += "-";
    else
      string += " ";
    string += "p" + this.period;
  }
  return string;
};

WorksheetDate.prototype.to_object = function() {
  return {year: this.year, month: this.month, day: this.day, period: this.period};
};

WorksheetDate.today = function(delta) {
  var today = new Date();
  if (delta != null)
    today.setDate(today.getDate() + delta);
  var year = today.getFullYear(), month = (today.getMonth()+1), day = today.getDate();
  return new WorksheetDate(year, month, day);
};

WorksheetDate.prototype.compare = function(other) {
  if (this.year < other.year)
    return -1;
  if (this.year > other.year)
    return +1;
  if (this.month < other.month)
    return -1;
  if (this.month > other.month)
    return +1;
  if (this.day < other.day)
    return -1;
  if (this.day > other.day)
    return +1;
  if (this.period != null && other.period != null) {
    if (this.period < other.period)
      return -1;
    if (this.period > other.period)
      return +1;
  }
  return 0;
}

return WorksheetDate;
}(); // end WorksheetDate namespace
