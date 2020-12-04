const expand_columns_property_key = "last_expand_columns"

function expand_columns_now() {
  const spreadsheet = MainSpreadsheet.get();
  var now = new Date();
  var current_time = now.getHours() * 60 + now.getMinutes();
  var last_expanded; last_expanded: {
    let last_expanded_s = PropertiesService.getDocumentProperties()
      .getProperty(expand_columns_property_key);
    if (last_expanded_s == null) {
      last_expanded = null;
      break last_expanded;
    }
    let date = new Date(last_expanded_s);
    if (
      date.getFullYear() != now.getFullYear() ||
      date.getMonth() != now.getMonth() ||
      date.getDate() != now.getDate()
    ) {
      last_expanded = null;
      break last_expanded;
    }
    last_expanded = date.getHours() * 60 + date.getMinutes();
  }
  for (let group of StudyGroup.list(spreadsheet)) {
    let timetable = group.get_today_timetable();
    if (timetable == null)
      continue;
    let eligible = false;
    for (let [period, {time, duration}] of Object.entries(timetable)) {
      if (time == null)
        continue;
      if (duration != null)
        time += duration;
      if (last_expanded != null && time < last_expanded)
        continue;
      if (time > current_time)
        continue;
      eligible = true;
      break;
    }
    if (eligible)
      group.sheet.expandAllColumnGroups();
  }
  PropertiesService.getDocumentProperties()
    .setProperty(expand_columns_property_key, now.toISOString());
}

function expand_columns_today() {
  // assume its early morning
  // schedule expand_columns_timely() calls for the day
  const spreadsheet = MainSpreadsheet.get();
  var times = new Set();
  var busy = false;
  for (let group of StudyGroup.list(spreadsheet)) {
    let timetable = group.get_today_timetable();
    if (timetable == null)
      continue;
    busy = true;
    for (let [period, {time, duration}] of Object.entries(timetable)) {
      if (time == null)
        continue;
      if (duration != null)
        time += duration;
      times.add(time);
    }
  }
  var now = new Date();
  for (let time of times) {
    now.setHours(0);
    now.setMinutes(time);
    ScriptApp.newTrigger("expand_columns_now")
      .timeBased().at(now)
      .create();
  }
  if (!busy)
    throw new Error("no group defines a timetable for today");
}

function expand_columns_forever() {
  // XXX set up trigger for expand_columns_today()
}
