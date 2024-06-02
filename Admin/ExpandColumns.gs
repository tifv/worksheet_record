var expand_columns = new Scheduler(
  /** @param {string[]} group_names */
  function expand_columns(group_names) {
    let iteratee = (group) => {
      if (!group_names.includes(group.name))
        return;
      group.sheet.expandAllColumnGroups();
    };
    const main_spreadsheet = MainSpreadsheet.get();
    for (let group of StudyGroup.list(main_spreadsheet))
      iteratee(group);
    const hidden_spreadsheet = HiddenSpreadsheet.get();
    if (hidden_spreadsheet !== null) {
      for (let group of StudyGroup.list(hidden_spreadsheet))
        iteratee(group);
    }
  },
  "expand_columns.schedule",
  function generate_schedule() {
    /** @type {Map<number, string[]>} */
    var schedule_map = new Map();
    let iteratee = (group) => {
      let timetable = group.get_today_timetable();
      if (timetable == null)
        return;
      for (let [_period, {time, duration}] of Object.entries(timetable)) {
        if (time == null)
          continue;
        if (duration == null)
          continue;
        time += duration + 5;
        let group_names = schedule_map.get(time);
        if (group_names == null) {
          group_names = [];
          schedule_map.set(time, group_names);
        }
        group_names.push(group.name);
      }
    };
    const main_spreadsheet = MainSpreadsheet.get();
    for (let group of StudyGroup.list(main_spreadsheet))
      iteratee(group);
    const hidden_spreadsheet = HiddenSpreadsheet.get();
    if (hidden_spreadsheet !== null) {
      for (let group of StudyGroup.list(hidden_spreadsheet))
        iteratee(group);
    }
    if (schedule_map.size == 0) {
      console.log("no group defines a timetable for today");
    }
    return Array.from(schedule_map.entries())
      .sort(([t1], [t2]) => t1 - t2)
      .map(([time, group_names]) => {
        let date = new Date();
        date.setHours(0);
        date.setMinutes(time);
        date.setSeconds(0);
        return {date, args: [group_names]};
      });
  }
);

function expand_columns_forever() {
  expand_columns.never();
  ScriptApp.newTrigger("expand_columns.today")
    .timeBased()
      .everyDays(1)
      .atHour(1)
      .nearMinute(15)
    .create();
}

