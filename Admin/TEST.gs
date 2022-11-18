function test_set_timetable_for_today() {
  const spreadsheet = MainSpreadsheet.get();
  var now_s = WorksheetDate.today().format();
  for (let group of StudyGroup.list(spreadsheet)) {
    console.log(group.name);
    let timetable = {};
    for (var i = 1; i < 10; ++i) {
      timetable[i] = {time: 1030 + i * 10, duration: 5};
    }
    group.set_timetable({[now_s]: timetable});
    console.log(group.get_timetable());
    console.log(group.get_today_timetable());
  }
}

function test_set_worksheet_plan_for_today() {
  const spreadsheet = MainSpreadsheet.get();
  var now_s = WorksheetDate.today().format();
  for (let group of StudyGroup.list(spreadsheet)) {
    console.log(group.name);
    let worksheet_plan = [{period: "1", category: "a"}, {period: "2", category: "g"}, {period: "3", category: "c"}];
    group.set_worksheet_plan({[now_s]: worksheet_plan});
    console.log(group.get_worksheet_plan());
    console.log(group.get_today_worksheet_plan());
  }
}

