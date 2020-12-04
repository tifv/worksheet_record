function test_set_timetable_for_today() {
  const spreadsheet = MainSpreadsheet.get();
  var now_s = WorksheetDate.today().format();
  console.log(JSON.stringify(now_s));
  for (let group of StudyGroup.list(spreadsheet)) {
    console.log(group.name);
    let tt = {}
    for (var i = 1; i < 10; ++i) {
      tt[i] = {time: 1030 + i * 10, duration: 5};
    }
    group.set_timetable({[now_s]: tt});
    console.log(group.get_timetable());
    console.log(group.get_today_timetable());
  }
}
