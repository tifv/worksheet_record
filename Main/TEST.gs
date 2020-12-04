/*

function test_whatever_spreadsheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = spreadsheet.getActiveSheet();
  const range = sheet.getActiveRange();
}

function test_whatever_range() {
  var range = SpreadsheetApp.getActiveRange();
}

function test_whatever() {
  console.log("whatever");
}

*/

function test_add_study_group_antirow(iteratee) {
  for (var i = 0; i < 16; ++i) {
    let flags = {
      mirror:   (i & 1) > 0,
      category: (i & 2) > 0,
      max:      (i & 4) > 0,
      weight:   (i & 8) > 0,
    };
    let name = "test_" + (i+1).toString().padStart(2, "0") + "_" +
      (flags.mirror ? "Z" : "z") +
      (flags.category ? "C" : "c") +
      (flags.max ? "M" : "m") +
      (flags.weight ? "W" : "w");
    (iteratee || test_add_study_group_antirow_)(name, flags);
  }
}

function test_add_study_group_antirow_clean() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  test_add_study_group_antirow((name, flags) => {
    test_worksheet_clear_(spreadsheet, name);
  });
}

function test_add_study_group_antirow_(name, flags) {
  if (!flags.mirror || !flags.category)
    return;
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var color_schemes = ColorSchemes.get(spreadsheet);
  function named_scheme(name) {
    return Object.assign({name: name}, color_schemes[name]);
  }
  test_worksheet_clear_(spreadsheet, name);
  var group = StudyGroupBuilder.build(spreadsheet, name, {
    rows: {
      data_row:     4 + flags.mirror + flags.category + flags.max + flags.weight,
      title_row:    2 + flags.mirror + flags.category + flags.max + flags.weight,
      label_row:    3 + flags.mirror + flags.category + flags.max + flags.weight,
      mirror_row:   flags.mirror ? 1 : null,
      category_row: flags.category ? 2 + flags.mirror + flags.max + flags.weight : null,
      max_row:      flags.max ? 1 + flags.mirror + flags.weight : null,
      weight_row:   flags.weight ? 1 + flags.mirror : null,
    },
    data_height: 20,
    rating: true, sum: true,
    categories: [
      {code: "a"},
      {code: "g"},
      {code: "c"},
      {code: "o", rating: false}
    ],
    category_musthave: true,
    attendance: {
      columns: {
        date_lists: [
          {
            title: "I",
            start: new Date(Date.parse("2020-09-01 ")),
            end:   new Date(Date.parse("2021-01-01 ")),
            weekdays: [true, false, false, true, false, false, false]
          },
          {
            title: "II",
            start: new Date(Date.parse("2021-01-01 ")),
            end:   new Date(Date.parse("2021-06-01 ")),
            weekdays: [true, false, false, true, false, false, false]
          }
        ] } },
    color_scheme: named_scheme("lotus"),
  });
  var sheet = group.sheet;
  sheet.collapseAllColumnGroups();
  group = new StudyGroup(sheet);
  sheet.getRange(group.dim.data_row + 1, 2, group.dim.data_height - 1, 1)
    .setValues(test_sample_names(20));
  test_worksheet_add_random_(group, {
    data_width: 20, title: "Алгебра", category: "a",
  });
  test_worksheet_add_random_(group, {
    data_width: 20, title: "Геометрия", category: "g",
  });
  test_worksheet_add_random_(group, {
    data_width: 20, title: "Комбинаторика", category: "c",
  });
}

function test_add_study_groups() {
  test_add_study_group_("test");
  test_add_study_group_color_("test_color");
}

function test_worksheet_clear_(spreadsheet, name) {
  var sheet = spreadsheet.getSheetByName(name);
  if (sheet != null)
    spreadsheet.deleteSheet(sheet);
}

function test_add_study_group_(name) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var color_schemes = ColorSchemes.get(spreadsheet);
  function named_scheme(name) {
    return Object.assign({name: name}, color_schemes[name]);
  }
  test_worksheet_clear_(spreadsheet, name);
  var group = StudyGroupBuilder.build(spreadsheet, name, {
    data_height: 20,
    rating: true, sum: true,
    categories: [
      {code: "a"},
      {code: "g"},
      {code: "c"},
      {code: "o", rating: false}
    ],
    category_musthave: true,
    attendance: {
      columns: {
        date_lists: [
          {
            title: "I",
            start: new Date(Date.parse("2020-02-01 ")),
            end:   new Date(Date.parse("2020-06-01 ")),
            weekdays: [true, false, false, true, false, false, false]
          },
          {
            title: "II",
            start: new Date(Date.parse("2020-09-01 ")),
            end:   new Date(Date.parse("2021-01-01 ")),
            weekdays: [true, false, false, true, false, false, false]
          }
        ] } },
    color_scheme: named_scheme("lotus"),
  });
  var sheet = group.sheet;
  sheet.collapseAllColumnGroups();
  group = new StudyGroup(sheet);
  sheet.getRange(group.dim.data_row + 1, 2, group.dim.data_height - 1, 1)
    .setValues(test_sample_names(20));
  test_worksheet_add_random_(group, {
    data_width: 20, title: "Алгебра", category: "a",
  });
  test_worksheet_add_random_(group, {
    data_width: 20, title: "Геометрия", category: "g",
  });
  test_worksheet_add_random_(group, {
    data_width: 20, title: "Комбинаторика", category: "c",
  });
}

function test_add_study_group_color_(name) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var color_schemes = ColorSchemes.get(spreadsheet);
  function named_scheme(name) {
    return Object.assign({name: name}, color_schemes[name]);
  }
  test_worksheet_clear_(spreadsheet, name);
  var group = StudyGroupBuilder.build(spreadsheet, name, {
    data_height: 20,
    rating: true, sum: true,
    categories: [
      {code: "a"},
      {code: "g"},
      {code: "c"},
      {code: "o", rating: {integrate: false}}
    ],
    category_musthave: false,
    //color_scheme: named_scheme("lotus*"),
  });
  var sheet = group.sheet;
  group = new StudyGroup(sheet);
  sheet.getRange(group.dim.data_row + 1, 2, group.dim.data_height - 1, 1)
    .setValues(test_sample_names(20));
  function add_worksheet(name, shift=0) {
    test_worksheet_add_random_(group, {
      data_width: 20, title: name + (shift != 0 ? " " + shift : ""),
      color_scheme: shift == 0 ? named_scheme(name) : scheme_shift(named_scheme(name), shift),
    });
  }
  function hsl_shift(hsl, i=0) {
    let {h, s, l} = hsl;
    return {h: (h + 6*i) % 360, s: s, l: l};
  }
  function scheme_shift(scheme, i=0) {
    let {mark, rating_mid, rating_top} = scheme;
    return {
      mark: hsl_shift(mark, i),
      rating_mid: hsl_shift(rating_mid, i),
      rating_top: hsl_shift(rating_top, i),
    };
  }
  add_worksheet("default");
  for (let name in color_schemes) {
    add_worksheet(name);
  }
  add_worksheet("default");
}

function test_sample_values(n, random) {
  var x = [], y = [];
  for (var i = 0; i < n; ++i) {
    x.push([i, Math.random()]);
    y.push([i, Math.random()]);
  }
  if (random) {
    x.sort(([i, u], [j, v]) => u - v);
    y.sort(([i, u], [j, v]) => u - v);
  }
  return x.map(([i,]) => y.map(([j,]) => (i >= j ? 1 : null)));
}

function test_sample_labels(n) {
  var l = [];
  for (var i = 0; i < n; ++i) {
    l.push(i + 1);
  }
  return [l];
}

function test_sample_names(n) {
  var names = [
    "Новоселова Янина", "Решетилова Анна", "Ноздрёв Андрей", "Корявова Зоя", "Елизаров Наум",
    "Уваров Валерий", "Яромеева Мария", "Газинский Дементий", "Позон Евлампий", "Золотова Алиса",
    "Сигачёв Севастьян", "Никешин Зиновий", "Зыкин Егор", "Граббе Анисья", "Куроптева Валерия",
    "Чемоданов Прокл", "Язынина Лада", "Нырцева Маргарита", "Маркин Никита", "Дорохов Серафим",
    "Пивоварова Евдокия", "Бурков Карл", "Сафонова Надежда", "Дёмина Аза", "Эсце Жанна",
    "Исакова Дарья", "Гаголина Василиса", "Ёжикова Вероника", "Фотеев Родион", "Набатникова Оксана",
    "Магазинер Роза", "Азаренков Моисей", "Якурин Николай", "Дорофеев Данил", "Шкуратов Евсей",
    "Кабинова Яна", "Васютин Якуб", "Тюшняков Евстигней", "Никитина Дина", "Грибова Анастасия",
    "Проскуркин Аркадий", "Буданов Харитон", "Курчатова Софья", "Гришин Лаврентий", "Грачёв Пимен",
    "Ерофеева Христина", "Серёгина Рада", "Лихачева Бронислава", "Шпикалов Семён", "Ягужинский Фома" ];
  if (!(typeof n == "number") || isNaN(n) || n < 0 || n > names.length)
    throw "wat";
  names.length = n;
  names.sort();
  return names.map(v => [v]);
}

function test_worksheet_add_(group, options) {
  const sheet = group.sheet;
  var worksheet = WorksheetBuilder.build(group, sheet.getRange(1,sheet.getMaxColumns()), options);
  if (group.dim.data_height != 20 || worksheet.dim.data_width != 20)
    throw "wat";
  sheet.getRange(group.dim.data_row + 1, worksheet.dim.data_start, group.dim.data_height - 1, worksheet.dim.data_width)
    .setValues(test_sample_values(20, false));
  sheet.getRange(group.dim.label_row, worksheet.dim.data_start, 1, worksheet.dim.data_width)
    .setValues(test_sample_labels(20));
  if (options.check)
    worksheet.check();
}

function test_worksheet_add_random_(group, options) {
  const sheet = group.sheet;
  var worksheet = WorksheetBuilder.build(group, sheet.getRange(1,sheet.getMaxColumns()), options);
  sheet.getRange(group.dim.data_row + 1, worksheet.dim.data_start, group.dim.data_height - 1, worksheet.dim.data_width)
    .setValues(test_sample_values(20, true));
  sheet.getRange(group.dim.label_row, worksheet.dim.data_start, 1, worksheet.dim.data_width)
    .setValues(test_sample_labels(20));
  if (options.check)
    worksheet.check();
}

function test_set_upload_config() {
  const ui = SpreadsheetApp.getUi();
  function get_value(label) {
    var response = ui.prompt( "Загрузка (тест)",
      label + ":", ui.ButtonSet.OK_CANCEL );
    if (response.getSelectedButton() == ui.Button.CANCEL)
      throw "wat";
    return response.getResponseText();
  }

  UploadConfig.set({
    access_key: get_value("Access key"),
    secret_key: get_value("Secret key"),
    region:     get_value("Region"),
    bucket_url: get_value("Bucket URL"),
    enable_solutions: true,
  });
}

function test_set_timetables() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  function get_metadata(name) {
    var range = spreadsheet.getRangeByName(name);
    if (range == null)
      return null;
    return load_antijson_(range.getValues());
  }
  for (let group of StudyGroup.list(spreadsheet)) {
    let name = group.name;
    let timetable = get_metadata("timetable_" + name);
    if (timetable)
      group.set_timetable(timetable);
    console.log( name + "'s timetable: " +
      JSON.stringify(group.get_timetable()) );
    let worksheet_plan = get_metadata("worksheet_plan_" + name);
    group.set_worksheet_plan(worksheet_plan);
    console.log( name + "'s worksheet plan: " +
      JSON.stringify(group.get_worksheet_plan()) );
  }
}

