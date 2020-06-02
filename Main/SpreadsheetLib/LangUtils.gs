function lazy_property_factory_(name, generator) {
  return {configurable: true, get: function() {
    var value = generator.call(this);
    Object.defineProperty( this, name,
      {configurable: true, value: value} );
    return value;
  }};
}

function define_lazy_property_(object, name, generator) {
  Object.defineProperty( object, name, 
    lazy_property_factory_.call(object, name, generator) );
}

function lazy_properties_factory_(generators) {
  var properties = {};
  for (let name in generators) {
    properties[name] = lazy_property_factory_.call(this, name, generators[name]);
  }
  return properties;
}

function define_lazy_properties_(object, generators) {
  Object.defineProperties( object,
    lazy_properties_factory_.call(object, generators) );
}

function* group_by_(array, valuater = (x => x)) {
  if (array.length < 1)
    return;
  var v = valuater(array[0]);
  var len = 1;
  var result = [];
  for (let i = 1; i < array.length; ++i) {
    var vv = valuater(array[i]);
    if (vv === v) {
      ++len;
    } else {
      yield {start: i - len, length: len, end: i, value: v};
      v = vv;
      len = 1;
    }
  }
  yield {start: array.length - len, length: len, end: array.length, value: v};
}

function load_antijson_(data_rows) {
  function extract_value(regex, datum) {
    var match = regex.exec(datum);
    if (match == null)
      return null;
    return match[1];
  }
  var container = {};
  for (let data_row of data_rows) {
    let current_obj = container;
    let current_key = "result";
    for (let datum of data_row) {
      let value;
      if (typeof datum == "number") {
        current_obj[current_key] = datum;
        break;
      } else if ((value = extract_value(/^"(.*)"$/, datum)) != null) {
        current_obj[current_key] = value;
        break;
      } else if ((value = extract_value(/^\["(.*)"\]$/, datum)) != null) {
        if (current_obj[current_key] == null) {
          current_obj[current_key] = {};
        }
        current_obj = current_obj[current_key];
        current_key = value;
        continue;
      } else if ((value = extract_value(/^\[(\d*)\]$/, datum)) != null) {
        if ( current_obj[current_key] == null ||
          !(current_obj[current_key] instanceof Array)
        ) {
          current_obj[current_key] = Object.assign([], current_obj[current_key]);
        }
        current_obj = current_obj[current_key];
        current_key = parseInt(value, 10);
        continue;
      } else {
        // malformed data? ignore
        break;
      }
    }
  }
  return container.result;
}

function include_html_(filename) {
    return HtmlService.createHtmlOutputFromFile(filename)
        .getContent();
}

