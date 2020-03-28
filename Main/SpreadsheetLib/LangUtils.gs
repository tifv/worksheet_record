function* group_by_(array, valuater = (x => x)) {
  if (array.length < 1)
    return;
  var v = valuater(array[0]);
  var len = 1;
  var result = [];
  for (var i = 1; i < array.length; ++i) {
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
  for (var name in generators) {
    properties[name] = lazy_property_factory_.call(this, name, generators[name]);
  }
  return properties;
}

function define_lazy_properties_(object, generators) {
  Object.defineProperties( object,
    lazy_properties_factory_.call(object, generators) );
}


