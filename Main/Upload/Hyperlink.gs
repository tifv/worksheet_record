function decode_hyperlink_formula_(formula) {
  var hyperlink_filter_match = /^=(?:hyperlink|HYPERLINK)\((?:filter|FILTER)\((?:uploads|'uploads')!R\d+C\d+:C\d+[,;](?:uploads|'uploads')!R\d+C\d+:C\d+="([^"]*)"\)[,;]"([^"]*)"\)$/
    .exec(formula);
  if (hyperlink_filter_match != null) {
    return [{filter: hyperlink_filter_match[1]}, hyperlink_filter_match[2]];
  }
  var hyperlink_match = /^=\s*hyperlink\s*\(\s*"([^"]*)"\s*[,;]\s*"([^"]*)"\s*\)\s*$/i
    .exec(formula);
  if (hyperlink_match != null) {
    return [{url: hyperlink_match[1]}, hyperlink_match[2]];
  }
  return null;
}


