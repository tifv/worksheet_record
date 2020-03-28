// Documentation:
// https://docs.aws.amazon.com/AmazonS3/latest/API/sig-v4-header-based-auth.html

S3Signer = function() { // begin namespace

const service = "s3";

function S3Signer({region, bucket_url, access_key, secret_key}) {
  this.region = region;
  var {
    scheme: bucket_protocol,
    host: bucket_host,
    path: bucket_path,
    query: bucket_query,
    fragment: bucket_fragment
  } = split_url(bucket_url);
  if (!bucket_path.endsWith("/")) {
    throw "S3Signer: bucket URL must end with slash";
  }
  if (bucket_query != "") {
    throw "S3Signer: bucket URL must not contain query";
  }
  if (bucket_fragment != "") {
    throw "S3Signer: bucket URL must not contain fragment";
  }
  this.bucket_protocol = bucket_protocol;
  this.bucket_host = bucket_host;
  this.bucket_path = bucket_path;
  this.access_key = access_key;
  this.secret_key = secret_key;
}

const url_regex = new RegExp( "^" +
  "(https?):\\/\\/((?:[0-9a-z](?:[0-9a-z-]*[0-9a-z]|))(?:\\.(?:[0-9a-z](?:[0-9a-z-]*[0-9a-z]|)))*(?::\\d+)?)" +
  "(\\/[!*'();:@&=+$,\\/\\[\\]%A-Za-z0-9\-_.~]*)" +
  "(?:\\?([!*'();:@&=+$,\\/\\?\\[\\]%A-Za-z0-9\\-_.~]*))?(?:#(.*))?" +
  "$" );

function split_url(url) {
  var [, scheme, host, path, query = "", fragment = ""] = url_regex.exec(url);
  return {
    scheme: scheme, host: host, path: path,
    query: query, fragment: fragment };
}

S3Signer.prototype.get_url = function(path) {
  return this.bucket_protocol + "://" + this.bucket_host + percent_encode(path);
}

S3Signer.prototype.sign = function(method, path, query_string, pre_header_list) {
  if (query_string != "")
    throw "XXX not implemented";
  var request_path = percent_encode(this.bucket_path + path);
  var request_url = this.bucket_protocol + "://" + this.bucket_host + request_path;
  var date = new Date();
  var signing_date = Utilities.formatDate(date, "GMT", "yyyyMMdd");
  var amz_date = Utilities.formatDate(date, "GMT", "yyyyMMdd'T'HHmmss'Z'");
  var payload_hash = null;
  var content_type = null;
  var header_list = [];
  for (let [name, value] of pre_header_list) {
    let name_lower = name.toLowerCase();
    header_list.push([name, value]);
    if (name_lower == "x-amz-content-sha256")
      payload_hash = value;
    if (name_lower == "content-type")
      content_type = value;
    if (name_lower == "host")
      throw "S3Signer().sign: 'Host' header is disallowed in input headers";
    if (name_lower == "x-amz-date")
      throw "S3Signer().sign: 'x-amz-date' header is disallowed in input headers";
  }
  header_list.push(["host", this.bucket_host]);
  header_list.push(["x-amz-date", amz_date]);
  if (content_type == null) {
    throw "S3Signer().sign: 'Content-Type' header must be present in input headers";
  }
  if (payload_hash == null) {
    payload_hash = "UNSIGNED-PAYLOAD";
    header_list.push(["x-amz-content-sha256", payload_hash]);
  }
  var signing_key = generate_signing_key.call(this, signing_date);
  var [canonical_headers, signed_headers] = generate_canonical_headers(header_list);
  var canonical_request = method + "\n" +
    request_path + "\n" + query_string + "\n" +
    canonical_headers + "\n" +
    signed_headers + "\n" +
    payload_hash;
  console.log(canonical_request);
  var string_to_sign = "AWS4-HMAC-SHA256" + "\n" +
    amz_date + "\n" +
    signing_date + "/" + this.region + "/" + service + "/aws4_request\n" +
    bytes_to_hex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, canonical_request));
  console.log(string_to_sign);
  header_list.push([ "Authorization",
    "AWS4-HMAC-SHA256" + " " +
    "Credential=" + this.access_key + "/" + signing_date + "/" +
      this.region + "/" + service + "/aws4_request" + ", " +
    "SignedHeaders=" + signed_headers + ", " +
    "Signature=" + bytes_to_hex(Utilities.computeHmacSha256Signature(
      string_to_bytes(string_to_sign), signing_key ))
  ]);
  return [request_url, header_list];
}

function generate_signing_key(date) { // applied to S3Signer
  if (/\d{8}/.exec(date) == null)
    throw "S3Signer().generate_signing_key: date is not in a correct format";
  var date_key = Utilities.computeHmacSha256Signature(
    string_to_bytes(date), string_to_bytes("AWS4" + this.secret_key) );
  var region_key = Utilities.computeHmacSha256Signature(
    string_to_bytes(this.region), date_key );
  var service_key = Utilities.computeHmacSha256Signature(
    string_to_bytes(service), region_key );
  var signing_key = Utilities.computeHmacSha256Signature(
    string_to_bytes("aws4_request"), service_key );
  return signing_key;
}

function generate_canonical_headers(header_list) {
  // header_list = [[`name`, `value`], ...]
  // Return [canonical_headers, signed_headers]
  var headers = new Map();
  for (let [name, value] of header_list) {
    name = name.toLowerCase();
    var values = headers.get(name);
    if (values == null) {
      values = [];
      headers.set(name, values);
    }
    values.push(value);
  }
  var canonical_headers_pieces = [];
  var signed_headers_pieces = [];
  for (let name of Array.from(headers.keys()).sort()) {
    canonical_headers_pieces.push(name, ":", headers.get(name).join(","), "\n");
    signed_headers_pieces.push(name);
  }
  return [
    canonical_headers_pieces.join(""),
    signed_headers_pieces.join(";") ];
}

function percent_encode(string) {
  var bytes = Utilities.newBlob(string).getBytes().map(x => (x + 256) % 256);
  var result = [];
  for (let i = 0; i < bytes.length; ++i) {
    let b = bytes[i];
    if (
      b >= 65 && b <  91 ||   // A-Z
      b >= 97 && b < 123 ||   // a-z
      b >= 48 && b <  58 ||   // 0-9
      b ===  45 ||   // -
      b ===  46 ||   // .
      b ===  47 ||   // /
      b ===  95 ||   // _
      b === 126      // ~
    ) {
      result.push(String.fromCharCode(b));
      continue;
    }
    percent: if (b === 37) {   // %
      if (i + 2 >= bytes.length)
        break percent;
      let b1 = bytes[i+1];
      let b2 = bytes[i+2];
      if (
        (b1 >= 65 && b1 < 91 || b1 >= 48 && b1 < 58) && // 0-9A-F
        (b2 >= 65 && b2 < 91 || b2 >= 48 && b2 < 58)
      ) {
        result.push(String.fromCharCode(b, b1, b2));
        i += 2;
        continue;
      }
    }
    result.push("%" + b.toString(16).toUpperCase().padStart(2,"0"));
  }
  return result.join("");
}

function string_to_bytes(string) {
  return Utilities.newBlob(string).getBytes();
}

function bytes_to_hex(bytes) {
  var pieces = [];
  for (var i = 0; i < bytes.length; ++i) {
    pieces.push(
      ((bytes[i] + 256) % 256).toString(16).padStart(2, "0")
    );
  }
  return pieces.join("");
}

return S3Signer;

}(); // end S3Signer namespace
