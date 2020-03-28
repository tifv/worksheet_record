/* HSL
 *   .to_hsl([H,S,L]) → [H,S,L]
 *       validating copy; may return null
 *   .to_rgb([H,S,L]) → [R,G,B]
 *       return the color as three 0-255 components
 *   .to_hex([H,S,L]) → hex string
 *       return color as hex value
 *   .to_css([H,S,L]) → CSS color string
 *       return color as CSS hsl() function
 */

var HSL = function() { // namespace

function to_hsl(hsl) {
  if (hsl == null)
    return null;
  var [H, S, L] = hsl;
  if (typeof H != "number" || typeof S != "number" || typeof L != "number") {
    return null;
  }
  return [H, S, L];
}

function deepen(hsl, factor=1) {
  var [H, S, L] = hsl;
  return [H, S, 1 - (1 - L) * factor];
}

function to_rgb(hsl) {
  hsl = to_hsl(hsl);
  if (hsl == null)
    return null;
  var [H, S, L] = hsl;
  var m2 = (L <= 0.5) ? L * (S + 1) : L + S - L * S;
  var m1 = 2 * L - m2;
  return [
    Math.round(255 * hue2g(m1, m2, H + 120)),
    Math.round(255 * hue2g(m1, m2, H)),
    Math.round(255 * hue2g(m1, m2, H - 120))
  ]
}

function hue2g(m1, m2, H) {
  H = (H + 360) % 360;
  if (H <  60) return m1 + (m2 - m1) * H / 60;
  if (H < 180) return m2;
  if (H < 240) return m1 + (m2 - m1) * (240 - H) / 60;
               return m1;
}

function hex2(n) {
  return n.toString(16).padStart(2, "0");
}

function to_hex(hsl) {
  var rgb = to_rgb(hsl);
  if (rgb == null)
    return null;
  return "#" + hex2(rgb[0]) + hex2(rgb[1]) + hex2(rgb[2]);
}

function to_css(hsl) {
  return "hsl(" + hsl[0] + "," + (100 * hsl[1]) + "%," + (100 * hsl[2]) + "%)";
}

return {to_hsl: to_hsl, deepen: deepen, to_rgb: to_rgb, to_hex: to_hex, to_css: to_css};
}(); // end HSL namespace

// XXX add function that sets hyperlink color to hsl(220, 75%, 40%)
