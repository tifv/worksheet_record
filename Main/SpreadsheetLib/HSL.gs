/* HSL
 *   .copy({h,s,l}) → {h,s,l}
 *       validating copy
 *   .to_rgb({h,s,l}) → {r,g,b}
 *       return the color as three 0-255 components
 *   .to_hex({h,s,l}) → hex string
 *       return color as hex value
 *   .to_css({h,s,l}) → CSS color string
 *       return color as CSS hsl() function
 */

var HSL = function() { // namespace

function copy(hsl) {
  var {h, s, l} = hsl;
  if (typeof h != "number" || typeof s != "number" || typeof l != "number") {
    throw new Error("invalid HSL value");
  }
  return {h: h, s: s, l: l};
}

function deepen(hsl, factor=1) {
  if (!(factor > 0)) {
    throw new Error("factor must be a positive number");
  }
  var {h, s, l} = copy(hsl);
  var lratio = l / (1 - l);
  if (!isFinite(lratio) || !(lratio > 0))
    return {h: h, s: s, l: l}
  lratio /= factor;
  return {h: h, s: s, l: lratio / (lratio + 1)};
}

function to_rgb(hsl) {
  var {h, s, l} = copy(hsl);
  var m2 = (l <= 0.5) ? l * (s + 1) : l + s - l * s;
  var m1 = 2 * l - m2;
  return {
    r: Math.round(255 * hue2g(m1, m2, h + 120)),
    g: Math.round(255 * hue2g(m1, m2, h)),
    b: Math.round(255 * hue2g(m1, m2, h - 120))
  };
}

function hue2g(m1, m2, h) {
  h = (h + 360) % 360;
  if (h <  60) return m1 + (m2 - m1) * h / 60;
  if (h < 180) return m2;
  if (h < 240) return m1 + (m2 - m1) * (240 - h) / 60;
               return m1;
}

function hex2(n) {
  return n.toString(16).padStart(2, "0");
}

function to_hex(hsl) {
  var rgb = to_rgb(hsl);
  return "#" + hex2(rgb.r) + hex2(rgb.g) + hex2(rgb.b);
}

function to_css(hsl) {
  return "hsl(" + hsl.h + "," + (100 * hsl.s) + "%," + (100 * hsl.l) + "%)";
}

return {copy: copy, deepen: deepen, to_rgb: to_rgb, to_hex: to_hex, to_css: to_css};
}(); // end HSL namespace

// XXX add function that sets hyperlink color to hsl(220, 75%, 40%)

