'use strict';
/* =========================================================
   Programming Paradigms · PowerPoint version
   Alexandru Ioan · Student ID 25002954

   Builds the .pptx with pptxgenjs (needs pptxgenjs and sharp):
     node make-pptx.js [output.pptx]
   ========================================================= */
const pptxgen = require('pptxgenjs');
const sharp = require('sharp');
const JSZip = require('jszip');
const fs = require('fs');

const OUT = process.argv[2] || 'Programming-Paradigms-Alexandru-Ioan.pptx';
const NAME = 'Alexandru Ioan';
const STUDENT_ID = '25002954';

/* ---------- colours: the same as the web version ---------- */
const C = {
  ink: '1B1F1D', ink2: '4A534E', muted: '5F6964', faint: '8A948F', line: 'DFE3DE', line2: 'B7C0BA',
  paper: 'F5F6F3', white: 'FFFFFF', grey: 'EEF0EE', grey2: 'E2E6E3',
  good: '0CA30C', bad: 'D03B3B', goodSoft: 'E3F3E5', badSoft: 'FBE8E7',
  codeBg: 'FBFCFA', key: 'B8323A', str: '1D7A4E', num: 'A35A00', fn: '3456C0',
  term: '1D2320', termInk: 'E6ECE8', termMuted: '9AA7A0', termOut: 'A6E3BF',
  night: '1B1F1D', night2: '262D29', nightLine: '3A4440', nightMuted: 'A9B5AE', nightInk2: 'C9D3CD',
};
const P = { 1: '2A78D6', 2: 'DF8700', 3: 'C2449A' };      // procedural, event-driven, object-oriented
const T8 = { 1: 'EEF4FC', 2: 'FCF5EB', 3: 'FAF0F7' };      // 8% tints
const T14 = { 1: 'E1ECF9', 2: 'FAEEDB', 3: 'F6E5F1' };     // 14% tints
const T25 = { 1: 'CADDF5', 2: 'F7E1BF', 3: 'F0D0E6' };     // 25% tints
const PNAME = { 1: 'Procedural', 2: 'Event-driven', 3: 'Object-oriented' };
const PGLYPH = { 1: 'proc', 2: 'event', 3: 'oop' };
const F = { head: 'Cambria', body: 'Calibri', code: 'Consolas' };

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5 inches
pres.author = NAME;
pres.company = `Student ID ${STUDENT_ID}`;
pres.title = 'Programming Paradigms';
pres.subject = 'Procedural, event-driven and object-oriented programming';
pres.theme = { headFontFace: F.head, bodyFontFace: F.body };
const S = pres.shapes;
const W = 13.333;
const MX = 0.6;
const CW = W - 2 * MX;

/* ---------- icons: the three paradigms and four kinds of event ---------- */
const GLYPHS = {
  proc: '<rect x="4" y="3" width="16" height="4.5" rx="1.6"/><rect x="4" y="9.75" width="16" height="4.5" rx="1.6"/><rect x="4" y="16.5" width="16" height="4.5" rx="1.6"/>',
  event: '<path d="M13.6 2.5 5 13.6h6.3l-1.3 7.9 8.9-11.3h-6.4l1.1-7.7z"/>',
  oop: '<path d="M12 2.6 3.7 7.1v9.8l8.3 4.5 8.3-4.5V7.1L12 2.6z"/><path d="M3.7 7.1 12 11.6l8.3-4.5M12 11.6v9.8"/>',
  click: '<path d="M6 3.5 18.5 10l-5.6 1.7-2.4 5.6L6 3.5z"/><path d="m13 12 5 5"/>',
  key: '<rect x="2.5" y="6" width="19" height="12" rx="2.2"/><path d="M6.5 10h1M10.5 10h1M14.5 10h1M7 14h10"/>',
  timer: '<circle cx="12" cy="13.5" r="7.5"/><path d="M12 9.5v4l2.6 1.8M9.5 2.8h5M12 2.8V6"/>',
  door: '<path d="M5 21V4.5A1.5 1.5 0 0 1 6.5 3H14v18"/><path d="M3 21h13"/><path d="M11 12.5h.01"/><path d="M17.5 8.5a4.5 4.5 0 0 1 0 7M19.8 6.2a7.8 7.8 0 0 1 0 11.6"/>',
};
const ICON = {};
async function makeIcons() {
  for (const [name, body] of Object.entries(GLYPHS)) {
    for (const [tone, hex] of [['ink', C.ink], ['white', C.white]]) {
      const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="256" height="256" fill="none" stroke="#${hex}" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">${body}</svg>`;
      const png = await sharp(Buffer.from(svg)).png().toBuffer();
      ICON[`${name}-${tone}`] = 'image/png;base64,' + png.toString('base64');
    }
  }
}

/* ---------- drawing helpers ---------- */
function rect(slide, x, y, w, h, o = {}) {
  slide.addShape(o.r ? S.ROUNDED_RECTANGLE : S.RECTANGLE, {
    x, y, w, h,
    ...(o.r ? { rectRadius: o.r } : {}),
    fill: o.fill ? { color: o.fill } : { type: 'none' },
    line: o.line ? { color: o.line, width: o.lw || 1, ...(o.dash ? { dashType: o.dash } : {}) } : { type: 'none' },
    ...(o.shadow ? { shadow: { type: 'outer', color: '000000', blur: 8, offset: 2, angle: 90, opacity: 0.14 } } : {}),
  });
}
function oval(slide, x, y, w, h, o = {}) {
  slide.addShape(S.OVAL, {
    x, y, w, h,
    fill: o.fill ? { color: o.fill } : { type: 'none' },
    line: o.line ? { color: o.line, width: o.lw || 1 } : { type: 'none' },
  });
}
// a straight line from (x1, y1) to (x2, y2); `end` puts an arrowhead at (x2, y2)
function line(slide, x1, y1, x2, y2, o = {}) {
  slide.addShape(S.LINE, {
    x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1),
    flipH: x2 < x1, flipV: y2 < y1,
    line: {
      color: o.color || C.ink2, width: o.w || 1.25,
      ...(o.dash ? { dashType: o.dash } : {}),
      ...(o.end ? { endArrowType: o.end } : {}),
    },
  });
}
function tx(slide, text, x, y, w, h, o = {}) {
  slide.addText(text, {
    x, y, w, h, margin: 0, isTextBox: true,
    fontFace: F.body, fontSize: 14, color: C.ink, valign: 'top',
    ...o,
  });
}
// "**bold**" and "`code`" inside a string, as text runs
function md(str, base = {}) {
  const runs = [];
  str.split(/(\*\*[^*]+\*\*|`[^`]+`)/).forEach(part => {
    if (!part) return;
    if (part.startsWith('**')) runs.push({ text: part.slice(2, -2), options: { ...base, bold: true } });
    else if (part.startsWith('`')) runs.push({ text: part.slice(1, -1), options: { ...base, fontFace: F.code, fontSize: (base.fontSize || 14) * 0.92 } });
    else runs.push({ text: part, options: { ...base } });
  });
  return runs;
}
function bullets(items, base) {
  const runs = [];
  items.forEach((item, i) => {
    const parts = md(item, base);
    parts[0].options.bullet = { indent: 14 };
    if (i < items.length - 1) parts[parts.length - 1].options.breakLine = true;
    runs.push(...parts);
  });
  return runs;
}
function label(slide, text, x, y, w, o = {}) {
  tx(slide, text.toUpperCase(), x, y, w, 0.24, { fontSize: 10, color: o.dark ? C.nightMuted : C.muted, charSpacing: 1.5, valign: 'middle', ...o.opts });
}
// width of a short Calibri label, in inches (a little generous)
const estW = (s, size) => s.replace(/[`*]/g, '').length * size * 0.53 / 72;

/* ---------- syntax colours for the code ---------- */
const SYNTAX = {
  python: {
    re: /(#.*$)|([rRfFbB]?(?:"(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*'))|(\b\d+(?:\.\d+)?\b)|([A-Za-z_]\w*)|(\s+)|([\s\S])/g,
    keywords: 'def class return if elif else for while in import from as global nonlocal and or not is None True False pass try except with lambda',
    builtins: 'print len range str int float super sum input',
    defs: 'def class',
  },
  js: {
    re: /(\/\/.*$)|(`(?:[^`\\]|\\.)*`|"(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*')|(\b\d+(?:\.\d+)?\b)|([A-Za-z_$][\w$]*)|(\s+)|([\s\S])/g,
    keywords: 'let const var function return if else for while new true false null of in',
    builtins: 'document console',
    defs: 'function',
  },
  c: {
    re: /(\/\/.*$|^\s*#.*$)|("(?:[^"\\]|\\.)*"|'(?:[^'\\]|\\.)*')|(\b\d+(?:\.\d+)?\b)|([A-Za-z_]\w*)|(\s+)|([\s\S])/g,
    keywords: 'int float double char void return for while if else struct',
    builtins: 'printf',
    defs: '',
  },
};
Object.values(SYNTAX).forEach(s => {
  s.keywords = new Set(s.keywords.split(' '));
  s.builtins = new Set(s.builtins.split(' '));
  s.defs = new Set(s.defs.split(' ').filter(Boolean));
});
const STYLE = {
  k: { color: C.key, bold: true }, s: { color: C.str }, n: { color: C.num }, f: { color: C.fn, bold: true },
  b: { color: C.fn }, self: { color: C.ink2 }, c: { color: C.ink2, italic: true }, '': { color: C.ink },
};
function tokens(line, lang) {
  const syn = SYNTAX[lang];
  const re = new RegExp(syn.re.source, 'g');
  const out = [];
  const push = (text, cls) => {
    const last = out[out.length - 1];
    if (last && last.cls === cls) last.text += text;
    else out.push({ text, cls });
  };
  let prev = '';
  let m;
  while ((m = re.exec(line)) !== null) {
    const [tok, comment, str, num, word] = m;
    if (tok === '') { re.lastIndex++; continue; }
    if (comment !== undefined) {
      push(comment, lang === 'c' && comment.trimStart().startsWith('#') ? 'k' : 'c');
      prev = '';
    } else if (str !== undefined) {
      push(str, 's'); prev = '';
    } else if (num !== undefined) {
      push(num, 'n'); prev = '';
    } else if (word !== undefined) {
      const callNext = /^\s*\(/.test(line.slice(re.lastIndex));
      let cls = '';
      if (syn.defs.has(prev) || (lang === 'c' && /^(int|float|double|char|void)$/.test(prev) && callNext)) cls = 'f';
      else if (syn.keywords.has(word)) cls = 'k';
      else if (word === 'self') cls = 'self';
      else if (syn.builtins.has(word)) cls = 'b';
      push(word, cls);
      prev = word;
    } else {
      push(tok, '');
      if (!/^\s+$/.test(tok)) prev = '';
    }
  }
  return out;
}
function codeRuns(code, lang, numbers) {
  const lines = code.split('\n');
  const width = String(lines.length).length;
  const runs = [];
  lines.forEach((ln, i) => {
    const parts = [];
    if (numbers) parts.push({ text: String(i + 1).padStart(width, ' ') + '  ', options: { color: C.faint } });
    const toks = tokens(ln, lang);
    if (!toks.length) parts.push({ text: ' ', options: { color: C.ink } });
    toks.forEach(t => parts.push({ text: t.text, options: { ...STYLE[t.cls] } }));
    if (i < lines.length - 1) parts[parts.length - 1].options.breakLine = true;
    runs.push(...parts);
  });
  return runs;
}

/* ---------- an editor window with code ---------- */
function editor(slide, o) {
  const { x, y, w } = o;
  const p = o.p || 0;
  const size = o.size || 11;
  const lineH = o.lineH || 15; // points
  const lines = o.code.split('\n');
  const barH = 0.36;
  const padT = 0.1;
  const padB = 0.13;
  const railW = o.rails ? 0.36 : 0;
  const padL = 0.16 + railW;
  const lh = lineH / 72;
  const codeH = lines.length * lh;
  const out = o.out || [];
  const outH = out.length ? 0.18 + out.length * 0.22 + 0.08 : 0;
  const h = barH + padT + codeH + padB + outH;
  const codeY = y + barH + padT;
  const lineY = n => codeY + (n - 1) * lh;

  rect(slide, x, y, w, h, { r: 0.08, fill: C.codeBg, shadow: true });
  rect(slide, x, y, w, barH + 0.12, { r: 0.08, fill: p ? T14[p] : C.grey });
  rect(slide, x, y + barH, w, 0.12, { fill: C.codeBg });
  line(slide, x, y + barH, x + w, y + barH, { color: C.line, w: 0.75 });
  oval(slide, x + 0.17, y + barH / 2 - 0.06, 0.12, 0.12, { fill: p ? P[p] : C.ink });
  tx(slide, o.file, x + 0.38, y, 3.6, barH, { fontFace: F.code, fontSize: 10.5, valign: 'middle' });
  tx(slide, o.tag || '', x + w - 3.3, y, 3.12, barH, { align: 'right', valign: 'middle', fontSize: 10.5, color: C.muted });

  (o.hot || []).forEach(n => rect(slide, x + 0.03, lineY(n), w - 0.06, lh, { fill: p ? T14[p] : C.grey }));
  if (o.cur) rect(slide, x + 0.03, lineY(o.cur), w - 0.06, lh, { fill: p ? T25[p] : C.grey2 });
  (o.rails || []).forEach(r => {
    rect(slide, x + 0.14 + r.col * 0.11, lineY(r.from) + 0.01, 0.06, (r.to - r.from + 1) * lh - 0.02, { fill: P[r.p] });
  });

  slide.addText(codeRuns(o.code, o.lang, o.numbers !== false), {
    x: x + padL, y: codeY, w: w - padL - 0.08, h: codeH,
    margin: 0, isTextBox: true, fontFace: F.code, fontSize: size, color: C.ink, valign: 'top',
    lineSpacing: lineH, paraSpaceBefore: 0, paraSpaceAfter: 0,
  });

  if (out.length) {
    const oy = y + h - outH;
    rect(slide, x, oy, w, outH, { r: 0.08, fill: C.term });
    rect(slide, x, oy, w, 0.12, { fill: C.term });
    tx(slide, 'OUTPUT', x + 0.18, oy + 0.12, 0.9, 0.22, { fontSize: 8.5, color: C.termMuted, charSpacing: 1.5, valign: 'middle' });
    tx(slide, out.join('\n'), x + 1.05, oy + 0.11, w - 1.2, out.length * 0.22, { fontFace: F.code, fontSize: 11, color: C.termOut, lineSpacing: 15.8 });
  }
  rect(slide, x, y, w, h, { r: 0.08, line: C.line, lw: 1 });
  return { h, bottom: y + h, lineY };
}

/* a dark terminal panel: lines are [kind, text] with kind 'in' | 'out' | 'cmd' */
function terminal(slide, x, y, w, title, lines, o = {}) {
  const lineH = o.lineH || 14.5;
  const h = 0.36 + lines.length * lineH / 72 + 0.14;
  rect(slide, x, y, w, h, { r: 0.1, fill: C.term });
  tx(slide, title.toUpperCase(), x + 0.2, y + 0.1, w - 0.4, 0.22, { fontSize: 8.5, color: C.termMuted, charSpacing: 1.5, valign: 'middle' });
  const runs = lines.map(([kind, t], i) => ({
    text: kind === 'in' ? `>>> ${t}` : t,
    options: { color: kind === 'in' ? C.termInk : kind === 'cmd' ? C.termMuted : C.termOut, ...(i < lines.length - 1 ? { breakLine: true } : {}) },
  }));
  slide.addText(runs, { x: x + 0.2, y: y + 0.36, w: w - 0.4, h: lines.length * lineH / 72, margin: 0, isTextBox: true, fontFace: F.code, fontSize: o.size || 10.5, valign: 'top', lineSpacing: lineH, paraSpaceAfter: 0 });
  return y + h;
}

/* ---------- repeated slide parts ---------- */
function badge(slide, p, x, y, d, o = {}) {
  oval(slide, x, y, d, d, { fill: o.soft ? (o.dark ? C.night2 : C.white) : T25[p], line: P[p], lw: d >= 0.5 ? 2.5 : 1.75 });
  const s = d * 0.52;
  slide.addImage({ data: ICON[`${PGLYPH[p]}-${o.dark && o.soft ? 'white' : 'ink'}`], x: x + (d - s) / 2, y: y + (d - s) / 2, w: s, h: s, altText: `${PNAME[p]} icon` });
}
// the three paradigm icons at the top right of each paradigm slide
function pstrip(slide, on) {
  const sizes = [1, 2, 3].map(p => (p === on ? 0.56 : 0.42));
  const gap = 0.2;
  const total = sizes[0] + sizes[1] + sizes[2] + 2 * gap;
  const cy = 1.02;
  const x0 = W - MX - total;
  const centers = [];
  let x = x0;
  sizes.forEach(d => { centers.push(x + d / 2); x += d + gap; });
  line(slide, centers[0], cy, centers[2], cy, { color: C.line2, w: 1.5 });
  x = x0;
  [1, 2, 3].forEach((p, i) => {
    badge(slide, p, x, cy - sizes[i] / 2, sizes[i], { soft: p !== on });
    x += sizes[i] + gap;
  });
}
function newSlide(o = {}) {
  const slide = pres.addSlide();
  slide.background = { color: o.dark ? C.night : C.paper };
  return slide;
}
function header(slide, { num, kicker, title, p = 0, dark = false }) {
  oval(slide, MX, 0.42, 0.3, 0.3, { fill: p ? T25[p] : (dark ? C.night2 : C.white), line: p ? P[p] : (dark ? C.white : C.ink), lw: p ? 2 : 1.25 });
  tx(slide, String(num), MX, 0.42, 0.3, 0.3, { align: 'center', valign: 'middle', fontFace: F.code, fontSize: 10.5, color: dark ? C.white : C.ink });
  tx(slide, kicker.toUpperCase(), MX + 0.42, 0.42, 8, 0.3, { valign: 'middle', fontSize: 11, color: dark ? C.nightMuted : C.ink2, charSpacing: 1.5 });
  tx(slide, title, MX, 0.8, p ? 10.0 : CW, 0.66, { fontFace: F.head, fontSize: 32, bold: true, color: dark ? C.white : C.ink, valign: 'middle' });
  if (p) pstrip(slide, p);
}
function footer(slide, dark = false) {
  tx(slide, `${NAME} · Student ID ${STUDENT_ID}`, MX, 7.04, 6, 0.26, { fontSize: 10, color: dark ? C.nightMuted : C.muted, valign: 'middle' });
  slide.slideNumber = { x: W - MX - 0.5, y: 7.04, w: 0.5, h: 0.26, fontFace: F.body, fontSize: 10, color: dark ? C.nightMuted : C.muted, align: 'right' };
}
function dict(slide, { x, y, w, h, p = 0, word, pos = 'noun', def, cite, also, size = 14 }) {
  rect(slide, x, y, w, h, { r: 0.12, fill: p ? T8[p] : C.grey, line: p ? T25[p] : C.line });
  tx(slide, [
    { text: word, options: { fontFace: F.head, fontSize: 19, bold: true, color: C.ink } },
    { text: `   ${pos}`, options: { fontFace: F.body, fontSize: 13, italic: true, color: C.muted } },
  ], x + 0.24, y + 0.15, w - 0.48, 0.42, { valign: 'middle' });
  tx(slide, [...md(`${def} `, { fontSize: size, color: C.ink }), { text: `(${cite.replace(/ /g, '\u00A0')})`, options: { fontSize: size - 3, color: C.muted } }],
    x + 0.24, y + 0.62, w - 0.48, h - 0.62 - (also ? 0.52 : 0.14), { lineSpacingMultiple: 1.05 });
  if (also) tx(slide, also, x + 0.24, y + h - 0.5, w - 0.48, 0.38, { fontFace: F.code, fontSize: 9.5, color: C.ink2, valign: 'bottom' });
}
function features(slide, items, { x, y, colW, rowH, gapX = 0.28 }) {
  items.forEach(([title, desc], i) => {
    const fx = x + (i % 2) * (colW + gapX);
    const fy = y + Math.floor(i / 2) * rowH;
    tx(slide, [{ text: title, options: { bold: true, fontSize: 13, color: C.ink, breakLine: true } }, ...md(desc, { fontSize: 11, color: C.ink2 })],
      fx, fy, colW, rowH - 0.04, { paraSpaceAfter: 1 });
  });
}
function prosCons(slide, { x, y, w, h, good, items, size = 12.5 }) {
  rect(slide, x, y, w, h, { r: 0.12, fill: good ? C.goodSoft : C.badSoft });
  oval(slide, x + 0.2, y + 0.16, 0.32, 0.32, { fill: good ? C.good : C.bad });
  tx(slide, good ? '+' : '−', x + 0.2, y + 0.16, 0.32, 0.32, { align: 'center', valign: 'middle', fontSize: 16, bold: true, color: C.white });
  tx(slide, good ? 'Advantages' : 'Disadvantages', x + 0.64, y + 0.12, w - 0.8, 0.4, { fontFace: F.head, fontSize: 17, bold: true, valign: 'middle' });
  tx(slide, bullets(items, { fontSize: size, color: C.ink }), x + 0.24, y + 0.6, w - 0.42, h - 0.66, { paraSpaceAfter: 4 });
}
function chips(slide, labels, { x, y, maxW, size = 11.5, h = 0.32 }) {
  let cx = x;
  let cy = y;
  labels.forEach(l => {
    const w = estW(l, size) + 0.32;
    if (cx + w > x + maxW + 0.001) { cx = x; cy += h + 0.1; }
    rect(slide, cx, cy, w, h, { r: h / 2, fill: C.white, line: C.line });
    tx(slide, md(l, { fontSize: size, color: C.ink }), cx, cy, w, h, { align: 'center', valign: 'middle' });
    cx += w + 0.1;
  });
  return cy + h;
}
function card(slide, x, y, w, h, o = {}) {
  rect(slide, x, y, w, h, { r: o.r || 0.12, fill: o.fill || C.white, line: o.line === undefined ? C.line : o.line, shadow: o.shadow !== false });
}

/* ---------- the code examples (the same as the web version) ---------- */
const CODE = {
  proc: `# Procedural: a list of steps, split into functions

def add_up(prices):              # a procedure (function)
    total = 0                    # a local variable
    for price in prices:         # iteration (a loop)
        total = total + price
    return total                 # send the answer back

def apply_offer(total):
    if total >= 5:               # selection (a decision)
        total = total - 1        # £1 off orders of £5+
    return total

basket = [1.25, 3.50, 0.75]      # croissant, sourdough, cookie
subtotal = add_up(basket)        # call the function
to_pay = apply_offer(subtotal)
print(f"To pay: £{to_pay:.2f}")`,
  c: `#include <stdio.h>

float add_up(float prices[], int count) {
    float total = 0;
    for (int i = 0; i < count; i++) {
        total = total + prices[i];
    }
    return total;
}

int main(void) {
    float basket[] = {1.25, 3.50, 0.75};
    printf("Total: %.2f\\n", add_up(basket, 3));
    return 0;
}`,
  till: `import tkinter as tk

total = 0                            # a global variable

def add_croissant():                 # event handler
    global total
    total = total + 1.25
    label.config(text=f"Total: £{total:.2f}")

def clear(event):                    # event handler
    global total
    total = 0
    label.config(text="Total: £0.00")

window = tk.Tk()                     # set up the window
window.title("Bakery till")
label = tk.Label(window, text="Total: £0.00")
label.pack()
button = tk.Button(window, text="Add croissant",
                   command=add_croissant)  # click → handler
button.pack()
window.bind("c", clear)              # key press C → handler

window.mainloop()                    # event loop: wait for events`,
  js: `// the web page has <p id="total"> and <button id="add">
let total = 0;
const button = document.querySelector("#add");
const label = document.querySelector("#total");

function addCroissant() {                        // event handler
  total = total + 1.25;
  label.textContent = \`Total: £\${total.toFixed(2)}\`;
}

button.addEventListener("click", addCroissant);  // event listener`,
  oop: `class Product:
    def __init__(self, name, price):   # constructor
        self.name = name               # attributes
        self.price = price

    def describe(self):                # a method
        return f"{self.name} costs £{self.price:.2f}"

class Basket:
    def __init__(self):
        self.items = []
        self.total = 0

    def add(self, product):            # a method
        self.items.append(product)
        self.total = self.total + product.price

croissant = Product("Croissant", 1.25) # objects
sourdough = Product("Sourdough", 3.50)
basket = Basket()
basket.add(croissant)
basket.add(sourdough)
print(croissant.describe())            # Croissant costs £1.25
print(basket.total)                    # 4.75`,
  poly: `class Product:
    def __init__(self, name, price):
        self.name = name
        self.__price = price             # private (encapsulation)

    def get_price(self):                 # the way to read it
        return self.__price

    def describe(self):
        return f"{self.name}: £{self.__price:.2f}"

class Cake(Product):                     # inherits from Product
    def __init__(self, name, price, slices):
        super().__init__(name, price)    # Product's constructor
        self.slices = slices             # an extra attribute

    def describe(self):                  # overrides describe()
        return f"{self.name} cake, {self.slices} slices"

items = [Product("Croissant", 1.25), Cake("Lemon", 12, 8)]
for item in items:
    print(item.describe())               # polymorphism
print(items[1].get_price())              # inherited from Product`,
  together: `import tkinter as tk

class TillApp:                                # a class
    def __init__(self, window):
        self.total = 0                        # an attribute
        self.label = tk.Label(window, text="Total: £0.00")
        self.label.pack()
        button = tk.Button(window, text="Add croissant",
                           command=self.add)  # click → handler
        button.pack()

    def add(self):                            # a method + event handler
        self.total = self.total + 1.25        # step 1
        if self.total >= 5:                   # step 2: a decision
            self.label.config(fg="green")
        self.label.config(text=f"Total: £{self.total:.2f}")

window = tk.Tk()
app = TillApp(window)                         # an object
window.mainloop()                             # the event loop`,
};

/* =========================================================
   THE SLIDES
   ========================================================= */
function build() {
  /* ---------- 1. title ---------- */
  {
    const s = newSlide({ dark: true });
    tx(s, 'LEVEL 3 IT · PROGRAMMING', 0.8, 1.35, 5.6, 0.3, { fontSize: 12, color: C.nightMuted, charSpacing: 3 });
    tx(s, 'Programming\nparadigms', 0.8, 1.75, 5.9, 2.05, { fontFace: F.head, fontSize: 54, bold: true, color: C.white, lineSpacingMultiple: 0.92 });
    tx(s, 'Procedural, event-driven and object-oriented programming: what each one is, how it works, and the same small program written all three ways.',
      0.8, 3.9, 5.5, 1.1, { fontSize: 17, color: C.nightInk2 });
    tx(s, NAME, 0.8, 5.2, 5, 0.38, { fontSize: 20, bold: true, color: C.white });
    tx(s, `Student ID ${STUDENT_ID}`, 0.8, 5.6, 5, 0.3, { fontFace: F.code, fontSize: 12.5, color: C.nightMuted });
    tx(s, '1 Introduction   ·   2 Procedural   ·   3 Event-driven   ·   4 Object-oriented   ·   5 Comparison   ·   6 Conclusion   ·   7 References',
      0.8, 6.62, 11.7, 0.3, { fontSize: 11, color: C.nightMuted });
    const files = [
      { p: 1, file: 'procedural.py', code: 'def add_up(prices):\n    total = 0\n    for price in prices:\n        total = total + price\n    return total' },
      { p: 2, file: 'event_driven.py', code: 'button = tk.Button(window, text="Add croissant",\n                   command=add_croissant)\nwindow.bind("c", clear)\nwindow.mainloop()        # wait for events' },
      { p: 3, file: 'oop.py', code: 'class Product:\n    def __init__(self, name, price):\n        self.name = name\n        self.price = price' },
    ];
    let fy = 0.62;
    files.forEach((f, i) => {
      const e = editor(s, { x: 6.85 + i * 0.3, y: fy, w: 5.3, p: f.p, code: f.code, lang: 'python', file: f.file, tag: PNAME[f.p], size: 11.5, lineH: 16, numbers: false });
      fy = e.bottom + 0.2;
    });
    tx(s, 'the same basket, three ways', 7.45, fy - 0.02, 5.3, 0.32, { fontFace: F.head, italic: true, fontSize: 14, color: C.nightMuted, align: 'right' });
    s.addNotes(`Hello, I'm ${NAME}. My presentation is about three programming paradigms: procedural, event-driven and object-oriented programming. For each one I'll give a definition, explain the key features, show a simple code example and go through the advantages and disadvantages. Every example does the same small job, adding up a bakery basket, so it's easy to compare them.`);
  }

  /* ---------- 2. introduction ---------- */
  {
    const s = newSlide();
    header(s, { num: 1, kicker: 'Introduction', title: 'What is a programming paradigm?' });
    dict(s, { x: MX, y: 1.7, w: 6.4, h: 1.55, word: 'programming paradigm', def: 'A style, or approach, to programming. It is based on a set of ideas about how a program should be organised and how it runs', cite: 'Van Roy, 2009', size: 15 });
    tx(s, 'A paradigm isn\'t a programming language. It\'s a way of thinking about a problem: how to split it up, where the data is kept and what decides the order the code runs in. Many languages support more than one paradigm. Python, for example, can be written in all three styles in this presentation.',
      MX, 3.42, 6.4, 1.5, { fontSize: 14.5, lineSpacingMultiple: 1.05 });

    // the receipt
    const rx = MX, ry0 = 4.8, rw = 2.85;
    card(s, rx, ry0, rw, 1.96, { r: 0.04 });
    tx(s, 'THE BAKERY', rx, ry0 + 0.12, rw, 0.24, { align: 'center', fontFace: F.code, fontSize: 11, bold: true, charSpacing: 2, valign: 'middle' });
    line(s, rx + 0.16, ry0 + 0.44, rx + rw - 0.16, ry0 + 0.44, { color: C.line2, w: 0.75, dash: 'dash' });
    let ry = ry0 + 0.5;
    const row = (l, r, o = {}) => {
      tx(s, l, rx + 0.18, ry, rw - 1.2, 0.2, { fontFace: F.code, fontSize: o.size || 10.5, bold: o.bold, valign: 'middle' });
      tx(s, r, rx + rw - 1.1, ry, 0.92, 0.2, { fontFace: F.code, fontSize: o.size || 10.5, bold: o.bold, align: 'right', valign: 'middle' });
      ry += 0.205;
    };
    row('Croissant', '£1.25');
    row('Sourdough loaf', '£3.50');
    row('Cookie', '£0.75');
    ry += 0.04;
    line(s, rx + 0.16, ry, rx + rw - 0.16, ry, { color: C.line2, w: 0.75, dash: 'dash' });
    ry += 0.05;
    row('Subtotal', '£5.50');
    row('£1 off £5+', '−£1.00');
    ry += 0.04;
    line(s, rx + 0.16, ry, rx + rw - 0.16, ry, { color: C.ink, w: 1.25 });
    ry += 0.05;
    row('To pay', '£4.50', { bold: true, size: 11.5 });

    label(s, 'My example all the way through', 3.75, 5.08, 3.3);
    tx(s, 'To compare the three fairly, every code example does the same small job: it adds up a bakery basket and takes £1 off orders of £5 or more.',
      3.75, 5.4, 3.25, 1.2, { fontSize: 13.5, color: C.ink2 });

    label(s, 'The three paradigms in this presentation', 7.35, 1.7, 5.4);
    const items = {
      1: ['A list of steps that run in order, grouped into procedures (functions).', 'like following a recipe'],
      2: ['The program waits for events, like a click, and reacts to each one when it happens.', 'like a shop assistant waiting for customers'],
      3: ['The program is built from objects that keep their own data and actions together.', 'like the till, oven and products in a bakery'],
    };
    [1, 2, 3].forEach((p, i) => {
      const cy = 2.05 + i * 1.57;
      card(s, 7.35, cy, 5.38, 1.42);
      badge(s, p, 7.6, cy + 0.24, 0.66);
      tx(s, PNAME[p], 8.5, cy + 0.16, 4.0, 0.36, { fontFace: F.head, fontSize: 18, bold: true, valign: 'middle' });
      tx(s, items[p][0], 8.5, cy + 0.54, 4.05, 0.46, { fontSize: 13, color: C.ink2 });
      tx(s, items[p][1], 8.5, cy + 1.03, 4.05, 0.28, { fontFace: F.head, italic: true, fontSize: 12.5, color: C.ink2 });
    });
    footer(s);
    s.addNotes('A programming paradigm is a style or approach to programming. It isn\'t a language: it\'s a way of organising a program and deciding how it runs. Some languages support more than one paradigm. Python can be written in all three styles, which is why I use it for most of my examples. My example adds up a basket: a croissant, a sourdough loaf and a cookie come to £5.50, and orders of £5 or more get £1 off, so the customer pays £4.50.');
  }

  /* ---------- 3. procedural: what it is ---------- */
  {
    const s = newSlide();
    header(s, { num: 2, kicker: 'Procedural · what it is', title: 'Procedural programming', p: 1 });
    dict(s, {
      x: MX, y: 1.7, w: 6.25, h: 2.42, p: 1, word: 'procedural programming',
      def: 'A style of programming where a program is a list of instructions that the computer carries out in order, from top to bottom. The instructions are grouped into procedures (also called functions or subroutines) that can be called whenever they\'re needed',
      cite: 'Sebesta, 2016', also: 'a type of imperative programming · e.g. C, Pascal, BASIC, Fortran, COBOL, or Python written this way',
    });
    label(s, 'Key features', MX, 4.28, 6.25);
    features(s, [
      ['Sequence, selection, iteration', 'Steps in order, decisions with `if`, and loops with `for` and `while`.'],
      ['Procedures (functions)', 'Named blocks of code that do one job and can be called again and again.'],
      ['Parameters and return values', 'Data goes in as parameters, and the answer comes back with `return`.'],
      ['Local and global variables', 'Local variables only exist inside their procedure; global ones anywhere.'],
      ['Top-down design', 'A big problem is broken down into smaller procedures (decomposition).'],
      ['Data is kept separate', 'Data is passed from procedure to procedure; it doesn\'t belong to them.'],
    ], { x: MX, y: 4.58, colW: 2.99, rowH: 0.74 });

    // structure chart
    card(s, 7.1, 1.7, 5.63, 5.05);
    label(s, 'Top-down design: a structure chart of my example', 7.32, 1.86, 5.2);
    const cx = [8.22, 9.92, 11.62];
    line(s, 9.92, 2.96, 9.92, 3.25);
    line(s, cx[0], 3.25, cx[2], 3.25);
    cx.forEach(c => line(s, c, 3.25, c, 3.9));
    line(s, cx[0], 4.56, cx[0], 4.95);
    line(s, cx[1], 4.56, cx[1], 4.95);
    rect(s, 8.67, 2.3, 2.5, 0.66, { r: 0.08, fill: T14[1], line: P[1], lw: 2 });
    tx(s, [{ text: 'Work out what to pay', options: { bold: true, fontSize: 13, breakLine: true } }, { text: 'the main program', options: { fontSize: 10.5, color: C.ink2 } }],
      8.67, 2.3, 2.5, 0.66, { align: 'center', valign: 'middle' });
    const kids = [['add_up(prices)', 'add up the basket'], ['apply_offer(total)', 'take £1 off?'], ['print(...)', 'show the answer']];
    kids.forEach(([code, sub], i) => {
      rect(s, cx[i] - 0.81, 3.9, 1.62, 0.66, { r: 0.06, fill: C.white, line: C.ink2, lw: 1.25 });
      tx(s, [{ text: code, options: { fontFace: F.code, fontSize: 10, bold: true, breakLine: true } }, { text: sub, options: { fontSize: 10, color: C.ink2 } }],
        cx[i] - 0.81, 3.9, 1.62, 0.66, { align: 'center', valign: 'middle' });
    });
    const flow = (t, x, y) => tx(s, t, x, y, 0.9, 0.22, { fontFace: F.code, fontSize: 9, color: C.muted, valign: 'middle' });
    flow('↓ basket', cx[0] + 0.07, 3.3); flow('↑ subtotal', cx[0] + 0.07, 3.57);
    flow('↓ subtotal', cx[1] + 0.07, 3.3); flow('↑ to_pay', cx[1] + 0.07, 3.57);
    flow('↓ to_pay', cx[2] + 0.07, 3.45);
    const grand = [['for price in prices:', 'loop: add each price', '*'], ['if total >= 5:', 'yes: take £1 off', 'o']];
    grand.forEach(([code, sub, mark], i) => {
      rect(s, cx[i] - 0.81, 4.95, 1.62, 0.74, { r: 0.06, fill: C.white, line: C.ink2, lw: 1.25 });
      tx(s, [{ text: code, options: { fontFace: F.code, fontSize: 9.5, breakLine: true } }, { text: sub, options: { fontSize: 10, color: C.ink2 } }],
        cx[i] - 0.81, 5.0, 1.62, 0.66, { align: 'center', valign: 'middle' });
      tx(s, mark, cx[i] + 0.6, 4.96, 0.18, 0.2, { fontFace: F.code, fontSize: 12, bold: true, align: 'center', valign: 'middle' });
    });
    tx(s, 'Read it left to right: the main program calls each procedure in turn. ↓ is data going in (a parameter), ↑ is data coming back (a return value), * is a loop and o is a decision.',
      7.32, 5.88, 5.2, 0.75, { fontSize: 11.5, color: C.ink2 });
    footer(s);
    s.addNotes('Procedural programming is the most direct style. The program is a list of instructions that run in order, from top to bottom, and the instructions are grouped into procedures, also called functions. Its key features are sequence, selection and iteration; procedures with parameters and return values; local and global variables; and top-down design, where a big problem is broken down into smaller procedures. The structure chart shows my example: the main program calls add_up, then apply_offer, then print, from left to right.');
  }

  /* ---------- 4. procedural: code + trace table ---------- */
  {
    const s = newSlide();
    header(s, { num: 2, kicker: 'Procedural · code example', title: 'Example: adding up the basket', p: 1 });
    editor(s, { x: MX, y: 1.7, w: 6.5, p: 1, code: CODE.proc, lang: 'python', file: 'basket_procedural.py', tag: 'Python 3', out: ['To pay: £4.50'] });
    label(s, 'Trace table: the values as the program runs', 7.4, 1.7, 5.33);
    const hd = t => ({ text: t, options: { bold: true, fill: { color: T25[1] }, fontFace: F.body, fontSize: 12 } });
    const v = (t, inFn) => ({ text: t, options: { fontFace: F.code, fontSize: 11.5, fill: { color: inFn ? T8[1] : C.white } } });
    const rows = [
      ['4', '', '0', '', '', '', 1],
      ['5–6', '1.25', '1.25', '', '', '', 1],
      ['5–6', '3.5', '4.75', '', '', '', 1],
      ['5–6', '0.75', '5.5', '', '', '', 1],
      ['15', '', '', '5.5', '', '', 0],
      ['10–11', '', '4.5', '', '', '', 1],
      ['16', '', '', '', '4.5', '', 0],
      ['17', '', '', '', '', 'To pay: £4.50', 0],
    ];
    s.addTable([
      ['Line', 'price', 'total', 'subtotal', 'to_pay', 'Output'].map(hd),
      ...rows.map(r => r.slice(0, 6).map(t => v(t, r[6]))),
    ], {
      x: 7.4, y: 2.0, w: 5.33, colW: [0.78, 0.68, 0.68, 0.86, 0.76, 1.57], rowH: 0.31,
      border: { type: 'solid', pt: 0.75, color: C.line2 }, color: C.ink, valign: 'middle', margin: [0.03, 0.08, 0.03, 0.08],
    });
    tx(s, bullets([
      'The two functions only run when they\'re called, on lines 15 and 16.',
      '`add_up()` uses a loop and `apply_offer()` uses an `if`.',
      '`total` is a local variable, so each function has its own. Shaded rows run inside a function.',
    ], { fontSize: 12.5, color: C.ink2 }), 7.4, 5.0, 5.33, 1.6, { paraSpaceAfter: 5 });
    footer(s);
    s.addNotes('Here is the procedural version in Python. There are two functions: add_up uses a loop to add up the prices, and apply_offer uses an if statement to take £1 off. The main program at the bottom calls them in order. The trace table shows what happens line by line: total goes from 0 to 5.5, so subtotal is 5.5, then the offer brings it down to 4.5, and the program prints "To pay: £4.50". Notice that total is a local variable: each function has its own one.');
  }

  /* ---------- 5. procedural: pros & cons ---------- */
  {
    const s = newSlide();
    header(s, { num: 2, kicker: 'Procedural · pros & cons', title: 'Procedural: pros and cons', p: 1 });
    const e = editor(s, { x: MX, y: 1.7, w: 6.0, p: 1, code: CODE.c, lang: 'c', file: 'basket.c', tag: 'C', out: ['Total: 5.50'] });
    tx(s, [...md('The same `add_up()` in C, a classic procedural language. A C program is just functions and data, with no classes or objects ', { fontSize: 12, color: C.ink2 }), { text: '(Kernighan and Ritchie, 1988).', options: { fontSize: 10.5, color: C.muted } }],
      MX, e.bottom + 0.14, 6.0, 0.6);
    prosCons(s, { x: 6.95, y: 1.7, w: 5.78, h: 1.62, good: true, items: [
      '**Easy to follow:** it runs in the order you read it',
      '**Fast and light:** good for scripts and small devices',
      '**Reusable functions:** call them and test them one at a time',
    ] });
    prosCons(s, { x: 6.95, y: 3.44, w: 5.78, h: 1.62, good: false, items: [
      '**Messy when it grows:** too many functions and globals',
      '**Unprotected data:** any function can change a global',
      '**A poor fit for real things** like customers and products',
    ] });
    label(s, 'Where it\'s used', 6.95, 5.24, 5.78);
    chips(s, ['scripts & automation', 'calculations & data processing', 'embedded systems (Arduino)', 'operating systems (Linux)'], { x: 6.95, y: 5.54, maxW: 5.78 });
    footer(s);
    s.addNotes('This is the same add_up function in C, a classic procedural language: there are no classes or objects, just functions and data. Procedural code is easy to follow, fast and light, and its functions can be reused and tested one at a time. But big procedural programs get messy, global data isn\'t protected, and it\'s not a natural way to model real-world things like customers and products. It\'s used for scripts, calculations, embedded systems like Arduino, and operating systems: Linux is mostly written in C.');
  }

  /* ---------- 6. event-driven: what it is ---------- */
  {
    const s = newSlide();
    header(s, { num: 3, kicker: 'Event-driven · what it is', title: 'Event-driven programming', p: 2 });
    dict(s, {
      x: MX, y: 1.7, w: 6.0, h: 2.45, p: 2, word: 'event-driven programming',
      def: 'A style of programming where the flow of the program is decided by events: things that happen, like a click, a key press, a sensor reading or a timer running out. The program waits for events, and when one happens it runs the code linked to it, called an event handler',
      cite: 'MDN Web Docs, no date a', also: 'e.g. Visual Basic, C# with Windows Forms, JavaScript, Python with tkinter', size: 13.5,
    });
    label(s, 'Key features', MX, 4.3, 6.0);
    features(s, [
      ['Events', 'User actions (click, key press), sensors, timers or messages from other programs.'],
      ['Event handlers', 'Functions that run when a certain event happens (also called callbacks).'],
      ['Listeners (triggers)', 'Code that links an event on an object, like a button, to its handler.'],
      ['The event loop', 'Runs all the time, waiting for the next event and sending it to its handler.'],
      ['An event queue', 'Events that arrive together wait in a queue and are handled one at a time.'],
      ['The user is in control', 'Usually a GUI with buttons and menus. The user decides what happens next.'],
    ], { x: MX, y: 4.6, colW: 2.86, rowH: 0.73 });

    // the event loop diagram
    card(s, 6.95, 1.7, 5.78, 5.05);
    label(s, 'The event loop', 7.15, 1.86, 5.3);
    const srcs = [['click', 'Mouse click'], ['key', 'Key press'], ['timer', 'Timer'], ['door', 'Door sensor']];
    let sx = 7.15;
    srcs.forEach(([icon, name]) => {
      const w = estW(name, 10.5) + 0.55;
      rect(s, sx, 2.2, w, 0.36, { r: 0.18, fill: C.white, line: C.line2 });
      s.addImage({ data: ICON[`${icon}-ink`], x: sx + 0.12, y: 2.28, w: 0.2, h: 0.2, altText: `${name} icon` });
      tx(s, name, sx + 0.38, 2.2, w - 0.42, 0.36, { fontSize: 10.5, bold: true, valign: 'middle' });
      sx += w + 0.08;
    });
    label(s, 'Event queue', 7.15, 2.84, 1.4);
    rect(s, 7.15, 3.12, 1.3, 2.25, { r: 0.1, line: C.line2, lw: 1.25, dash: 'dash' });
    [['key', 'key C'], ['timer', 'timer'], ['door', 'door']].forEach(([icon, t], i) => {
      const ty = 3.22 + i * 0.42;
      rect(s, 7.23, ty, 1.14, 0.34, { r: 0.06, fill: T25[2] });
      s.addImage({ data: ICON[`${icon}-ink`], x: 7.3, y: ty + 0.07, w: 0.2, h: 0.2, altText: `${t} event icon` });
      tx(s, t, 7.55, ty, 0.8, 0.34, { fontFace: F.code, fontSize: 10.5, valign: 'middle' });
    });
    tx(s, 'next', 7.15, 4.55, 1.3, 0.22, { fontSize: 9.5, italic: true, color: C.muted, align: 'center' });
    line(s, 8.5, 4.25, 8.8, 4.25, { color: C.ink2, w: 1.5, end: 'triangle' });
    oval(s, 8.88, 3.5, 1.5, 1.5, { fill: C.white, line: C.line2, lw: 1.25 });
    s.addShape(S.ARC, { x: 8.8, y: 3.42, w: 1.66, h: 1.66, angleRange: [200, 150], line: { color: P[2], width: 3, endArrowType: 'triangle' } });
    tx(s, [
      { text: 'EVENT LOOP', options: { fontSize: 8.5, color: C.muted, charSpacing: 1.2, breakLine: true } },
      { text: 'running', options: { fontSize: 10, color: C.ink2, breakLine: true } },
      { text: 'on_click()', options: { fontFace: F.code, fontSize: 11, bold: true } },
    ], 8.88, 3.5, 1.5, 1.5, { align: 'center', valign: 'middle' });
    line(s, 10.44, 4.25, 10.66, 4.25, { color: C.ink2, w: 1.5, end: 'triangle' });
    label(s, 'Event handlers', 10.7, 2.84, 1.85);
    [['on_click()', 'add a croissant'], ['on_key()', 'clear the basket'], ['on_timer()', 'update the clock'], ['on_sensor()', 'ring the shop bell']].forEach(([fn, what], i) => {
      const hy = 3.12 + i * 0.58;
      rect(s, 10.7, hy, 1.83, 0.5, { r: 0.07, fill: i === 0 ? T25[2] : C.white, line: i === 0 ? P[2] : C.line, lw: i === 0 ? 1.5 : 1 });
      tx(s, [{ text: fn, options: { fontFace: F.code, fontSize: 10.5, bold: true, breakLine: true } }, { text: what, options: { fontSize: 9.5, color: C.ink2 } }],
        10.82, hy, 1.66, 0.5, { valign: 'middle' });
    });
    rect(s, 7.15, 5.55, 5.38, 0.36, { r: 0.07, fill: T8[2], line: T25[2] });
    tx(s, 'Mouse click  →  on_click() runs: a croissant is added', 7.3, 5.55, 5.1, 0.36, { fontFace: F.code, fontSize: 10.5, valign: 'middle' });
    tx(s, 'Events wait in the queue and the loop handles them one at a time, in the order they arrived.', 7.15, 6.05, 5.38, 0.55, { fontSize: 11.5, color: C.ink2 });
    footer(s);
    s.addNotes('In event-driven programming, the order the code runs in is decided by events, like a mouse click, a key press, a timer or a sensor. The program waits in a loop called the event loop. When an event happens it goes into a queue, and the loop runs the event handler linked to it. The handlers are just functions. The diagram shows four kinds of event: the loop is running on_click for a mouse click, and the other events wait in the queue to be handled one at a time.');
  }

  /* ---------- 7. event-driven: code ---------- */
  {
    const s = newSlide();
    header(s, { num: 3, kicker: 'Event-driven · code example', title: 'Example: a bakery till with a button', p: 2 });
    editor(s, { x: MX, y: 1.7, w: 6.75, p: 2, code: CODE.till, lang: 'python', file: 'till_events.py', tag: 'Python 3 · tkinter', size: 10.5, lineH: 13.5,
      hot: [5, 6, 7, 8, 10, 11, 12, 13, 19, 20, 22], cur: 24 });
    label(s, 'The window the code makes', 7.7, 1.7, 5.0);
    // a mock-up of the tkinter window
    const wx = 7.7, wy = 2.0;
    card(s, wx, wy, 3.3, 1.72, { r: 0.06, line: C.line2 });
    rect(s, wx, wy, 3.3, 0.34, { r: 0.06, fill: C.paper });
    rect(s, wx, wy + 0.22, 3.3, 0.12, { fill: C.paper });
    line(s, wx, wy + 0.34, wx + 3.3, wy + 0.34, { color: C.line2, w: 0.75 });
    tx(s, 'Bakery till', wx + 0.14, wy, 2, 0.34, { fontSize: 11, bold: true, valign: 'middle' });
    [0, 1, 2].forEach(i => rect(s, wx + 2.55 + i * 0.22, wy + 0.1, 0.15, 0.15, { r: 0.03, line: C.line2, lw: 0.75 }));
    tx(s, 'Total: £2.50', wx, wy + 0.46, 3.3, 0.5, { fontFace: F.head, fontSize: 24, bold: true, align: 'center', valign: 'middle' });
    rect(s, wx + 0.82, wy + 1.08, 1.66, 0.42, { r: 0.06, fill: C.paper, line: C.line2 });
    tx(s, 'Add croissant', wx + 0.82, wy + 1.08, 1.66, 0.42, { fontSize: 12, bold: true, align: 'center', valign: 'middle' });
    tx(s, 'This is the window after two clicks.', 11.2, 2.45, 1.5, 0.8, { fontFace: F.head, italic: true, fontSize: 12.5, color: C.ink2 });

    label(s, 'What happens, event by event', 7.7, 3.95, 5.0);
    const hd = t => ({ text: t, options: { bold: true, fill: { color: T25[2] }, fontSize: 12 } });
    const cell = (t, code) => ({ text: t, options: { fontFace: code ? F.code : F.body, fontSize: code ? 11 : 12, fill: { color: C.white } } });
    s.addTable([
      ['Event', 'Handler that runs', 'Total'].map(hd),
      [cell('click'), cell('add_croissant()', true), cell('£1.25', true)],
      [cell('click'), cell('add_croissant()', true), cell('£2.50', true)],
      [cell('key press C'), cell('clear(event)', true), cell('£0.00', true)],
    ], { x: 7.7, y: 4.25, w: 5.03, colW: [1.5, 2.2, 1.33], rowH: 0.31, border: { type: 'solid', pt: 0.75, color: C.line2 }, color: C.ink, valign: 'middle', margin: [0.03, 0.08, 0.03, 0.08] });
    tx(s, bullets([
      'Lines 15–22 run once to build the window and link the events.',
      'Then `mainloop()` waits until an event happens.',
      'The user decides the order; the code decides what each event does.',
    ], { fontSize: 12, color: C.ink2 }), 7.7, 5.62, 5.03, 1.15, { paraSpaceAfter: 4 });
    footer(s);
    s.addNotes('This is a till written with tkinter, Python\'s library for windows and buttons. Lines 15 to 22 run once to build the window and link the events to their handlers: the button\'s command is add_croissant, and the C key is bound to clear. Then mainloop starts the event loop and waits. When the user clicks the button twice, the total goes to £1.25 and then £2.50. Pressing C clears it. The user decides the order; the programmer only decides what each event does.');
  }

  /* ---------- 8. event-driven: pros & cons ---------- */
  {
    const s = newSlide();
    header(s, { num: 3, kicker: 'Event-driven · pros & cons', title: 'Event-driven: pros and cons', p: 2 });
    const e = editor(s, { x: MX, y: 1.7, w: 6.6, p: 2, code: CODE.js, lang: 'js', file: 'till.js', tag: 'JavaScript', size: 10.5, lineH: 14.5 });
    tx(s, [...md('The same idea on a web page. Browsers are event-driven: `addEventListener()` links an event to its handler ', { fontSize: 12, color: C.ink2 }), { text: '(MDN Web Docs, no date b).', options: { fontSize: 10.5, color: C.muted } }],
      MX, e.bottom + 0.12, 6.6, 0.5);
    const py = e.bottom + 0.72;
    card(s, MX, py, 6.6, 1.38, { fill: T8[2], line: T25[2], shadow: false });
    tx(s, 'This presentation is event-driven too', MX + 0.22, py + 0.12, 6.1, 0.32, { fontFace: F.head, fontSize: 14, bold: true, valign: 'middle' });
    tx(s, 'Every click or key press is an event. PowerPoint\'s event handler reacts by showing the next slide.', MX + 0.22, py + 0.46, 6.1, 0.42, { fontSize: 12, color: C.ink2 });
    let lx = MX + 0.22;
    ['click  →  next slide', 'arrow key  →  next slide'].forEach(t => {
      const w = t.length * 10 * 0.6 / 72 + 0.3;
      rect(s, lx, py + 0.92, w, 0.32, { r: 0.06, fill: C.white, line: T25[2] });
      tx(s, t, lx, py + 0.92, w, 0.32, { fontFace: F.code, fontSize: 10, align: 'center', valign: 'middle' });
      lx += w + 0.12;
    });
    prosCons(s, { x: 7.55, y: 1.7, w: 5.18, h: 1.62, good: true, items: [
      '**Made for interaction:** it reacts straight away',
      '**Easy to extend:** a new button needs a new handler',
      '**No wasted work:** nothing runs until an event',
    ] });
    prosCons(s, { x: 7.55, y: 3.44, w: 5.18, h: 1.62, good: false, items: [
      '**Harder to test:** events can come in any order',
      '**Harder to follow:** the code jumps between handlers',
      '**A slow handler freezes** the whole app',
    ] });
    label(s, 'Where it\'s used', 7.55, 5.24, 5.18);
    chips(s, ['desktop apps', 'websites (JavaScript)', 'games', 'mobile apps', 'sensors & alarms', 'web servers'], { x: 7.55, y: 5.54, maxW: 5.18 });
    footer(s);
    s.addNotes('On the web, JavaScript works the same way: addEventListener links a click to a handler function. Even PowerPoint is event-driven: when I click or press the arrow key, that\'s an event, and its handler shows the next slide. Event-driven programs are made for interaction, easy to extend, and don\'t waste work. But they\'re harder to test and follow, because events can come in any order, and one slow handler can freeze the whole app.');
  }

  /* ---------- 9. OOP: what it is ---------- */
  {
    const s = newSlide();
    header(s, { num: 4, kicker: 'Object-oriented · what it is', title: 'Object-oriented programming (OOP)', p: 3 });
    dict(s, {
      x: MX, y: 1.7, w: 6.1, h: 2.25, p: 3, word: 'object-oriented programming', pos: 'noun · OOP for short',
      def: 'A style of programming where the program is built from objects. Each object keeps its own data (attributes) together with the actions that use that data (methods). Objects are made from classes, which work like blueprints',
      cite: 'Oracle, no date', also: 'e.g. Java, C#, C++, Python, Ruby, Kotlin', size: 13.5,
    });
    const term = (t, d) => [
      { text: t, options: { bold: true, fontSize: 12.5 } },
      { text: md(d, { fontSize: 12, color: C.ink2 }), options: {} },
    ];
    s.addTable([
      term('Class', 'The blueprint for making objects, e.g. `Product`.'),
      term('Object', 'One thing made from a class (an instance), e.g. `croissant`.'),
      term('Attribute', 'A piece of data an object keeps, e.g. its `name` and `price`.'),
      term('Method', 'A function that belongs to an object, e.g. `describe()`.'),
      term('Constructor', 'The method that sets up a new object: `__init__()` in Python.'),
    ], {
      x: MX, y: 4.08, w: 6.1, colW: [1.35, 4.75], rowH: 0.29, color: C.ink, valign: 'middle', fontFace: F.body,
      border: [{ type: 'solid', pt: 0.75, color: C.line }, { type: 'none' }, { type: 'solid', pt: 0.75, color: C.line }, { type: 'none' }],
      margin: [0.02, 0.06, 0.02, 0], fill: { color: C.paper },
    });

    // UML class diagram + two objects
    card(s, 7.0, 1.7, 5.73, 3.85);
    label(s, 'One class, many objects', 7.2, 1.86, 5.3);
    const side = (t, y) => tx(s, t, 7.12, y, 0.8, 0.24, { fontSize: 9.5, color: C.muted, align: 'right', valign: 'middle' });
    side('class', 2.35); side('attributes', 2.92); side('methods', 3.62);
    rect(s, 7.98, 2.25, 1.95, 0.44, { fill: T25[3], line: P[3], lw: 1.5 });
    tx(s, 'Product', 7.98, 2.25, 1.95, 0.44, { fontFace: F.head, fontSize: 14, bold: true, align: 'center', valign: 'middle' });
    rect(s, 7.98, 2.69, 1.95, 0.68, { fill: C.white, line: C.ink2, lw: 1 });
    rect(s, 7.98, 3.37, 1.95, 0.68, { fill: C.white, line: C.ink2, lw: 1 });
    tx(s, '+ name: str\n+ price: float', 8.06, 2.74, 1.85, 0.58, { fontFace: F.code, fontSize: 9.5, lineSpacing: 16, valign: 'middle' });
    tx(s, '+ __init__(name, price)\n+ describe(): str', 8.06, 3.42, 1.85, 0.58, { fontFace: F.code, fontSize: 9.5, lineSpacing: 16, valign: 'middle' });
    const obj = (name, nm, price, y) => {
      rect(s, 10.75, y, 1.78, 0.36, { fill: C.white, line: C.ink2 });
      tx(s, name, 10.75, y, 1.78, 0.36, { fontFace: F.code, fontSize: 10, underline: { style: 'sng' }, align: 'center', valign: 'middle' });
      rect(s, 10.75, y + 0.36, 1.78, 0.6, { fill: C.white, line: C.ink2 });
      tx(s, `name = '${nm}'\nprice = ${price}`, 10.84, y + 0.4, 1.66, 0.52, { fontFace: F.code, fontSize: 10, lineSpacing: 16, valign: 'middle' });
    };
    obj('croissant : Product', 'Croissant', '1.25', 2.15);
    obj('sourdough : Product', 'Sourdough', '3.5', 3.45);
    line(s, 10.75, 2.63, 9.95, 3.0, { color: C.ink2, w: 1.1, dash: 'dash', end: 'arrow' });
    line(s, 10.75, 3.93, 9.95, 3.5, { color: C.ink2, w: 1.1, dash: 'dash', end: 'arrow' });
    tx(s, 'instance of', 9.95, 3.13, 0.8, 0.24, { fontSize: 9, italic: true, color: C.muted, align: 'center', valign: 'middle' });
    tx(s, 'the class = a blueprint', 7.98, 4.12, 1.95, 0.3, { fontFace: F.head, italic: true, fontSize: 11.5, color: C.ink2, align: 'center' });
    tx(s, 'objects made from it', 10.75, 4.47, 1.78, 0.3, { fontFace: F.head, italic: true, fontSize: 11.5, color: C.ink2, align: 'center' });
    tx(s, 'A UML class diagram (left) and two objects (right). Both objects have the same attributes, but each one has its own values.', 7.2, 4.86, 5.33, 0.6, { fontSize: 11.5, color: C.ink2 });

    label(s, 'The four main ideas of OOP  (Booch et al., 2007)', MX, 5.7, 8);
    const pillars = [
      ['Encapsulation', 'Data and methods live together in an object, and data can be kept private.'],
      ['Inheritance', 'A child class gets everything from a parent class, e.g. a Cake is a Product.'],
      ['Polymorphism', 'The same method name does different things in different classes.'],
      ['Abstraction', 'You use an object without seeing how it works inside, e.g. `basket.add()`.'],
    ];
    const pw = (CW - 3 * 0.2) / 4;
    pillars.forEach(([t, d], i) => {
      const px = MX + i * (pw + 0.2);
      rect(s, px, 5.98, pw, 0.92, { r: 0.1, fill: T14[3] });
      tx(s, [{ text: t, options: { fontFace: F.head, fontSize: 13.5, bold: true, breakLine: true } }, ...md(d, { fontSize: 11, color: C.ink2 })],
        px + 0.16, 6.05, pw - 0.3, 0.82, { paraSpaceAfter: 2 });
    });
    footer(s);
    s.addNotes('Object-oriented programming builds a program out of objects. Each object keeps its own data, called attributes, and its own actions, called methods. Objects are made from classes, which are like blueprints. The UML diagram shows the Product class with its attributes and methods, and two objects made from it, a croissant and a sourdough loaf, each with its own values. OOP has four main ideas: encapsulation, inheritance, polymorphism and abstraction.');
  }

  /* ---------- 10. OOP: code ---------- */
  {
    const s = newSlide();
    header(s, { num: 4, kicker: 'Object-oriented · code example', title: 'Example: products and a basket', p: 3 });
    editor(s, { x: MX, y: 1.7, w: 6.3, p: 3, code: CODE.oop, lang: 'python', file: 'basket_oop.py', tag: 'Python 3', size: 10.5, lineH: 13.5, hot: [18, 19, 20] });
    label(s, 'The objects after line 22', 7.25, 1.7, 5.48);
    const objCard = (x, y, w, head, attrs, o = {}) => {
      card(s, x, y, w, 0.98, { fill: o.fill || C.white, line: o.line || C.line, shadow: !o.fill });
      tx(s, [{ text: head[0], options: { underline: { style: 'sng' } } }, { text: ` : ${head[1]}`, options: {} }], x + 0.16, y + 0.1, w - 0.3, 0.28, { fontFace: F.code, fontSize: 11, valign: 'middle' });
      attrs.forEach(([k, v], i) => {
        tx(s, k, x + 0.16, y + 0.44 + i * 0.24, 1.0, 0.22, { fontFace: F.code, fontSize: 10.5, color: C.ink2, valign: 'middle' });
        tx(s, v, x + 1.0, y + 0.44 + i * 0.24, w - 1.16, 0.22, { fontFace: F.code, fontSize: 10.5, align: 'right', valign: 'middle' });
      });
    };
    objCard(7.25, 2.0, 2.6, ['croissant', 'Product'], [['name', "'Croissant'"], ['price', '1.25']]);
    objCard(10.13, 2.0, 2.6, ['sourdough', 'Product'], [['name', "'Sourdough'"], ['price', '3.5']]);
    line(s, 8.55, 3.38, 8.55, 3.0, { color: C.ink2, w: 1.25, end: 'triangle' });
    line(s, 11.43, 3.38, 11.43, 3.0, { color: C.ink2, w: 1.25, end: 'triangle' });
    tx(s, 'basket.items points to both objects', 8.75, 3.08, 2.6, 0.26, { fontFace: F.head, italic: true, fontSize: 11, color: C.ink2, align: 'center', valign: 'middle' });
    objCard(7.25, 3.38, 5.48, ['basket', 'Basket'], [['items', '[croissant, sourdough]'], ['total', '4.75']], { fill: T8[3], line: T25[3] });
    terminal(s, 7.25, 4.58, 5.48, 'Python shell', [
      ['cmd', '$ python -i basket_oop.py'],
      ['out', 'Croissant costs £1.25'],
      ['out', '4.75'],
      ['in', 'sourdough.describe()'],
      ['out', "'Sourdough costs £3.50'"],
      ['in', 'basket.add(Product("Cookie", 0.75))'],
      ['in', 'basket.total'],
      ['out', '5.5'],
    ], { lineH: 13.6 });
    footer(s);
    s.addNotes('This is the OOP version. Product is a class with a constructor and a describe method. Basket is a class with an add method that updates its own total. At the bottom I create objects from the classes and call their methods. On the right you can see the objects in memory: the basket\'s items list points to the two product objects. In the Python shell I can keep using the same objects, for example adding a cookie makes the total 5.5.');
  }

  /* ---------- 11. OOP: inheritance, pros & cons ---------- */
  {
    const s = newSlide();
    header(s, { num: 4, kicker: 'Object-oriented · pros & cons', title: 'OOP: inheritance, pros and cons', p: 3 });
    editor(s, { x: MX, y: 1.7, w: 6.75, p: 3, code: CODE.poly, lang: 'python', file: 'products.py', tag: 'Python 3', size: 10.5, lineH: 13.5, hot: [4, 6, 7, 12, 14, 17, 18, 22] });
    const ty = terminal(s, 7.7, 1.7, 5.03, 'Output', [['out', 'Croissant: £1.25'], ['out', 'Lemon cake, 8 slices'], ['out', '12']], { size: 11, lineH: 14.5 });
    prosCons(s, { x: 7.7, y: ty + 0.1, w: 5.03, h: 1.4, good: true, size: 12, items: [
      '**Organised:** each class looks after its own data',
      '**Reusable:** inheritance builds on existing classes',
      '**Safer data:** encapsulation stops accidental changes',
    ] });
    prosCons(s, { x: 7.7, y: ty + 1.6, w: 5.03, h: 1.4, good: false, size: 12, items: [
      '**More to learn:** classes, inheritance, polymorphism',
      '**Too big for small jobs:** scripts don\'t need classes',
      '**Can be slower:** objects use more memory',
    ] });
    label(s, 'Where it\'s used', 7.7, ty + 3.08, 5.03);
    chips(s, ['games (Unity, C#)', 'Android apps', 'business systems', 'GUI toolkits'], { x: 7.7, y: ty + 3.36, maxW: 5.03, size: 11 });
    footer(s);
    s.addNotes('This example shows the other OOP ideas. The price is private, so other code has to use get_price: that\'s encapsulation. Cake inherits from Product, so it gets get_price for free: that\'s inheritance. Both classes have a describe method, but they give different results: that\'s polymorphism. The output shows each object using its own describe. OOP keeps big programs organised, makes code reusable and protects data, but there\'s more to learn, it\'s too much for small jobs, and it can be slower.');
  }

  /* ---------- 12. comparison ---------- */
  {
    const s = newSlide();
    header(s, { num: 5, kicker: 'Comparison', title: 'Comparing the three paradigms' });
    const colW = [2.05, 3.36, 3.36, 3.36];
    const rowsData = [
      ['Main idea', 'A list of steps that run in order', 'Wait for events, then react to them', 'Objects that hold their own data and methods'],
      ['Built from', 'Procedures (functions)', 'Events, listeners and handlers', 'Classes and objects'],
      ['Who decides the order', 'The programmer: top to bottom', 'The user or the system: whatever happens next', 'Objects calling each other\'s methods'],
      ['Where the data lives', 'Separate from the code, passed around or global', 'In variables that the handlers share', 'Inside the objects, protected by encapsulation'],
      ['Best for', 'Small programs, scripts and calculations', 'GUIs, websites, games and sensors', 'Big programs, team projects, modelling real things'],
      ['Example languages', 'C, Pascal, BASIC', 'JavaScript, Visual Basic, C#', 'Java, C#, C++, Python'],
      ['In my example', '`subtotal = add_up(basket)`', '`click → add_croissant()`', '`basket.add(croissant)`'],
    ];
    const rows = [[
      { text: '', options: { fill: { color: C.paper } } },
      ...[1, 2, 3].map(p => ({ text: PNAME[p], options: { fill: { color: T25[p] }, fontFace: F.head, fontSize: 15, bold: true, margin: [0.05, 0.1, 0.05, 0.62] } })),
    ]];
    rowsData.forEach(r => {
      rows.push([
        { text: r[0].toUpperCase(), options: { fontSize: 10, color: C.muted, fill: { color: C.paper } } },
        ...[1, 2, 3].map(p => {
          const t = r[p];
          const isCode = t.startsWith('`');
          return { text: isCode ? t.slice(1, -1) : t, options: { fill: { color: T8[p] }, fontFace: isCode ? F.code : F.body, fontSize: isCode ? 11.5 : 13 } };
        }),
      ]);
    });
    const rowH = [0.56, 0.6, 0.46, 0.6, 0.6, 0.6, 0.46, 0.46];
    s.addTable(rows, { x: MX, y: 1.7, w: CW, colW, rowH, color: C.ink, valign: 'middle', border: { type: 'solid', pt: 2.5, color: C.paper }, margin: [0.05, 0.12, 0.05, 0.12] });
    let hx = MX + colW[0];
    [1, 2, 3].forEach((p, i) => {
      badge(s, p, hx + 0.13, 1.7 + (0.56 - 0.4) / 2, 0.4, { soft: true });
      hx += colW[i + 1];
    });
    tx(s, 'Python appears in all three examples because it\'s multi-paradigm: it supports all three styles.', MX, 6.45, CW, 0.3, { fontSize: 12, color: C.ink2 });
    footer(s);
    s.addNotes('This table compares the three side by side. The biggest difference is who decides the order the code runs in: in procedural code it\'s the programmer, in event-driven code it\'s the user or the system, and in OOP it\'s objects calling each other\'s methods. Python appears in all three examples because it\'s multi-paradigm.');
  }

  /* ---------- 13. how they fit together ---------- */
  {
    const s = newSlide();
    header(s, { num: 5, kicker: 'Comparison', title: 'How they fit together' });
    tx(s, 'The three paradigms aren\'t rivals. Most real programs use all three at the same time, like this version of my till. The coloured bars show which paradigm each line belongs to.',
      MX, 1.5, CW, 0.55, { fontSize: 14, color: C.ink2 });
    editor(s, { x: MX, y: 2.15, w: 7.0, code: CODE.together, lang: 'python', file: 'till_app.py', tag: 'Python 3 · tkinter', size: 10, lineH: 13,
      rails: [
        { p: 3, col: 0, from: 3, to: 16 }, { p: 3, col: 0, from: 19, to: 19 },
        { p: 2, col: 1, from: 8, to: 9 }, { p: 2, col: 1, from: 12, to: 16 }, { p: 2, col: 1, from: 20, to: 20 },
        { p: 1, col: 2, from: 13, to: 16 },
      ] });
    const legend = [
      [3, 'Object-oriented: the outside', 'The program is a class, `TillApp`, and `app` is an object made from it.', 'lines 3–16 and 19'],
      [2, 'Event-driven: when code runs', '`add()` is linked to the button, so it runs on every click. `mainloop()` waits for clicks.', 'lines 8–9, 12–16 and 20'],
      [1, 'Procedural: the inside', 'Inside `add()`, the code is plain steps in order, with an `if` to make a decision.', 'lines 13–16'],
    ];
    legend.forEach(([p, t, d, ln], i) => {
      const ly = 2.15 + i * 1.27;
      card(s, 7.95, ly, 4.78, 1.15);
      badge(s, p, 8.13, ly + 0.2, 0.5);
      tx(s, t, 8.82, ly + 0.12, 3.8, 0.3, { fontFace: F.head, fontSize: 14, bold: true, valign: 'middle' });
      tx(s, md(d, { fontSize: 11.5, color: C.ink2 }), 8.82, ly + 0.43, 3.8, 0.45);
      tx(s, ln, 8.82, ly + 0.86, 3.8, 0.2, { fontFace: F.code, fontSize: 9.5, color: C.muted, valign: 'middle' });
    });
    tx(s, 'Python, C#, Java and JavaScript are multi-paradigm, so the question isn\'t which paradigm is best, but which one fits each part of a program.',
      7.95, 6.0, 4.78, 0.65, { fontSize: 11.5, color: C.ink2 });
    footer(s);
    s.addNotes('In real programs the paradigms work together. Here the till is a class, so the outside is object-oriented. The add method is linked to the button, so it\'s an event handler. And inside add, the code is plain procedural steps with an if. Python, C#, Java and JavaScript all let you mix the three, so the real question is which one fits each part of a program.');
  }

  /* ---------- 14. conclusion ---------- */
  {
    const s = newSlide({ dark: true });
    header(s, { num: 6, kicker: 'Conclusion', title: 'Conclusion', dark: true });
    tx(s, 'A paradigm is a way of organising a program. Procedural programming is about the steps, event-driven programming is about when the code runs, and object-oriented programming is about how the code and data are grouped together.',
      MX, 1.55, 11.6, 0.8, { fontSize: 15, color: C.nightInk2 });
    const rows = {
      1: ['Steps in order, grouped into functions.', 'A recipe.', 'Small programs, scripts and calculations.'],
      2: ['Waits for events and runs a handler for each one.', 'A shop assistant waiting for customers.', 'Anything with buttons, keys or sensors.'],
      3: ['Objects made from classes, each with its own data and methods.', 'The things in a bakery, each with its own job.', 'Big programs that need to stay organised.'],
    };
    const cw = (CW - 2 * 0.25) / 3;
    [1, 2, 3].forEach((p, i) => {
      const x = MX + i * (cw + 0.25);
      rect(s, x, 2.5, cw, 2.88, { r: 0.12, fill: C.night2, line: C.nightLine });
      badge(s, p, x + 0.22, 2.7, 0.5);
      tx(s, PNAME[p], x + 0.86, 2.7, cw - 1.0, 0.5, { fontFace: F.head, fontSize: 17, bold: true, color: C.white, valign: 'middle' });
      ['In one line', 'Think of it as', 'Best for'].forEach((k, j) => {
        tx(s, [{ text: k.toUpperCase(), options: { fontSize: 9, color: C.nightMuted, charSpacing: 1.2, breakLine: true } }, { text: rows[p][j], options: { fontSize: 12.5, color: C.white } }],
          x + 0.22, 3.36 + j * 0.66, cw - 0.4, 0.62);
      });
    });
    tx(s, 'Which one I would use', MX, 5.55, 6, 0.32, { fontFace: F.head, fontSize: 15, bold: true, color: C.white, valign: 'middle' });
    tx(s, 'For a quick calculation I\'d write procedural code. As soon as there\'s a window with buttons, I need event-driven code. Once a program grows, I\'d organise it into classes. In Python I can do all three in the same program, like the till on the last slide.',
      MX, 5.9, CW, 0.52, { fontSize: 13, color: C.nightInk2 });
    tx(s, 'My opinion: learn procedural first, because the other two are built on top of it.', MX, 6.5, CW, 0.32, { fontFace: F.head, italic: true, fontSize: 13.5, color: 'E19B2C', valign: 'middle' });
    footer(s, true);
    s.addNotes('To sum up: procedural programming is about the steps, event-driven programming is about when the code runs, and object-oriented programming is about how the code and data are grouped. For a quick calculation I\'d write procedural code; anything with buttons needs event-driven code; and once a program grows I\'d organise it into classes. In my opinion, procedural is the one to learn first, because the other two are built on top of it.');
  }

  /* ---------- 15. references + thank you ---------- */
  {
    const s = newSlide();
    header(s, { num: 7, kicker: 'References', title: 'References' });
    const refs = [
      ['Booch, G., Maksimchuk, R.A., Engle, M.W., Young, B.J., Conallen, J. and Houston, K.A. (2007) ', 'Object-oriented analysis and design with applications', '. 3rd edn. Upper Saddle River, NJ: Addison-Wesley.'],
      ['Kernighan, B.W. and Ritchie, D.M. (1988) ', 'The C programming language', '. 2nd edn. Englewood Cliffs, NJ: Prentice Hall.'],
      ['MDN Web Docs (no date a) ', 'Introduction to events', '. Available at: ', 'https://developer.mozilla.org/en-US/docs/Learn_web_development/Core/Scripting/Events', ' (Accessed: 1 October 2026).'],
      ['MDN Web Docs (no date b) ', 'EventTarget: addEventListener() method', '. Available at: ', 'https://developer.mozilla.org/en-US/docs/Web/API/EventTarget/addEventListener', ' (Accessed: 1 October 2026).'],
      ['Oracle (no date) ', 'The Java tutorials: object-oriented programming concepts', '. Available at: ', 'https://docs.oracle.com/javase/tutorial/java/concepts/', ' (Accessed: 1 October 2026).'],
      ['Python Software Foundation (no date a) ', 'The Python tutorial: 9. Classes', '. Available at: ', 'https://docs.python.org/3/tutorial/classes.html', ' (Accessed: 1 October 2026).'],
      ['Python Software Foundation (no date b) ', 'tkinter: Python interface to Tcl/Tk', '. Available at: ', 'https://docs.python.org/3/library/tkinter.html', ' (Accessed: 1 October 2026).'],
      ['Sebesta, R.W. (2016) ', 'Concepts of programming languages', '. 11th edn. Boston, MA: Pearson.'],
      ['Van Roy, P. (2009) \'Programming paradigms for dummies: what every programmer should know\', in Assayag, G. and Gerzso, A. (eds) ', 'New computational paradigms for computer music', '. Paris: IRCAM/Delatour France. Available at: ', 'https://www.info.ucl.ac.be/~pvr/VanRoyChapter.pdf', ' (Accessed: 1 October 2026).'],
    ];
    const runs = [];
    refs.forEach((r, i) => {
      const [before, title, after, url, tail] = r;
      const base = { fontSize: 10.5, color: C.ink };
      runs.push({ text: before, options: { ...base } });
      runs.push({ text: title, options: { ...base, italic: true } });
      runs.push({ text: after, options: { ...base } });
      if (url) {
        runs.push({ text: url, options: { ...base, color: '1F5FB4', hyperlink: { url } } });
        runs.push({ text: tail, options: { ...base } });
      }
      if (i < refs.length - 1) runs[runs.length - 1].options.breakLine = true;
    });
    tx(s, runs, MX, 1.62, 7.25, 4.6, { paraSpaceAfter: 5 });
    tx(s, 'Harvard style. The citations on the slides, like (Oracle, no date), point to this list.', MX, 5.55, 7.25, 0.3, { fontSize: 10.5, color: C.muted });

    // thank you + feedback
    rect(s, 8.3, 0.55, 4.43, 6.28, { r: 0.16, fill: C.night });
    label(s, 'That\'s the end', 8.62, 1.0, 3.8, { dark: true });
    tx(s, 'Thank you for listening!', 8.62, 1.32, 3.85, 1.1, { fontFace: F.head, fontSize: 28, bold: true, color: C.white });
    tx(s, 'Any questions?', 8.62, 2.52, 3.85, 0.42, { fontFace: F.head, fontSize: 19, color: C.nightInk2, valign: 'middle' });
    rect(s, 8.62, 3.4, 1.65, 1.65, { r: 0.1, fill: C.white, line: C.line2, lw: 1.25, dash: 'dash' });
    tx(s, 'Paste the QR code for your feedback form here', 8.72, 3.4, 1.45, 1.65, { fontSize: 10, color: C.muted, align: 'center', valign: 'middle' });
    tx(s, 'How was my presentation?', 10.45, 3.4, 2.05, 0.6, { fontFace: F.head, fontSize: 14, bold: true, color: C.white });
    tx(s, 'Scan the QR code with your phone camera to fill in my short feedback form on Microsoft Forms.', 10.45, 4.03, 2.05, 1.05, { fontSize: 11.5, color: C.nightInk2 });
    tx(s, `${NAME.toUpperCase()}  ·  STUDENT ID ${STUDENT_ID}`, 8.62, 6.3, 3.9, 0.3, { fontFace: F.code, fontSize: 10, color: C.nightMuted, valign: 'middle' });
    footer(s);
    s.addNotes('These are the sources I used, in Harvard style. Thank you for listening. Are there any questions? If you have a minute, please scan the QR code and fill in my feedback form.');
  }
}

// pptxgenjs writes an <a:pPr> before every run, so a paragraph with mixed
// formatting gets several (out of spec, and the later ones switch bullets off).
// Keep only the first one in each paragraph.
async function onePPrPerParagraph(buf) {
  const zip = await JSZip.loadAsync(buf);
  const parts = Object.keys(zip.files).filter(f => /^ppt\/(slides|notesSlides)\/[^/]+\.xml$/.test(f));
  let removed = 0;
  for (const f of parts) {
    const xml = await zip.file(f).async('string');
    const fixed = xml.replace(/<a:p>([\s\S]*?)<\/a:p>/g, (m, inner) => {
      let seen = false;
      const body = inner.replace(/<a:pPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:pPr>)/g, pPr => {
        if (!seen) { seen = true; return pPr; }
        removed++;
        return '';
      });
      return `<a:p>${body}</a:p>`;
    });
    zip.file(f, fixed);
  }
  return { buf: await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }), removed };
}

(async () => {
  await makeIcons();
  build();
  const raw = await pres.write({ outputType: 'nodebuffer' });
  const { buf, removed } = await onePPrPerParagraph(raw);
  fs.writeFileSync(OUT, buf);
  console.log('wrote', OUT, `(removed ${removed} extra paragraph-property blocks)`);
})().catch(err => { console.error(err); process.exit(1); });
