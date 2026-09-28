// MathML (đã chuẩn hoá) → OMML — định dạng phương trình gốc của Word.
// Tự viết để kiểm soát chất lượng: ngoặc co giãn (m:d), tổng/tích phân có thân (m:nary),
// hàm số (m:func), hệ phương trình, ma trận, dấu mũ/vector (m:acc), gạch trên (m:bar)…

import { ln, kids, txt } from './dom.js';
import { NARY_RE, APPLY_FN, isNaryEl, isStopper, latexToWordMathDoc } from './mathml.js';

export const OMML_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const INVISIBLE = /[⁡-⁤​]/g;

const VARIANTS = {
  normal: { sty: 'p' },
  bold: { sty: 'b' },
  italic: { sty: 'i' },
  'bold-italic': { sty: 'bi' },
  'double-struck': { scr: 'double-struck', sty: 'p' },
  script: { scr: 'script', sty: 'p' },
  'bold-script': { scr: 'script', sty: 'b' },
  fraktur: { scr: 'fraktur', sty: 'p' },
  'bold-fraktur': { scr: 'fraktur', sty: 'b' },
  'sans-serif': { scr: 'sans-serif', sty: 'p' },
  'bold-sans-serif': { scr: 'sans-serif', sty: 'b' },
  'sans-serif-italic': { scr: 'sans-serif', sty: 'i' },
  monospace: { scr: 'monospace', sty: 'p' },
};

// ký tự phía trên → ký tự "kết hợp" mà Word dùng cho m:acc
const ACCENT_CHR = {
  '^': '̂', 'ˆ': '̂', '̂': '̂',
  '~': '̃', '˜': '̃', '̃': '̃',
  '¯': '̅', '‾': '̅', '̅': '̅', '̄': '̄', '_': '̅', '−': '̅', '-': '̅',
  '˙': '̇', '.': '̇', '̇': '̇',
  '¨': '̈', '̈': '̈',
  'ˇ': '̌', '̌': '̌',
  '˘': '̆', '̆': '̆',
  '´': '́', '́': '́',
  '`': '̀', '̀': '̀',
  '˚': '̊', '̊': '̊',
  '→': '⃗', '⟶': '⃗', '⃗': '⃗',
  '←': '⃖', '⟵': '⃖', '⃖': '⃖',
  '↔': '⃡', '⟷': '⃡', '⃡': '⃡',
};
const BAR_CHARS = new Set(['¯', '‾', '_', '−', '-', '̲', '̅']);
const BRACES_TOP = new Set(['⏞', '⏜', '⎴', '︷']);
const BRACES_BOT = new Set(['⏟', '⏝', '⎵', '︸']);
const ARROWS = /^[←-⇿⟰-⟿⤀-⥿]$/;

function tokenText(el) {
  return txt(el).replace(INVISIBLE, '');
}

// ---------- Định dạng khi đưa vào Word (font, cỡ chữ) ----------
// fmt = { font: 'Cambria Math' | 'Times New Roman' | …, size: 12 (pt), subSize: 9 (pt, 0 = Word tự tính) }
let FMT = null;
let SCRIPT = 0; // đang ở chỉ số trên/dưới, bậc căn, cận… — Word TỰ thu nhỏ chữ ở đây
let SMALL = 0;  // chữ ghi trên/dưới mũi tên — Word KHÔNG tự thu nhỏ

/** Font toán OpenType mà Word dùng được làm font của phương trình. */
export const MATH_FONTS = ['Cambria Math', 'STIX Two Math', 'Latin Modern Math', 'TeX Gyre Termes Math', 'XITS Math', 'Asana Math'];
export const isMathFont = (f) => !f || /\bmath$/i.test(String(f).trim());

function normFmt(f) {
  if (!f) return null;
  const font = String(f.font || '').trim();
  const size = +f.size > 0 ? +f.size : 0;
  const subSize = +f.subSize > 0 ? +f.subSize : 0;
  if (!font && !size && !subSize) return null;
  // Font văn bản (Times New Roman…): chữ & số dùng font đó ở chế độ "Normal Text" như MathType,
  // còn dấu phép toán, ngoặc, ∑… vẫn dùng font toán để Word giữ khoảng cách chuẩn.
  const textMode = !!font && !isMathFont(font);
  const mathFont = textMode ? 'Cambria Math' : font;
  // tỉ lệ Word thu nhỏ chỉ số (bảng MATH của font toán)
  const ratio = !mathFont || /cambria/i.test(mathFont) ? 0.73 : 0.7;
  return { font, size, subSize, textMode, mathFont, ratio };
}

function halfPts() {
  if (!FMT) return 0;
  let pt = FMT.size;
  if (SCRIPT > 0 && FMT.subSize) pt = FMT.subSize / FMT.ratio;
  else if (SMALL > 0 && FMT.subSize) pt = FMT.subSize;
  return pt ? Math.round(pt * 2) : 0;
}

function wRPr({ font, bold, italic } = {}) {
  let x = '';
  if (font) x += `<w:rFonts w:ascii="${esc(font)}" w:hAnsi="${esc(font)}" w:cs="${esc(font)}"/>`;
  if (bold) x += '<w:b/>';
  if (italic) x += '<w:i/>';
  const hp = halfPts();
  if (hp) x += `<w:sz w:val="${hp}"/><w:szCs w:val="${hp}"/>`;
  return x ? `<w:rPr>${x}</w:rPr>` : '';
}

/** m:ctrlPr — định dạng cho ký hiệu của chính cấu trúc (∑, ngoặc, dấu căn, gạch phân số…). */
function ctrl() {
  if (!FMT) return '';
  const w = wRPr({ font: FMT.mathFont });
  return w ? `<m:ctrlPr>${w}</m:ctrlPr>` : '';
}
/** Khối thuộc tính <m:xxxPr> (ctrlPr luôn đứng cuối theo lược đồ OMML). */
function pr(name, inner = '') {
  const c = ctrl();
  return inner || c ? `<m:${name}Pr>${inner}${c}</m:${name}Pr>` : '';
}
function inScript(fn) {
  SCRIPT++;
  try { return fn(); } finally { SCRIPT--; }
}
function inSmall(fn) {
  SMALL++;
  try { return fn(); } finally { SMALL--; }
}

/**
 * Một "run" chữ trong phương trình.
 * kind: 'ident' (mi) | 'num' (mn) | 'op' (mo) | 'text' (mtext)
 */
function run(text, rpr = {}, kind = 'op') {
  if (!text) return '';
  let { nor, sty, scr, bold } = rpr;
  let italic = false;
  let font = null;
  if (FMT) {
    // ký tự toán đặc biệt (𝐯, ℝ…) font văn bản không có → giữ ở chế độ toán
    const special = /[\u{1D400}-\u{1D7FF}\u2100-\u214F]/u.test(text);
    if (FMT.textMode && kind !== 'op' && !scr && !special) {
      if (!nor) {
        // kiểu MathType: biến một chữ cái Latin in nghiêng; số, hàm, chữ Hy Lạp đứng
        italic = sty === 'i' || sty === 'bi' || (!sty && kind === 'ident' && /^[A-Za-z]$/.test(text));
        bold = bold || sty === 'b' || sty === 'bi';
        nor = true;
        sty = undefined;
      }
      font = FMT.font;
    } else if (nor) {
      font = FMT.textMode ? FMT.font : null; // chữ thường theo font của đoạn văn
    } else {
      font = FMT.mathFont || null;
    }
  }
  let r = '';
  if (nor) r += '<m:nor/>';
  else {
    if (scr) r += `<m:scr m:val="${scr}"/>`;
    if (sty) r += `<m:sty m:val="${sty}"/>`;
  }
  if (rpr.aln) r += '<m:aln/>';
  const mrpr = r ? `<m:rPr>${r}</m:rPr>` : '';
  return `<m:r>${mrpr}${wRPr({ font, bold, italic })}<m:t xml:space="preserve">${esc(text)}</m:t></m:r>`;
}

function token(el) {
  const name = ln(el);
  const t = tokenText(el);
  if (!t) return '';
  const mv = el.getAttribute('mathvariant');
  if (name === 'mtext' || name === 'ms') {
    const q = name === 'ms' ? `"${t}"` : t;
    return run(q, { nor: true, bold: !!(mv && mv.includes('bold')) }, 'text');
  }
  const kind = name === 'mi' ? 'ident' : name === 'mn' ? 'num' : 'op';
  if (mv && VARIANTS[mv]) return run(t, VARIANTS[mv], kind);
  if (name === 'mi') {
    // MathML: mi nhiều ký tự (sin, lim, max…) mặc định là chữ đứng
    return run(t, [...t].length > 1 ? { sty: 'p' } : {}, kind);
  }
  if (name === 'mo' && /[A-Za-z]/.test(t)) return run(t, { sty: 'p' }, kind);
  return run(t, {}, kind);
}

const wrap = (tag, inner) => (inner ? `<m:${tag}>${inner}</m:${tag}>` : `<m:${tag}/>`);

function naryChar(el) {
  if (ln(el) === 'mo') return tokenText(el).trim();
  return tokenText(kids(el)[0]).trim();
}

function nary(el, bodyXml) {
  const n = ln(el);
  const chr = naryChar(el);
  const k = kids(el);
  let sub = '', sup = '', loc = 'undOvr';
  if (n === 'msub' || n === 'munder') sub = inScript(() => conv(k[1]));
  if (n === 'msup' || n === 'mover') sup = inScript(() => conv(k[1]));
  if (n === 'msubsup' || n === 'munderover') { sub = inScript(() => conv(k[1])); sup = inScript(() => conv(k[2])); }
  if (n === 'msub' || n === 'msup' || n === 'msubsup') loc = 'subSup';
  if (n === 'mo') loc = /[∫-∳⨌⨐-⨗]/.test(chr) ? 'subSup' : 'undOvr';
  let p = `<m:chr m:val="${esc(chr)}"/><m:limLoc m:val="${loc}"/>`;
  if (!sub) p += '<m:subHide m:val="1"/>';
  if (!sup) p += '<m:supHide m:val="1"/>';
  return `<m:nary>${pr('nary', p)}${wrap('sub', sub)}${wrap('sup', sup)}${wrap('e', bodyXml)}</m:nary>`;
}

function isEmptyBase(el) {
  const b = kids(el)[0];
  return !b || (ln(b) === 'mrow' && kids(b).length === 0 && !tokenText(b).trim());
}

function isApply(el) {
  return el && ln(el) === 'mo' && txt(el).trim() === APPLY_FN;
}

function sPre(subEl, supEl, baseXml) {
  const sub = subEl ? inScript(() => conv(subEl)) : '';
  const sup = supEl ? inScript(() => conv(supEl)) : '';
  return `<m:sPre>${pr('sPre')}${wrap('sub', sub)}${wrap('sup', sup)}${wrap('e', baseXml)}</m:sPre>`;
}

/** Chuyển một dãy phần tử cùng hàng. */
function convRow(nodes) {
  let out = '';
  for (let i = 0; i < nodes.length; i++) {
    const n = nodes[i];
    const name = ln(n);

    // chỉ số trước: {}_{a}^{b}X → m:sPre
    if ((name === 'msub' || name === 'msup' || name === 'msubsup') && isEmptyBase(n) && i + 1 < nodes.length && !isStopper(nodes[i + 1])) {
      const k = kids(n);
      const subEl = name === 'msup' ? null : k[1];
      const supEl = name === 'msub' ? null : k[name === 'msubsup' ? 2 : 1];
      out += sPre(subEl, supEl, conv(nodes[i + 1]));
      i++;
      continue;
    }

    // toán tử lớn + thân
    if (isNaryEl(n)) {
      const body = [];
      let j = i + 1;
      while (j < nodes.length && !isStopper(nodes[j])) body.push(nodes[j++]);
      out += nary(n, convRow(body));
      i = j - 1;
      continue;
    }

    // hàm số: sin⁡x, log₂⁡x, lim_{x→0}⁡f
    if (i + 1 < nodes.length && isApply(nodes[i + 1])) {
      const fname = conv(n);
      const arg = i + 2 < nodes.length ? conv(nodes[i + 2]) : '';
      out += `<m:func>${pr('func')}<m:fName>${fname}</m:fName>${wrap('e', arg)}</m:func>`;
      i += 2;
      continue;
    }
    out += conv(n);
  }
  return out;
}

function conv(el) {
  if (!el) return '';
  const name = ln(el);
  const k = kids(el);
  switch (name) {
    case 'mi': case 'mn': case 'mo': case 'mtext': case 'ms':
      return token(el);
    case 'mspace': case 'none': case 'mprescripts': case 'annotation': case 'annotation-xml':
      return '';
    case 'mfrac': {
      const lt = (el.getAttribute('linethickness') || '').trim();
      const noBar = /^0(\.0*)?(px|pt|em|ex)?$/.test(lt);
      const type = noBar ? 'noBar' : el.getAttribute('bevelled') === 'true' ? 'skw' : '';
      const p = pr('f', type ? `<m:type m:val="${type}"/>` : '');
      return `<m:f>${p}${wrap('num', conv(k[0]))}${wrap('den', conv(k[1]))}</m:f>`;
    }
    case 'msqrt':
      return `<m:rad>${pr('rad', '<m:degHide m:val="1"/>')}<m:deg/>${wrap('e', convRow(k))}</m:rad>`;
    case 'mroot':
      return `<m:rad>${pr('rad')}${wrap('deg', inScript(() => conv(k[1])))}${wrap('e', conv(k[0]))}</m:rad>`;
    case 'msub':
      if (isNaryEl(el)) return nary(el, '');
      return `<m:sSub>${pr('sSub')}${wrap('e', conv(k[0]))}${wrap('sub', inScript(() => conv(k[1])))}</m:sSub>`;
    case 'msup':
      if (isNaryEl(el)) return nary(el, '');
      return `<m:sSup>${pr('sSup')}${wrap('e', conv(k[0]))}${wrap('sup', inScript(() => conv(k[1])))}</m:sSup>`;
    case 'msubsup':
      if (isNaryEl(el)) return nary(el, '');
      return `<m:sSubSup>${pr('sSubSup')}${wrap('e', conv(k[0]))}${wrap('sub', inScript(() => conv(k[1])))}${wrap('sup', inScript(() => conv(k[2])))}</m:sSubSup>`;
    case 'mover': case 'munder': case 'munderover':
      if (isNaryEl(el)) return nary(el, '');
      return underOver(el);
    case 'mfenced': {
      const open = el.hasAttribute('open') ? el.getAttribute('open') : '(';
      const close = el.hasAttribute('close') ? el.getAttribute('close') : ')';
      const inner = convRow(k);
      return `<m:d>${pr('d', `<m:begChr m:val="${esc(open)}"/><m:endChr m:val="${esc(close)}"/>`)}${wrap('e', inner)}</m:d>`;
    }
    case 'mtable':
      return table(el);
    case 'menclose':
      return enclose(el);
    case 'mphantom':
      return `<m:phant>${pr('phant', '<m:show m:val="0"/>')}${wrap('e', convRow(k))}</m:phant>`;
    case 'mmultiscripts':
      return multiscripts(el);
    case 'semantics':
      return conv(k[0]);
    default:
      // math, mrow, mstyle, mpadded, merror, mtd…
      return convRow(k);
  }
}

function singleMo(el) {
  return el && ln(el) === 'mo' ? tokenText(el).trim() : null;
}

function groupChr(chr, inner, { pos, vertJc } = {}) {
  let p = `<m:chr m:val="${esc(chr)}"/>`;
  if (pos) p += `<m:pos m:val="${pos}"/>`;
  if (vertJc) p += `<m:vertJc m:val="${vertJc}"/>`;
  return `<m:groupChr>${pr('groupChr', p)}${wrap('e', inner)}</m:groupChr>`;
}
const limLow = (e, lim) => `<m:limLow>${pr('limLow')}${wrap('e', e)}${wrap('lim', lim)}</m:limLow>`;
const limUpp = (e, lim) => `<m:limUpp>${pr('limUpp')}${wrap('e', e)}${wrap('lim', lim)}</m:limUpp>`;
const bar = (pos, e) => `<m:bar>${pr('bar', `<m:pos m:val="${pos}"/>`)}${wrap('e', e)}</m:bar>`;

function underOver(el) {
  const name = ln(el);
  const [base, s1, s2] = kids(el);
  const baseArrow = singleMo(base);
  const isArrowBase = baseArrow && ARROWS.test(baseArrow);

  if (name === 'mover') {
    const ch = singleMo(s1);
    // mũi tên có chữ phía trên: \xrightarrow{t°}
    if (isArrowBase) return groupChr(baseArrow, inSmall(() => conv(s1)), { vertJc: 'bot' });
    if (ch && (BRACES_TOP.has(ch) || ch === '⌢')) {
      // \overbrace, cung \overparen{AB} → dấu nhóm co giãn phía trên
      const gc = ch === '⌢' ? '⏜' : ch;
      return groupChr(gc, conv(base), { pos: 'top', vertJc: 'bot' });
    }
    if (ch && BAR_CHARS.has(ch) && el.getAttribute('accent') !== 'true') return bar('top', conv(base));
    if (ch && ACCENT_CHR[ch] && [...ch].length === 1) {
      return `<m:acc>${pr('acc', `<m:chr m:val="${ACCENT_CHR[ch]}"/>`)}${wrap('e', conv(base))}</m:acc>`;
    }
    return limUpp(conv(base), inScript(() => conv(s1)));
  }

  if (name === 'munder') {
    const ch = singleMo(s1);
    if (isArrowBase) return groupChr(baseArrow, inSmall(() => conv(s1)), { pos: 'top', vertJc: 'top' });
    if (ch && BRACES_BOT.has(ch)) return groupChr(ch, conv(base));
    if (ch && BAR_CHARS.has(ch)) return bar('bot', conv(base));
    return limLow(conv(base), inScript(() => conv(s1)));
  }

  // munderover
  if (isArrowBase) {
    const top = groupChr(baseArrow, inSmall(() => conv(s2)), { vertJc: 'bot' });
    return limLow(top, inScript(() => conv(s1)));
  }
  const low = limLow(conv(base), inScript(() => conv(s1)));
  return limUpp(low, inScript(() => conv(s2)));
}

function injectAln(xml) {
  if (xml.startsWith('<m:r><m:rPr>')) return xml.replace('</m:rPr>', '<m:aln/></m:rPr>');
  if (xml.startsWith('<m:r>')) return xml.replace('<m:r>', '<m:r><m:rPr><m:aln/></m:rPr>');
  return '<m:r><m:rPr><m:aln/></m:rPr><m:t></m:t></m:r>' + xml;
}

function table(el) {
  const rows = kids(el).filter((r) => ln(r) === 'mtr' || ln(r) === 'mlabeledtr');
  const cellsOf = (r) => kids(r).filter((c) => ln(c) === 'mtd').slice(ln(r) === 'mlabeledtr' ? 1 : 0);
  const ncols = Math.max(1, ...rows.map((r) => cellsOf(r).length));
  const aligns = (el.getAttribute('columnalign') || '').split(/\s+/).filter(Boolean);
  const alignOf = (i) => {
    const a = aligns[Math.min(i, aligns.length - 1)] || 'center';
    return a === 'left' || a === 'right' ? a : 'center';
  };

  // môi trường aligned: cột phải|trái xen kẽ → m:eqArr với điểm căn (m:aln)
  const isAligned = ncols >= 2 && ncols % 2 === 0 && alignOf(0) === 'right' && alignOf(1) === 'left';
  if (isAligned) {
    const es = rows.map((r) => {
      const cells = cellsOf(r);
      let x = '';
      cells.forEach((c, i) => {
        const cx = convRow(kids(c));
        x += i % 2 === 1 ? injectAln(cx) : cx;
      });
      return wrap('e', x);
    });
    return `<m:eqArr>${pr('eqArr')}${es.join('')}</m:eqArr>`;
  }

  let mcs = '';
  for (let i = 0; i < ncols; i++) {
    mcs += `<m:mc><m:mcPr><m:count m:val="1"/><m:mcJc m:val="${alignOf(i)}"/></m:mcPr></m:mc>`;
  }
  const mrs = rows.map((r) => {
    const cells = cellsOf(r);
    let x = '';
    for (let i = 0; i < ncols; i++) x += `<m:e>${cells[i] ? convRow(kids(cells[i])) : ''}</m:e>`;
    return `<m:mr>${x}</m:mr>`;
  });
  return `<m:m>${pr('m', `<m:plcHide m:val="1"/><m:mcs>${mcs}</m:mcs>`)}${mrs.join('')}</m:m>`;
}

function enclose(el) {
  const notation = (el.getAttribute('notation') || 'box').split(/\s+/);
  const inner = convRow(kids(el));
  if (notation.includes('top') && notation.length === 1) return bar('top', inner);
  if (notation.includes('bottom') && notation.length === 1) return bar('bot', inner);
  if (notation.includes('radical')) return `<m:rad>${pr('rad', '<m:degHide m:val="1"/>')}<m:deg/>${wrap('e', inner)}</m:rad>`;
  const strike = [];
  if (notation.includes('updiagonalstrike')) strike.push('<m:strikeBLTR m:val="1"/>');
  if (notation.includes('downdiagonalstrike')) strike.push('<m:strikeTLBR m:val="1"/>');
  if (notation.includes('horizontalstrike')) strike.push('<m:strikeH m:val="1"/>');
  if (notation.includes('verticalstrike')) strike.push('<m:strikeV m:val="1"/>');
  const boxed = notation.some((n) => ['box', 'roundedbox', 'circle', 'left', 'right'].includes(n));
  if (strike.length && !boxed) {
    const hide = '<m:hideTop m:val="1"/><m:hideBot m:val="1"/><m:hideLeft m:val="1"/><m:hideRight m:val="1"/>';
    return `<m:borderBox>${pr('borderBox', hide + strike.join(''))}${wrap('e', inner)}</m:borderBox>`;
  }
  return `<m:borderBox>${pr('borderBox', strike.join(''))}${wrap('e', inner)}</m:borderBox>`;
}

function multiscripts(el) {
  const k = kids(el);
  const base = k[0];
  const pi = k.findIndex((c) => ln(c) === 'mprescripts');
  const post = pi < 0 ? k.slice(1) : k.slice(1, pi);
  const pre = pi < 0 ? [] : k.slice(pi + 1);
  let x = conv(base);
  if (post.length >= 2) {
    x = `<m:sSubSup>${pr('sSubSup')}${wrap('e', x)}${wrap('sub', inScript(() => conv(post[0])))}${wrap('sup', inScript(() => conv(post[1])))}</m:sSubSup>`;
  }
  if (pre.length >= 2) x = sPre(pre[0], pre[1], x);
  return x;
}

/**
 * Phần tử <math> đã chuẩn hoá → OMML (<m:oMath> hoặc <m:oMathPara>).
 * format: { font, size, subSize } — định dạng chữ khi đưa vào Word (tuỳ chọn).
 */
export function mathRootToOmml(root, { display = true, format = null } = {}) {
  FMT = normFmt(format);
  SCRIPT = 0;
  SMALL = 0;
  try {
    const body = convRow(kids(root));
    const ns = `xmlns:m="${OMML_NS}" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;
    if (display) return `<m:oMathPara ${ns}><m:oMath>${body}</m:oMath></m:oMathPara>`;
    return `<m:oMath ${ns}>${body}</m:oMath>`;
  } finally {
    FMT = null;
  }
}

/** Tiện: LaTeX → OMML (không kèm khai báo namespace nếu bare=true). */
export function latexToOmml(latex, { display = true, bare = false, format = null } = {}) {
  const { root } = latexToWordMathDoc(latex, { display });
  const x = mathRootToOmml(root, { display, format });
  return bare ? stripNs(x) : x;
}

export function stripNs(x) {
  return x.replace(/ xmlns:m="[^"]*"/, '').replace(/ xmlns:w="[^"]*"/, '');
}

export { NARY_RE };
