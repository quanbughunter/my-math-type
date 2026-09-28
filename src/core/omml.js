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

function run(text, rpr = {}) {
  if (!text) return '';
  let r = '';
  if (rpr.nor) r += '<m:nor/>';
  else {
    if (rpr.scr) r += `<m:scr m:val="${rpr.scr}"/>`;
    if (rpr.sty) r += `<m:sty m:val="${rpr.sty}"/>`;
  }
  if (rpr.aln) r += '<m:aln/>';
  const mrpr = r ? `<m:rPr>${r}</m:rPr>` : '';
  const wrpr = rpr.bold ? '<w:rPr><w:b/></w:rPr>' : '';
  return `<m:r>${mrpr}${wrpr}<m:t xml:space="preserve">${esc(text)}</m:t></m:r>`;
}

function token(el) {
  const name = ln(el);
  const t = tokenText(el);
  if (!t) return '';
  const mv = el.getAttribute('mathvariant');
  if (name === 'mtext' || name === 'ms') {
    const q = name === 'ms' ? `"${t}"` : t;
    return run(q, { nor: true, bold: mv && mv.includes('bold') });
  }
  if (mv && VARIANTS[mv]) return run(t, VARIANTS[mv]);
  if (name === 'mi') {
    // MathML: mi nhiều ký tự (sin, lim, max…) mặc định là chữ đứng
    return run(t, [...t].length > 1 ? { sty: 'p' } : {});
  }
  if (name === 'mo' && /[A-Za-z]/.test(t)) return run(t, { sty: 'p' });
  return run(t);
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
  if (n === 'msub' || n === 'munder') sub = conv(k[1]);
  if (n === 'msup' || n === 'mover') sup = conv(k[1]);
  if (n === 'msubsup' || n === 'munderover') { sub = conv(k[1]); sup = conv(k[2]); }
  if (n === 'msub' || n === 'msup' || n === 'msubsup') loc = 'subSup';
  if (n === 'mo') loc = /[∫-∳⨌⨐-⨗]/.test(chr) ? 'subSup' : 'undOvr';
  let pr = `<m:chr m:val="${esc(chr)}"/><m:limLoc m:val="${loc}"/>`;
  if (!sub) pr += '<m:subHide m:val="1"/>';
  if (!sup) pr += '<m:supHide m:val="1"/>';
  return `<m:nary><m:naryPr>${pr}</m:naryPr>${wrap('sub', sub)}${wrap('sup', sup)}${wrap('e', bodyXml)}</m:nary>`;
}

function isEmptyBase(el) {
  const b = kids(el)[0];
  return !b || (ln(b) === 'mrow' && kids(b).length === 0 && !tokenText(b).trim());
}

function isApply(el) {
  return el && ln(el) === 'mo' && txt(el).trim() === APPLY_FN;
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
      const sub = name === 'msup' ? '' : conv(k[1]);
      const sup = name === 'msub' ? '' : conv(k[name === 'msubsup' ? 2 : 1]);
      out += `<m:sPre>${wrap('sub', sub)}${wrap('sup', sup)}${wrap('e', conv(nodes[i + 1]))}</m:sPre>`;
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
      out += `<m:func><m:fName>${fname}</m:fName>${wrap('e', arg)}</m:func>`;
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
      const pr = type ? `<m:fPr><m:type m:val="${type}"/></m:fPr>` : '';
      return `<m:f>${pr}${wrap('num', conv(k[0]))}${wrap('den', conv(k[1]))}</m:f>`;
    }
    case 'msqrt':
      return `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>${wrap('e', convRow(k))}</m:rad>`;
    case 'mroot':
      return `<m:rad>${wrap('deg', conv(k[1]))}${wrap('e', conv(k[0]))}</m:rad>`;
    case 'msub':
      if (isNaryEl(el)) return nary(el, '');
      return `<m:sSub>${wrap('e', conv(k[0]))}${wrap('sub', conv(k[1]))}</m:sSub>`;
    case 'msup':
      if (isNaryEl(el)) return nary(el, '');
      return `<m:sSup>${wrap('e', conv(k[0]))}${wrap('sup', conv(k[1]))}</m:sSup>`;
    case 'msubsup':
      if (isNaryEl(el)) return nary(el, '');
      return `<m:sSubSup>${wrap('e', conv(k[0]))}${wrap('sub', conv(k[1]))}${wrap('sup', conv(k[2]))}</m:sSubSup>`;
    case 'mover': case 'munder': case 'munderover':
      if (isNaryEl(el)) return nary(el, '');
      return underOver(el);
    case 'mfenced': {
      const open = el.hasAttribute('open') ? el.getAttribute('open') : '(';
      const close = el.hasAttribute('close') ? el.getAttribute('close') : ')';
      const inner = convRow(k);
      return `<m:d><m:dPr><m:begChr m:val="${esc(open)}"/><m:endChr m:val="${esc(close)}"/></m:dPr>${wrap('e', inner)}</m:d>`;
    }
    case 'mtable':
      return table(el);
    case 'menclose':
      return enclose(el);
    case 'mphantom':
      return `<m:phant><m:phantPr><m:show m:val="0"/></m:phantPr>${wrap('e', convRow(k))}</m:phant>`;
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

function underOver(el) {
  const name = ln(el);
  const [base, s1, s2] = kids(el);
  const baseArrow = singleMo(base);
  const isArrowBase = baseArrow && ARROWS.test(baseArrow);

  if (name === 'mover') {
    const ch = singleMo(s1);
    // mũi tên có chữ phía trên: \xrightarrow{t°}
    if (isArrowBase) {
      return `<m:groupChr><m:groupChrPr><m:chr m:val="${esc(baseArrow)}"/><m:vertJc m:val="bot"/></m:groupChrPr>${wrap('e', conv(s1))}</m:groupChr>`;
    }
    if (ch && (BRACES_TOP.has(ch) || ch === '\u2322')) {
      // \overbrace, cung \overset{\frown}{AB} → dấu nhóm co giãn phía trên
      const gc = ch === '\u2322' ? '\u23DC' : ch;
      return `<m:groupChr><m:groupChrPr><m:chr m:val="${esc(gc)}"/><m:pos m:val="top"/><m:vertJc m:val="bot"/></m:groupChrPr>${wrap('e', conv(base))}</m:groupChr>`;
    }
    if (ch && BAR_CHARS.has(ch) && el.getAttribute('accent') !== 'true') {
      return `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr>${wrap('e', conv(base))}</m:bar>`;
    }
    if (ch && ACCENT_CHR[ch] && [...ch].length === 1) {
      return `<m:acc><m:accPr><m:chr m:val="${ACCENT_CHR[ch]}"/></m:accPr>${wrap('e', conv(base))}</m:acc>`;
    }
    return `<m:limUpp>${wrap('e', conv(base))}${wrap('lim', conv(s1))}</m:limUpp>`;
  }

  if (name === 'munder') {
    const ch = singleMo(s1);
    if (isArrowBase) {
      return `<m:groupChr><m:groupChrPr><m:chr m:val="${esc(baseArrow)}"/><m:pos m:val="top"/><m:vertJc m:val="top"/></m:groupChrPr>${wrap('e', conv(s1))}</m:groupChr>`;
    }
    if (ch && BRACES_BOT.has(ch)) {
      return `<m:groupChr><m:groupChrPr><m:chr m:val="${esc(ch)}"/></m:groupChrPr>${wrap('e', conv(base))}</m:groupChr>`;
    }
    if (ch && BAR_CHARS.has(ch)) {
      return `<m:bar><m:barPr><m:pos m:val="bot"/></m:barPr>${wrap('e', conv(base))}</m:bar>`;
    }
    return `<m:limLow>${wrap('e', conv(base))}${wrap('lim', conv(s1))}</m:limLow>`;
  }

  // munderover
  if (isArrowBase) {
    const top = `<m:groupChr><m:groupChrPr><m:chr m:val="${esc(baseArrow)}"/><m:vertJc m:val="bot"/></m:groupChrPr>${wrap('e', conv(s2))}</m:groupChr>`;
    return `<m:limLow>${wrap('e', top)}${wrap('lim', conv(s1))}</m:limLow>`;
  }
  const low = `<m:limLow>${wrap('e', conv(base))}${wrap('lim', conv(s1))}</m:limLow>`;
  return `<m:limUpp>${wrap('e', low)}${wrap('lim', conv(s2))}</m:limUpp>`;
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
    return `<m:eqArr>${es.join('')}</m:eqArr>`;
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
  return `<m:m><m:mPr><m:plcHide m:val="1"/><m:mcs>${mcs}</m:mcs></m:mPr>${mrs.join('')}</m:m>`;
}

function enclose(el) {
  const notation = (el.getAttribute('notation') || 'box').split(/\s+/);
  const inner = convRow(kids(el));
  if (notation.includes('top') && notation.length === 1) return `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr>${wrap('e', inner)}</m:bar>`;
  if (notation.includes('bottom') && notation.length === 1) return `<m:bar><m:barPr><m:pos m:val="bot"/></m:barPr>${wrap('e', inner)}</m:bar>`;
  if (notation.includes('radical')) return `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>${wrap('e', inner)}</m:rad>`;
  const strike = [];
  if (notation.includes('updiagonalstrike')) strike.push('<m:strikeBLTR m:val="1"/>');
  if (notation.includes('downdiagonalstrike')) strike.push('<m:strikeTLBR m:val="1"/>');
  if (notation.includes('horizontalstrike')) strike.push('<m:strikeH m:val="1"/>');
  if (notation.includes('verticalstrike')) strike.push('<m:strikeV m:val="1"/>');
  const boxed = notation.some((n) => ['box', 'roundedbox', 'circle', 'left', 'right'].includes(n));
  if (strike.length && !boxed) {
    const hide = '<m:hideTop m:val="1"/><m:hideBot m:val="1"/><m:hideLeft m:val="1"/><m:hideRight m:val="1"/>';
    return `<m:borderBox><m:borderBoxPr>${hide}${strike.join('')}</m:borderBoxPr>${wrap('e', inner)}</m:borderBox>`;
  }
  const pr = strike.length ? `<m:borderBoxPr>${strike.join('')}</m:borderBoxPr>` : '';
  return `<m:borderBox>${pr}${wrap('e', inner)}</m:borderBox>`;
}

function multiscripts(el) {
  const k = kids(el);
  const base = k[0];
  const pi = k.findIndex((c) => ln(c) === 'mprescripts');
  const post = pi < 0 ? k.slice(1) : k.slice(1, pi);
  const pre = pi < 0 ? [] : k.slice(pi + 1);
  let x = conv(base);
  if (post.length >= 2) {
    x = `<m:sSubSup>${wrap('e', x)}${wrap('sub', conv(post[0]))}${wrap('sup', conv(post[1]))}</m:sSubSup>`;
  }
  if (pre.length >= 2) {
    x = `<m:sPre>${wrap('sub', conv(pre[0]))}${wrap('sup', conv(pre[1]))}${wrap('e', x)}</m:sPre>`;
  }
  return x;
}

/** Phần tử <math> đã chuẩn hoá → OMML (<m:oMath> hoặc <m:oMathPara>). */
export function mathRootToOmml(root, { display = true } = {}) {
  const body = convRow(kids(root));
  const ns = `xmlns:m="${OMML_NS}" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"`;
  if (display) return `<m:oMathPara ${ns}><m:oMath>${body}</m:oMath></m:oMathPara>`;
  return `<m:oMath ${ns}>${body}</m:oMath>`;
}

/** Tiện: LaTeX → OMML (không kèm khai báo namespace nếu bare=true). */
export function latexToOmml(latex, { display = true, bare = false } = {}) {
  const { root } = latexToWordMathDoc(latex, { display });
  const x = mathRootToOmml(root, { display });
  return bare ? stripNs(x) : x;
}

export function stripNs(x) {
  return x.replace(/ xmlns:m="[^"]*"/, '').replace(/ xmlns:w="[^"]*"/, '');
}

export { NARY_RE };
