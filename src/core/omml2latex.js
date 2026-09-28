// OMML (phương trình Word) → LaTeX, để mở lại công thức có sẵn trong Word và sửa tiếp.

import { ln, kids, parseXml } from './dom.js';

const CHAR_TEX = {
  'α': '\\alpha', 'β': '\\beta', 'γ': '\\gamma', 'δ': '\\delta', 'ε': '\\varepsilon', 'ϵ': '\\epsilon', 'ζ': '\\zeta',
  'η': '\\eta', 'θ': '\\theta', 'ϑ': '\\vartheta', 'ι': '\\iota', 'κ': '\\kappa', 'λ': '\\lambda', 'μ': '\\mu',
  'ν': '\\nu', 'ξ': '\\xi', 'π': '\\pi', 'ϖ': '\\varpi', 'ρ': '\\rho', 'ϱ': '\\varrho', 'σ': '\\sigma', 'ς': '\\varsigma',
  'τ': '\\tau', 'υ': '\\upsilon', 'φ': '\\varphi', 'ϕ': '\\phi', 'χ': '\\chi', 'ψ': '\\psi', 'ω': '\\omega',
  'Γ': '\\Gamma', 'Δ': '\\Delta', 'Θ': '\\Theta', 'Λ': '\\Lambda', 'Ξ': '\\Xi', 'Π': '\\Pi', 'Σ': '\\Sigma',
  'Υ': '\\Upsilon', 'Φ': '\\Phi', 'Ψ': '\\Psi', 'Ω': '\\Omega',
  '±': '\\pm', '∓': '\\mp', '×': '\\times', '÷': '\\div', '·': '\\cdot', '⋅': '\\cdot', '∘': '\\circ', '∗': '\\ast',
  '≤': '\\le', '≥': '\\ge', '⩽': '\\leqslant', '⩾': '\\geqslant', '≠': '\\ne', '≈': '\\approx', '≡': '\\equiv',
  '≅': '\\cong', '∼': '\\sim', '≃': '\\simeq', '∝': '\\propto', '≪': '\\ll', '≫': '\\gg',
  '∞': '\\infty', '∂': '\\partial', '∇': '\\nabla', '∈': '\\in', '∉': '\\notin', '∋': '\\ni',
  '⊂': '\\subset', '⊃': '\\supset', '⊆': '\\subseteq', '⊇': '\\supseteq', '∪': '\\cup', '∩': '\\cap',
  '∅': '\\emptyset', '∖': '\\setminus', '→': '\\to', '←': '\\leftarrow', '↔': '\\leftrightarrow',
  '⇒': '\\Rightarrow', '⇐': '\\Leftarrow', '⇔': '\\Leftrightarrow', '⟶': '\\longrightarrow', '↦': '\\mapsto',
  '⇌': '\\rightleftharpoons', '↑': '\\uparrow', '↓': '\\downarrow',
  '∀': '\\forall', '∃': '\\exists', '¬': '\\neg', '∧': '\\land', '∨': '\\lor', '∠': '\\angle', '⊥': '\\perp',
  '∥': '\\parallel', '△': '\\triangle', '°': '^{\\circ}', '′': "'", '″': "''", '…': '\\ldots', '⋯': '\\cdots',
  '⋮': '\\vdots', '⋱': '\\ddots', 'ℏ': '\\hbar', 'ℓ': '\\ell', '∴': '\\therefore', '∵': '\\because',
  '−': '-', '∣': '\\mid', '‖': '\\Vert', '⟨': '\\langle', '⟩': '\\rangle', '⌊': '\\lfloor', '⌋': '\\rfloor',
  '⌈': '\\lceil', '⌉': '\\rceil', 'ℝ': '\\mathbb{R}', 'ℕ': '\\mathbb{N}', 'ℤ': '\\mathbb{Z}', 'ℚ': '\\mathbb{Q}',
  'ℂ': '\\mathbb{C}', ' ': '\\,', ' ': '\\:', ' ': '\\;', ' ': '\\ ', ' ': '\\quad',
  '{': '\\{', '}': '\\}', '%': '\\%', '#': '\\#', '&': '\\&', '$': '\\$', '_': '\\_',
};

const NARY_TEX = {
  '∑': '\\sum', '∏': '\\prod', '∐': '\\coprod', '∫': '\\int', '∬': '\\iint', '∭': '\\iiint', '∮': '\\oint',
  '∯': '\\oiint', '∰': '\\oiiint', '⋃': '\\bigcup', '⋂': '\\bigcap', '⋁': '\\bigvee', '⋀': '\\bigwedge',
  '⨁': '\\bigoplus', '⨂': '\\bigotimes', '⨀': '\\bigodot',
};

const ACC_TEX = {
  '̂': ['\\hat', '\\widehat'], '̃': ['\\tilde', '\\widetilde'], '̅': ['\\bar', '\\overline'],
  '̄': ['\\bar', '\\overline'], '⃗': ['\\vec', '\\overrightarrow'], '⃖': ['\\overleftarrow', '\\overleftarrow'],
  '⃡': ['\\overleftrightarrow', '\\overleftrightarrow'], '̇': ['\\dot', '\\dot'], '̈': ['\\ddot', '\\ddot'],
  '̌': ['\\check', '\\check'], '̆': ['\\breve', '\\breve'], '́': ['\\acute', '\\acute'],
  '̀': ['\\grave', '\\grave'], '̊': ['\\mathring', '\\mathring'],
};

const FUNCS = new Set(['sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'arcsin', 'arccos', 'arctan', 'sinh', 'cosh', 'tanh',
  'coth', 'log', 'ln', 'lg', 'exp', 'lim', 'max', 'min', 'sup', 'inf', 'det', 'dim', 'ker', 'deg', 'gcd', 'arg', 'Pr',
  'limsup', 'liminf', 'hom']);

const FENCE_TEX = {
  '(': '(', ')': ')', '[': '[', ']': ']', '{': '\\{', '}': '\\}', '|': '|', '‖': '\\Vert', '⟨': '\\langle',
  '⟩': '\\rangle', '⌊': '\\lfloor', '⌋': '\\rfloor', '⌈': '\\lceil', '⌉': '\\rceil', '': '.',
};

function child(el, name) {
  return kids(el).find((k) => ln(k) === name) || null;
}
function prop(el, prName, key) {
  const pr = el && child(el, prName);
  const p = pr && child(pr, key);
  if (!p) return undefined;
  return p.getAttribute('m:val') ?? p.getAttributeNS?.('http://schemas.openxmlformats.org/officeDocument/2006/math', 'val') ?? '';
}
const onOff = (v) => v !== undefined && v !== '0' && v !== 'off' && v !== 'false';

function charsToTex(s) {
  let out = '';
  for (const ch of s) {
    const t = CHAR_TEX[ch];
    if (t) out += /^\\[a-zA-Z]+$/.test(t) ? t + ' ' : t;
    else if (ch === '\\') out += '\\backslash ';
    else out += ch;
  }
  return out;
}

const M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const mRPrOf = (r) => kids(r).find((k) => ln(k) === 'rPr' && k.namespaceURI === M_NS) || null;
const wRPrOf = (r) => kids(r).find((k) => ln(k) === 'rPr' && k.namespaceURI !== M_NS) || null;
const valOf = (el, key) => {
  const p = el && child(el, key);
  if (!p) return undefined;
  return p.getAttribute('m:val') ?? p.getAttributeNS?.(M_NS, 'val') ?? '';
};

function runTex(r) {
  const tEls = kids(r).filter((k) => ln(k) === 't');
  const text = tEls.map((t) => t.textContent || '').join('');
  if (!text) return '';
  const mrpr = mRPrOf(r);
  const wrpr = wRPrOf(r);
  const nor = mrpr && child(mrpr, 'nor');
  if (nor) {
    if (/^[\s -  ]+$/.test(text)) {
      const SP = { ' ': '\\,', ' ': '\\:', ' ': '\;', ' ': '\\quad ', ' ': '\\enspace ' };
      return [...text].map((c) => SP[c] || '\\ ').join('');
    }
    // Chữ "Normal Text" kiểu MathType (Times New Roman): biến nghiêng, số, tên hàm → trở lại dạng toán
    const italic = !!(wrpr && child(wrpr, 'i'));
    const bold = !!(wrpr && child(wrpr, 'b'));
    const wrapB = (x) => (bold ? `\\mathbf{${x}}` : x);
    if (/^[A-Za-z]$/.test(text)) return wrapB(italic ? text : `\\mathrm{${text}}`);
    if (italic && /^[A-Za-z]+$/.test(text)) return wrapB(text);
    if (/^\d+(?:[.,]\d+)*$/.test(text)) return text;
    if (FUNCS.has(text)) return `\\${text} `;
    if (/^[Ͱ-Ͽ]+$/.test(text)) return charsToTex(text);
    const clean = text.replace(/[\\{}$&#%_^~]/g, (c) => '\\' + c);
    return `\\text{${clean}}`;
  }
  const sty = valOf(mrpr, 'sty');
  const scr = valOf(mrpr, 'scr');
  let body = charsToTex(text);
  if (scr === 'double-struck') return `\\mathbb{${text}}`;
  if (scr === 'script') return `\\mathcal{${text}}`;
  if (scr === 'fraktur') return `\\mathfrak{${text}}`;
  if (scr === 'sans-serif') return `\\mathsf{${body}}`;
  if (scr === 'monospace') return `\\mathtt{${body}}`;
  if (sty === 'b') return `\\mathbf{${body}}`;
  if (sty === 'bi') return `\\boldsymbol{${body}}`;
  if (sty === 'p' && /[A-Za-z]/.test(text)) {
    if (FUNCS.has(text)) return `\\${text} `;
    return `\\mathrm{${body}}`;
  }
  return body;
}

function row(el) {
  return kids(el).map(o2l).join('');
}
function arg(el, name) {
  const c = child(el, name);
  return c ? row(c) : '';
}
const grp = (s) => `{${s}}`;

function o2l(el) {
  const n = ln(el);
  switch (n) {
    case 'r': return runTex(el);
    case 'f': {
      const type = prop(el, 'fPr', 'type');
      const num = arg(el, 'num'), den = arg(el, 'den');
      if (type === 'lin') return `${grp(num)}/${grp(den)}`;
      if (type === 'noBar') return `\\genfrac{}{}{0pt}{}${grp(num)}${grp(den)}`;
      return `\\frac${grp(num)}${grp(den)}`;
    }
    case 'rad': {
      const deg = arg(el, 'deg');
      const hide = onOff(prop(el, 'radPr', 'degHide'));
      return !hide && deg ? `\\sqrt[${deg}]${grp(arg(el, 'e'))}` : `\\sqrt${grp(arg(el, 'e'))}`;
    }
    case 'sSub': return `${grp(arg(el, 'e'))}_${grp(arg(el, 'sub'))}`;
    case 'sSup': return `${grp(arg(el, 'e'))}^${grp(arg(el, 'sup'))}`;
    case 'sSubSup': return `${grp(arg(el, 'e'))}_${grp(arg(el, 'sub'))}^${grp(arg(el, 'sup'))}`;
    case 'sPre': return `{}_${grp(arg(el, 'sub'))}^${grp(arg(el, 'sup'))}${grp(arg(el, 'e'))}`;
    case 'nary': {
      const chr = prop(el, 'naryPr', 'chr') ?? '∫';
      let cmd = NARY_TEX[chr] || charsToTex(chr);
      const sub = onOff(prop(el, 'naryPr', 'subHide')) ? '' : arg(el, 'sub');
      const sup = onOff(prop(el, 'naryPr', 'supHide')) ? '' : arg(el, 'sup');
      let s = cmd;
      if (sub) s += `_${grp(sub)}`;
      if (sup) s += `^${grp(sup)}`;
      return `${s} ${arg(el, 'e')}`;
    }
    case 'd': {
      const beg = prop(el, 'dPr', 'begChr') ?? '(';
      const end = prop(el, 'dPr', 'endChr') ?? ')';
      const sep = prop(el, 'dPr', 'sepChr') ?? '|';
      const es = kids(el).filter((k) => ln(k) === 'e');
      // hệ phương trình / cases
      if (beg === '{' && end === '' && es.length === 1) {
        const only = kids(es[0]);
        if (only.length === 1 && (ln(only[0]) === 'eqArr' || ln(only[0]) === 'm')) {
          return `\\begin{cases}${rowsOf(only[0], true)}\\end{cases}`;
        }
      }
      const inner = es.map(row).join(`\\middle${FENCE_TEX[sep] ?? sep}`);
      return `\\left${FENCE_TEX[beg] ?? beg}${inner}\\right${FENCE_TEX[end] ?? end}`;
    }
    case 'm': return `\\begin{matrix}${rowsOf(el)}\\end{matrix}`;
    case 'eqArr': {
      const hasAln = el.getElementsByTagName('m:aln').length > 0;
      return hasAln ? `\\begin{aligned}${rowsOf(el)}\\end{aligned}` : `\\begin{array}{l}${rowsOf(el)}\\end{array}`;
    }
    case 'func': {
      const fn = child(el, 'fName');
      return `${fn ? row(fn) : ''}${fn && /\\[a-z]+ $/.test(row(fn)) ? '' : ' '}${grp(arg(el, 'e'))}`;
    }
    case 'limLow': {
      const e = arg(el, 'e').trim();
      const lim = arg(el, 'lim');
      if (/^\\(lim|max|min|sup|inf|limsup|liminf)$/.test(e)) return `${e}_${grp(lim)}`;
      return `\\underset${grp(lim)}${grp(e)}`;
    }
    case 'limUpp': return `\\overset${grp(arg(el, 'lim'))}${grp(arg(el, 'e'))}`;
    case 'acc': {
      const chr = prop(el, 'accPr', 'chr') ?? '̂';
      const e = arg(el, 'e');
      const pair = ACC_TEX[chr];
      if (!pair) return `\\overset{${charsToTex(chr)}}${grp(e)}`;
      const wide = e.replace(/\\[a-zA-Z]+\s?|[{}\s]/g, 'x').length > 1;
      return `${pair[wide ? 1 : 0]}${grp(e)}`;
    }
    case 'bar': {
      const pos = prop(el, 'barPr', 'pos') ?? 'bot';
      return `${pos === 'top' ? '\\overline' : '\\underline'}${grp(arg(el, 'e'))}`;
    }
    case 'groupChr': {
      const chr = prop(el, 'groupChrPr', 'chr') ?? '⏟';
      const pos = prop(el, 'groupChrPr', 'pos') ?? 'bot';
      const e = arg(el, 'e');
      if (chr === '⏟') return `\\underbrace${grp(e)}`;
      if (chr === '⏞') return `\\overbrace${grp(e)}`;
      if (chr === '⏜' && pos === 'top') return `\\overparen${grp(e)}`;
      if (/[←-⇿⟰-⟿]/.test(chr)) {
        const cmd = /[←⟵⇐]/.test(chr) ? '\\xleftarrow' : '\\xrightarrow';
        return pos === 'top' ? `${cmd}[${e}]{}` : `${cmd}${grp(e)}`;
      }
      return pos === 'top' ? `\\overset{${charsToTex(chr)}}${grp(e)}` : `\\underset{${charsToTex(chr)}}${grp(e)}`;
    }
    case 'borderBox': return `\\boxed${grp(arg(el, 'e'))}`;
    case 'box': return arg(el, 'e');
    case 'phant': return `\\phantom${grp(arg(el, 'e'))}`;
    case 'oMathPara': return kids(el).filter((k) => ln(k) === 'oMath').map(row).join(' \\\\ ');
    case 'e': case 'oMath': case 'num': case 'den': case 'sub': case 'sup': case 'deg': case 'lim': case 'fName':
      return row(el);
    default:
      if (/Pr$/.test(n) || n === 'ctrlPr') return '';
      return row(el);
  }
}

function rowsOf(el, casesMode = false) {
  const rows = [];
  if (ln(el) === 'm') {
    for (const mr of kids(el).filter((k) => ln(k) === 'mr')) {
      rows.push(kids(mr).filter((k) => ln(k) === 'e').map(row).join(' & '));
    }
  } else {
    for (const e of kids(el).filter((k) => ln(k) === 'e')) {
      // điểm căn m:aln → &
      rows.push(casesMode ? row(e) : alnRow(e));
    }
  }
  return rows.join(' \\\\ ');
}

function alnRow(e) {
  return kids(e).map((k) => {
    const s = o2l(k);
    if (ln(k) === 'r') {
      const rpr = child(k, 'rPr');
      if (rpr && child(rpr, 'aln')) return '&' + s;
    }
    return s;
  }).join('');
}

/** Chuỗi OOXML/OMML bất kỳ → mảng { latex, display } các phương trình tìm được. */
export function ommlToLatexList(xmlString) {
  const doc = parseXml(xmlString);
  const found = [];
  const walk = (n) => {
    for (const k of kids(n)) {
      const name = ln(k);
      if (name === 'oMathPara') {
        found.push({ latex: tidy(o2l(k)), display: true });
      } else if (name === 'oMath') {
        found.push({ latex: tidy(o2l(k)), display: false });
      } else walk(k);
    }
  };
  walk(doc);
  return found;
}

function tidy(s) {
  return s.replace(/\s{2,}/g, ' ').replace(/ ([}_^])/g, '$1').trim();
}
