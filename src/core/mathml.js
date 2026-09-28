// LaTeX → MathML (Temml) và chuẩn hoá MathML cho Microsoft Word.
//
// Word nhận MathML dán vào (Ctrl+V) và tự đổi thành phương trình gốc (OMML) bằng
// bộ chuyển MML2OMML của nó. Bộ đó "thích" kiểu MathML 2: <mfenced> cho ngoặc co giãn,
// không có class/style, khoảng trắng là ký tự… Hàm wordifyMathML() làm đúng việc đó.

import temml from 'temml';
import { prepareLatex, temmlMacros } from './latex.js';
import { MML_NS, ln, kids, txt, allEls, renameEl, unwrapEl, parseXml, serializeXml } from './dom.js';

export { MML_NS };

/** LaTeX → chuỗi MathML thô (Temml). Ném lỗi nếu LaTeX sai. */
export function latexToRawMathML(latex, { display = true } = {}) {
  const tex = prepareLatex(latex);
  return temml.renderToString(tex, {
    displayMode: display,
    xml: true,
    annotate: false,
    throwOnError: true,
    wrap: 'none',
    macros: temmlMacros(),
  });
}

// Toán tử lớn (tổng, tích phân, hợp, giao…)
export const NARY_RE = /^[∑∏∐∫-∳⋀-⋃⨀-⨆⨉⨌⨐-⨗]$/;
// Quan hệ / dấu ngăn: thân của toán tử lớn dừng tại đây
const STOPPERS = new Set(['=', '≠', '<', '>', '≤', '≥', '≦', '≧', '⩽', '⩾', '≈', '≡', '≅', '∼', '≃', '∝', '≪', '≫',
  '⇒', '⇐', '⇔', '⟹', '⟸', '⟺', '→', '←', '↔', '⟶', '⟵', '↦', ',', ';', '≔', '⇌', '⇄', '⇌']);
const ROWLIKE = new Set(['math', 'mrow', 'mtd', 'msqrt', 'mstyle', 'mpadded', 'mphantom', 'menclose', 'merror', 'mfenced']);
const SCRIPTS = new Set(['msub', 'msup', 'msubsup', 'munder', 'mover', 'munderover']);
const FIXED_ARITY = new Set(['mfrac', 'mroot', 'msub', 'msup', 'msubsup', 'munder', 'mover', 'munderover', 'mmultiscripts']);
export const APPLY_FN = '⁡';

export function isNaryEl(el) {
  if (!el) return false;
  const n = ln(el);
  if (n === 'mo') return NARY_RE.test(txt(el).trim());
  if (SCRIPTS.has(n)) {
    const b = kids(el)[0];
    return !!b && ln(b) === 'mo' && NARY_RE.test(txt(b).trim());
  }
  return false;
}

export function isStopper(el) {
  return !!el && ln(el) === 'mo' && STOPPERS.has(txt(el).trim());
}

function parseEm(v) {
  if (!v) return null;
  const m = /^(-?[\d.]+)\s*(em|pt|px|ex|mu)?$/.exec(String(v).trim());
  if (!m) return null;
  const x = parseFloat(m[1]);
  switch (m[2]) {
    case 'pt': return x / 10;
    case 'px': return x / 16;
    case 'ex': return x * 0.43;
    case 'mu': return x / 18;
    default: return x;
  }
}

function spaceFor(em) {
  if (em == null || em <= 0.05) return null;
  if (em < 0.2) return ' ';      // \,  thin space
  if (em < 0.25) return ' ';     // \:  medium space
  if (em < 0.4) return ' ';      // \;  thick space
  if (em < 0.8) return ' ';      // en space
  if (em < 1.5) return ' ';      // \quad
  return '  ';              // \qquad
}

/**
 * Chuẩn hoá cây MathML (tại chỗ) cho Word. Trả về phần tử <math>.
 */
export function normalizeMathDoc(doc) {
  const root = doc.documentElement;

  // 0) <semantics> → giữ phần trình bày đầu tiên
  for (const s of allEls(root, 'semantics')) {
    const [first, ...rest] = kids(s);
    rest.forEach((r) => s.removeChild(r));
    if (first) renameEl(doc, s, 'mrow');
  }

  // 1) Căn cột bảng: Temml ghi bằng class tml-left/tml-right trên <mtd>
  for (const t of allEls(root, 'mtable')) {
    if (t.getAttribute('columnalign')) continue;
    const firstRow = kids(t).find((r) => ln(r) === 'mtr' || ln(r) === 'mlabeledtr');
    if (!firstRow) continue;
    const aligns = kids(firstRow).map((td) => {
      const c = td.getAttribute('class') || '';
      const ca = td.getAttribute('columnalign');
      if (ca) return ca;
      if (/tml-left/.test(c)) return 'left';
      if (/tml-right/.test(c)) return 'right';
      return 'center';
    });
    if (aligns.some((a) => a !== 'center')) t.setAttribute('columnalign', aligns.join(' '));
  }

  // 1b) \boxed: Temml vẽ khung bằng CSS border → <menclose notation="box">
  for (const el of allEls(root, 'mrow')) {
    if (/border\s*:/.test(el.getAttribute('style') || '')) {
      const box = renameEl(doc, el, 'menclose');
      box.setAttribute('notation', 'box');
    }
  }

  // 1c) Dấu mũ nhỏ (\hat \bar \vec \dot…): Temml ghi stretchy="false" thay vì accent="true"
  for (const el of allEls(root, 'mover')) {
    const s = kids(el)[1];
    if (!s || ln(s) !== 'mo') continue;
    const c = s.getAttribute('class') || '';
    if (s.getAttribute('stretchy') === 'false' || /(^|\s)(tml-vec|wbk-sml-acc|wbk-sml-vec)/.test(c)) {
      el.setAttribute('accent', 'true');
    }
  }

  // 2) Bỏ thuộc tính trình duyệt
  for (const el of [root, ...allEls(root)]) {
    for (const a of ['class', 'style', 'lspace', 'rspace', 'height', 'depth']) {
      if (ln(el) === 'mspace' && (a === 'height' || a === 'depth')) continue;
      el.removeAttribute(a);
    }
  }

  // 3) <mspace> → ký tự khoảng trắng Unicode (Word hiểu) hoặc bỏ
  const UO = new Set(['mover', 'munder', 'munderover']);
  for (const sp of allEls(root, 'mspace')) {
    if (sp.getAttribute('linebreak')) continue;
    const parent = sp.parentNode;
    const pp = parent.parentNode;
    const siblings = kids(parent);
    const prev = siblings[siblings.indexOf(sp) - 1];
    // khoảng đệm quanh chữ trên mũi tên (\xrightarrow), sau tên hàm (sin⁡ x) → bỏ, Word tự căn
    const padInScript = ln(parent) === 'mrow' && pp && UO.has(ln(pp)) &&
      (sp === siblings[0] || sp === siblings[siblings.length - 1]);
    const afterApply = prev && ln(prev) === 'mo' && txt(prev) === APPLY_FN;
    const ch = UO.has(ln(parent)) || padInScript || afterApply ? null : spaceFor(parseEm(sp.getAttribute('width')));
    if (ch) {
      const t = doc.createElementNS(MML_NS, 'mtext');
      t.appendChild(doc.createTextNode(ch));
      parent.replaceChild(t, sp);
    } else if (FIXED_ARITY.has(ln(parent))) {
      parent.replaceChild(doc.createElementNS(MML_NS, 'mrow'), sp);
    } else {
      parent.removeChild(sp);
    }
  }

  // 4) mstyle / mpadded → mrow (Word bỏ qua các thuộc tính đó)
  for (const el of allEls(root)) {
    if (ln(el) === 'mstyle' || ln(el) === 'mpadded') renameEl(doc, el, 'mrow');
  }

  // 5) mover/munder mà phần trên/dưới rỗng → chỉ còn phần gốc
  for (const el of allEls(root).reverse()) {
    const n = ln(el);
    if (n !== 'mover' && n !== 'munder') continue;
    const k = kids(el);
    const script = k[1];
    if (k.length === 1 || (script && isEmptyRow(script))) {
      if (script) el.removeChild(script);
      renameEl(doc, el, 'mrow');
    }
  }

  // 6) Temml bọc "sin⁡" thành <mrow><mi>sin</mi><mo>⁡</mo></mrow> tách khỏi đối số → trải phẳng ra
  for (const el of allEls(root, 'mrow').reverse()) {
    const parent = el.parentNode;
    if (!parent || !ROWLIKE.has(ln(parent))) continue;
    const k = kids(el);
    const last = k[k.length - 1];
    if (last && ln(last) === 'mo' && txt(last) === APPLY_FN) unwrapEl(el);
  }

  // 7) mrow chỉ có một con → thay bằng chính con đó; mrow rỗng trong hàng → bỏ
  collapseRows(root);

  // 8) Ngoặc \left … \right → <mfenced> (Word đổi thành ngoặc co giãn m:d)
  for (const el of allEls(root, 'mrow').reverse()) {
    const k = kids(el);
    if (k.length < 2) continue;
    const a = k[0];
    const b = k[k.length - 1];
    if (!isFence(a, 'prefix') || !isFence(b, 'postfix')) continue;
    const fenced = doc.createElementNS(MML_NS, 'mfenced');
    fenced.setAttribute('open', txt(a).trim());
    fenced.setAttribute('close', txt(b).trim());
    fenced.setAttribute('separators', '');
    const inner = doc.createElementNS(MML_NS, 'mrow');
    k.slice(1, -1).forEach((c) => inner.appendChild(c));
    fenced.appendChild(inner);
    el.removeChild(a);
    el.removeChild(b);
    el.parentNode.replaceChild(fenced, el);
  }

  // 9) Mũ/ngã rộng (\widehat, \widetilde) → accent để Word dùng m:acc
  for (const el of allEls(root, 'mover')) {
    const s = kids(el)[1];
    if (s && ln(s) === 'mo' && /^[ˆ^˜~]$/.test(txt(s).trim())) el.setAttribute('accent', 'true');
  }

  // 10) Thân của ∑ ∫ …: gom các phần tử phía sau (tới dấu quan hệ) vào một <mrow>
  for (const row of [root, ...allEls(root)]) {
    if (!ROWLIKE.has(ln(row))) continue;
    for (const k of kids(row)) {
      if (!isNaryEl(k)) continue;
      const body = [];
      for (let s = nextEl(k); s && !isStopper(s); s = nextEl(s)) body.push(s);
      if (!body.length) continue;
      if (body.length === 1 && ln(body[0]) === 'mrow') continue;
      const wrap = doc.createElementNS(MML_NS, 'mrow');
      row.insertBefore(wrap, body[0]);
      body.forEach((b) => wrap.appendChild(b));
    }
  }
  return root;
}

function nextEl(n) {
  let s = n.nextSibling;
  while (s && s.nodeType !== 1) s = s.nextSibling;
  return s;
}

function isEmptyRow(el) {
  return ln(el) === 'mrow' && kids(el).length === 0 && !txt(el).trim();
}

function isFence(el, form) {
  if (!el || ln(el) !== 'mo') return false;
  if (el.getAttribute('fence') === 'true') {
    const f = el.getAttribute('form');
    return !f || f === form;
  }
  return false;
}

function collapseRows(root) {
  for (const el of allEls(root, 'mrow').reverse()) {
    const parent = el.parentNode;
    if (!parent) continue;
    const k = kids(el);
    if (k.length === 1 && !hasOwnText(el)) {
      parent.replaceChild(k[0], el);
    } else if (k.length === 0 && !txt(el).trim() && ROWLIKE.has(ln(parent))) {
      parent.removeChild(el);
    }
  }
}

function hasOwnText(el) {
  for (let c = el.firstChild; c; c = c.nextSibling) if (c.nodeType === 3 && c.data.trim()) return true;
  return false;
}

/** LaTeX → { doc, root } MathML đã chuẩn hoá cho Word. */
export function latexToWordMathDoc(latex, { display = true } = {}) {
  const raw = latexToRawMathML(latex, { display });
  const doc = parseXml(raw);
  const root = normalizeMathDoc(doc);
  if (display) root.setAttribute('display', 'block');
  else root.removeAttribute('display');
  return { doc, root };
}

/** LaTeX → chuỗi MathML để dán vào Word. */
export function latexToWordMathML(latex, opts = {}) {
  const { root } = latexToWordMathDoc(latex, opts);
  let s = serializeXml(root);
  // đảm bảo có khai báo namespace (Word nhận diện bằng thẻ <math xmlns=…>)
  if (!/^<math[^>]*xmlns=/.test(s)) s = s.replace(/^<math/, `<math xmlns="${MML_NS}"`);
  return s;
}
