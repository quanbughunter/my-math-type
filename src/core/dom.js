// Tiện ích DOM nhỏ, chạy được cả trong trình duyệt lẫn Node (xmldom) để kiểm thử.

export const MML_NS = 'http://www.w3.org/1998/Math/MathML';

export const ln = (n) => n.localName || n.nodeName.replace(/^.*:/, '');

export function kids(n) {
  const a = [];
  for (let c = n.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) a.push(c);
  return a;
}

export const txt = (n) => (n ? n.textContent || '' : '');

/** Mọi phần tử con cháu (thứ tự trước), lọc theo tên nếu có. */
export function allEls(root, name) {
  const out = [];
  (function walk(n) {
    for (const k of kids(n)) {
      if (!name || ln(k) === name) out.push(k);
      walk(k);
    }
  })(root);
  return out;
}

export function renameEl(doc, el, name, ns = MML_NS) {
  const ne = doc.createElementNS(ns, name);
  const attrs = el.attributes;
  for (let i = 0; i < attrs.length; i++) {
    const a = attrs.item(i);
    if (a.name !== 'xmlns') ne.setAttribute(a.name, a.value);
  }
  while (el.firstChild) ne.appendChild(el.firstChild);
  el.parentNode.replaceChild(ne, el);
  return ne;
}

export function unwrapEl(el) {
  const p = el.parentNode;
  while (el.firstChild) p.insertBefore(el.firstChild, el);
  p.removeChild(el);
}

export function parseXml(str) {
  const P = globalThis.DOMParser;
  if (!P) throw new Error('DOMParser không khả dụng');
  const doc = new P().parseFromString(str, 'application/xml');
  const err = doc.getElementsByTagName('parsererror')[0];
  if (err) throw new Error('XML lỗi: ' + (err.textContent || '').slice(0, 200));
  return doc;
}

export function serializeXml(node) {
  return new globalThis.XMLSerializer().serializeToString(node);
}
