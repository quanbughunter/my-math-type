// Khung soạn công thức dùng chung cho ứng dụng và Word add-in:
// thanh mẫu (kiểu MathType) + ô soạn MathLive.

import { MathfieldElement, convertLatexToMarkup } from 'mathlive';
import 'mathlive/fonts.css';
import 'mathlive/static.css';
import { PALETTES, previewLatex } from './palettes.js';

MathfieldElement.fontsDirectory = null; // phông đã nạp qua CSS ở trên (chạy offline)
MathfieldElement.soundsDirectory = null;

const escHtml = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

/** Vẽ LaTeX thành HTML tĩnh (cho nút mẫu, thư viện). */
export function mathMarkup(latex, display = false) {
  try {
    return convertLatexToMarkup(latex, { defaultMode: display ? 'math' : 'inline-math' });
  } catch {
    return `<code>${escHtml(latex)}</code>`;
  }
}

/**
 * @param {HTMLElement} host
 * @param {{ settings: object, onChange?: (latex:string)=>void, onSubmit?: ()=>void, compact?: boolean }} opts
 */
export function createEditor(host, opts) {
  const { settings } = opts;
  host.classList.add('mm-editor');
  host.innerHTML = `
    <div class="mm-tabs" role="tablist" aria-label="Nhóm mẫu công thức"></div>
    <div class="mm-palette" role="toolbar" aria-label="Mẫu công thức"></div>
    <div class="mm-field-wrap">
      <div class="mm-field-slot"></div>
    </div>`;

  const tabsEl = host.querySelector('.mm-tabs');
  const palEl = host.querySelector('.mm-palette');
  const slot = host.querySelector('.mm-field-slot');

  // ----- ô soạn thảo -----
  const mf = new MathfieldElement();
  mf.className = 'mm-field';
  mf.smartFence = true;
  mf.smartSuperscript = true;
  mf.mathModeSpace = '\\,';
  mf.setAttribute('aria-label', 'Ô soạn công thức');
  slot.appendChild(mf);
  applyKeyboardPolicy(mf, settings.keyboard);
  setFontSize(settings.fontSize);

  mf.addEventListener('input', () => opts.onChange?.(getLatex()));
  mf.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
      e.preventDefault();
      opts.onSubmit?.();
    }
  }, { capture: true });

  // ----- bảng mẫu -----
  const cache = new Map();
  let active = PALETTES[0].id;
  tabsEl.innerHTML = PALETTES.map((p) =>
    `<button type="button" role="tab" class="mm-tab" data-id="${p.id}">${escHtml(p.label)}</button>`).join('');
  tabsEl.addEventListener('click', (e) => {
    const b = e.target.closest('.mm-tab');
    if (b) showPalette(b.dataset.id);
  });
  // giữ vùng chọn trong ô soạn khi bấm nút mẫu
  palEl.addEventListener('mousedown', (e) => {
    if (e.target.closest('.mm-sym')) e.preventDefault();
  });
  palEl.addEventListener('click', (e) => {
    const b = e.target.closest('.mm-sym');
    if (!b) return;
    const pal = PALETTES.find((p) => p.id === active);
    const item = pal.items[+b.dataset.i];
    insertTemplate(item.t);
  });

  function showPalette(id) {
    active = id;
    for (const b of tabsEl.children) {
      const on = b.dataset.id === id;
      b.classList.toggle('on', on);
      b.setAttribute('aria-selected', on ? 'true' : 'false');
    }
    if (!cache.has(id)) {
      const pal = PALETTES.find((p) => p.id === id);
      cache.set(id, pal.items.map((it, i) =>
        `<button type="button" class="mm-sym" data-i="${i}" title="${escHtml(it.h || it.t)}">${mathMarkup(previewLatex(it))}</button>`).join(''));
    }
    palEl.innerHTML = cache.get(id);
    palEl.scrollTop = 0;
    // mẫu rộng (ma trận 3×3…) chiếm 2 ô
    requestAnimationFrame(() => {
      for (const b of palEl.children) {
        const inner = b.firstElementChild;
        const w = inner ? inner.getBoundingClientRect().width : 0;
        if (w > b.clientWidth - 10 || b.scrollWidth > b.clientWidth + 2) b.classList.add('wide');
      }
    });
  }
  showPalette(active);

  function insertTemplate(t) {
    const hasSlot = /#0|#\?|#@/.test(t);
    mf.focus();
    mf.insert(t, { format: 'latex', selectionMode: hasSlot ? 'placeholder' : 'after', focus: true, scrollIntoView: true });
    opts.onChange?.(getLatex());
  }

  function getLatex() {
    return mf.getValue('latex');
  }
  /** LaTeX chuẩn (đã mở rộng macro của MathLive) — dùng để chuyển sang Word. */
  function getPortableLatex() {
    return mf.getValue('latex-expanded');
  }
  function setLatex(latex, { focus = true } = {}) {
    mf.setValue(latex || '', { silenceNotifications: false });
    if (focus) mf.focus();
    opts.onChange?.(getLatex());
  }
  function setFontSize(px) {
    mf.style.fontSize = `${px}px`;
  }

  return {
    mf,
    getLatex,
    getPortableLatex,
    setLatex,
    setFontSize,
    setKeyboardPolicy: (p) => applyKeyboardPolicy(mf, p),
    undo: () => { mf.executeCommand('undo'); opts.onChange?.(getLatex()); },
    redo: () => { mf.executeCommand('redo'); opts.onChange?.(getLatex()); },
    focus: () => mf.focus(),
    insertTemplate,
  };
}

function applyKeyboardPolicy(mf, policy) {
  if (policy === 'off') {
    mf.mathVirtualKeyboardPolicy = 'manual';
    mf.classList.add('no-vk');
  } else {
    mf.mathVirtualKeyboardPolicy = policy === 'manual' ? 'manual' : 'auto';
    mf.classList.remove('no-vk');
  }
}

/** Đẩy nội dung lên khi bàn phím ảo của MathLive hiện ra (điện thoại). */
export function followVirtualKeyboard() {
  const vk = window.mathVirtualKeyboard;
  if (!vk) return;
  vk.addEventListener('geometrychange', () => {
    const h = vk.visible ? vk.boundingRect.height : 0;
    document.documentElement.style.setProperty('--vk-h', `${h}px`);
  });
}
