// Thanh định dạng: font và cỡ chữ (pt) của công thức khi đưa vào Word.

const FONTS = [
  ['Cambria Math', 'Cambria Math'],
  ['Times New Roman', 'Times New Roman'],
  ['STIX Two Math', 'STIX Two Math'],
  ['Latin Modern Math', 'Latin Modern Math'],
  ['TeX Gyre Termes Math', 'TeX Gyre Termes Math'],
  ['Arial', 'Arial'],
];
const CUSTOM = '__custom';

const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

/** Định dạng hiện tại để truyền cho bộ chuyển OMML. */
export function formatFromSettings(s) {
  return { font: s.fmtFont || '', size: +s.fmtSize || 0, subSize: +s.fmtSub || 0 };
}

/**
 * @param {HTMLElement} host
 * @param {object} settings  đối tượng cài đặt (sẽ được cập nhật tại chỗ)
 * @param {(s:object)=>void} onChange
 */
export function createFormatBar(host, settings, onChange) {
  const uid = Math.random().toString(36).slice(2, 7);
  host.classList.add('fmt-bar');
  host.innerHTML = `
    <label class="fmt-item fmt-font" title="Font của công thức khi đưa vào Word. Cambria Math: mặc định của Word. Times New Roman: chữ và số theo Times New Roman, biến in nghiêng (giống MathType).">
      <span>Font</span>
      <select class="fmt-select"></select>
    </label>
    <input class="fmt-custom" type="text" placeholder="Tên font…" hidden />
    <label class="fmt-item" title="Cỡ chữ chính của công thức (pt)">
      <span>Cỡ</span>
      <input class="fmt-size" type="number" min="6" max="72" step="0.5" list="pt-${uid}" inputmode="decimal" />
      <em>pt</em>
    </label>
    <label class="fmt-item" title="Cỡ chỉ số trên/dưới (pt). Để trống: Word tự thu nhỏ còn khoảng 3/4 cỡ chính.">
      <span>Chỉ số</span>
      <input class="fmt-sub" type="number" min="4" max="48" step="0.5" list="sub-${uid}" inputmode="decimal" />
      <em>pt</em>
    </label>
    <datalist id="pt-${uid}"><option value="10"><option value="11"><option value="12"><option value="13"><option value="14"></datalist>
    <datalist id="sub-${uid}"><option value="7"><option value="8"><option value="9"><option value="10"></datalist>`;

  const sel = host.querySelector('.fmt-select');
  const custom = host.querySelector('.fmt-custom');
  const size = host.querySelector('.fmt-size');
  const sub = host.querySelector('.fmt-sub');

  function fillFonts() {
    const cur = settings.fmtFont || 'Cambria Math';
    const list = FONTS.some(([v]) => v === cur) ? FONTS : [...FONTS, [cur, cur]];
    sel.innerHTML = list.map(([v, l]) => `<option value="${esc(v)}">${esc(l)}</option>`).join('') +
      `<option value="${CUSTOM}">Font khác…</option>`;
    sel.value = cur;
  }
  function paint() {
    fillFonts();
    size.value = settings.fmtSize || '';
    sub.value = settings.fmtSub || '';
    const auto = settings.fmtSize ? (settings.fmtSize * 0.73).toFixed(1).replace(/\.0$/, '') : '';
    sub.placeholder = auto ? `~${auto}` : 'tự động';
  }
  function commit() {
    onChange?.(settings);
    paint();
  }

  sel.addEventListener('change', () => {
    if (sel.value === CUSTOM) {
      custom.hidden = false;
      custom.value = '';
      custom.focus();
      return;
    }
    custom.hidden = true;
    settings.fmtFont = sel.value;
    commit();
  });
  custom.addEventListener('change', () => {
    const v = custom.value.trim();
    custom.hidden = true;
    if (v) settings.fmtFont = v;
    commit();
  });
  custom.addEventListener('keydown', (e) => { if (e.key === 'Enter') custom.blur(); });
  size.addEventListener('change', () => {
    const v = parseFloat(size.value);
    settings.fmtSize = v > 0 ? Math.min(72, v) : 0;
    commit();
  });
  sub.addEventListener('change', () => {
    const v = parseFloat(sub.value);
    settings.fmtSub = v > 0 ? Math.min(48, v) : 0;
    commit();
  });

  paint();
  return { refresh: paint };
}
