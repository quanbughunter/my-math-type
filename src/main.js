import './app.css';
import { createEditor, mathMarkup, followVirtualKeyboard } from './ui/editor.js';
import { ICONS } from './ui/icons.js';
import { toast, applyTheme, setupDialogs, debounce } from './ui/common.js';
import { loadSettings, saveSettings, loadLibrary, remember, toggleFav, removeItem, exportLibraryJson, importLibraryJson } from './core/storage.js';
import { latexToWordMathML } from './core/mathml.js';
import { latexToOmml } from './core/omml.js';
import { isBlankLatex } from './core/latex.js';
import { buildDocx } from './core/ooxml.js';
import { copyText, copyImage, saveFile, pickTextFile, isCapacitor } from './core/platform.js';
import { createFormatBar, formatFromSettings } from './ui/format-bar.js';

const $ = (s) => document.querySelector(s);
const settings = loadSettings();
applyTheme(settings.theme);

// ---------- biểu tượng & nhãn nút ----------
const label = (id, icon, text) => { $(id).innerHTML = `${ICONS[icon]}${text ? `<span>${text}</span>` : ''}`; };
label('#btnUndo', 'undo'); label('#btnRedo', 'redo'); label('#btnNew', 'plus');
label('#btnLib', 'library'); label('#btnSettings', 'settings'); label('#btnHelp', 'help');
label('#btnWord', 'word', 'Copy vào Word');
label('#btnLatex', 'tex', 'Copy LaTeX');
label('#btnImage', 'image', 'Copy ảnh');
label('#btnDocx', 'download', 'Tải .docx');
label('#btnFav', 'star', 'Lưu');
label('#btnCloseLib', 'close');
label('#btnLibDocx', 'download', 'Xuất ra Word');
document.querySelectorAll('[data-close]').forEach((b) => { b.innerHTML = ICONS.close; });

// ---------- ô soạn ----------
const saveDraft = debounce((latex) => { try { localStorage.setItem('mymath.draft', latex); } catch { /* bỏ qua */ } }, 400);
const validate = debounce(() => checkConvertible(), 350);

const ed = createEditor($('#editorHost'), {
  settings,
  onChange: (latex) => {
    saveDraft(latex);
    validate();
    refreshSource();
  },
  onSubmit: () => copyForWord(),
});
followVirtualKeyboard();

const params = new URLSearchParams(location.search);
let initial = params.get('latex');
if (initial == null) { try { initial = localStorage.getItem('mymath.draft') || ''; } catch { initial = ''; } }
if (initial) ed.setLatex(initial, { focus: false });
setTimeout(() => ed.focus(), 50);

// ---------- chuyển đổi ----------
function current() {
  return { latex: ed.getLatex(), portable: ed.getPortableLatex(), display: settings.display };
}
function convertOrThrow(kind = 'mathml') {
  const { portable, display } = current();
  if (isBlankLatex(portable)) throw new Error('Công thức đang trống.');
  try {
    return kind === 'omml' ? latexToOmml(portable, { display, format: formatFromSettings(settings) }) : latexToWordMathML(portable, { display });
  } catch (e) {
    throw new Error('Chưa chuyển được sang Word: ' + cleanErr(e));
  }
}
function cleanErr(e) {
  return String(e?.message || e).replace(/^ParseError:\s*/, '').slice(0, 160);
}
function checkConvertible() {
  const w = $('#warn');
  const { portable, display } = current();
  if (isBlankLatex(portable)) { w.hidden = true; return; }
  try {
    latexToWordMathML(portable, { display });
    w.hidden = true;
  } catch (e) {
    w.textContent = 'Lệnh này chưa chuyển được sang Word: ' + cleanErr(e);
    w.hidden = false;
  }
}
function rememberCurrent(fav = false) {
  if (!fav && !settings.autoSave) return;
  const { latex, display } = current();
  remember(latex, { display, fav });
  if (!$('#drawer').hidden) renderLibrary();
}

// ---------- các nút xuất ----------
async function copyForWord() {
  try {
    const mathml = convertOrThrow('mathml');
    await copyText(mathml);
    rememberCurrent();
    toast(isCapacitor() ? 'Đã copy. Word trên Android có thể chưa nhận — nếu vậy hãy dùng “Tải .docx”.' : 'Đã copy — sang Word nhấn Ctrl+V');
  } catch (e) {
    toast(e.message, true);
  }
}
async function copyLatexCode() {
  const { portable } = current();
  if (isBlankLatex(portable)) return toast('Công thức đang trống.', true);
  try {
    await copyText(portable);
    rememberCurrent();
    toast('Đã copy mã LaTeX');
  } catch (e) { toast(e.message, true); }
}
async function copyAsImage() {
  const { portable, display } = current();
  if (isBlankLatex(portable)) return toast('Công thức đang trống.', true);
  try {
    toast('Đang tạo ảnh…');
    const { latexToPng } = await import('./core/image.js');
    const png = await latexToPng(portable, { display, scale: +settings.pngScale, background: settings.pngWhite ? '#ffffff' : null });
    const ok = await copyImage(png).catch(() => false);
    if (ok) toast('Đã copy ảnh — dán vào nơi cần dùng');
    else {
      const r = await saveFile(png, 'cong-thuc.png');
      toast(r === 'shared' ? 'Đã mở bảng chia sẻ ảnh' : 'Đã tải ảnh cong-thuc.png');
    }
    rememberCurrent();
  } catch (e) { toast('Không tạo được ảnh: ' + cleanErr(e), true); }
}
async function downloadDocx() {
  try {
    const omml = convertOrThrow('omml');
    const blob = await buildDocx([{ omml }], { title: 'Công thức' });
    const r = await saveFile(blob, 'cong-thuc.docx');
    rememberCurrent();
    toast(r === 'shared' ? 'Chọn Word để mở tệp' : 'Đã tải cong-thuc.docx');
  } catch (e) { toast(e.message, true); }
}

$('#btnWord').addEventListener('click', copyForWord);
$('#btnLatex').addEventListener('click', copyLatexCode);
$('#btnImage').addEventListener('click', copyAsImage);
$('#btnDocx').addEventListener('click', downloadDocx);
$('#btnFav').addEventListener('click', () => {
  if (isBlankLatex(ed.getLatex())) return toast('Công thức đang trống.', true);
  rememberCurrent(true);
  toast('Đã lưu vào thư viện');
});
$('#btnUndo').addEventListener('click', () => ed.undo());
$('#btnRedo').addEventListener('click', () => ed.redo());
$('#btnNew').addEventListener('click', () => {
  const l = ed.getLatex();
  if (l && !isBlankLatex(l)) remember(l, { display: settings.display });
  ed.setLatex('');
});
document.addEventListener('keydown', (e) => {
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
    e.preventDefault();
    $('#btnFav').click();
  }
});

// ---------- kiểu hiển thị & cỡ chữ ----------
function paintMode() {
  for (const b of $('#modeSeg').children) b.classList.toggle('on', (b.dataset.display === '1') === settings.display);
}
$('#modeSeg').addEventListener('click', (e) => {
  const b = e.target.closest('button');
  if (!b) return;
  settings.display = b.dataset.display === '1';
  saveSettings(settings);
  paintMode();
  refreshSource();
});
paintMode();
createFormatBar($('#fmtHost'), settings, () => { saveSettings(settings); refreshSource(); });
const fs = $('#fontSize');
fs.value = settings.fontSize;
fs.addEventListener('input', () => {
  settings.fontSize = +fs.value;
  ed.setFontSize(settings.fontSize);
  saveSettings(settings);
});

// ---------- mã nguồn ----------
let srcKind = 'latex';
const srcText = $('#srcText');
function refreshSource() {
  if (!$('#source').open || document.activeElement === srcText) return;
  const { latex, portable, display } = current();
  srcText.readOnly = srcKind !== 'latex';
  if (srcKind === 'latex') { srcText.value = latex; return; }
  try {
    srcText.value = isBlankLatex(portable) ? '' : srcKind === 'omml' ? latexToOmml(portable, { display, format: formatFromSettings(settings) }) : latexToWordMathML(portable, { display });
  } catch (e) { srcText.value = '⚠ ' + cleanErr(e); }
}
$('#source').addEventListener('toggle', refreshSource);
$('.src-tabs').addEventListener('click', (e) => {
  const b = e.target.closest('button[data-k]');
  if (!b) return;
  srcKind = b.dataset.k;
  for (const x of $('.src-tabs').querySelectorAll('button[data-k]')) x.classList.toggle('on', x === b);
  refreshSource();
});
srcText.addEventListener('input', debounce(() => {
  if (srcKind !== 'latex') return;
  ed.setLatex(srcText.value, { focus: false });
}, 350));
srcText.addEventListener('blur', refreshSource);
$('#btnCopySrc').addEventListener('click', async () => {
  try { await copyText(srcText.value); toast('Đã copy'); } catch (e) { toast(e.message, true); }
});

// ---------- thư viện ----------
let libFilter = 'recent';
function renderLibrary() {
  const all = loadLibrary();
  const list = libFilter === 'fav' ? all.filter((x) => x.fav) : all;
  const ul = $('#libList');
  if (!list.length) {
    ul.innerHTML = `<li class="lib-empty">${libFilter === 'fav' ? 'Chưa lưu công thức nào. Bấm ★ Lưu để giữ công thức hay dùng.' : 'Các công thức đã copy sẽ hiện ở đây.'}</li>`;
    return;
  }
  ul.innerHTML = list.map((x) => `
    <li data-id="${x.id}">
      <button class="pick" title="Mở để sửa">${mathMarkup(x.latex)}</button>
      <button class="icon-btn fav ${x.fav ? 'fav-on' : ''}" title="${x.fav ? 'Bỏ lưu' : 'Lưu'}">${x.fav ? ICONS.starFill : ICONS.star}</button>
      <button class="icon-btn del" title="Xoá">${ICONS.trash}</button>
    </li>`).join('');
}
$('#libList').addEventListener('click', (e) => {
  const li = e.target.closest('li[data-id]');
  if (!li) return;
  const id = li.dataset.id;
  if (e.target.closest('.fav')) { toggleFav(id); renderLibrary(); return; }
  if (e.target.closest('.del')) { removeItem(id); renderLibrary(); return; }
  if (e.target.closest('.pick')) {
    const it = loadLibrary().find((x) => x.id === id);
    if (!it) return;
    settings.display = it.display !== false;
    paintMode();
    ed.setLatex(it.latex);
    if (window.innerWidth < 900) $('#drawer').hidden = true;
  }
});
$('#libSeg').addEventListener('click', (e) => {
  const b = e.target.closest('button');
  if (!b) return;
  libFilter = b.dataset.f;
  for (const x of $('#libSeg').children) x.classList.toggle('on', x === b);
  renderLibrary();
});
$('#btnLib').addEventListener('click', () => { const d = $('#drawer'); d.hidden = !d.hidden; if (!d.hidden) renderLibrary(); });
$('#btnCloseLib').addEventListener('click', () => { $('#drawer').hidden = true; });
$('#btnLibDocx').addEventListener('click', async () => {
  const all = loadLibrary();
  const list = libFilter === 'fav' ? all.filter((x) => x.fav) : all;
  const items = [];
  for (const x of list) {
    try { items.push({ omml: latexToOmml(x.latex, { display: x.display !== false, format: formatFromSettings(settings) }) }); } catch { /* bỏ mục lỗi */ }
  }
  if (!items.length) return toast('Không có công thức để xuất.', true);
  const blob = await buildDocx(items, { title: 'Thư viện công thức', heading: 'Thư viện công thức MyMath' });
  const r = await saveFile(blob, 'thu-vien-cong-thuc.docx');
  toast(r === 'shared' ? 'Chọn Word để mở tệp' : `Đã xuất ${items.length} công thức`);
});
$('#btnLibExport').addEventListener('click', async () => {
  const blob = new Blob([exportLibraryJson()], { type: 'application/json' });
  await saveFile(blob, 'mymath-thu-vien.json');
});
$('#btnLibImport').addEventListener('click', async () => {
  const text = await pickTextFile();
  if (!text) return;
  try { importLibraryJson(text); renderLibrary(); toast('Đã nhập thư viện'); } catch (e) { toast('Tệp không hợp lệ: ' + e.message, true); }
});

// ---------- cài đặt & hướng dẫn ----------
setupDialogs();
$('#btnSettings').addEventListener('click', () => {
  $('#setKeyboard').value = settings.keyboard;
  $('#setTheme').value = settings.theme;
  $('#setScale').value = String(settings.pngScale);
  $('#setWhite').checked = !!settings.pngWhite;
  $('#setAutoSave').checked = !!settings.autoSave;
  $('#dlgSettings').showModal();
});
$('#dlgSettings').addEventListener('change', () => {
  settings.keyboard = $('#setKeyboard').value;
  settings.theme = $('#setTheme').value;
  settings.pngScale = +$('#setScale').value;
  settings.pngWhite = $('#setWhite').checked;
  settings.autoSave = $('#setAutoSave').checked;
  saveSettings(settings);
  applyTheme(settings.theme);
  ed.setKeyboardPolicy(settings.keyboard);
});
$('#btnHelp').addEventListener('click', () => $('#dlgHelp').showModal());
$('#aboutLine').textContent = `MyMath ${__APP_VERSION__} · chạy offline · dữ liệu chỉ lưu trên máy này.`;

// ---------- PWA: cài như ứng dụng, chạy offline ----------
if (import.meta.env.PROD && 'serviceWorker' in navigator && location.protocol === 'https:' && !isCapacitor()) {
  navigator.serviceWorker.register('./sw.js').catch(() => {});
}
