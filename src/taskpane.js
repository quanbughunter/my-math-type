// Word add-in: soạn công thức trong khung bên phải Word và chèn thẳng thành phương trình gốc (OMML).
/* global Office, Word */
import './taskpane.css';
import { createEditor, mathMarkup } from './ui/editor.js';
import { ICONS } from './ui/icons.js';
import { toast, applyTheme, debounce } from './ui/common.js';
import { loadSettings, saveSettings, loadLibrary, remember } from './core/storage.js';
import { latexToWordMathML } from './core/mathml.js';
import { latexToOmml } from './core/omml.js';
import { ommlToLatexList } from './core/omml2latex.js';
import { flatOpcForInsert } from './core/ooxml.js';
import { isBlankLatex } from './core/latex.js';
import { copyText } from './core/platform.js';
import { parseXml, ln } from './core/dom.js';
import { createFormatBar, formatFromSettings } from './ui/format-bar.js';

const $ = (s) => document.querySelector(s);
const settings = loadSettings();
applyTheme(settings.theme);

let inWord = false;
let editing = false;          // true: nút chính sẽ THAY công thức đang chọn
let ignoreSelUntil = 0;       // bỏ qua sự kiện đổi vùng chọn do chính add-in gây ra

$('#btnUndo').innerHTML = ICONS.undo;
$('#btnNew').innerHTML = ICONS.plus;
$('#btnPick').innerHTML = `${ICONS.pick}<span>Lấy công thức đang chọn</span>`;
$('#btnLatex').innerHTML = `${ICONS.tex}<span>Copy LaTeX</span>`;

const validate = debounce(checkConvertible, 350);
const ed = createEditor($('#editorHost'), {
  settings: { ...settings, fontSize: Math.min(settings.fontSize, 26), keyboard: settings.keyboard === 'auto' ? 'manual' : settings.keyboard },
  onChange: () => validate(),
  onSubmit: () => insert(),
});
ed.setFontSize(Math.min(settings.fontSize, 26));

function paintPrimary() {
  $('#btnInsert').innerHTML = editing
    ? `${ICONS.insert}<span>Cập nhật công thức trong Word</span>`
    : `${ICONS.insert}<span>Chèn vào Word</span>`;
  $('#editing').hidden = !editing;
}
paintPrimary();

function paintMode() {
  for (const b of $('#modeSeg').children) b.classList.toggle('on', (b.dataset.display === '1') === settings.display);
}
$('#modeSeg').addEventListener('click', (e) => {
  const b = e.target.closest('button');
  if (!b) return;
  settings.display = b.dataset.display === '1';
  saveSettings(settings);
  paintMode();
});
paintMode();
createFormatBar($('#fmtHost'), settings, () => saveSettings(settings));

function checkConvertible() {
  const w = $('#warn');
  const latex = ed.getPortableLatex();
  if (isBlankLatex(latex)) { w.hidden = true; return; }
  try {
    latexToWordMathML(latex, { display: settings.display });
    w.hidden = true;
  } catch (e) {
    w.textContent = 'Lệnh này chưa chuyển được sang Word: ' + String(e.message || e).slice(0, 140);
    w.hidden = false;
  }
}

// ---------- kết nối Word ----------
function setStatus(text, ok) {
  const s = $('#status');
  s.textContent = text;
  s.classList.toggle('ok', !!ok);
}
if (typeof Office === 'undefined') {
  setStatus('Không ở trong Word', false);
} else {
  Office.onReady((info) => {
    inWord = info.host === Office.HostType.Word;
    if (!inWord) return setStatus('Không ở trong Word', false);
    if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
      return setStatus('Word quá cũ', false);
    }
    setStatus('Đã kết nối Word', true);
    try {
      Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => {
        if (Date.now() < ignoreSelUntil || !editing) return;
        editing = false;
        paintPrimary();
      });
    } catch { /* không bắt buộc */ }
  });
}

function requireWord() {
  if (inWord) return true;
  toast('Hãy mở khung này bên trong Word: tab MyMath → Công thức.', true);
  return false;
}

/** Trong đoạn OOXML có chữ thường nằm ngoài công thức không? */
function hasTextOutsideMath(ooxml) {
  const doc = parseXml(ooxml);
  const ts = doc.getElementsByTagName('w:t');
  for (let i = 0; i < ts.length; i++) {
    let p = ts[i].parentNode;
    let inMath = false;
    while (p && p.nodeType === 1) {
      const n = ln(p);
      if (n === 'oMath' || n === 'oMathPara') { inMath = true; break; }
      p = p.parentNode;
    }
    if (!inMath && (ts[i].textContent || '').trim()) return true;
  }
  return false;
}

// ---------- chèn / cập nhật ----------
async function insert({ asNew = false } = {}) {
  if (!requireWord()) return;
  const latex = ed.getPortableLatex();
  if (isBlankLatex(latex)) return toast('Công thức đang trống.', true);
  let omml;
  try {
    omml = latexToOmml(latex, { display: settings.display, format: formatFromSettings(settings) });
  } catch (e) {
    return toast('Chưa chuyển được: ' + String(e.message || e).slice(0, 140), true);
  }
  const ooxml = flatOpcForInsert(omml);
  const replacing = editing && !asNew;
  try {
    await Word.run(async (ctx) => {
      const sel = ctx.document.getSelection();
      sel.load('text');
      await ctx.sync();
      // đang bôi đen chữ thường mà không ở chế độ sửa → chèn phía sau để không mất chữ
      const where = replacing || !sel.text ? 'Replace' : 'After';
      const r = sel.insertOoxml(ooxml, where);
      ignoreSelUntil = Date.now() + 1500;
      r.select('End');
      await ctx.sync();
    });
    remember(ed.getLatex(), { display: settings.display });
    renderRecent();
    toast(replacing ? 'Đã cập nhật công thức' : 'Đã chèn công thức');
    editing = false;
    paintPrimary();
  } catch (e) {
    toast('Word báo lỗi: ' + (e.message || e), true);
  }
}
$('#btnInsert').addEventListener('click', () => insert());
$('#asNew').addEventListener('click', (e) => { e.preventDefault(); insert({ asNew: true }); });

// ---------- lấy công thức có sẵn trong Word để sửa ----------
async function pickFromWord() {
  if (!requireWord()) return;
  try {
    let found = [];
    let canReplace = false;
    await Word.run(async (ctx) => {
      const sel = ctx.document.getSelection();
      const ox = sel.getOoxml();
      await ctx.sync();
      found = ommlToLatexList(ox.value);
      if (found.length) {
        canReplace = found.length === 1 && !hasTextOutsideMath(ox.value);
        return;
      }
      // con trỏ chỉ đặt trong công thức: thử cả đoạn chứa nó
      const para = sel.paragraphs.getFirst();
      const pox = para.getOoxml();
      await ctx.sync();
      const list = ommlToLatexList(pox.value);
      if (list.length === 1) {
        found = list;
        if (!hasTextOutsideMath(pox.value)) {
          ignoreSelUntil = Date.now() + 1500;
          para.getRange('Content').select();
          await ctx.sync();
          canReplace = true;
        }
      } else if (list.length > 1) {
        throw new Error('Đoạn này có nhiều công thức — hãy bôi đen đúng công thức cần sửa.');
      }
    });
    if (!found.length) {
      return toast('Chưa thấy công thức. Bôi đen công thức trong Word rồi bấm lại.', true);
    }
    const eq = found[0];
    settings.display = eq.display;
    paintMode();
    ed.setLatex(eq.latex);
    editing = canReplace;
    paintPrimary();
    toast(canReplace ? 'Đã mở công thức — sửa xong bấm “Cập nhật”' : 'Đã mở công thức. Vùng chọn có cả chữ nên sẽ chèn thành công thức mới.');
  } catch (e) {
    toast(e.message || String(e), true);
  }
}
$('#btnPick').addEventListener('click', pickFromWord);

$('#btnLatex').addEventListener('click', async () => {
  const latex = ed.getPortableLatex();
  if (isBlankLatex(latex)) return toast('Công thức đang trống.', true);
  try { await copyText(latex); toast('Đã copy mã LaTeX'); } catch (e) { toast(e.message, true); }
});
$('#btnUndo').addEventListener('click', () => ed.undo());
$('#btnNew').addEventListener('click', () => { editing = false; paintPrimary(); ed.setLatex(''); });

// ---------- gần đây ----------
function renderRecent() {
  const list = loadLibrary().slice(0, 12);
  $('#recent').innerHTML = list.length
    ? list.map((x) => `<li><button data-id="${x.id}" title="Mở để sửa">${mathMarkup(x.latex)}</button></li>`).join('')
    : '<li class="empty">Công thức bạn chèn sẽ hiện ở đây.</li>';
}
$('#recent').addEventListener('click', (e) => {
  const b = e.target.closest('button[data-id]');
  if (!b) return;
  const it = loadLibrary().find((x) => x.id === b.dataset.id);
  if (!it) return;
  editing = false;
  paintPrimary();
  settings.display = it.display !== false;
  paintMode();
  ed.setLatex(it.latex);
});
renderRecent();
setTimeout(() => ed.focus(), 100);
