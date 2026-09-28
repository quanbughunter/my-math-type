// Lưu lịch sử, thư viện công thức và cài đặt trên máy (localStorage, bọc try/catch).

const KEY_LIB = 'mymath.library.v1';
const KEY_SET = 'mymath.settings.v1';
const MAX_RECENT = 150;

function read(key, fallback) {
  try {
    const v = localStorage.getItem(key);
    return v ? JSON.parse(v) : fallback;
  } catch {
    return fallback;
  }
}
function write(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch {
    /* bộ nhớ trình duyệt bị chặn → bỏ qua */
  }
}

export const DEFAULT_SETTINGS = {
  display: true,          // công thức riêng dòng (display) hay cùng dòng (inline)
  fontSize: 30,           // cỡ chữ ô soạn thảo (px)
  keyboard: 'auto',       // bàn phím ảo: auto | manual | off
  pngScale: 4,            // độ nét ảnh PNG
  pngWhite: false,        // nền trắng cho ảnh
  theme: 'auto',          // auto | light | dark
  autoSave: true,         // tự lưu vào "Gần đây" mỗi lần copy/chèn
  fmtFont: 'Cambria Math', // font khi đưa vào Word
  fmtSize: 12,            // cỡ chữ chính (pt)
  fmtSub: 0,              // cỡ chỉ số (pt), 0 = Word tự tính
};

export function loadSettings() {
  return { ...DEFAULT_SETTINGS, ...read(KEY_SET, {}) };
}
export function saveSettings(s) {
  write(KEY_SET, s);
}

/** @returns {{id:string, latex:string, display:boolean, fav:boolean, time:number}[]} */
export function loadLibrary() {
  const v = read(KEY_LIB, []);
  return Array.isArray(v) ? v : [];
}
function saveLibrary(list) {
  write(KEY_LIB, list);
}

/** Thêm/đưa công thức lên đầu. fav=true để đánh dấu "Đã lưu". */
export function remember(latex, { display = true, fav = false } = {}) {
  latex = (latex || '').trim();
  if (!latex) return loadLibrary();
  const list = loadLibrary();
  const i = list.findIndex((x) => x.latex === latex);
  let item;
  if (i >= 0) {
    item = list.splice(i, 1)[0];
    item.time = Date.now();
    item.display = display;
    if (fav) item.fav = true;
  } else {
    item = { id: Math.random().toString(36).slice(2, 10), latex, display, fav, time: Date.now() };
  }
  list.unshift(item);
  // giữ mọi mục "Đã lưu", cắt bớt mục gần đây
  let recent = 0;
  const kept = list.filter((x) => x.fav || ++recent <= MAX_RECENT);
  saveLibrary(kept);
  return kept;
}

export function toggleFav(id) {
  const list = loadLibrary();
  const it = list.find((x) => x.id === id);
  if (it) it.fav = !it.fav;
  saveLibrary(list);
  return list;
}

export function removeItem(id) {
  const list = loadLibrary().filter((x) => x.id !== id);
  saveLibrary(list);
  return list;
}

export function clearRecent() {
  const list = loadLibrary().filter((x) => x.fav);
  saveLibrary(list);
  return list;
}

/** Xuất / nhập thư viện (tệp JSON) để chuyển giữa máy tính và điện thoại. */
export function exportLibraryJson() {
  return JSON.stringify({ app: 'MyMath', version: 1, items: loadLibrary() }, null, 1);
}
export function importLibraryJson(text) {
  const data = JSON.parse(text);
  const items = Array.isArray(data) ? data : data.items;
  if (!Array.isArray(items)) throw new Error('Tệp không đúng định dạng');
  const list = loadLibrary();
  for (const it of items.reverse()) {
    if (!it || !it.latex) continue;
    if (!list.some((x) => x.latex === it.latex)) {
      list.unshift({ id: Math.random().toString(36).slice(2, 10), latex: it.latex, display: it.display !== false, fav: !!it.fav, time: it.time || Date.now() });
    }
  }
  saveLibrary(list);
  return list;
}
