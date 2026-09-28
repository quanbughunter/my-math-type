// Khác biệt giữa các nền tảng: web/PWA, Windows (Electron), Android (Capacitor), Word add-in.

export const isElectron = () => typeof window !== 'undefined' && !!window.mymathNative;
export const isCapacitor = () => typeof window !== 'undefined' && !!window.Capacitor?.isNativePlatform?.();
export const isAndroid = () => /Android/i.test(navigator.userAgent);

function blobToBase64(blob) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(String(r.result).split(',')[1]);
    r.onerror = rej;
    r.readAsDataURL(blob);
  });
}
function blobToDataUrl(blob) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(String(r.result));
    r.onerror = rej;
    r.readAsDataURL(blob);
  });
}

function legacyCopy(text) {
  const ta = document.createElement('textarea');
  ta.value = text;
  ta.setAttribute('readonly', '');
  ta.style.cssText = 'position:fixed;top:-1000px;opacity:0';
  document.body.appendChild(ta);
  ta.select();
  let ok = false;
  try {
    ok = document.execCommand('copy');
  } catch {
    ok = false;
  }
  ta.remove();
  if (!ok) throw new Error('Trình duyệt không cho phép copy');
}

/** Copy chữ thuần. */
export async function copyText(text) {
  if (isElectron()) return window.mymathNative.copyText(text);
  if (isCapacitor()) {
    const { Clipboard } = await import('@capacitor/clipboard');
    return Clipboard.write({ string: text });
  }
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    legacyCopy(text);
  }
}

/** Copy ảnh PNG. Trả về false nếu nền tảng không hỗ trợ (khi đó nên chia sẻ/tải tệp). */
export async function copyImage(pngBlob) {
  if (isElectron()) {
    await window.mymathNative.copyImage(await blobToDataUrl(pngBlob));
    return true;
  }
  if (isCapacitor()) return false;
  if (!window.ClipboardItem || !navigator.clipboard?.write) return false;
  await navigator.clipboard.write([new ClipboardItem({ 'image/png': pngBlob })]);
  return true;
}

/** Lưu / chia sẻ tệp. Trên Android mở bảng Chia sẻ (gửi sang Word, Zalo, Drive…). */
export async function saveFile(blob, filename) {
  if (isCapacitor()) {
    const { Filesystem, Directory } = await import('@capacitor/filesystem');
    const { Share } = await import('@capacitor/share');
    const data = await blobToBase64(blob);
    const res = await Filesystem.writeFile({ path: filename, data, directory: Directory.Cache });
    await Share.share({ title: filename, files: [res.uri], dialogTitle: 'Gửi tệp tới…' });
    return 'shared';
  }
  // Trên điện thoại (PWA) thử bảng chia sẻ của hệ điều hành trước
  if (isAndroid() && navigator.canShare) {
    const file = new File([blob], filename, { type: blob.type });
    if (navigator.canShare({ files: [file] })) {
      try {
        await navigator.share({ files: [file], title: filename });
        return 'shared';
      } catch (e) {
        if (e?.name === 'AbortError') return 'cancelled';
      }
    }
  }
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
  return 'downloaded';
}

/** Mở tệp văn bản do người dùng chọn. */
export function pickTextFile(accept = '.json,application/json') {
  return new Promise((resolve) => {
    const inp = document.createElement('input');
    inp.type = 'file';
    inp.accept = accept;
    inp.onchange = () => {
      const f = inp.files?.[0];
      if (!f) return resolve(null);
      f.text().then(resolve, () => resolve(null));
    };
    inp.click();
  });
}
