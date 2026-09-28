# MyMath — soạn công thức kiểu MathType cho Word

Một mã nguồn, ba cách dùng:

| Nền tảng | Dạng cài | Đưa công thức vào Word |
|---|---|---|
| **Windows** | Bộ cài `.exe` (Electron) hoặc cài từ web (PWA) | **Copy vào Word** → `Ctrl+V` ⇒ thành phương trình gốc (Equation) |
| **Word (Windows, Mac, Word trên web)** | **Word add-in**: tab riêng **MyMath** trên thanh công cụ | Chèn thẳng phương trình gốc, lấy công thức cũ ra sửa |
| **Android** | Cài từ web (PWA) hoặc `.apk` (Capacitor) | **Tải .docx** mở bằng Word, hoặc **Copy ảnh** |

> Tên ứng dụng là *MyMath* (MathType là thương hiệu của Wiris — tránh trùng tên khi đưa lên cửa hàng).

---

## 1. Tính năng

- Ô soạn trực quan (MathLive): gõ `/` ra phân số, `^` số mũ, `_` chỉ số, `sqrt`, `int`, `sum`, `lim`, `alpha`, `>=`, `!=`, `->`…; `Tab` nhảy ô trống.
- **12 nhóm mẫu kiểu MathType** (≈280 nút): phân số, căn, ngoặc co giãn, tổng/tích phân/giới hạn/đạo hàm, ma trận, định thức, **hệ phương trình, tuyển “hoặc”**, vector `AB→`, góc `ABC^`, tổ hợp `C_n^k`, chỉnh hợp, **hoá học** (`t°`, xúc tác, ⇌, ↑↓, số oxi hoá, đồng vị), Hy Lạp, tập hợp, logic, kiểu chữ, khoảng cách.
- Xuất: **Copy vào Word** (MathML — Word tự đổi thành phương trình gốc), **Copy LaTeX**, **Copy ảnh PNG** nét (MathJax, chạy offline), **Tải .docx** (OMML gốc).
- **Thư viện**: tự lưu công thức gần đây, đánh dấu ★, xuất cả thư viện ra một tệp Word, sao lưu/nhập `.json` để chuyển giữa máy tính và điện thoại.
- Kiểu **Riêng dòng / Cùng dòng**, giao diện sáng/tối, bàn phím toán trên màn hình (điện thoại).
- **Font và cỡ chữ (pt)** khi đưa vào Word: Cambria Math (mặc định Word), **Times New Roman kiểu MathType** (chữ, số Times; biến in nghiêng), các font toán STIX Two Math / Latin Modern Math / TeX Gyre Termes Math, hoặc font tự nhập; cỡ chính (vd 12 pt) và cỡ chỉ số (vd 9 pt, để trống = Word tự tính ≈ 3/4). Áp dụng khi chèn bằng add-in và khi tải .docx.
- Add-in Word: **Chèn vào Word**, **Lấy công thức đang chọn** (đọc phương trình Word → LaTeX → sửa → **Cập nhật**).
- Chạy **offline** hoàn toàn (trừ Office.js của add-in phải tải từ Microsoft).

## 2. Cấu trúc thư mục

```
my-math-type/
├─ index.html            giao diện chính (Windows / Android / web)
├─ taskpane.html         khung Word add-in
├─ commands.html         tệp lệnh bắt buộc của add-in
├─ src/
│  ├─ main.js            logic ứng dụng
│  ├─ taskpane.js        logic Word add-in (Office.js)
│  ├─ core/
│  │  ├─ latex.js        chuẩn hoá LaTeX của MathLive
│  │  ├─ mathml.js       LaTeX → MathML (Temml) + chuẩn hoá "kiểu Word"
│  │  ├─ omml.js         MathML → OMML (phương trình gốc của Word) — tự viết
│  │  ├─ omml2latex.js   OMML → LaTeX (mở lại công thức Word để sửa)
│  │  ├─ ooxml.js        đóng gói .docx và gói chèn cho add-in
│  │  ├─ image.js        LaTeX → SVG/PNG (MathJax)
│  │  ├─ platform.js     clipboard / lưu tệp / chia sẻ theo nền tảng
│  │  └─ storage.js      thư viện, cài đặt (lưu trên máy)
│  └─ ui/                bảng mẫu, ô soạn, biểu tượng, CSS
├─ public/               icon, manifest PWA, service worker
├─ addin/                manifest Word add-in + script cài/gỡ trên Windows
├─ electron/             vỏ ứng dụng Windows
├─ android/              dự án Android (Capacitor) — mở bằng Android Studio
├─ release/              bộ cài Windows đã build sẵn
├─ tests/                kiểm thử chuyển đổi (npm test)
└─ .github/workflows/    tự deploy web/add-in và build .exe/.apk trên GitHub
```

## 3. Windows

**Dùng bộ cài có sẵn:** mở `release/MyMath-Setup-1.0.0.exe` (hoặc `MyMath-1.0.0-portable.exe` — chạy luôn, không cần cài).
Bộ cài chưa ký số nên Windows SmartScreen có thể cảnh báo → bấm **More info → Run anyway**.

**Từ mã nguồn** (cần [Node.js 22+](https://nodejs.org)):

```bat
cd D:\C1.CODING\myAPP\my-math-type
npm install
npm run dev        :: chạy thử trên trình duyệt: http://localhost:3000
npm run app        :: chạy thử dạng ứng dụng Windows
npm run dist:win   :: tạo bộ cài vào thư mục release\
npm test           :: kiểm thử 31 công thức mẫu (LaTeX → MathML → OMML → LaTeX)
```

**Dán vào Word:** bấm **Copy vào Word** (hoặc `Ctrl+Enter`) → sang Word đặt con trỏ → `Ctrl+V`.
Công thức trở thành phương trình gốc của Word (sửa được bằng công cụ Equation, in nét, không phải ảnh).
Nếu Word dán ra một đoạn chữ bắt đầu bằng `<math…`: dùng Word 2013 trở lên, hoặc dùng **Tải .docx** / add-in.

## 4. Word add-in (nút ngay trong Word)

Add-in là một trang web chạy trong khung bên phải Word, nên cần được host qua **https**. Cách đơn giản nhất là GitHub Pages (miễn phí):

1. Tạo repo **`my-math-type`** trên GitHub (ví dụ tài khoản `quanbughunter`), rồi đẩy mã lên:
   ```bat
   cd D:\C1.CODING\myAPP\my-math-type
   git init -b main
   git add .
   git commit -m "MyMath 1.0"
   git remote add origin https://github.com/quanbughunter/my-math-type.git
   git push -u origin main
   ```
2. Trên GitHub: **Settings → Pages → Source: GitHub Actions**. Workflow `Deploy web + Word add-in` tự build và đưa lên
   `https://quanbughunter.github.io/my-math-type/` (manifest ở `…/manifest.xml`).
   *Tài khoản/tên repo khác?* Sửa `mymath.addinUrl` trong `package.json` rồi chạy `npm run manifest`
   (hoặc `npm run manifest -- --url https://ten.github.io/ten-repo/`).
3. Cài vào Word trên Windows (làm **một lần**):
   - Đóng hết Word → **bấm đúp `addin\cai-addin-word.cmd`**. Windows hỏi quyền quản trị một lần (để chia sẻ thư mục add-in, chỉ đọc, chỉ trên máy này) → **Yes**.
   - Mở Word → **Home → Add-ins → More Add-ins** → thẻ **SHARED FOLDER** → chọn **MyMath** → **Add**.
   - Từ đó mỗi lần mở Word đều có sẵn tab **MyMath** → **Công thức**; không phải thêm lại.
   - Không thấy MyMath trong SHARED FOLDER: đóng Word, bấm đúp `addin\cai-addin-word-xoa-cache.cmd` rồi thử lại. Gỡ: `addin\go-addin-word.cmd`.
   - Cách này là "danh mục add-in tin cậy" (Trusted Add-in Catalog) của Office: script tạo thư mục `%LOCALAPPDATA%\MyMath\WordAddin`, chia sẻ thành `\\<tên máy>\MyMathAddin` và khai báo nó trong Word.
4. **Word trên web / Mac:** Chèn → Bổ trợ → *Tải bổ trợ lên (Upload My Add-in)* → chọn `addin/manifest.xml`.

Trong khung add-in:
- **Chèn vào Word** (`Ctrl+Enter`): chèn tại con trỏ (nếu đang bôi đen chữ thường thì chèn phía sau, không xoá chữ).
- **Lấy công thức đang chọn**: bôi đen một phương trình trong Word (hoặc đặt con trỏ trong phương trình nằm riêng dòng) → mở ra sửa → **Cập nhật công thức trong Word**. Dùng được cả với phương trình tạo bằng Word.

*Phát triển add-in trên máy (không cần GitHub):* `npm run addin:certs` (một lần) → `npm run addin:dev` →
cài `addin\manifest.localhost.xml` bằng `addin\install-addin.ps1 -Manifest addin\manifest.localhost.xml`.

## 5. Android

- **Nhanh nhất (PWA):** mở `https://quanbughunter.github.io/my-math-type/` bằng Chrome → menu ⋮ → **Cài đặt ứng dụng / Thêm vào màn hình chính**. Có biểu tượng riêng, mở toàn màn hình, chạy offline.
- **Tệp .apk:** trên GitHub → tab **Actions → Build Windows + Android → Run workflow** → tải `MyMath-Android` ở mục *Artifacts* → cài lên máy (cho phép cài từ nguồn không xác định).
  Hoặc tự build bằng Android Studio: `npm run android:sync` → `npm run android:open` → *Build → Build APK(s)*.
- Word cho Android **không hỗ trợ add-in**. Trên điện thoại dùng **Tải .docx** (mở bảng chia sẻ → chọn Word; công thức trong tệp là phương trình gốc, copy sang tài liệu khác được) hoặc **Copy ảnh**.

## 6. Cách chuyển đổi hoạt động

```
MathLive (soạn) ──LaTeX──▶ Temml ──MathML──▶ chuẩn hoá cho Word ──┬─▶ Clipboard (Ctrl+V → Word tự đổi)
                                                                 └─▶ omml.js ──OMML──▶ .docx / add-in
Word (phương trình có sẵn) ──OMML──▶ omml2latex.js ──LaTeX──▶ MathLive (sửa lại)
```

Phần chuẩn hoá lo những chỗ Word hay làm hỏng: ngoặc `\left…\right` → ngoặc co giãn (`m:d`), thân của ∑/∫ được gom đúng (`m:nary`),
`sin x`, `log₂`, `lim` thành hàm (`m:func`), hệ/ma trận căn trái (`m:m`), `aligned` căn dấu = (`m:eqArr`), vector/góc/mũ (`m:acc`),
gạch trên (`m:bar`), mũi tên có chữ (`m:groupChr`), khung (`m:borderBox`), chỉ số trái (`m:sPre`), khoảng trắng LaTeX → ký tự Unicode.

## 7. Giới hạn đã biết

- Bộ cài Windows chưa ký số (SmartScreen cảnh báo lần đầu).
- Chưa có `.apk` dựng sẵn (cần Android SDK) — dùng workflow GitHub hoặc Android Studio như mục 5, hoặc cài bản PWA.
- Màu chữ (`\color`) và một số lệnh LaTeX hiếm không chuyển sang Word; app sẽ báo ngay dưới ô soạn.
- Word trên Android có thể không nhận MathML khi dán — dùng **Tải .docx**.
- Đã kiểm thử tự động: 31 công thức qua cả chuỗi chuyển đổi, 284 nút mẫu, tệp .docx mở được (LibreOffice), luồng add-in (Office.js giả lập).
  Chưa chạy trên Microsoft Word thật — nếu công thức nào dán/chèn sai, gửi lại mã LaTeX (nút *Copy LaTeX*) để sửa bộ chuyển.

Thư viện dùng: MathLive (MIT), Temml (MIT), MathJax (Apache-2.0), JSZip (MIT), Electron (MIT), Capacitor (MIT).
