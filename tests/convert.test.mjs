// Kiểm thử chuyển đổi LaTeX → MathML (Word) → OMML → LaTeX, và tạo .docx mẫu.
// Chạy: node tests/convert.test.mjs [thư_mục_xuất]
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { writeFileSync, mkdirSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
globalThis.DOMParser = DOMParser;
globalThis.XMLSerializer = XMLSerializer;

const { latexToWordMathML } = await import('../src/core/mathml.js');
const { latexToOmml } = await import('../src/core/omml.js');
const { ommlToLatexList } = await import('../src/core/omml2latex.js');
const { buildDocx } = await import('../src/core/ooxml.js');

export const CASES = [
  ['Nghiệm phương trình bậc hai', String.raw`x_{1,2}=\frac{-b\pm\sqrt{b^2-4ac}}{2a}`],
  ['Ngoặc co giãn', String.raw`\left(\frac{a}{b}\right)^{2}+\left|x-1\right|`],
  ['Tích phân', String.raw`\int_{0}^{1} x^2\,\mathrm{d}x=\frac{1}{3}`],
  ['Tổng', String.raw`\sum_{i=1}^{n} i^2=\frac{n(n+1)(2n+1)}{6}`],
  ['Giới hạn', String.raw`\lim_{x\to 0}\frac{\sin x}{x}=1`],
  ['Ma trận', String.raw`A=\begin{pmatrix}1&2&3\\4&5&6\end{pmatrix}`],
  ['Định thức', String.raw`\det A=\begin{vmatrix}a&b\\c&d\end{vmatrix}=ad-bc`],
  ['Hệ phương trình', String.raw`\begin{cases}x+y=1\\x-y=2\end{cases}`],
  ['Hàm từng khúc', String.raw`f(x)=\begin{cases}x^2 & \text{khi } x\ge 0\\-x & \text{khi } x<0\end{cases}`],
  ['Tuyển (hoặc)', String.raw`\left[\begin{array}{l}x=1\\x=2\end{array}\right.`],
  ['Vector, góc', String.raw`\overrightarrow{AB}+\vec{u}=\vec{0},\ \widehat{ABC}=90^{\circ}`],
  ['Tổ hợp', String.raw`\mathrm{C}_{n}^{k}=\frac{n!}{k!(n-k)!}`],
  ['Hoá học', String.raw`2\mathrm{H}_2+\mathrm{O}_2\xrightarrow{t^{\circ}}2\mathrm{H}_2\mathrm{O}`],
  ['Căn bậc n', String.raw`\sqrt[3]{x+1}+\sqrt{2}`],
  ['Logarit', String.raw`\log_{2}8=3,\ \ln e=1`],
  ['Đạo hàm riêng', String.raw`\frac{\partial^2 u}{\partial x^2}+\frac{\partial^2 u}{\partial y^2}=0`],
  ['Biến đổi', String.raw`\begin{aligned}(a+b)^2&=a^2+2ab+b^2\\&\ge 4ab\end{aligned}`],
  ['Tập hợp', String.raw`A=\left\{x\in\mathbb{R}\mid x^2<4\right\}`],
  ['Gạch trên, mũ', String.raw`\overline{z}=a-bi,\ \hat{x},\ \tilde{y},\ \bar{x}`],
  ['Ngoặc nhọn trên/dưới', String.raw`\underbrace{1+1+\cdots+1}_{n}=n`],
  ['Tích phân kép', String.raw`\iint_{D} f(x,y)\,\mathrm{d}x\,\mathrm{d}y`],
  ['Đồng vị', String.raw`{}_{6}^{14}\mathrm{C}`],
  ['Chữ tiếng Việt', String.raw`S=\pi r^2\ \text{(diện tích hình tròn)}`],
  ['Khung', String.raw`\boxed{E=mc^2}`],
  ['Đạo hàm cấp 2 (MathLive)', String.raw`f^{\doubleprime}(x)=6x`],
  ['Nhị thức', String.raw`(a+b)^n=\sum_{k=0}^{n}\binom{n}{k}a^{n-k}b^k`],
  ['Rút gọn', String.raw`\frac{\cancel{2}x}{\cancel{2}}=x`],
  ["Cung", String.raw`\overparen{AB}`],
  ['Phản ứng có xúc tác', String.raw`\mathrm{N}_2+3\mathrm{H}_2\xrightarrow[\text{xt}]{t^{\circ}}2\mathrm{NH}_3`],
  ['Thuận nghịch', String.raw`\mathrm{CH_3COOH}\rightleftharpoons\mathrm{CH_3COO^-}+\mathrm{H^+}`],
  ['Macro MathLive', String.raw`\int\exponentialE^{x}\,\differentialD x=\exponentialE^{x}+C`],
];

const out = process.argv[2] || join(tmpdir(), 'mymath-test');
mkdirSync(out, { recursive: true });
let fail = 0;
const items = [];
for (const [name, tex] of CASES) {
  try {
    const mml = latexToWordMathML(tex, { display: true });
    const omml = latexToOmml(tex, { display: true });
    new DOMParser({ onError: (lvl, msg) => { if (lvl !== 'warning') throw new Error(msg); } }).parseFromString(omml, 'application/xml');
    const back = ommlToLatexList(omml)[0].latex;
    // vòng lại lần 2 để chắc LaTeX quay về vẫn chuyển được
    latexToOmml(back, { display: true });
    items.push({ caption: `${name}:  ${tex}`, omml });
    console.log(`✓ ${name}\n   MML : ${mml.slice(0, 160)}${mml.length > 160 ? '…' : ''}\n   BACK: ${back}`);
  } catch (e) {
    fail++;
    console.log(`✗ ${name}: ${e.message}`);
  }
}
const bytes = await buildDocx(items, { heading: 'MyMath — kiểm thử chuyển đổi', type: 'nodebuffer' });
writeFileSync(join(out, 'mymath-test.docx'), bytes);
console.log(`\n${CASES.length - fail}/${CASES.length} đạt. Đã ghi ${out}/mymath-test.docx`);
process.exit(fail ? 1 : 0);
