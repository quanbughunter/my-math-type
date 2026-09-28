// Tạo manifest Word add-in từ mẫu, trỏ tới nơi đang host MyMath (phải là https).
//   node scripts/make-manifest.mjs                       → dùng "addinUrl" trong package.json
//   node scripts/make-manifest.mjs --url https://ten.github.io/my-math-type/
// Ghi ra: addin/manifest.xml, addin/manifest.localhost.xml và dist/manifest.xml (nếu đã build).
import { readFileSync, writeFileSync, existsSync } from 'node:fs';

const pkg = JSON.parse(readFileSync('package.json', 'utf8'));
const argIdx = process.argv.indexOf('--url');
let url = argIdx > 0 ? process.argv[argIdx + 1] : process.env.ADDIN_URL || pkg.mymath?.addinUrl;
if (!url || !/^https:\/\//.test(url)) {
  console.error('Cần URL https, ví dụ: node scripts/make-manifest.mjs --url https://ten.github.io/my-math-type/');
  process.exit(1);
}
if (!url.endsWith('/')) url += '/';
const tpl = readFileSync('addin/manifest.template.xml', 'utf8');
const version = pkg.version.split('.').concat(['0', '0', '0']).slice(0, 4).join('.');

function make(base, id) {
  return tpl
    .replaceAll('{{BASE_URL}}', base)
    .replaceAll('{{ORIGIN}}', new URL(base).origin)
    .replaceAll('{{VERSION}}', version)
    .replaceAll('{{ADDIN_ID}}', id);
}

const ID = pkg.mymath?.addinId || '7f3c2a51-9b4e-4d8a-a6c1-5e2f0b9d4c73';
const ID_DEV = pkg.mymath?.addinIdDev || '7f3c2a51-9b4e-4d8a-a6c1-5e2f0b9d4c74';

writeFileSync('addin/manifest.xml', make(url, ID));
writeFileSync('addin/manifest.localhost.xml', make('https://localhost:3000/', ID_DEV));
if (existsSync('dist')) writeFileSync('dist/manifest.xml', make(url, ID));
console.log(`Manifest Word add-in → ${url}`);
