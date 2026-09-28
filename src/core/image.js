// LaTeX → SVG/PNG bằng MathJax (chỉ tải khi cần, chạy hoàn toàn offline).

import mathjaxUrl from 'mathjax/es5/tex-svg-full.js?url';
import { prepareLatex, LATEX_MACROS } from './latex.js';

let loading = null;

function loadMathJax() {
  if (loading) return loading;
  loading = new Promise((resolve, reject) => {
    const macros = {};
    for (const [k, v] of Object.entries(LATEX_MACROS)) macros[k.slice(1)] = v;
    window.MathJax = {
      startup: { typeset: false },
      tex: { macros },
      svg: { fontCache: 'none' },
      options: { enableMenu: false },
    };
    const s = document.createElement('script');
    s.src = mathjaxUrl;
    s.async = true;
    s.onload = () => window.MathJax.startup.promise.then(() => resolve(window.MathJax), reject);
    s.onerror = () => reject(new Error('Không tải được MathJax'));
    document.head.appendChild(s);
  });
  return loading;
}

/** LaTeX → chuỗi <svg> độc lập (dùng đường vẽ, không cần phông). */
export async function latexToSvg(latex, { display = true, color = '#000000' } = {}) {
  const MJ = await loadMathJax();
  const node = await MJ.tex2svgPromise(prepareLatex(latex), { display });
  const svg = node.querySelector('svg');
  svg.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
  svg.style.color = color;
  let str = new XMLSerializer().serializeToString(svg);
  str = str.replace(/currentColor/g, color);
  return str;
}

/** SVG → PNG Blob, scale = độ phóng. */
export async function svgToPng(svgStr, { scale = 4, background = null, exPx = 8.5, padding = 6 } = {}) {
  const w = parseFloat((/width="([\d.]+)ex"/.exec(svgStr) || [])[1] || '10') * exPx;
  const h = parseFloat((/height="([\d.]+)ex"/.exec(svgStr) || [])[1] || '4') * exPx;
  const sized = svgStr
    .replace(/width="[\d.]+ex"/, `width="${w}px"`)
    .replace(/height="[\d.]+ex"/, `height="${h}px"`);
  const url = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(sized);
  const img = new Image();
  await new Promise((res, rej) => {
    img.onload = res;
    img.onerror = () => rej(new Error('Không vẽ được ảnh'));
    img.src = url;
  });
  const canvas = document.createElement('canvas');
  canvas.width = Math.ceil((w + padding * 2) * scale);
  canvas.height = Math.ceil((h + padding * 2) * scale);
  const ctx = canvas.getContext('2d');
  if (background) {
    ctx.fillStyle = background;
    ctx.fillRect(0, 0, canvas.width, canvas.height);
  }
  ctx.drawImage(img, padding * scale, padding * scale, w * scale, h * scale);
  return new Promise((res) => canvas.toBlob(res, 'image/png'));
}

export async function latexToPng(latex, opts = {}) {
  const svg = await latexToSvg(latex, { display: opts.display ?? true });
  return svgToPng(svg, opts);
}
