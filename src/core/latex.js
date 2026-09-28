// Chuẩn hoá LaTeX do MathLive sinh ra để các bộ chuyển đổi khác (Temml, MathJax) hiểu được.

/** Các lệnh riêng của MathLive → LaTeX chuẩn. Dùng làm macro cho Temml và MathJax. */
export const LATEX_MACROS = {
  '\\differentialD': '\\mathrm{d}',
  '\\capitalDifferentialD': '\\mathrm{D}',
  '\\exponentialE': '\\mathrm{e}',
  '\\imaginaryI': '\\mathrm{i}',
  '\\imaginaryJ': '\\mathrm{j}',
  '\\mleft': '\\left',
  '\\mright': '\\right',
  '\\R': '\\mathbb{R}',
  '\\N': '\\mathbb{N}',
  '\\Z': '\\mathbb{Z}',
  '\\Q': '\\mathbb{Q}',
  '\\C': '\\mathbb{C}',
  '\\P': '\\mathbb{P}',
  '\\degree': '{^\\circ}',
  '\\lparen': '(',
  '\\rparen': ')',
  '\\doubleprime': '\\prime\\prime',
  '\\tripleprime': '\\prime\\prime\\prime',
  '\\roundimplies': '\\mathrel{\u2970}',
  '\\biconditional': '\\leftrightarrow',
  '\\Colon': '\\mathrel{::}',
};

/**
 * Bỏ ô trống (placeholder) của MathLive và dọn các cấu trúc lạ.
 * `\placeholder[id]{nội dung}` → `{nội dung}`
 */
export function prepareLatex(latex) {
  if (!latex) return '';
  let s = String(latex);
  // \placeholder[...]{...} (nội dung không lồng ngoặc sâu trong thực tế)
  s = s.replace(/\\placeholder(?:\[[^\]]*\])?\{([^{}]*)\}/g, '{$1}');
  s = s.replace(/\\placeholder(?:\[[^\]]*\])?/g, '{}');
  // MathLive đôi khi xuất \mathrm{...} rỗng
  s = s.replace(/\\(mathrm|mathit|mathbf)\{\}/g, '{}');
  return s.trim();
}

/** Macro cho Temml (khoá không có dấu gạch chéo ngược vẫn được chấp nhận). */
export function temmlMacros() {
  return { ...LATEX_MACROS };
}

/** Kiểm tra nhanh xem công thức có rỗng không. */
export function isBlankLatex(latex) {
  return !prepareLatex(latex).replace(/[{}\s]/g, '');
}
