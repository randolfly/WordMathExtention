import { AsciiMath } from "asciimath-parser";
import katex from "katex";
import { mml2omml } from "@hungknguyen/mathml2omml";
import "katex/dist/katex.min.css";

export function asciiMathToLatex(ascii: string): string {
  const trimmed = ascii.trim();
  if (!trimmed) return "";
  const am = new AsciiMath();
  return am.toTex(trimmed);
}

export function latexToPreviewHtml(tex: string): string {
  if (!tex.trim()) return "";
  return katex.renderToString(tex, {
    throwOnError: false,
    strict: "ignore",
    displayMode: true,
    output: "htmlAndMathml"
  });
}

export function latexToMathml(tex: string): string {
  if (!tex.trim()) return "";
  return katex.renderToString(tex, {
    throwOnError: false,
    strict: "ignore",
    displayMode: true,
    output: "mathml"
  });
}

/**
 * 从 KaTeX 生成的 HTML 中提取纯 MathML 内容
 * @param html 包含 <span class="katex">...<math>...</math></span> 的 HTML 字符串
 * @returns 纯 <math>...</math> 字符串，移除 <annotation> 标签
 */
function extractPureMathml(html: string): string {
  // 使用正则表达式提取 <math> 标签及其内容
  const mathMatch = html.match(/<math[^>]*>[\s\S]*?<\/math>/i);
  if (mathMatch && mathMatch[0]) {
    let pureMathml = mathMatch[0];
    // 移除 <annotation> 标签
    pureMathml = pureMathml.replace(/<annotation[^>]*>[\s\S]*?<\/annotation>/gi, "");
    return pureMathml;
  }
  // 如果没有找到 <math> 标签，返回原始字符串（可能已经是纯 MathML）
  return html;
}

export function mathmlToOmml(mathml: string): string {
  const trimmed = mathml.trim();
  if (!trimmed) return "";

  // 提取纯的 <math> 元素，移除外层的 <span class="katex"> 包装
  const pureMathml = extractPureMathml(trimmed);
  console.log("pure mathml:", pureMathml);

  try {
    const omml = mml2omml(pureMathml);
    // 确保 OMML 包含正确的根节点
    if (!omml.includes("<m:oMath") && !omml.includes("<oMath")) {
      console.warn("OMML 缺少 <m:oMath> 根节点，尝试修复");
      const mNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/math";
      return `<m:oMath xmlns:m="${mNamespace}">${omml}</m:oMath>`;
    }
    return omml;
  } catch (error) {
    console.error("MathML 转换为 OMML 失败:", error);
    throw error;
  }
}

export function asciiMathToOmml(ascii: string): { tex: string; mathml: string; omml: string } {
  const tex = asciiMathToLatex(ascii);
  const mathml = latexToMathml(tex);
  // 这里的mathml包含一个外层的<span class="katex">，需要提取其中的<math>元素才能正确转换为OMML
  // <span class="katex"><math xmlns="http://www.w3.org/1998/Math/MathML" display="block"><semantics><mrow><mstyle scriptlevel="0" displaystyle="true"><mrow><mi>x</mi><mo>+</mo><mn>1</mn></mrow></mstyle></mrow><annotation encoding="application/x-tex">\displaystyle{ x + 1 }</annotation></semantics></math></span>
  const omml = mathmlToOmml(mathml);
  console.log("omml:", omml);
  return { tex, mathml, omml };
}