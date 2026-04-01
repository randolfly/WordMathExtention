const META_PREFIX = "WMASCII:";

function xmlEscapeText(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function toBase64Utf8(value: string): string {
  const bytes = new TextEncoder().encode(value);
  let binary = "";
  for (const b of bytes) binary += String.fromCharCode(b);
  return btoa(binary);
}

export function extractAsciiMathFromOoxml(ooxml: string): string | null {
  const match = ooxml.match(/WMASCII:([A-Za-z0-9+/=]+)/);
  if (!match) return null;
  const b64 = match[1];
  try {
    const binary = atob(b64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    return new TextDecoder().decode(bytes);
  } catch {
    return null;
  }
}

/**
 * 构建用于插入 Word 的 OOXML（Flat OPC / pkg:package）。
 *
 * Word JavaScript API 的 `Range.insertOoxml()` 官方示例使用 Flat OPC 格式（`<pkg:package ...>`），
 * 其中包含 `/_rels/.rels` 与 `/word/document.xml` 两个最小 part。
 */
export function buildOoxmlForOmml(omml: string, asciiMath: string): string {
  const m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
  const w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";

  const trimmed = omml.trim();
  if (!trimmed) {
    throw new Error("OMML 为空，无法生成 OOXML。");
  }

  // 兼容不同来源的 OMML：
  // - 可能是 <m:oMathPara>...</m:oMathPara>
  // - 可能是 <m:oMath>...</m:oMath>
  // - 也可能是 <m:...> 片段（无根节点）
  let mathNode = trimmed;
  const lower = trimmed.toLowerCase();
  const hasOMathPara = lower.startsWith("<m:omathpara") || lower.startsWith("<omathpara");
  const hasOMath = lower.startsWith("<m:omath") || lower.startsWith("<omath");

  if (!hasOMathPara) {
    if (!hasOMath) {
      mathNode = `<m:oMath xmlns:m="${m}">${mathNode}</m:oMath>`;
    }
    mathNode = `<m:oMathPara>${mathNode}</m:oMathPara>`;
  }

  const meta = `${META_PREFIX}${toBase64Utf8(asciiMath)}`;
  const metaEscaped = xmlEscapeText(meta);

  const documentXml = `
<w:document xmlns:w="${w}" xmlns:m="${m}">
  <w:body>
    <w:p>
      ${mathNode}
      <w:r>
        <w:rPr>
          <w:vanish />
        </w:rPr>
        <w:t xml:space="preserve">${metaEscaped}</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`.trim();

  return `
<pkg:package xmlns:pkg="${pkg}">
  <pkg:part pkg:name="/_rels/.rels"
            pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"
            pkg:padding="512">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1"
                      Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                      Target="word/document.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml"
            pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      ${documentXml}
    </pkg:xmlData>
  </pkg:part>
</pkg:package>`.trim();
}
