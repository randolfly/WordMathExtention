# Office API 调研：如何由 OMML 生成 Word 公式

## 结论（现阶段最推荐）

Office.js / Word JavaScript API **没有**提供“直接插入 OMML（如 insertOmml）”的专用 API。要把 OMML 变成 Word 原生公式，主流且可控的方式是：

1. 将 OMML 包装进 **WordprocessingML（OOXML）**，通常放在 `<w:document><w:body><w:p> ... </w:p></w:body></w:document>` 中；
2. 再用 **`Range.insertOoxml()`** 把 OOXML 插入到文档中。官方示例使用的是 **Flat OPC（`<pkg:package ...>`）** 格式（包含 `/_rels/.rels` 与 `/word/document.xml`）。见 Microsoft Learn 的 `Word.Range.insertOoxml` 示例。  

因此：**“OMML → 公式” 本质上就是 “构造正确的 OOXML → insertOoxml”。**

## insertOoxml 的官方示例要点

Microsoft Learn 的 `Word.Range.insertOoxml()` 示例里，传入的是：

- `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"> ... </pkg:package>`
- 通常至少包含这些 part（不同示例/容错程度可能略有差异）：
  - `/[Content_Types].xml`
  - `/_rels/.rels`：指向 `word/document.xml`
  - `/word/document.xml`：包含 `<w:document ...><w:body>...</w:body></w:document>`

这说明官方“推荐/至少可靠”的形态是 **Flat OPC 包**，而不仅仅是一个 `<w:p>` 片段。

## 另一个插入入口：setSelectedDataAsync（旧式 Office.js）

在较早期的 Office.js 中，也可以通过：

- `Office.context.document.setSelectedDataAsync(ooxml, { coercionType: Office.CoercionType.Ooxml }, cb)`

把 OOXML 写入选区。它同样依赖 “OOXML 字符串” 的正确性。

## 对本项目的影响

我们已将 `buildOoxmlForOmml()` 输出调整为 **Flat OPC（pkg:package）**，以贴近官方示例并提高兼容性。

- 生成 OOXML：`src/core/ooxml.ts`
- 插入：`src/core/word.ts`（`Range.insertOoxml(..., Word.InsertLocation.replace)`）

## 风险与待验证点（建议你我一起验证）

1. **Word 桌面端 vs Word Web**：不同宿主对 `insertOoxml` 的支持/容错不同，建议分别验证。
2. **插入位置语义**：Flat OPC 中的内容通常以段落 `<w:p>` 形式出现；若用户期望“行内公式”，需要改用 `<m:oMath>`（而不是 `<m:oMathPara>`）并构造更合适的段落结构。
3. **字体/渲染差异**：相同公式在 Word 与 KaTeX 预览可能有排版差异，这是正常现象（不同渲染引擎）。
