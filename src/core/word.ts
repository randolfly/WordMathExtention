import { asciiMathToOmml } from "./convert";
import { buildOoxmlForOmml, extractAsciiMathFromOoxml } from "./ooxml";

const WORDMATH_TAG = "wordmath";

function ensureOffice(): void {
  if (typeof Office === "undefined" || typeof Word === "undefined") {
    throw new Error("当前环境未加载 Office.js（需要在 Word 插件任务窗格中运行）。");
  }
}

async function getSelectedWordMathContentControl(context: Word.RequestContext): Promise<Word.ContentControl | null> {
  const range = context.document.getSelection();
  const contentControls = range.getContentControls();
  contentControls.load("items/id/tag");
  await context.sync();

  const first = contentControls.items[0];
  if (!first) return null;
  if (first.tag !== WORDMATH_TAG) return null;
  return first;
}

export async function loadAsciiMathFromSelection(): Promise<string | null> {
  ensureOffice();
  return Word.run(async (context) => {
    const cc = await getSelectedWordMathContentControl(context);
    if (!cc) return null;
    const contentRange = cc.getRange(Word.RangeLocation.content);
    const ooxmlResult = contentRange.getOoxml();
    await context.sync();
    return extractAsciiMathFromOoxml(ooxmlResult.value);
  });
}

export async function insertOrUpdateEquation(ascii: string): Promise<void> {
  ensureOffice();
  const trimmed = ascii.trim();
  if (!trimmed) {
    throw new Error("AsciiMath 为空。");
  }

  const { tex, omml } = asciiMathToOmml(trimmed);
  if (!tex.trim()) {
    throw new Error("AsciiMath 转换得到的 LaTeX 为空。");
  }
  if (!omml.trim()) {
    throw new Error("转换失败：未得到 OMML。");
  }

  const ooxml = buildOoxmlForOmml(omml, trimmed);

  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    let cc = await getSelectedWordMathContentControl(context);

    if (!cc) {
      cc = selection.insertContentControl();
      cc.tag = WORDMATH_TAG;
      cc.title = "WordMath";
      cc.appearance = Word.ContentControlAppearance.boundingBox;
    }

    const contentRange = cc.getRange(Word.RangeLocation.content);
    contentRange.insertOoxml(ooxml, Word.InsertLocation.replace);
    await context.sync();
  });
}
