import { readFile, mkdir, mkdtemp, link, rm } from "node:fs/promises";
import path from "node:path";
import { createRequire } from "node:module";

const [input, output] = process.argv.slice(2);
if (!input || !output || path.extname(output).toLowerCase() !== ".xlsx") {
  throw new Error("用法：node export.mjs 输入.json 输出.xlsx");
}
const modules = process.env.DIRECTOR_NODE_MODULES ||
  path.resolve(path.dirname(process.execPath), "../node_modules");
const require = createRequire(path.join(modules, "_director_runtime.cjs"));
const { Workbook, SpreadsheetFile } = await import(require.resolve("@oai/artifact-tool"));
const { title, rows } = JSON.parse(await readFile(input, "utf8"));
const headers = ["集", "场次", "场景", "剧情核心（分集节拍／本场）", "视觉重点", "必需镜头", "情绪重点", "服装",
  "场景氛围要求", "Key Prop", "特殊要求说明", "导演问题/说明"];
const widths = [54, 76, 218, 330, 330, 380, 260, 224, 245, 205, 260, 260];
if (typeof title !== "string" || !title.trim() || !Array.isArray(rows) || !rows.length) {
  throw new Error("请提供表名与逐场数据。");
}
for (const [index, row] of rows.entries()) {
  if (!Array.isArray(row) || row.length !== headers.length - 1 ||
      row.some((v, i) => typeof v !== "string" && !(i === 0 && Number.isFinite(v))) ||
      !row[2].trim() || !row[3].trim() || !row[4].trim() || !row[6].trim()) {
    throw new Error(`第 ${index + 1} 条场次需要十一个字段，场景、剧情、视觉和情绪重点必填。旧稿须在视觉重点后插入必需镜头。`);
  }
}

// Count explicit line breaks as well as wrapping. Excel CJK glyphs at 11 pt
// occupy about 15 pixels; allow extra room for font substitution and word wraps.
function wrappedLines(value, width) {
  return String(value).split(/\r?\n/).reduce((count, line) => {
    let pixels = 0;
    for (const ch of line) pixels += ch.codePointAt(0) <= 0x7f ? 9 : 16;
    return count + Math.max(1, Math.ceil(pixels / (width - 18)));
  }, 0);
}

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Scene by Scene");
const first = 3;
const last = first + rows.length - 1;
const all = sheet.getRange(`A1:L${last}`);
all.format.font = { name: "Arial", size: 11, color: "#242124" };
for (const [i, width] of widths.entries()) {
  sheet.getRangeByIndexes(0, i, 1, 1).format.columnWidthPx = width;
}
const banner = sheet.getRange("A1:L1");
banner.merge();
sheet.getRange("A1").values = [[title]];
banner.format.fill = "#6E1F36";
banner.format.font = { name: "Arial", bold: true, size: 18, color: "#FFFFFF" };
banner.format.verticalAlignment = "center";
banner.format.rowHeightPx = 46;
const header = sheet.getRange("A2:L2");
header.values = [headers];
header.format.fill = "#B54866";
header.format.font = { name: "Arial", bold: true, size: 11, color: "#FFFFFF" };
header.format.horizontalAlignment = "center";
header.format.verticalAlignment = "center";
header.format.rowHeightPx = 32;
const body = sheet.getRange(`A${first}:L${last}`);
body.format.verticalAlignment = "top";
body.format.wrapText = true;
body.format.borders = { preset: "all", style: "thin", color: "#D9CED2" };
const labels = sheet.getRange(`A${first}:B${last}`);
labels.setNumberFormat("@");
labels.format.horizontalAlignment = "center";
labels.format.fill = "#FBF4F2";
labels.format.font = { name: "Arial", bold: true, size: 11, color: "#6E1F36" };
sheet.getRange(`A${first}:K${last}`).values =
  rows.map(row => row.map(v => v === "" ? null : String(v)));
for (const [i, row] of rows.entries()) {
  const lines = Math.max(...row.map((v, j) => wrappedLines(v, widths[j])));
  const height = Math.max(56, lines * 24 + 16);
  if (height > 540) throw new Error(`第 ${i + 1} 场文字超出 Excel 单行可读高度，请精简或调宽列。`);
  sheet.getRangeByIndexes(i + first - 1, 0, 1, headers.length).format.rowHeightPx = height;
}
sheet.freezePanes.freezeRows(first - 1);

// Publish a complete file without replacing a prior deliverable on a retry.
const folder = path.dirname(path.resolve(output));
await mkdir(folder, { recursive: true });
const staging = await mkdtemp(path.join(folder, ".director-"));
try {
  const staged = path.join(staging, "output.xlsx");
  await (await SpreadsheetFile.exportXlsx(workbook)).save(staged);
  await link(staged, output);
} finally {
  await rm(staging, { recursive: true, force: true });
}
