import {
  Document, Packer, Paragraph, TextRun, AlignmentType, BorderStyle,
} from "docx";
import { saveAs } from "file-saver";

const FONT = "游ゴシック";
const run = (t, opt = {}) => new TextRun({ text: t, font: FONT, size: 21, ...opt });
const para = (t, opt = {}, runOpt = {}) => new Paragraph({ children: [run(t, runOpt)], ...opt });

// 職務経歴書のみ（履歴書はExcel出力を使用）
export async function exportWord({ basic, jobs, gaps, gapTexts, dateStr }) {
  const children = [
    new Paragraph({ alignment: AlignmentType.CENTER, children: [run("職 務 経 歴 書", { bold: true, size: 32 })] }),
    new Paragraph({ alignment: AlignmentType.RIGHT, children: [run(dateStr)] }),
    new Paragraph({ alignment: AlignmentType.RIGHT, children: [run(`氏名: ${basic.name}`)] }),
    para(""),
  ];
  jobs.filter((j) => j.company).forEach((j) => {
    children.push(new Paragraph({
      children: [run(`${j.company}（${j.from} 〜 ${j.to || "現在"}）`, { bold: true })],
      border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: "2F5D8A" } },
      spacing: { before: 240, after: 60 },
    }));
    if (j.role) children.push(para(j.role, {}, { color: "555555", size: 18 }));
    (j.detail || "").split("\n").forEach((l) => children.push(para(l)));
  });
  const gapEntries = gaps.filter((g) => gapTexts[g.key]);
  gapEntries.forEach((g) => {
    children.push(new Paragraph({
      children: [run(`離職期間について（${g.from}〜${g.to}）`, { bold: true })],
      border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: "2F5D8A" } },
      spacing: { before: 240, after: 60 },
    }));
    gapTexts[g.key].split("\n").forEach((l) => children.push(para(l)));
  });

  const doc = new Document({
    sections: [{
      properties: { page: { margin: { top: 1000, bottom: 1000, left: 1100, right: 1100 } } },
      children,
    }],
  });
  const blob = await Packer.toBlob(doc);
  saveAs(blob, `職務経歴書_${basic.name || "無題"}.docx`);
}
