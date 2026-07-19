import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, PageBreak, VerticalAlign,
} from "docx";
import { saveAs } from "file-saver";

const FONT = "游ゴシック";
const bAll = { style: BorderStyle.SINGLE, size: 6, color: "000000" };
const borders = { top: bAll, bottom: bAll, left: bAll, right: bAll };
const noB = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noB, bottom: noB, left: noB, right: noB };

const run = (t, opt = {}) => new TextRun({ text: t, font: FONT, size: 21, ...opt });
const para = (t, opt = {}, runOpt = {}) =>
  new Paragraph({ children: [run(t, runOpt)], ...opt });

const cell = (text, { w, bold, size, align, gray, vAlign, colSpan, rowSpan, lines } = {}) =>
  new TableCell({
    borders,
    verticalAlign: vAlign || VerticalAlign.CENTER,
    columnSpan: colSpan,
    rowSpan,
    width: w ? { size: w, type: WidthType.DXA } : undefined,
    children: (lines || [text]).map((t) =>
      new Paragraph({
        alignment: align,
        children: [run(t, { bold, size: size || 21, color: gray ? "888888" : undefined })],
      })
    ),
  });

export async function exportWord({ basic, historyRows, quals, motive, wish, jobs, gaps, gapTexts, dateStr }) {
  const W = 9600; // 全幅(DXA)
  const yrW = 800, moW = 550;

  // ---- 履歴書: 基本情報 ----
  const infoTable = new Table({
    width: { size: W, type: WidthType.DXA },
    rows: [
      new TableRow({ children: [
        cell("ふりがな", { w: 1300, size: 16 }), cell(basic.kana, { colSpan: 2 }),
        cell("写真貼付欄", { w: 1900, rowSpan: 4, align: AlignmentType.CENTER, gray: true, size: 16,
          lines: ["写真貼付欄", "", "縦40mm×横30mm", "本人単身・胸から上"] }),
      ]}),
      new TableRow({ children: [ cell("氏名", { size: 16 }), cell(basic.name, { colSpan: 2, size: 28 }) ]}),
      new TableRow({ children: [ cell("生年月日", { size: 16 }), cell(`${basic.birth}${basic.age ? `　（満 ${basic.age}歳）` : ""}`, { colSpan: 2 }) ]}),
      new TableRow({ children: [ cell("性別", { size: 16 }), cell(basic.gender || "", { colSpan: 2 }) ]}),
      new TableRow({ children: [ cell("ふりがな", { size: 16 }), cell(basic.addrKana, { colSpan: 3 }) ]}),
      new TableRow({ children: [ cell("現住所", { size: 16 }), cell(`${basic.zip ? `〒${basic.zip}　` : ""}${basic.address}`, { colSpan: 3 }) ]}),
      new TableRow({ children: [
        cell("電話", { size: 16 }), cell(basic.tel),
        cell("E-mail", { size: 16, w: 900 }), cell(basic.email),
      ]}),
    ],
  });

  // ---- 学歴・職歴 ----
  const histRows = [...historyRows];
  while (histRows.length < 14) histRows.push({ y: "", m: "", t: "" });
  const histTable = new Table({
    width: { size: W, type: WidthType.DXA },
    rows: [
      new TableRow({ children: [
        cell("年", { w: yrW, align: AlignmentType.CENTER, size: 18 }),
        cell("月", { w: moW, align: AlignmentType.CENTER, size: 18 }),
        cell("学歴・職歴（各別にまとめて書く）", { align: AlignmentType.CENTER, size: 18 }),
      ]}),
      ...histRows.map((r) => new TableRow({ children: [
        cell(String(r.y || ""), { w: yrW, align: AlignmentType.CENTER }),
        cell(String(r.m || ""), { w: moW, align: AlignmentType.CENTER }),
        cell(r.t || "", { bold: r.center, align: r.center ? AlignmentType.CENTER : r.right ? AlignmentType.RIGHT : undefined }),
      ]})),
    ],
  });

  // ---- 資格 ----
  const qualList = quals.filter((q) => q.text);
  while (qualList.length < 4) qualList.push({ y: "", m: "", text: "" });
  const qualTable = new Table({
    width: { size: W, type: WidthType.DXA },
    rows: [
      new TableRow({ children: [
        cell("年", { w: yrW, align: AlignmentType.CENTER, size: 18 }),
        cell("月", { w: moW, align: AlignmentType.CENTER, size: 18 }),
        cell("免許・資格", { align: AlignmentType.CENTER, size: 18 }),
      ]}),
      ...qualList.map((q) => new TableRow({ children: [
        cell(String(q.y || ""), { w: yrW, align: AlignmentType.CENTER }),
        cell(String(q.m || ""), { w: moW, align: AlignmentType.CENTER }),
        cell(q.text || ""),
      ]})),
    ],
  });

  const bigBox = (title, body, minLines) => new Table({
    width: { size: W, type: WidthType.DXA },
    rows: [
      new TableRow({ children: [cell(title, { align: AlignmentType.CENTER, size: 18 })] }),
      new TableRow({ children: [
        new TableCell({
          borders, verticalAlign: VerticalAlign.TOP,
          children: [...body.split("\n"), ...Array(Math.max(0, minLines - body.split("\n").length)).fill("")]
            .map((l) => para(l)),
        }),
      ]}),
    ],
  });

  // ---- 職務経歴書 ----
  const cvChildren = [
    new Paragraph({ alignment: AlignmentType.CENTER, children: [run("職 務 経 歴 書", { bold: true, size: 32 })] }),
    new Paragraph({ alignment: AlignmentType.RIGHT, children: [run(dateStr)] }),
    new Paragraph({ alignment: AlignmentType.RIGHT, children: [run(`氏名: ${basic.name}`)] }),
    para(""),
  ];
  jobs.filter((j) => j.company).forEach((j) => {
    cvChildren.push(new Paragraph({
      children: [run(`${j.company}（${j.from} 〜 ${j.to || "現在"}）`, { bold: true })],
      border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: "2F5D8A" } },
      spacing: { before: 200 },
    }));
    if (j.role) cvChildren.push(para(j.role, {}, { color: "555555", size: 18 }));
    (j.detail || "").split("\n").forEach((l) => cvChildren.push(para(l)));
  });
  const gapEntries = gaps.filter((g) => gapTexts[g.key]);
  if (gapEntries.length) {
    gapEntries.forEach((g) => {
      cvChildren.push(new Paragraph({
        children: [run(`離職期間について（${g.from}〜${g.to}）`, { bold: true })],
        border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: "2F5D8A" } },
        spacing: { before: 200 },
      }));
      gapTexts[g.key].split("\n").forEach((l) => cvChildren.push(para(l)));
    });
  }

  const doc = new Document({
    sections: [{
      properties: { page: { margin: { top: 720, bottom: 720, left: 850, right: 850 } } },
      children: [
        new Table({ width: { size: W, type: WidthType.DXA }, rows: [new TableRow({ children: [
          new TableCell({ borders: noBorders, children: [new Paragraph({ children: [run("履 歴 書", { bold: true, size: 32 })] })] }),
          new TableCell({ borders: noBorders, children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [run(dateStr, { size: 18 })] })] }),
        ]})]}),
        para(""),
        infoTable,
        para(""),
        histTable,
        para(""),
        qualTable,
        para(""),
        bigBox("志望の動機、特技、好きな学科、アピールポイントなど", motive || "", 8),
        para(""),
        bigBox("本人希望記入欄（特に給料・職種・勤務時間・勤務地・その他についての希望などがあれば記入）", wish || "", 5),
        para("※「性別」欄：記載は任意です。未記載とすることも可能です。", {}, { size: 16 }),
        new Paragraph({ children: [new PageBreak()] }),
        ...cvChildren,
      ],
    }],
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, `履歴書_${basic.name || "無題"}.docx`);
}
