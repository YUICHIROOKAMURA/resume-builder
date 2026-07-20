import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

// A4縦2ページに余白なく収まるよう、倍率100%固定で列幅・行高を実寸から逆算して構築
// 用紙: A4 (210x297mm) / 余白: 左右10mm・上下12mm
// 印刷可能幅 190mm ≒ 列幅合計102文字幅 / 印刷可能高さ 273mm ≒ 774pt per page

const THIN = { style: "thin", color: { argb: "FF000000" } };
const BOX = { top: THIN, left: THIN, bottom: THIN, right: THIN };
const DASH = { style: "dashed", color: { argb: "FF808080" } };
const PHOTO_BOX = { top: DASH, left: DASH, bottom: DASH, right: DASH };
const F = (opt = {}) => ({ name: "游ゴシック", size: 10.5, ...opt });

const PAGE_PT = 774; // 1ページあたりの高さ予算(pt)

export function historyOverflow(historyRows) {
  return historyRows.length > 22;
}

export async function exportExcel({ basic, historyRows, quals, motive, wishLines, dateStr }) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("履歴書", {
    pageSetup: {
      paperSize: 9, orientation: "portrait", scale: 100,
      margins: { left: 0.394, right: 0.394, top: 0.472, bottom: 0.472, header: 0.2, footer: 0.2 },
    },
  });

  // 列幅合計 ≒ 102文字幅(190mm)
  ws.columns = [
    { width: 2 },   // A 余白
    { width: 10 },  // B 項目名/年
    { width: 7 },   // C 月
    { width: 28 },  // D
    { width: 26 },  // E
    { width: 2 },   // F 隙間
    { width: 18 },  // G 写真列
  ];

  const set = (addr, v, { font, align, border } = {}) => {
    const c = ws.getCell(addr);
    c.value = v;
    c.font = F(font);
    if (align) c.alignment = align;
    if (border) c.border = border;
  };
  const outerBox = (r1, r2) => {
    for (let rr = r1; rr <= r2; rr++)
      for (let cc = 2; cc <= 7; cc++)
        ws.getCell(rr, cc).border = {
          top: rr === r1 ? THIN : undefined, bottom: rr === r2 ? THIN : undefined,
          left: cc === 2 ? THIN : undefined, right: cc === 7 ? THIN : undefined,
        };
  };

  let r = 1;
  let used = 0; // 1ページ目の高さ消費(pt)
  const H = (row, h) => { ws.getRow(row).height = h; used += h; };

  // ---- 1ページ目 ----
  set(`B${r}`, "履 歴 書", { font: { size: 16, bold: true } });
  ws.mergeCells(`E${r}:G${r}`);
  set(`E${r}`, dateStr, { font: { size: 9 }, align: { horizontal: "right", vertical: "middle" } });
  H(r, 28); r++;
  H(r, 8); r++; // 余白行

  // 基本情報(写真枠を右に)
  const info = [
    ["ふりがな", basic.kana, 18],
    ["氏名", basic.name, 34],
    ["生年月日", `${basic.birth}${basic.age ? `　（満 ${basic.age}歳）` : ""}`, 22],
    ["性別", basic.gender || "", 20],
    ["ふりがな", basic.addrKana, 18],
    ["現住所", `${basic.zip ? `〒${basic.zip}　` : ""}${basic.address}`, 36],
    ["電話", basic.tel, 20],
    ["E-mail", basic.email, 20],
  ];
  const infoStart = r;
  info.forEach(([k, v, h]) => {
    set(`B${r}`, k, { font: { size: 8 }, border: BOX, align: { vertical: "middle" } });
    ws.mergeCells(`C${r}:E${r}`);
    set(`C${r}`, v, {
      border: BOX,
      font: k === "氏名" ? { size: 14 } : {},
      align: { vertical: "middle", wrapText: true },
    });
    H(r, h); r++;
  });
  // 写真枠(基本情報の右にF列の隙間を挟んでG列へ、上下も1行ずつ空ける)
  const photoTop = infoStart + 1, photoBottom = r - 2;
  ws.mergeCells(`G${photoTop}:G${photoBottom}`);
  set(`G${photoTop}`, "写真貼付欄\n\n縦 40mm × 横 30mm\n本人単身・胸から上", {
    font: { size: 8, color: { argb: "FF808080" } },
    align: { horizontal: "center", vertical: "middle", wrapText: true },
  });
  for (let rr = photoTop; rr <= photoBottom; rr++) ws.getCell(`G${rr}`).border = PHOTO_BOX;
  H(r, 8); r++; // 余白行

  // 学歴・職歴ヘッダ
  set(`B${r}`, "年", { border: BOX, align: { horizontal: "center", vertical: "middle" }, font: { size: 9 } });
  set(`C${r}`, "月", { border: BOX, align: { horizontal: "center", vertical: "middle" }, font: { size: 9 } });
  ws.mergeCells(`D${r}:G${r}`);
  set(`D${r}`, "学歴・職歴（各別にまとめて書く）", { border: BOX, align: { horizontal: "center", vertical: "middle" }, font: { size: 9 } });
  H(r, 18); r++;

  // 残り高さを学歴・職歴の行で敷き詰める(最低18行)
  const histRows = [...historyRows];
  const histCount = Math.max(18, Math.min(histRows.length, 22));
  while (histRows.length < histCount) histRows.push({ y: "", m: "", t: "" });
  const remain1 = PAGE_PT - used;
  const histH = Math.floor((remain1 / histCount) * 10) / 10; // 行高を均等割り
  histRows.slice(0, histCount).forEach((h) => {
    set(`B${r}`, h.y ? Number(h.y) || h.y : "", { border: BOX, align: { horizontal: "center", vertical: "middle" } });
    set(`C${r}`, h.m ? Number(h.m) || h.m : "", { border: BOX, align: { horizontal: "center", vertical: "middle" } });
    ws.mergeCells(`D${r}:G${r}`);
    set(`D${r}`, h.t || "", {
      border: BOX,
      font: h.center ? { bold: true } : {},
      align: { horizontal: h.center ? "center" : h.right ? "right" : "left", vertical: "middle" },
    });
    ws.getRow(r).height = histH;
    r++;
  });

  // ---- 2ページ目 ----
  ws.getRow(r - 1).addPageBreak();
  let used2 = 0;
  const H2 = (row, h) => { ws.getRow(row).height = h; used2 += h; };

  H2(r, 12); r++; // 上余白行

  // 免許・資格
  set(`B${r}`, "年", { border: BOX, align: { horizontal: "center", vertical: "middle" }, font: { size: 9 } });
  set(`C${r}`, "月", { border: BOX, align: { horizontal: "center", vertical: "middle" }, font: { size: 9 } });
  ws.mergeCells(`D${r}:G${r}`);
  set(`D${r}`, "免許・資格", { border: BOX, align: { horizontal: "center", vertical: "middle" }, font: { size: 9 } });
  H2(r, 18); r++;
  const qualList = quals.filter((q) => q.text);
  while (qualList.length < 6) qualList.push({ y: "", m: "", text: "" });
  qualList.slice(0, 8).forEach((q) => {
    set(`B${r}`, q.y ? Number(q.y) || q.y : "", { border: BOX, align: { horizontal: "center", vertical: "middle" } });
    set(`C${r}`, q.m ? Number(q.m) || q.m : "", { border: BOX, align: { horizontal: "center", vertical: "middle" } });
    ws.mergeCells(`D${r}:G${r}`);
    set(`D${r}`, q.text || "", { border: BOX, align: { vertical: "middle" } });
    H2(r, 26); r++;
  });
  H2(r, 10); r++; // 余白行

  // 注記と下部固定分を先に確保して、残りを志望動機/本人希望に配分
  const noteH = 16, gapH = 10;
  const remain2 = PAGE_PT - used2 - noteH - gapH - 18 - 18; // 見出し2つ分
  const motiveH = Math.floor(remain2 * 0.55);
  const wishH = remain2 - motiveH;

  // 志望動機
  ws.mergeCells(`B${r}:G${r}`);
  set(`B${r}`, "志望の動機、特技、好きな学科、アピールポイントなど", { border: BOX, align: { horizontal: "center", vertical: "middle" }, font: { size: 9 } });
  H2(r, 18); r++;
  const mRows = 10;
  ws.mergeCells(`B${r}:G${r + mRows - 1}`);
  set(`B${r}`, motive || "", { align: { wrapText: true, vertical: "top" } });
  outerBox(r, r + mRows - 1);
  for (let i = 0; i < mRows; i++) { ws.getRow(r + i).height = motiveH / mRows; }
  used2 += motiveH; r += mRows;
  H2(r, gapH); r++;

  // 本人希望
  ws.mergeCells(`B${r}:G${r}`);
  set(`B${r}`, "本人希望記入欄（特に給料・職種・勤務時間・勤務地・その他についての希望などがあれば記入）", { border: BOX, align: { horizontal: "center", vertical: "middle" }, font: { size: 9 } });
  H2(r, 18); r++;
  const wRows = 7;
  ws.mergeCells(`B${r}:G${r + wRows - 1}`);
  set(`B${r}`, (wishLines || []).join("\n"), { align: { wrapText: true, vertical: "top" } });
  outerBox(r, r + wRows - 1);
  for (let i = 0; i < wRows; i++) { ws.getRow(r + i).height = wishH / wRows; }
  r += wRows;

  // 注記
  set(`B${r}`, "※「性別」欄：記載は任意です。未記載とすることも可能です。", { font: { size: 8 } });
  ws.getRow(r).height = noteH;

  ws.pageSetup.printArea = `A1:G${r}`;

  const out = await wb.xlsx.writeBuffer();
  saveAs(new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    `履歴書_${basic.name || "無題"}.xlsx`);
}
