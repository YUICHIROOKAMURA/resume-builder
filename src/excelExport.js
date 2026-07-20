import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

// exceljsでゼロから履歴書を構築（A4縦2ページ、JIS準拠レイアウト）
// テンプレート読み込みを行わないため、Excelでの破損警告が出ない

const THIN = { style: "thin", color: { argb: "FF000000" } };
const BOX = { top: THIN, left: THIN, bottom: THIN, right: THIN };
const DASH = { style: "dashed", color: { argb: "FF808080" } };
const PHOTO_BOX = { top: DASH, left: DASH, bottom: DASH, right: DASH };
const F = (opt = {}) => ({ name: "游ゴシック", size: 10.5, ...opt });

export function historyOverflow(historyRows) {
  return historyRows.length > 20;
}

export async function exportExcel({ basic, historyRows, quals, motive, wishLines, dateStr }) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("履歴書", {
    pageSetup: {
      paperSize: 9, orientation: "portrait",
      fitToPage: true, fitToWidth: 1, fitToHeight: 0,
      margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3 },
    },
  });

  // 列幅: A余白 / B項目名 / C-E内容 / F写真
  ws.columns = [
    { width: 2 },   // A
    { width: 12 },  // B 項目名/年
    { width: 7 },   // C 月
    { width: 22 },  // D
    { width: 22 },  // E
    { width: 17 },  // F 写真列
  ];

  const set = (addr, v, { font, align, border, fill } = {}) => {
    const c = ws.getCell(addr);
    c.value = v;
    c.font = F(font);
    if (align) c.alignment = align;
    if (border) c.border = border;
    if (fill) c.fill = fill;
  };
  const boxRange = (range, border = BOX) => {
    const [s, e] = range.split(":");
    const col = (a) => a.charCodeAt(0) - 64, row = (a) => parseInt(a.slice(1));
    for (let r = row(s); r <= row(e); r++)
      for (let cc = col(s[0]); cc <= col(e[0]); cc++)
        ws.getCell(r, cc).border = border;
  };

  let r = 1;
  // タイトル・日付
  set(`B${r}`, "履 歴 書", { font: { size: 16, bold: true } });
  set(`E${r}`, dateStr, { font: { size: 9 }, align: { horizontal: "right" } });
  ws.getRow(r).height = 24;
  r += 2; // r=3

  // 基本情報ブロック(行3〜13) + 写真枠(F3:F13)
  const info = [
    ["ふりがな", basic.kana],
    ["氏名", basic.name],
    ["生年月日", `${basic.birth}${basic.age ? `　（満 ${basic.age}歳）` : ""}`],
    ["性別", basic.gender || ""],
    ["ふりがな", basic.addrKana],
    ["現住所", `${basic.zip ? `〒${basic.zip}　` : ""}${basic.address}`],
    ["電話", basic.tel],
    ["E-mail", basic.email],
  ];
  const infoStart = r;
  info.forEach(([k, v], i) => {
    const row = infoStart + i;
    set(`B${row}`, k, { font: { size: 8 }, border: BOX, align: { vertical: "middle" } });
    ws.mergeCells(`C${row}:E${row}`);
    set(`C${row}`, v, {
      border: BOX,
      font: k === "氏名" ? { size: 14 } : {},
      align: { vertical: "middle", wrapText: true },
    });
    ws.getRow(row).height = k === "氏名" ? 26 : k === "現住所" ? 28 : 16.5;
  });
  // 写真枠: F3:F13相当 → 実寸調整はfitToWidth次第のため案内文のみ
  const photoEnd = infoStart + info.length - 1;
  ws.mergeCells(`F${infoStart}:F${photoEnd}`);
  set(`F${infoStart}`, "写真貼付欄\n\n縦 40mm\n横 30mm\n\n本人単身\n胸から上", {
    font: { size: 8, color: { argb: "FF808080" } },
    align: { horizontal: "center", vertical: "middle", wrapText: true },
  });
  for (let rr = infoStart; rr <= photoEnd; rr++) ws.getCell(`F${rr}`).border = PHOTO_BOX;
  r = photoEnd + 2;

  // 学歴・職歴
  const histHeader = r;
  set(`B${r}`, "年", { border: BOX, align: { horizontal: "center" }, font: { size: 9 } });
  set(`C${r}`, "月", { border: BOX, align: { horizontal: "center" }, font: { size: 9 } });
  ws.mergeCells(`D${r}:F${r}`);
  set(`D${r}`, "学歴・職歴（各別にまとめて書く）", { border: BOX, align: { horizontal: "center" }, font: { size: 9 } });
  r++;
  const histRows = [...historyRows];
  while (histRows.length < 18) histRows.push({ y: "", m: "", t: "" });
  histRows.slice(0, 20).forEach((h) => {
    set(`B${r}`, h.y ? Number(h.y) || h.y : "", { border: BOX, align: { horizontal: "center" } });
    set(`C${r}`, h.m ? Number(h.m) || h.m : "", { border: BOX, align: { horizontal: "center" } });
    ws.mergeCells(`D${r}:F${r}`);
    set(`D${r}`, h.t || "", {
      border: BOX,
      font: h.center ? { bold: true } : {},
      align: { horizontal: h.center ? "center" : h.right ? "right" : "left", vertical: "middle" },
    });
    ws.getRow(r).height = 17;
    r++;
  });

  // ---- 2ページ目への改ページ ----
  ws.getRow(r - 1).addPageBreak();
  r += 1;

  // 資格・免許
  set(`B${r}`, "年", { border: BOX, align: { horizontal: "center" }, font: { size: 9 } });
  set(`C${r}`, "月", { border: BOX, align: { horizontal: "center" }, font: { size: 9 } });
  ws.mergeCells(`D${r}:F${r}`);
  set(`D${r}`, "免許・資格", { border: BOX, align: { horizontal: "center" }, font: { size: 9 } });
  r++;
  const qualList = quals.filter((q) => q.text);
  while (qualList.length < 5) qualList.push({ y: "", m: "", text: "" });
  qualList.slice(0, 8).forEach((q) => {
    set(`B${r}`, q.y ? Number(q.y) || q.y : "", { border: BOX, align: { horizontal: "center" } });
    set(`C${r}`, q.m ? Number(q.m) || q.m : "", { border: BOX, align: { horizontal: "center" } });
    ws.mergeCells(`D${r}:F${r}`);
    set(`D${r}`, q.text || "", { border: BOX });
    ws.getRow(r).height = 17;
    r++;
  });
  r++;

  // 志望動機
  ws.mergeCells(`B${r}:F${r}`);
  set(`B${r}`, "志望の動機、特技、好きな学科、アピールポイントなど", { border: BOX, align: { horizontal: "center" }, font: { size: 9 } });
  r++;
  ws.mergeCells(`B${r}:F${r + 9}`);
  set(`B${r}`, motive || "", { align: { wrapText: true, vertical: "top" } });
  boxRange(`B${r}:F${r + 9}`);
  ws.getCell(`B${r}`).border = BOX;
  // 大ブロックの外枠のみ罫線
  for (let rr = r; rr <= r + 9; rr++)
    for (let cc = 2; cc <= 6; cc++)
      ws.getCell(rr, cc).border = {
        top: rr === r ? THIN : undefined, bottom: rr === r + 9 ? THIN : undefined,
        left: cc === 2 ? THIN : undefined, right: cc === 6 ? THIN : undefined,
      };
  r += 11;

  // 本人希望
  ws.mergeCells(`B${r}:F${r}`);
  set(`B${r}`, "本人希望記入欄（特に給料・職種・勤務時間・勤務地・その他についての希望などがあれば記入）", { border: BOX, align: { horizontal: "center" }, font: { size: 9 } });
  r++;
  ws.mergeCells(`B${r}:F${r + 6}`);
  set(`B${r}`, (wishLines || []).join("\n"), { align: { wrapText: true, vertical: "top" } });
  for (let rr = r; rr <= r + 6; rr++)
    for (let cc = 2; cc <= 6; cc++)
      ws.getCell(rr, cc).border = {
        top: rr === r ? THIN : undefined, bottom: rr === r + 6 ? THIN : undefined,
        left: cc === 2 ? THIN : undefined, right: cc === 6 ? THIN : undefined,
      };
  r += 8;
  set(`B${r}`, "※「性別」欄：記載は任意です。未記載とすることも可能です。", { font: { size: 8 } });

  ws.pageSetup.printArea = `A1:F${r}`;

  const out = await wb.xlsx.writeBuffer();
  saveAs(new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    `履歴書_${basic.name || "無題"}.xlsx`);
}
