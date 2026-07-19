import ExcelJS from "exceljs";
import { saveAs } from "file-saver";

// template.xlsx のセル座標に完全準拠して書き込む
// 左面: 基本情報 + 学歴・職歴(B38〜3行ごと、最大16ブロック)
// 右面: 学歴・職歴続き(L5〜) / 資格(L25〜) / 志望動機(L47) / 本人希望(L71〜)

export async function exportExcel({ basic, historyRows, quals, motive, wishLines, dateStr }) {
  const res = await fetch("/template.xlsx");
  const buf = await res.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  const ws = wb.getWorksheet("履歴書");

  const set = (addr, v, align) => {
    const c = ws.getCell(addr);
    c.value = v;
    if (align) c.alignment = { ...c.alignment, ...align };
  };

  // 日付・基本情報
  set("E3", `　${dateStr}`);
  set("C6", basic.kana);
  set("C9", basic.name);
  set("B14", `${basic.birth}${basic.age ? `　（満 ${basic.age}歳）` : ""}　`);
  if (basic.gender) set("F15", basic.gender);
  set("C16", basic.addrKana);
  set("C19", basic.zip ? `〒${basic.zip}` : "");
  set("C21", basic.address);
  set("I16", basic.tel);
  set("H21", basic.email);

  // 学歴・職歴: 左面 B38〜B83(3行ごと16ブロック) → 溢れたら右面 L5〜L40(12ブロック)
  const leftRows = Array.from({ length: 16 }, (_, i) => 38 + i * 3);
  const rightRows = Array.from({ length: 12 }, (_, i) => 5 + i * 3);
  historyRows.forEach((r, i) => {
    let yCol, mCol, tCol, row;
    if (i < leftRows.length) { row = leftRows[i]; yCol = "B"; mCol = "C"; tCol = "D"; }
    else if (i - leftRows.length < rightRows.length) { row = rightRows[i - leftRows.length]; yCol = "L"; mCol = "M"; tCol = "N"; }
    else return; // 収まらない分は切り捨て(UI側で警告)
    if (r.y) set(`${yCol}${row}`, Number(r.y) || r.y, { horizontal: "center" });
    if (r.m) set(`${mCol}${row}`, Number(r.m) || r.m, { horizontal: "center" });
    set(`${tCol}${row}`, r.t, r.center ? { horizontal: "center" } : r.right ? { horizontal: "right" } : undefined);
  });

  // 資格・免許: L25〜3行ごと6ブロック
  quals.filter((q) => q.text).slice(0, 6).forEach((q, i) => {
    const row = 25 + i * 3;
    if (q.y) set(`L${row}`, Number(q.y) || q.y, { horizontal: "center" });
    if (q.m) set(`M${row}`, Number(q.m) || q.m, { horizontal: "center" });
    set(`N${row}`, q.text);
  });

  // 志望動機(L47:R66 大ブロック)
  set("L47", motive, { wrapText: true, vertical: "top", horizontal: "left" });

  // 本人希望(L71,74,77,80 行ブロック)
  wishLines.slice(0, 4).forEach((line, i) => set(`L${71 + i * 3}`, line));

  const out = await wb.xlsx.writeBuffer();
  saveAs(new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    `履歴書_${basic.name || "無題"}.xlsx`);
}

export function historyOverflow(historyRows) {
  return historyRows.length > 28; // 16 + 12
}
