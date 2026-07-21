import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { saveAs } from "file-saver";

// 事業所指定のWordフォーマット(public/rireki_template.docx)に差し込んで履歴書を出力
export async function exportRirekiWord({ basic, historyRows, quals, motive, wishLines, dateStr }) {
  const res = await fetch("/rireki_template.docx");
  const buf = await res.arrayBuffer();
  const zip = new PizZip(buf);
  const doc = new Docxtemplater(zip, { linebreaks: true, nullGetter: () => "" });

  const data = {
    dateStr,
    kana: basic.kana || "",
    name: basic.name || "",
    birth: `${basic.birth || ""}${basic.age ? `　（満 ${basic.age}歳）` : ""}`,
    gender: basic.gender || "",
    addrKana: basic.addrKana || "",
    zip: basic.zip || "",
    address: basic.address || "",
    tel: basic.tel || "",
    email: basic.email || "",
    motive: motive || "",
  };
  // 学歴・職歴: 1ページ目15行 + 2ページ目の続き7行
  for (let i = 1; i <= 15; i++) {
    const h = historyRows[i - 1] || { y: "", m: "", t: "" };
    data[`h${i}y`] = h.y || ""; data[`h${i}m`] = h.m || ""; data[`h${i}t`] = h.t || "";
  }
  for (let i = 1; i <= 7; i++) {
    const h = historyRows[15 + i - 1] || { y: "", m: "", t: "" };
    data[`c${i}y`] = h.y || ""; data[`c${i}m`] = h.m || ""; data[`c${i}t`] = h.t || "";
  }
  // 資格・免許: 6行
  const qualList = quals.filter((q) => q.text);
  for (let i = 1; i <= 6; i++) {
    const q = qualList[i - 1] || { y: "", m: "", text: "" };
    data[`g${i}y`] = q.y || ""; data[`g${i}m`] = q.m || ""; data[`g${i}t`] = q.text || "";
  }
  // 本人希望: 4行
  for (let i = 1; i <= 4; i++) data[`w${i}`] = (wishLines || [])[i - 1] || "";

  doc.render(data);
  const out = doc.getZip().generate({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
  saveAs(out, `履歴書_${basic.name || "無題"}.docx`);
}

export function rirekiOverflow(historyRows, quals) {
  return historyRows.length > 22 || quals.filter((q) => q.text).length > 6;
}
