import React, { useState, useMemo, useRef } from "react";
import { exportRirekiWord, rirekiOverflow } from "./rirekiWordExport.js";
import { exportWord } from "./wordExport.js";

const GAP_TEMPLATES = {
  療養: {
    開示: "体調を崩し療養に専念しておりましたが、現在は回復し、就労移行支援事業所にて{months}安定して通所訓練を継続しております。生活リズムと体調管理の方法を身につけ、就労への準備が整いました。",
    非開示: "一身上の都合により離職し、心身のリフレッシュと今後のキャリアの見直しを行っておりました。現在は再就職に向けた準備を整えております。",
  },
  通所訓練: {
    開示: "就労移行支援事業所に通所し、ビジネススキルの訓練と自己理解を深めてまいりました。週{days}日の通所を{months}継続しており、安定した就労が可能な状態です。",
    非開示: "職業訓練機関にてビジネススキルの習得に取り組み、再就職に向けた準備を行っておりました。",
  },
  資格学習: {
    開示: "資格取得に向けた学習に取り組みつつ、就労移行支援事業所にて就労準備を進めておりました。",
    非開示: "資格取得に向けた学習期間として、計画的にスキルアップに取り組んでおりました。",
  },
  家庭事情: {
    開示: "家庭の事情により離職いたしましたが、現在は状況が落ち着き、就労移行支援事業所での訓練を経て就労できる環境が整っております。",
    非開示: "家庭の事情により離職いたしましたが、現在は状況が解消し、フルタイムでの勤務が可能です。",
  },
};

function formatBirth(raw) {
  if (!raw) return "";
  const s = raw.trim();
  // すでに「年」を含む場合はそのまま(末尾に「生」がなければ補う)
  if (/年/.test(s)) return /生\s*$/.test(s) ? s : s + "生";
  // 8桁数字 (19831014) / 区切り付き (1983/10/14, 1983-10-14) に対応
  const m = /^(\d{4})[\/\-年]?(\d{1,2})[\/\-月]?(\d{1,2})日?/.exec(s.replace(/\s/g, ""));
  if (m) return `${m[1]}年${parseInt(m[2])}月${parseInt(m[3])}日生`;
  return s;
}

const emptyJob = () => ({ id: Date.now() + Math.random(), company: "", role: "", from: "", to: "", detail: "" });
const emptyRow = () => ({ id: Date.now() + Math.random(), y: "", m: "", text: "" });

function parseYM(s) {
  const m = /^(\d{4})[-/年]?(\d{1,2})/.exec(s || "");
  return m ? parseInt(m[1]) * 12 + parseInt(m[2]) - 1 : null;
}
function splitYM(s) {
  const m = /^(\d{4})[-/年]?(\d{1,2})/.exec(s || "");
  return m ? [m[1], String(parseInt(m[2]))] : ["", ""];
}
function detectGaps(jobs) {
  const sorted = jobs.map((j) => ({ ...j, f: parseYM(j.from), t: parseYM(j.to) }))
    .filter((j) => j.f != null).sort((a, b) => a.f - b.f);
  const gaps = [];
  for (let i = 0; i < sorted.length - 1; i++) {
    const end = sorted[i].t, next = sorted[i + 1].f;
    if (end != null && next != null && next - end >= 4) {
      gaps.push({
        from: `${Math.floor((end + 1) / 12)}年${((end + 1) % 12) + 1}月`,
        to: `${Math.floor((next - 1) / 12)}年${((next - 1) % 12) + 1}月`,
        months: next - end - 1, key: `${end}-${next}`,
      });
    }
  }
  return gaps;
}

const S = {
  page: { fontFamily: "'Hiragino Sans','Yu Gothic',sans-serif", background: "#eef1f5", minHeight: "100vh", color: "#1e2a38" },
  wrap: { maxWidth: 960, margin: "0 auto", padding: 16 },
  card: { background: "#fff", borderRadius: 10, padding: 20, marginBottom: 16, boxShadow: "0 1px 3px rgba(30,42,56,.08)" },
  h2: { fontSize: 15, fontWeight: 700, borderLeft: "4px solid #2f5d8a", paddingLeft: 10, margin: "0 0 14px" },
  input: { width: "100%", boxSizing: "border-box", padding: "8px 10px", border: "1px solid #c7d0da", borderRadius: 6, fontSize: 14, marginBottom: 8 },
  label: { fontSize: 12, color: "#5a6b7d", display: "block", marginBottom: 3 },
  btn: { background: "#2f5d8a", color: "#fff", border: "none", borderRadius: 6, padding: "9px 16px", fontSize: 14, cursor: "pointer" },
  btnGreen: { background: "#1e7145", color: "#fff", border: "none", borderRadius: 6, padding: "9px 16px", fontSize: 14, cursor: "pointer" },
  btnBlue: { background: "#2b579a", color: "#fff", border: "none", borderRadius: 6, padding: "9px 16px", fontSize: 14, cursor: "pointer" },
  btnGhost: { background: "#fff", color: "#2f5d8a", border: "1px solid #2f5d8a", borderRadius: 6, padding: "8px 14px", fontSize: 13, cursor: "pointer" },
  tab: (on) => ({ flex: 1, padding: "11px 0", textAlign: "center", cursor: "pointer", fontSize: 13, fontWeight: on ? 700 : 400, background: on ? "#2f5d8a" : "#fff", color: on ? "#fff" : "#2f5d8a", border: "1px solid #2f5d8a" }),
  gapBox: { background: "#fff8ec", border: "1px solid #e8c98a", borderRadius: 8, padding: 14, marginTop: 10 },
};

export default function App() {
  const [tab, setTab] = useState("input");
  const [disclose, setDisclose] = useState("開示");
  const [basic, setBasic] = useState({ name: "", kana: "", birth: "", age: "", gender: "", zip: "", addrKana: "", address: "", tel: "", email: "" });
  const [edus, setEdus] = useState([emptyRow()]);
  const [jobs, setJobs] = useState([emptyJob()]);
  const [quals, setQuals] = useState([emptyRow()]);
  const [motive, setMotive] = useState("");
  const [summary, setSummary] = useState("");
  const [skills, setSkills] = useState("");
  const [selfPR, setSelfPR] = useState("");
  const [wish, setWish] = useState("貴社規定に従います。");
  const [gapTexts, setGapTexts] = useState({});
  const [months, setMonths] = useState("6ヶ月");
  const [days, setDays] = useState("5");
  const [busy, setBusy] = useState("");
  const fileRef = useRef(null);

  const gaps = useMemo(() => detectGaps(jobs), [jobs]);
  const today = new Date();
  const dateStr = `${today.getFullYear()}年${today.getMonth() + 1}月${today.getDate()}日現在`;

  const setJob = (id, k, v) => setJobs(jobs.map((j) => (j.id === id ? { ...j, [k]: v } : j)));
  const setRow = (list, setList) => (id, k, v) => setList(list.map((r) => (r.id === id ? { ...r, [k]: v } : r)));
  const setEdu = setRow(edus, setEdus), setQual = setRow(quals, setQuals);

  const applyTemplate = (gapKey, reason) => {
    const t = GAP_TEMPLATES[reason][disclose].replace("{months}", months).replace("{days}", days);
    setGapTexts({ ...gapTexts, [gapKey]: t });
  };

  const historyRows = useMemo(() => {
    const rows = [{ y: "", m: "", t: "学　　歴", center: true }];
    edus.filter((e) => e.text).forEach((e) => rows.push({ y: e.y, m: e.m, t: e.text }));
    rows.push({ y: "", m: "", t: "職　　歴", center: true });
    jobs.filter((j) => j.company).forEach((j) => {
      const [fy, fm] = splitYM(j.from);
      rows.push({ y: fy, m: fm, t: `${j.company} 入社` });
      if (j.to) {
        const [ty, tm] = splitYM(j.to);
        rows.push({ y: ty, m: tm, t: `${j.company} 一身上の都合により退社` });
      }
    });
    if (jobs.some((j) => j.company && !j.to)) rows.push({ y: "", m: "", t: "現在に至る", right: true });
    rows.push({ y: "", m: "", t: "以上", right: true });
    return rows;
  }, [edus, jobs]);

  const payload = () => ({
    basic: { ...basic, birth: formatBirth(basic.birth) }, historyRows, quals, motive,
    wish, wishLines: (wish || "").split("\n").filter(Boolean),
    jobs, gaps, gapTexts, dateStr,
    summary, skills, selfPR,
  });

  const doRireki = async () => {
    if (rirekiOverflow(historyRows, quals)) alert("学歴・職歴または資格が多いため、フォーマットに収まらない行は省略されます。");
    setBusy("rireki");
    try { await exportRirekiWord(payload()); } catch (e) { alert("履歴書出力に失敗しました: " + e.message); }
    setBusy("");
  };
  const doWord = async () => {
    setBusy("word");
    try { await exportWord(payload()); } catch (e) { alert("Word出力に失敗しました: " + e.message); }
    setBusy("");
  };

  const exportJSON = () => {
    const blob = new Blob([JSON.stringify({ basic, edus, jobs, quals, motive, wish, gapTexts, summary, skills, selfPR }, null, 2)], { type: "application/json" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = "resume-data.json"; a.click();
  };
  const importJSON = (e) => {
    const f = e.target.files[0];
    if (!f) return;
    const r = new FileReader();
    r.onload = () => {
      try {
        const d = JSON.parse(r.result);
        setBasic(d.basic || basic); setEdus(d.edus || edus); setJobs(d.jobs || jobs);
        setQuals(d.quals || quals); setMotive(d.motive || ""); setWish(d.wish || ""); setGapTexts(d.gapTexts || {});
        setSummary(d.summary || ""); setSkills(d.skills || ""); setSelfPR(d.selfPR || "");
      } catch { alert("読み込めませんでした。ファイルを確認してください。"); }
    };
    r.readAsText(f);
  };

  const rowEditor = (list, setter, removeSetter, ph) => list.map((r) => (
    <div key={r.id} style={{ display: "grid", gridTemplateColumns: "90px 60px 1fr 40px", gap: 6, marginBottom: 4 }}>
      <input style={S.input} placeholder="2011" value={r.y} onChange={(ev) => setter(r.id, "y", ev.target.value)} />
      <input style={S.input} placeholder="4" value={r.m} onChange={(ev) => setter(r.id, "m", ev.target.value)} />
      <input style={S.input} placeholder={ph} value={r.text} onChange={(ev) => setter(r.id, "text", ev.target.value)} />
      <button style={{ ...S.btnGhost, padding: "4px 8px" }} onClick={() => removeSetter(list.filter((x) => x.id !== r.id))}>✕</button>
    </div>
  ));

  return (
    <div style={S.page}>
      <div style={S.wrap}>
        <div style={{ padding: "18px 0 12px" }}>
          <div style={{ fontSize: 20, fontWeight: 800, color: "#2f5d8a" }}>履歴書・職務経歴書ビルダー</div>
          <div style={{ fontSize: 12, color: "#5a6b7d", marginTop: 4 }}>データはブラウザ内のみで処理されます。作業後は「JSONで保存」で保管してください。</div>
        </div>

        <div style={{ display: "flex", marginBottom: 16, borderRadius: 8, overflow: "hidden" }}>
          <div style={S.tab(tab === "input")} onClick={() => setTab("input")}>入力</div>
          <div style={S.tab(tab === "gap")} onClick={() => setTab("gap")}>空白期間 {gaps.length > 0 && `(${gaps.length})`}</div>
          <div style={S.tab(tab === "export")} onClick={() => setTab("export")}>出力</div>
        </div>

        {tab === "input" && (
          <div>
            <div style={S.card}>
              <h2 style={S.h2}>基本情報</h2>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 14px" }}>
                {[["kana", "ふりがな"], ["name", "氏名"], ["birth", "生年月日（例: 1992年4月1日 / 19920401 どちらも可）"], ["age", "満年齢"], ["gender", "性別（任意・空欄可）"], ["tel", "電話番号"], ["zip", "郵便番号（例: 160-0000）"], ["email", "E-mail"], ["addrKana", "住所ふりがな"], ["address", "住所"]].map(([k, l]) => (
                  <div key={k}><label style={S.label}>{l}</label><input style={S.input} value={basic[k]} onChange={(e) => setBasic({ ...basic, [k]: e.target.value })} /></div>
                ))}
              </div>
            </div>

            <div style={S.card}>
              <h2 style={S.h2}>学歴</h2>
              {rowEditor(edus, setEdu, setEdus, "〇〇高等学校 入学")}
              <button style={S.btnGhost} onClick={() => setEdus([...edus, emptyRow()])}>＋ 行を追加</button>
            </div>

            <div style={S.card}>
              <h2 style={S.h2}>職歴</h2>
              <div style={{ fontSize: 12, color: "#5a6b7d", marginBottom: 8 }}>入社・退社の年月から履歴書の職歴欄と空白期間を自動生成します</div>
              {jobs.map((j) => (
                <div key={j.id} style={{ border: "1px solid #dde4ec", borderRadius: 8, padding: 12, marginBottom: 10 }}>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "0 12px" }}>
                    <div><label style={S.label}>会社・組織名</label><input style={S.input} value={j.company} onChange={(e) => setJob(j.id, "company", e.target.value)} /></div>
                    <div><label style={S.label}>職種・役割</label><input style={S.input} value={j.role} onChange={(e) => setJob(j.id, "role", e.target.value)} /></div>
                    <div><label style={S.label}>入社（例: 2018/04）</label><input style={S.input} value={j.from} onChange={(e) => setJob(j.id, "from", e.target.value)} placeholder="2018/04" /></div>
                    <div><label style={S.label}>退社（在職中は空欄）</label><input style={S.input} value={j.to} onChange={(e) => setJob(j.id, "to", e.target.value)} placeholder="2021/03" /></div>
                  </div>
                  <label style={S.label}>業務内容（職務経歴書用）</label>
                  <textarea style={{ ...S.input, minHeight: 56, marginBottom: 4 }} value={j.detail} onChange={(e) => setJob(j.id, "detail", e.target.value)} />
                  {jobs.length > 1 && <button style={{ ...S.btnGhost, borderColor: "#b0433c", color: "#b0433c" }} onClick={() => setJobs(jobs.filter((x) => x.id !== j.id))}>この職歴を削除</button>}
                </div>
              ))}
              <button style={S.btnGhost} onClick={() => setJobs([...jobs, emptyJob()])}>＋ 職歴を追加</button>
            </div>

            <div style={S.card}>
              <h2 style={S.h2}>免許・資格</h2>
              {rowEditor(quals, setQual, setQuals, "MOS Excel 2019 取得")}
              <button style={S.btnGhost} onClick={() => setQuals([...quals, emptyRow()])}>＋ 行を追加</button>
            </div>

            <div style={S.card}>
              <h2 style={S.h2}>志望動機</h2>
              <textarea style={{ ...S.input, minHeight: 90 }} value={motive} onChange={(e) => setMotive(e.target.value)} />
              <h2 style={{ ...S.h2, marginTop: 14 }}>本人希望記入欄</h2>
              <textarea style={{ ...S.input, minHeight: 56 }} value={wish} onChange={(e) => setWish(e.target.value)} />
            </div>

            <div style={S.card}>
              <h2 style={S.h2}>職務経歴書用の項目</h2>
              <label style={S.label}>職務要約（3〜5行でこれまでの経歴の概要）</label>
              <textarea style={{ ...S.input, minHeight: 70 }} value={summary} onChange={(e) => setSummary(e.target.value)} />
              <label style={S.label}>活かせるスキル・経験（1行に1項目）</label>
              <textarea style={{ ...S.input, minHeight: 70 }} placeholder={"Excelでのデータ集計（VLOOKUP、ピボットテーブル）\n電話・来客応対\n正確な事務処理"} value={skills} onChange={(e) => setSkills(e.target.value)} />
              <label style={S.label}>自己PR</label>
              <textarea style={{ ...S.input, minHeight: 90 }} value={selfPR} onChange={(e) => setSelfPR(e.target.value)} />
            </div>

            <div style={{ display: "flex", gap: 10, marginBottom: 24 }}>
              <button style={S.btn} onClick={exportJSON}>JSONで保存</button>
              <button style={S.btnGhost} onClick={() => fileRef.current.click()}>JSONを読み込む</button>
              <input ref={fileRef} type="file" accept=".json" style={{ display: "none" }} onChange={importJSON} />
            </div>
          </div>
        )}

        {tab === "gap" && (
          <div style={S.card}>
            <h2 style={S.h2}>空白期間の説明サポート</h2>
            <div style={{ display: "flex", gap: 16, alignItems: "center", marginBottom: 12, flexWrap: "wrap" }}>
              <div>
                <label style={S.label}>応募スタイル</label>
                <div style={{ display: "flex", gap: 6 }}>
                  {["開示", "非開示"].map((d) => (
                    <button key={d} style={d === disclose ? S.btn : S.btnGhost} onClick={() => setDisclose(d)}>{d === "開示" ? "障害開示" : "非開示"}</button>
                  ))}
                </div>
              </div>
              <div><label style={S.label}>通所期間</label><input style={{ ...S.input, width: 110, marginBottom: 0 }} value={months} onChange={(e) => setMonths(e.target.value)} /></div>
              <div><label style={S.label}>週の通所日数</label><input style={{ ...S.input, width: 70, marginBottom: 0 }} value={days} onChange={(e) => setDays(e.target.value)} /></div>
            </div>
            {gaps.length === 0 && <div style={{ fontSize: 14, color: "#5a6b7d" }}>4ヶ月以上の空白期間は検出されていません。「入力」タブで職歴の年月を入力すると自動検出されます。</div>}
            {gaps.map((g) => (
              <div key={g.key} style={S.gapBox}>
                <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 8 }}>⚠ 空白期間: {g.from} 〜 {g.to}（約{g.months}ヶ月）</div>
                <label style={S.label}>理由を選んでテンプレートを挿入</label>
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 8 }}>
                  {Object.keys(GAP_TEMPLATES).map((r) => (
                    <button key={r} style={S.btnGhost} onClick={() => applyTemplate(g.key, r)}>{r}</button>
                  ))}
                </div>
                <textarea style={{ ...S.input, minHeight: 90 }} placeholder="テンプレートを選ぶか、自分の言葉で入力してください" value={gapTexts[g.key] || ""} onChange={(e) => setGapTexts({ ...gapTexts, [g.key]: e.target.value })} />
                <div style={{ fontSize: 12, color: "#8a6d3b" }}>※事実と異なる内容にならないよう、必ず自分の状況に合わせて編集してください。職務経歴書の「離職期間について」欄に反映されます。</div>
              </div>
            ))}
          </div>
        )}

        {tab === "export" && (
          <div style={S.card}>
            <h2 style={S.h2}>ファイル出力</h2>
            <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 12 }}>
              <button style={S.btnBlue} disabled={busy !== ""} onClick={doRireki}>{busy === "rireki" ? "生成中…" : "履歴書を出力（.docx / 指定フォーマット）"}</button>
              <button style={S.btnBlue} disabled={busy !== ""} onClick={doWord}>{busy === "word" ? "生成中…" : "Word出力（.docx / 職務経歴書）"}</button>
            </div>
            <div style={{ fontSize: 13, color: "#5a6b7d", lineHeight: 1.8 }}>
              履歴書は事業所指定のWordフォーマット（厚労省様式ベース・写真欄付き）に差し込んだ.docxファイルです。Wordで開いてそのまま編集・印刷できます。<br />
              職務経歴書は職務要約・スキル・自己PR・離職期間の説明を含む.docxファイルです。
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
