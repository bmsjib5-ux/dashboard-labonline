import { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ========== CONFIG ==========
const NOTION_DB_ID = "e5ab403a57af4659ba9e61764c2769d7";
const NOTION_MCP_URL = "https://mcp.notion.com/mcp";

// ========== HELPERS ==========
const scoreMap = { "5 - มากที่สุด": 5, "4 - มาก": 4, "3 - ปานกลาง": 3, "2 - น้อย": 2, "1 - น้อยที่สุด": 1 };

const scoreFields = [
  { key: "1.1", prop: "1.1 ระบบเข้าใช้งานง่าย", label: "ระบบเข้าใช้งานง่าย", cat: "usage" },
  { key: "1.2", prop: "1.2 ความแม่นยำและเสถียร", label: "ความแม่นยำและเสถียร", cat: "usage" },
  { key: "1.3", prop: "1.3 คู่มือเข้าใจง่าย", label: "คู่มือเข้าใจง่าย", cat: "usage" },
  { key: "2.1", prop: "2.1 ลดเวลาและค่าใช้จ่าย", label: "ลดเวลาและค่าใช้จ่าย", cat: "value" },
  { key: "2.2", prop: "2.2 ลดกระบวนการซ้ำซ้อน", label: "ลดกระบวนการซ้ำซ้อน", cat: "value" },
  { key: "2.3", prop: "2.3 ถูกต้องแม่นยำกว่าเดิม", label: "ถูกต้องแม่นยำกว่าเดิม", cat: "value" },
  { key: "2.4", prop: "2.4 ลดรายงานผลล่าช้า", label: "ลดรายงานผลล่าช้า", cat: "value" },
  { key: "3.1", prop: "3.1 ฟังก์ชันตรงความคาดหวัง", label: "ฟังก์ชันตรงความคาดหวัง", cat: "expect" },
  { key: "3.2", prop: "3.2 มั่นใจความถูกต้อง", label: "มั่นใจความถูกต้อง", cat: "expect" },
  { key: "3.3", prop: "3.3 ต้องการใช้งานต่อ", label: "ต้องการใช้งานต่อ", cat: "expect" },
  { key: "3.4", prop: "3.4 แนะนำให้ผู้อื่นใช้", label: "แนะนำให้ผู้อื่นใช้", cat: "expect" },
];

const catMeta = {
  usage: { label: "ด้านการใช้งาน", color: "#38bdf8", grad: "linear-gradient(135deg,#38bdf8,#22d3ee)" },
  value: { label: "ด้านความคุ้มค่า", color: "#818cf8", grad: "linear-gradient(135deg,#818cf8,#a78bfa)" },
  expect: { label: "ความคาดหวังและการใช้ต่อ", color: "#34d399", grad: "linear-gradient(135deg,#34d399,#22d3ee)" },
};

function avg(arr) { return arr.length ? arr.reduce((a, b) => a + b, 0) / arr.length : 0; }
function getLevel(s) {
  if (s >= 4.5) return { label: "ดีมาก", color: "#34d399", icon: "⭐" };
  if (s >= 4.0) return { label: "ดี", color: "#38bdf8", icon: "✅" };
  if (s >= 3.5) return { label: "ปานกลาง", color: "#fbbf24", icon: "🔶" };
  return { label: "ควรปรับปรุง", color: "#fb923c", icon: "⚠️" };
}

// ========== INITIAL DATA ==========
const initialData = [
  { name: "โรงพยาบาลแก่งหางแมว", scores: {"1.1":5,"1.2":5,"1.3":4,"2.1":4,"2.2":3,"2.3":5,"2.4":5,"3.1":3,"3.2":5,"3.3":5,"3.4":5}, strength:"ส่งตรวจพร้อมรายงานผลกลับได้แม่นยำรวดเร็ว", issue:"ขาดการดึง CC BP มาเพื่อความครอบคลุม", suggest:"", date:"2026-02-13" },
  { name: "ฮางโฮง อ.เมือง สกลนคร", scores: {"1.1":3,"1.2":4,"1.3":4,"2.1":3,"2.2":3,"2.3":5,"2.4":4,"3.1":4,"3.2":5,"3.3":4,"3.4":4}, strength:"ลดภาระการคีย์", issue:"ผล Lab ไม่ออกเป็นบางคน ต้องโทรตาม", suggest:"V/S นน ส่วนสูง รอบเอว ไม่ควรลงในงาน Lab", date:"2026-02-13" },
  { name: "ศสม.อ้อมน้อย", scores: {"1.1":4,"1.2":3,"1.3":4,"2.1":4,"2.2":3,"2.3":4,"2.4":4,"3.1":3,"3.2":3,"3.3":4,"3.4":3}, strength:"", issue:"", suggest:"ขอให้เสถียรในการดึงข้อมูลจาก HosXP", date:"2026-02-13" },
  { name: "สมเด็จพระยุพราชปัว", scores: {"1.1":5,"1.2":5,"1.3":2,"2.1":3,"2.2":5,"2.3":4,"2.4":5,"3.1":3,"3.2":5,"3.3":5,"3.4":5}, strength:"Realtime, Lean, Accuracy", issue:"แหล่งงบประมาณค่าเชื่อมระบบ", suggest:"พัฒนาขั้นตอนฝั่งห้องปฏิบัติการ", date:"2026-02-13" },
  { name: "รพ.สต.ปัว", scores: {"1.1":4,"1.2":4,"1.3":3,"2.1":4,"2.2":3,"2.3":5,"2.4":5,"3.1":4,"3.2":5,"3.3":5,"3.4":5}, strength:"รพ.สต. ไม่ต้องคีย์ผลเอง", issue:"ต้องคีย์ล่วงหน้า / Net ไม่เสถียร", suggest:"", date:"2026-02-13" },
  { name: "รพ.สต.บ้านเชียงเครือวัดใหญ่", scores: {"1.1":5,"1.2":4,"1.3":5,"2.1":5,"2.2":5,"2.3":5,"2.4":5,"3.1":5,"3.2":5,"3.3":5,"3.4":5}, strength:"ทราบผลตรวจแบบ Real-time", issue:"ดึงผลตรวจมาไม่ครบ", suggest:"ดูผลตรวจทุกรายการได้แบบ Real-time", date:"2026-02-13" },
];

// ========== EXCEL HELPERS ==========
function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws);
        const parsed = rows.filter((r) => r["ชื่อหน่วยบริการ"]).map((r) => {
          const scores = {};
          scoreFields.forEach((f) => {
            const raw = r[f.prop];
            if (raw == null) { scores[f.key] = 0; return; }
            if (typeof raw === "number") { scores[f.key] = raw; return; }
            scores[f.key] = scoreMap[String(raw).trim()] || parseInt(raw) || 0;
          });
          let dateVal = "";
          const rawDate = r["วันที่ตอบแบบประเมิน"];
          if (rawDate) {
            if (typeof rawDate === "number") {
              const d = XLSX.SSF.parse_date_code(rawDate);
              dateVal = `${d.y}-${String(d.m).padStart(2,"0")}-${String(d.d).padStart(2,"0")}`;
            } else {
              dateVal = String(rawDate);
            }
          }
          return {
            name: String(r["ชื่อหน่วยบริการ"]).trim(),
            scores,
            strength: String(r["จุดเด่นที่ชอบที่สุด"] || ""),
            issue: String(r["ปัญหา/อุปสรรค"] || ""),
            suggest: String(r["ข้อเสนอแนะ"] || ""),
            date: dateVal,
          };
        });
        resolve(parsed);
      } catch (err) {
        reject(new Error("ไม่สามารถอ่านไฟล์ Excel ได้: " + err.message));
      }
    };
    reader.onerror = () => reject(new Error("ไม่สามารถอ่านไฟล์ได้"));
    reader.readAsArrayBuffer(file);
  });
}

function downloadTemplate() {
  const headers = ["ชื่อหน่วยบริการ", ...scoreFields.map((f) => f.prop), "จุดเด่นที่ชอบที่สุด", "ปัญหา/อุปสรรค", "ข้อเสนอแนะ", "วันที่ตอบแบบประเมิน"];
  const example = ["โรงพยาบาลตัวอย่าง", "5 - มากที่สุด", "4 - มาก", "3 - ปานกลาง", "4 - มาก", "5 - มากที่สุด", "4 - มาก", "5 - มากที่สุด", "4 - มาก", "5 - มากที่สุด", "4 - มาก", "5 - มากที่สุด", "ระบบใช้ง่าย", "อินเทอร์เน็ตไม่เสถียร", "เพิ่มรายการตรวจ", "2026-02-15"];
  const ws = XLSX.utils.aoa_to_sheet([headers, example]);
  const colWidths = headers.map((h) => ({ wch: Math.max(h.length * 1.5, 18) }));
  ws["!cols"] = colWidths;
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "แบบประเมิน");
  XLSX.writeFile(wb, "template_lab_online.xlsx");
}

// ========== NOTION SYNC via Claude API ==========
async function fetchFromNotion() {
  const prompt = `Fetch the Notion database at URL https://www.notion.so/${NOTION_DB_ID}. Then search within data source collection://cfe2f2d1-dcd2-49b3-ac46-732bb240afe9 for all entries. For each page found, fetch its full properties. Return ONLY a JSON array (no markdown, no backticks, no explanation) where each element has this structure:
{"name":"...","scores":{"1.1":5,"1.2":4,...},"strength":"...","issue":"...","suggest":"...","date":"..."}
The score keys are: 1.1, 1.2, 1.3, 2.1, 2.2, 2.3, 2.4, 3.1, 3.2, 3.3, 3.4
Map the select values: "5 - มากที่สุด"=5, "4 - มาก"=4, "3 - ปานกลาง"=3, "2 - น้อย"=2, "1 - น้อยที่สุด"=1
"name" = ชื่อหน่วยบริการ, "strength" = จุดเด่นที่ชอบที่สุด, "issue" = ปัญหา/อุปสรรค, "suggest" = ข้อเสนอแนะ, "date" = วันที่ตอบแบบประเมิน start date.
Return ONLY the JSON array.`;

  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 4000,
      messages: [{ role: "user", content: prompt }],
      mcp_servers: [{ type: "url", url: NOTION_MCP_URL, name: "notion-mcp" }],
    }),
  });

  const data = await response.json();

  // Extract text from content blocks
  const textBlocks = data.content?.filter((b) => b.type === "text").map((b) => b.text) || [];
  const fullText = textBlocks.join("\n");

  // Try to parse JSON from response
  const jsonMatch = fullText.match(/\[[\s\S]*\]/);
  if (jsonMatch) {
    return JSON.parse(jsonMatch[0]);
  }

  throw new Error("ไม่สามารถแปลงผลลัพธ์เป็น JSON ได้");
}

// ========== COMPONENTS ==========
function Bar({ value, max = 5, color, label }) {
  const pct = (value / max) * 100;
  return (
    <div style={{ marginBottom: 12 }}>
      {label && (
        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 4, fontSize: 12 }}>
          <span style={{ color: "#cbd5e1" }}>{label}</span>
          <span style={{ fontFamily: "monospace", fontWeight: 600, color }}>{value.toFixed(2)}</span>
        </div>
      )}
      <div style={{ height: 22, background: "#1e293b", borderRadius: 6, overflow: "hidden" }}>
        <div style={{ height: "100%", width: `${pct}%`, background: color, borderRadius: 6, transition: "width 0.8s ease" }} />
      </div>
    </div>
  );
}

function Card({ title, dotColor, children, style }) {
  return (
    <div style={{ background: "#111827", border: "1px solid #2a3654", borderRadius: 14, padding: 20, ...style }}>
      {title && (
        <div style={{ fontSize: 13, fontWeight: 600, color: "#8892a8", marginBottom: 16, display: "flex", alignItems: "center", gap: 8, textTransform: "uppercase", letterSpacing: 0.4 }}>
          <span style={{ width: 8, height: 8, borderRadius: "50%", background: dotColor || "#38bdf8", display: "inline-block" }} />
          {title}
        </div>
      )}
      {children}
    </div>
  );
}

function ScoreCard({ label, value, color, accent }) {
  return (
    <div style={{ background: "#111827", border: "1px solid #2a3654", borderRadius: 14, padding: "18px 16px", textAlign: "center", position: "relative", overflow: "hidden", flex: 1, minWidth: 130 }}>
      <div style={{ position: "absolute", bottom: 0, left: 0, right: 0, height: 3, background: accent }} />
      <div style={{ fontSize: 11, color: "#8892a8", marginBottom: 6, fontWeight: 500 }}>{label}</div>
      <div style={{ fontFamily: "monospace", fontSize: 30, fontWeight: 600, color }}>{value.toFixed(2)}</div>
      <div style={{ fontSize: 10, color: "#64748b" }}>จาก 5.00</div>
    </div>
  );
}

function FeedbackTag({ type, children }) {
  const s = { strength: { bg: "rgba(52,211,153,0.12)", color: "#34d399" }, issue: { bg: "rgba(248,113,113,0.12)", color: "#f87171" }, suggest: { bg: "rgba(251,191,36,0.12)", color: "#fbbf24" } }[type];
  return <span style={{ display: "inline-block", fontSize: 10, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.5, padding: "2px 8px", borderRadius: 4, background: s.bg, color: s.color, marginBottom: 4 }}>{children}</span>;
}

// ========== MAIN DASHBOARD ==========
export default function LabOnlineDashboard() {
  const [data, setData] = useState(initialData);
  const [syncing, setSyncing] = useState(false);
  const [syncStatus, setSyncStatus] = useState(null);
  const [lastSync, setLastSync] = useState(null);
  const [activeTab, setActiveTab] = useState("overview");
  const [syncLog, setSyncLog] = useState([]);
  const [importPreview, setImportPreview] = useState(null);
  const [importError, setImportError] = useState(null);
  const [dragOver, setDragOver] = useState(false);
  const fileInputRef = useRef(null);

  const handleFile = useCallback(async (file) => {
    if (!file) return;
    setImportError(null);
    setImportPreview(null);
    try {
      const parsed = await parseExcelFile(file);
      if (parsed.length === 0) throw new Error("ไม่พบข้อมูลในไฟล์");
      setImportPreview(parsed);
    } catch (err) {
      setImportError(err.message);
    }
  }, []);

  const handleImportConfirm = useCallback(() => {
    if (!importPreview) return;
    setData(importPreview);
    setLastSync(new Date());
    setSyncStatus({ ok: true, msg: `นำเข้าจาก Excel สำเร็จ ${importPreview.length} รายการ` });
    setSyncLog((prev) => [...prev, { time: new Date(), type: "success", msg: `Import Excel: ${importPreview.length} หน่วยบริการ` }]);
    setImportPreview(null);
    setActiveTab("overview");
  }, [importPreview]);

  const handleSync = useCallback(async () => {
    setSyncing(true);
    setSyncStatus(null);
    setSyncLog((prev) => [...prev, { time: new Date(), type: "info", msg: "กำลังเชื่อมต่อ Notion Database..." }]);

    try {
      setSyncLog((prev) => [...prev, { time: new Date(), type: "info", msg: "กำลังดึงข้อมูลผ่าน Claude API + Notion MCP..." }]);
      const result = await fetchFromNotion();

      if (Array.isArray(result) && result.length > 0) {
        setData(result);
        setLastSync(new Date());
        setSyncStatus({ ok: true, msg: `ดึงข้อมูลสำเร็จ ${result.length} รายการ` });
        setSyncLog((prev) => [...prev, { time: new Date(), type: "success", msg: `สำเร็จ! ได้ข้อมูล ${result.length} หน่วยบริการ` }]);
      } else {
        throw new Error("ข้อมูลว่างเปล่า");
      }
    } catch (err) {
      setSyncStatus({ ok: false, msg: `เกิดข้อผิดพลาด: ${err.message}` });
      setSyncLog((prev) => [...prev, { time: new Date(), type: "error", msg: err.message }]);
    } finally {
      setSyncing(false);
    }
  }, []);

  // ========== STATS CALCULATION ==========
  const totalEntries = data.length;
  const questionAvgs = {};
  scoreFields.forEach((f) => { questionAvgs[f.key] = avg(data.map((d) => d.scores[f.key]).filter(Boolean)); });

  const catAvgs = {};
  ["usage", "value", "expect"].forEach((cat) => {
    const fields = scoreFields.filter((f) => f.cat === cat);
    catAvgs[cat] = avg(fields.flatMap((f) => data.map((d) => d.scores[f.key]).filter(Boolean)));
  });

  const allScores = data.flatMap((d) => Object.values(d.scores));
  const overallAvg = avg(allScores);
  const overallPct = ((overallAvg / 5) * 100).toFixed(1);

  const dist = { 5: 0, 4: 0, 3: 0, 2: 0, 1: 0 };
  allScores.forEach((s) => { if (dist[s] !== undefined) dist[s]++; });
  const totalScores = allScores.length;

  const facilityStats = data.map((d) => {
    const vals = Object.values(d.scores);
    return {
      ...d,
      total: avg(vals),
      usage: avg(scoreFields.filter((f) => f.cat === "usage").map((f) => d.scores[f.key])),
      value: avg(scoreFields.filter((f) => f.cat === "value").map((f) => d.scores[f.key])),
      expect: avg(scoreFields.filter((f) => f.cat === "expect").map((f) => d.scores[f.key])),
    };
  });

  const sortedQs = scoreFields.map((f) => ({ ...f, avg: questionAvgs[f.key] })).sort((a, b) => b.avg - a.avg);
  const bestQ = sortedQs[0];
  const worstQ = sortedQs[sortedQs.length - 1];

  // Gather strengths/issues/suggestions
  const strengths = data.filter((d) => d.strength).map((d) => d.strength);
  const issues = data.filter((d) => d.issue).map((d) => d.issue);
  const suggestions = data.filter((d) => d.suggest).map((d) => d.suggest);

  return (
    <div style={{ fontFamily: "'Noto Sans Thai','Segoe UI',sans-serif", background: "#0a0e1a", color: "#e2e8f0", minHeight: "100vh" }}>

      {/* ===== HEADER ===== */}
      <div style={{ background: "linear-gradient(135deg,#0f172a 0%,#1e1b4b 50%,#0f172a 100%)", borderBottom: "1px solid #2a3654", padding: "24px 24px 20px" }}>
        <div style={{ maxWidth: 1200, margin: "0 auto" }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
              <div style={{ width: 42, height: 42, background: "linear-gradient(135deg,#38bdf8,#818cf8)", borderRadius: 10, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18, fontWeight: 800, color: "#fff" }}>L+</div>
              <div>
                <h1 style={{ fontSize: 20, fontWeight: 700, background: "linear-gradient(135deg,#e2e8f0,#38bdf8)", WebkitBackgroundClip: "text", WebkitTextFillColor: "transparent", margin: 0 }}>Lab Online — แบบประเมินความพึงพอใจ</h1>
                <div style={{ color: "#64748b", fontSize: 11, marginTop: 2 }}>
                  Notion Database ID: {NOTION_DB_ID.slice(0, 8)}...
                  {lastSync && <span> | อัปเดตล่าสุด: {lastSync.toLocaleTimeString("th-TH")}</span>}
                </div>
              </div>
            </div>

            {/* SYNC BUTTON */}
            <button
              onClick={handleSync}
              disabled={syncing}
              style={{
                display: "flex", alignItems: "center", gap: 8,
                padding: "10px 20px", borderRadius: 10, border: "1px solid rgba(56,189,248,0.3)",
                background: syncing ? "rgba(56,189,248,0.1)" : "linear-gradient(135deg,rgba(56,189,248,0.15),rgba(129,140,248,0.15))",
                color: syncing ? "#64748b" : "#38bdf8",
                fontSize: 13, fontWeight: 600, cursor: syncing ? "wait" : "pointer",
                fontFamily: "inherit",
                transition: "all 0.3s",
              }}
            >
              <span style={{ fontSize: 18, display: "inline-block", animation: syncing ? "spin 1s linear infinite" : "none" }}>
                {syncing ? "⟳" : "🔄"}
              </span>
              {syncing ? "กำลังดึงข้อมูลจาก Notion..." : "ดึงข้อมูลจาก Notion"}
            </button>
          </div>

          {/* Sync Status Banner */}
          {syncStatus && (
            <div style={{
              marginTop: 12, padding: "10px 16px", borderRadius: 8, fontSize: 13, fontWeight: 500,
              background: syncStatus.ok ? "rgba(52,211,153,0.1)" : "rgba(248,113,113,0.1)",
              border: `1px solid ${syncStatus.ok ? "rgba(52,211,153,0.3)" : "rgba(248,113,113,0.3)"}`,
              color: syncStatus.ok ? "#34d399" : "#f87171",
              display: "flex", alignItems: "center", gap: 8,
            }}>
              <span>{syncStatus.ok ? "✅" : "❌"}</span>
              {syncStatus.msg}
              <button onClick={() => setSyncStatus(null)} style={{ marginLeft: "auto", background: "none", border: "none", color: "#64748b", cursor: "pointer", fontSize: 16 }}>✕</button>
            </div>
          )}

          {/* Summary Stats */}
          <div style={{ display: "flex", gap: 24, marginTop: 16, flexWrap: "wrap" }}>
            {[
              { val: totalEntries, label: "หน่วยบริการ", color: "#38bdf8" },
              { val: overallAvg.toFixed(2), label: "คะแนนเฉลี่ยรวม", color: "#34d399" },
              { val: `${overallPct}%`, label: "ความพึงพอใจ", color: "#818cf8" },
            ].map((s, i) => (
              <div key={i} style={{ display: "flex", alignItems: "baseline", gap: 6 }}>
                <span style={{ fontFamily: "monospace", fontSize: 22, fontWeight: 600, color: s.color }}>{s.val}</span>
                <span style={{ fontSize: 12, color: "#8892a8" }}>{s.label}</span>
              </div>
            ))}
          </div>
        </div>
      </div>

      {/* ===== TABS ===== */}
      <div style={{ maxWidth: 1200, margin: "0 auto" }}>
        <div style={{ display: "flex", gap: 2, padding: "12px 0 0", borderBottom: "1px solid #1e293b", overflowX: "auto" }}>
          {[
            { id: "overview", label: "📊 ภาพรวม" },
            { id: "detail", label: "📋 รายข้อ" },
            { id: "facility", label: "🏥 รายหน่วยบริการ" },
            { id: "feedback", label: "💬 ความคิดเห็น" },
            { id: "import", label: "📥 Import Excel" },
            { id: "log", label: "📡 Sync Log" },
          ].map((t) => (
            <button
              key={t.id} onClick={() => setActiveTab(t.id)}
              style={{
                padding: "10px 16px", fontSize: 13, fontWeight: 500, border: "none",
                borderRadius: "8px 8px 0 0", cursor: "pointer", fontFamily: "inherit",
                background: activeTab === t.id ? "#111827" : "transparent",
                color: activeTab === t.id ? "#38bdf8" : "#64748b",
                borderBottom: activeTab === t.id ? "2px solid #38bdf8" : "2px solid transparent",
                whiteSpace: "nowrap",
              }}
            >{t.label}</button>
          ))}
        </div>
      </div>

      {/* ===== CONTENT ===== */}
      <div style={{ padding: "20px 24px", maxWidth: 1200, margin: "0 auto" }}>

        {/* OVERVIEW */}
        {activeTab === "overview" && (
          <>
            <div style={{ display: "flex", gap: 12, marginBottom: 18, flexWrap: "wrap" }}>
              <ScoreCard label="ด้านการใช้งาน" value={catAvgs.usage} color="#38bdf8" accent={catMeta.usage.grad} />
              <ScoreCard label="ด้านความคุ้มค่า" value={catAvgs.value} color="#818cf8" accent={catMeta.value.grad} />
              <ScoreCard label="ความคาดหวัง/ใช้ต่อ" value={catAvgs.expect} color="#34d399" accent={catMeta.expect.grad} />
              <ScoreCard label="คะแนนรวม" value={overallAvg} color="#fbbf24" accent="linear-gradient(90deg,#fbbf24,#fb923c)" />
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 18 }}>
              <Card title="การกระจายตัวของคะแนน" dotColor="#fbbf24">
                {[5,4,3,2,1].map((s) => {
                  const count = dist[s]; const pct = totalScores ? ((count/totalScores)*100).toFixed(0) : 0;
                  const labels = {5:"มากที่สุด",4:"มาก",3:"ปานกลาง",2:"น้อย",1:"น้อยที่สุด"};
                  const colors = {5:"#34d399",4:"#38bdf8",3:"#fbbf24",2:"#fb923c",1:"#f87171"};
                  return (
                    <div key={s} style={{ display:"flex", alignItems:"center", gap:8, marginBottom:5, fontSize:12 }}>
                      <span style={{ width:80, textAlign:"right", color:"#8892a8", flexShrink:0 }}>{s} - {labels[s]}</span>
                      <div style={{ flex:1, height:17, background:"#1e293b", borderRadius:5, overflow:"hidden" }}>
                        <div style={{ height:"100%", width:`${pct}%`, background:colors[s], borderRadius:5, display:"flex", alignItems:"center", justifyContent:"flex-end", paddingRight:5, fontSize:10, fontWeight:600, color:"#fff", minWidth: pct>0?22:0 }}>{pct}%</div>
                      </div>
                      <span style={{ width:24, textAlign:"center", fontFamily:"monospace", fontWeight:600 }}>{count}</span>
                    </div>
                  );
                })}
                <div style={{ marginTop:12, paddingTop:10, borderTop:"1px solid #2a3654", display:"flex", alignItems:"center", gap:10 }}>
                  <div style={{ flex:1, height:7, background:"#1e293b", borderRadius:4, overflow:"hidden" }}>
                    <div style={{ width:`${overallPct}%`, height:"100%", background:"linear-gradient(90deg,#38bdf8,#34d399)", borderRadius:4 }} />
                  </div>
                  <span style={{ fontFamily:"monospace", fontSize:14, fontWeight:600, color:"#34d399" }}>{overallAvg.toFixed(2)}/5</span>
                </div>
              </Card>

              <div style={{ display:"flex", flexDirection:"column", gap:14 }}>
                <Card title="ข้อคะแนนสูงสุด" dotColor="#34d399">
                  <div style={{ textAlign:"center", padding:"6px 0" }}>
                    <div style={{ fontFamily:"monospace", fontSize:34, fontWeight:600, color:"#34d399" }}>{bestQ.avg.toFixed(2)}</div>
                    <div style={{ fontSize:13 }}>{bestQ.key} {bestQ.label}</div>
                  </div>
                </Card>
                <Card title="ข้อคะแนนต่ำสุด — ควรปรับปรุง" dotColor="#f87171">
                  <div style={{ textAlign:"center", padding:"6px 0" }}>
                    <div style={{ fontFamily:"monospace", fontSize:34, fontWeight:600, color:"#fb923c" }}>{worstQ.avg.toFixed(2)}</div>
                    <div style={{ fontSize:13 }}>{worstQ.key} {worstQ.label}</div>
                  </div>
                </Card>
              </div>
            </div>

            <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 }}>
              <Card title="จุดเด่นที่ได้รับชื่นชม" dotColor="#34d399">
                {strengths.length > 0 ? strengths.map((t,i) => (
                  <div key={i} style={{ background:"rgba(52,211,153,0.07)", border:"1px solid rgba(52,211,153,0.2)", borderRadius:8, padding:"9px 14px", fontSize:13, marginBottom:7 }}>🌟 {t}</div>
                )) : <div style={{ color:"#64748b", fontSize:13 }}>ไม่มีข้อมูล</div>}
              </Card>
              <Card title="ปัญหา/อุปสรรค" dotColor="#f87171">
                {issues.length > 0 ? issues.map((t,i) => (
                  <div key={i} style={{ background:"rgba(248,113,113,0.07)", border:"1px solid rgba(248,113,113,0.2)", borderRadius:8, padding:"9px 14px", fontSize:13, marginBottom:7 }}>⚠️ {t}</div>
                )) : <div style={{ color:"#64748b", fontSize:13 }}>ไม่มีข้อมูล</div>}
              </Card>
            </div>
          </>
        )}

        {/* DETAIL */}
        {activeTab === "detail" && (
          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:14 }}>
            {["usage","value","expect"].map((cat) => (
              <Card key={cat} title={catMeta[cat].label} dotColor={catMeta[cat].color}>
                {scoreFields.filter((f) => f.cat === cat).map((f) => (
                  <Bar key={f.key} value={questionAvgs[f.key]} label={`${f.key} ${f.label}`} color={catMeta[cat].color} />
                ))}
                <div style={{ marginTop:10, paddingTop:8, borderTop:"1px solid #2a3654", textAlign:"center" }}>
                  <span style={{ fontSize:11, color:"#8892a8" }}>เฉลี่ยหมวด: </span>
                  <span style={{ fontFamily:"monospace", fontSize:17, fontWeight:600, color:catMeta[cat].color }}>{catAvgs[cat].toFixed(2)}</span>
                </div>
              </Card>
            ))}
          </div>
        )}

        {/* FACILITY */}
        {activeTab === "facility" && (
          <Card title="เปรียบเทียบคะแนนรายหน่วยบริการ" dotColor="#f472b6">
            <div style={{ overflowX:"auto" }}>
              <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
                <thead>
                  <tr>
                    {["หน่วยบริการ","การใช้งาน","ความคุ้มค่า","ความคาดหวัง","เฉลี่ยรวม","ระดับ"].map((h) => (
                      <th key={h} style={{ textAlign:"left", padding:"10px 10px", fontWeight:600, color:"#8892a8", borderBottom:"2px solid #2a3654", fontSize:11, whiteSpace:"nowrap" }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {facilityStats.sort((a,b) => b.total-a.total).map((f,i) => {
                    const lv = getLevel(f.total);
                    return (
                      <tr key={i} style={{ borderBottom:"1px solid rgba(42,54,84,0.5)" }}>
                        <td style={{ padding:10, fontWeight:600, maxWidth:180 }}>{f.name}</td>
                        {[{v:f.usage,c:"#38bdf8"},{v:f.value,c:"#818cf8"},{v:f.expect,c:"#34d399"},{v:f.total,c:lv.color}].map((c,j) => (
                          <td key={j} style={{ padding:10 }}>
                            <div style={{ display:"flex", alignItems:"center", gap:5 }}>
                              <div style={{ width:48, height:5, background:"#1e293b", borderRadius:3, overflow:"hidden" }}>
                                <div style={{ width:`${(c.v/5)*100}%`, height:"100%", background:c.c, borderRadius:3 }} />
                              </div>
                              <span style={{ fontFamily:"monospace", fontSize:12, fontWeight:600, color:c.c }}>{c.v.toFixed(2)}</span>
                            </div>
                          </td>
                        ))}
                        <td style={{ padding:10, fontSize:12 }}>{lv.icon} {lv.label}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </Card>
        )}

        {/* FEEDBACK */}
        {activeTab === "feedback" && (
          <div style={{ display:"flex", flexDirection:"column", gap:12 }}>
            {facilityStats.map((f,i) => {
              const lv = getLevel(f.total);
              return (
                <Card key={i}>
                  <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10 }}>
                    <span style={{ fontSize:14, fontWeight:600, color:"#38bdf8" }}>{f.name}</span>
                    <span style={{ fontFamily:"monospace", fontSize:12, padding:"2px 10px", borderRadius:20, fontWeight:600, background:`${lv.color}22`, color:lv.color }}>{f.total.toFixed(2)}/5</span>
                  </div>
                  {f.strength && <div style={{ marginBottom:6 }}><FeedbackTag type="strength">จุดเด่น</FeedbackTag><div style={{ fontSize:13, lineHeight:1.6 }}>{f.strength}</div></div>}
                  {f.issue && <div style={{ marginBottom:6 }}><FeedbackTag type="issue">ปัญหา/อุปสรรค</FeedbackTag><div style={{ fontSize:13, lineHeight:1.6 }}>{f.issue}</div></div>}
                  {f.suggest && <div><FeedbackTag type="suggest">ข้อเสนอแนะ</FeedbackTag><div style={{ fontSize:13, lineHeight:1.6 }}>{f.suggest}</div></div>}
                </Card>
              );
            })}
          </div>
        )}

        {/* IMPORT EXCEL */}
        {activeTab === "import" && (
          <Card title="นำเข้าข้อมูลจากไฟล์ Excel" dotColor="#f472b6">
            {/* Buttons row */}
            <div style={{ display: "flex", gap: 10, marginBottom: 16 }}>
              <button onClick={downloadTemplate} style={{ padding: "8px 16px", borderRadius: 8, border: "1px solid rgba(56,189,248,0.3)", background: "rgba(56,189,248,0.1)", color: "#38bdf8", fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>
                📄 ดาวน์โหลด Template
              </button>
              <button onClick={() => { setImportPreview(null); setImportError(null); }} style={{ padding: "8px 16px", borderRadius: 8, border: "1px solid rgba(100,116,139,0.3)", background: "rgba(100,116,139,0.1)", color: "#94a3b8", fontSize: 13, fontWeight: 500, cursor: "pointer", fontFamily: "inherit", display: importPreview ? "block" : "none" }}>
                ล้างข้อมูล Preview
              </button>
            </div>

            {/* Drop zone */}
            <div
              onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
              onDragLeave={() => setDragOver(false)}
              onDrop={(e) => { e.preventDefault(); setDragOver(false); handleFile(e.dataTransfer.files[0]); }}
              onClick={() => fileInputRef.current?.click()}
              style={{
                border: `2px dashed ${dragOver ? "#38bdf8" : "#2a3654"}`,
                borderRadius: 12, padding: "40px 20px", textAlign: "center", cursor: "pointer",
                background: dragOver ? "rgba(56,189,248,0.05)" : "rgba(17,24,39,0.5)",
                transition: "all 0.3s",
              }}
            >
              <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={(e) => handleFile(e.target.files[0])} />
              <div style={{ fontSize: 40, marginBottom: 8 }}>📥</div>
              <div style={{ fontSize: 14, color: "#e2e8f0", fontWeight: 600 }}>ลากไฟล์มาวางที่นี่ หรือคลิกเพื่อเลือกไฟล์</div>
              <div style={{ fontSize: 12, color: "#64748b", marginTop: 6 }}>รองรับ .xlsx, .xls, .csv</div>
            </div>

            {/* Error */}
            {importError && (
              <div style={{ marginTop: 12, padding: "10px 16px", borderRadius: 8, background: "rgba(248,113,113,0.1)", border: "1px solid rgba(248,113,113,0.3)", color: "#f87171", fontSize: 13 }}>
                ❌ {importError}
              </div>
            )}

            {/* Preview */}
            {importPreview && (
              <div style={{ marginTop: 16 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
                  <div style={{ fontSize: 13, color: "#34d399", fontWeight: 600 }}>
                    ✅ พบข้อมูล {importPreview.length} หน่วยบริการ — ตรวจสอบก่อนนำเข้า
                  </div>
                  <button onClick={handleImportConfirm} style={{ padding: "8px 20px", borderRadius: 8, border: "none", background: "linear-gradient(135deg,#34d399,#22d3ee)", color: "#0a0e1a", fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>
                    ✓ นำเข้าข้อมูล
                  </button>
                </div>

                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                    <thead>
                      <tr>
                        {["#", "ชื่อหน่วยบริการ", ...scoreFields.map((f) => f.key), "เฉลี่ย", "จุดเด่น", "ปัญหา", "วันที่"].map((h) => (
                          <th key={h} style={{ textAlign: "left", padding: "8px 6px", fontWeight: 600, color: "#8892a8", borderBottom: "2px solid #2a3654", whiteSpace: "nowrap" }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {importPreview.map((row, i) => {
                        const vals = Object.values(row.scores);
                        const rowAvg = avg(vals);
                        const lv = getLevel(rowAvg);
                        return (
                          <tr key={i} style={{ borderBottom: "1px solid rgba(42,54,84,0.5)" }}>
                            <td style={{ padding: "6px", color: "#64748b" }}>{i + 1}</td>
                            <td style={{ padding: "6px", fontWeight: 600, maxWidth: 150, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{row.name}</td>
                            {scoreFields.map((f) => {
                              const v = row.scores[f.key];
                              const c = v >= 4 ? "#34d399" : v >= 3 ? "#fbbf24" : "#fb923c";
                              return <td key={f.key} style={{ padding: "6px", fontFamily: "monospace", fontWeight: 600, color: c, textAlign: "center" }}>{v}</td>;
                            })}
                            <td style={{ padding: "6px", fontFamily: "monospace", fontWeight: 600, color: lv.color }}>{rowAvg.toFixed(1)}</td>
                            <td style={{ padding: "6px", maxWidth: 120, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontSize: 11 }}>{row.strength || "-"}</td>
                            <td style={{ padding: "6px", maxWidth: 120, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", fontSize: 11 }}>{row.issue || "-"}</td>
                            <td style={{ padding: "6px", fontSize: 11, whiteSpace: "nowrap" }}>{row.date || "-"}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Instructions */}
            <div style={{ marginTop: 16, padding: "12px 16px", borderRadius: 8, background: "rgba(56,189,248,0.05)", border: "1px solid rgba(56,189,248,0.15)", fontSize: 12, color: "#8892a8" }}>
              <strong style={{ color: "#f472b6" }}>วิธีใช้:</strong> กดดาวน์โหลด Template → กรอกข้อมูลใน Excel → ลากไฟล์มาวาง → ตรวจสอบ Preview → กด "นำเข้าข้อมูล"
              <br />คะแนนรองรับทั้งตัวเลข (1-5) และข้อความ เช่น "5 - มากที่สุด"
            </div>
          </Card>
        )}

        {/* SYNC LOG */}
        {activeTab === "log" && (
          <Card title="Sync Log — ประวัติการดึงข้อมูล" dotColor="#22d3ee">
            {syncLog.length === 0 ? (
              <div style={{ textAlign:"center", padding:30, color:"#64748b" }}>
                <div style={{ fontSize:36, marginBottom:8 }}>📡</div>
                <div style={{ fontSize:14 }}>ยังไม่มีการดึงข้อมูล</div>
                <div style={{ fontSize:12, marginTop:4 }}>กดปุ่ม "ดึงข้อมูลจาก Notion" ด้านบนเพื่อเริ่มต้น</div>
              </div>
            ) : (
              <div style={{ display:"flex", flexDirection:"column", gap:6 }}>
                {syncLog.map((log, i) => {
                  const icons = { info:"ℹ️", success:"✅", error:"❌" };
                  const colors = { info:"#38bdf8", success:"#34d399", error:"#f87171" };
                  return (
                    <div key={i} style={{ display:"flex", alignItems:"flex-start", gap:10, padding:"8px 12px", borderRadius:8, background:"rgba(30,41,59,0.5)", fontSize:13 }}>
                      <span>{icons[log.type]}</span>
                      <span style={{ fontFamily:"monospace", fontSize:11, color:"#64748b", flexShrink:0 }}>{log.time.toLocaleTimeString("th-TH")}</span>
                      <span style={{ color: colors[log.type] }}>{log.msg}</span>
                    </div>
                  );
                })}
              </div>
            )}
            <div style={{ marginTop:16, padding:"12px 16px", borderRadius:8, background:"rgba(56,189,248,0.05)", border:"1px solid rgba(56,189,248,0.15)", fontSize:12, color:"#8892a8" }}>
              <strong style={{ color:"#38bdf8" }}>วิธีการทำงาน:</strong> เมื่อกดปุ่ม "ดึงข้อมูลจาก Notion" ระบบจะเรียก Claude API พร้อม Notion MCP Server เพื่อดึงข้อมูลล่าสุดจาก Database แล้วอัปเดต Dashboard อัตโนมัติ
            </div>
          </Card>
        )}
      </div>

      {/* FOOTER */}
      <div style={{ textAlign:"center", padding:"16px 24px", borderTop:"1px solid #1e293b", color:"#475569", fontSize:11 }}>
        Notion Database: แบบประเมินความพึงพอใจ Lab Online | ดึงข้อมูลผ่าน Claude API + Notion MCP
      </div>

      <style>{`@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>
    </div>
  );
}
