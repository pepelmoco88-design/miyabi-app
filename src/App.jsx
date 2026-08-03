import React, { useState, useEffect, useRef, useCallback } from "react";
import {
  BarChart, Bar, PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid,
  Tooltip, ResponsiveContainer, Legend
} from "recharts";
import * as XLSX from "xlsx";
import { Camera, Upload, Trash2, Loader2, Receipt, TrendingUp, TrendingDown, Wallet, X, Pencil, Download } from "lucide-react";

// ---- Design tokens ----
const COLORS = {
  navy: "#16213E",
  navyDeep: "#0F1830",
  navySoft: "#26365E",
  cream: "#FAF6EE",
  creamCard: "#FFFFFF",
  gold: "#B8935F",
  goldDeep: "#9C7A45",
  text: "#2A2A2A",
  textMuted: "#6B6458",
  border: "#E4DCC9",
};

const CATEGORY_LIST = [
  "食材費", "酒類・飲料費", "消耗品費", "水道光熱費", "通信費",
  "家賃", "人件費", "修繕費", "広告宣伝費", "雑費", "売上", "その他"
];

const CATEGORY_COLORS = {
  "食材費": "#B5654A",
  "酒類・飲料費": "#9C7A45",
  "消耗品費": "#7C8B65",
  "水道光熱費": "#5B7A9D",
  "通信費": "#8E7CC3",
  "家賃": "#4A6670",
  "人件費": "#C77B58",
  "修繕費": "#6E8B7A",
  "広告宣伝費": "#B8935F",
  "雑費": "#9C9284",
  "売上": "#2E7D5B",
  "その他": "#8A8378",
};

const PERIOD_TABS = [
  { value: "day", label: "日" },
  { value: "month", label: "月" },
  { value: "year", label: "年" },
];

const TYPE_TABS = [
  { value: "all", label: "すべて" },
  { value: "income", label: "売上" },
  { value: "expense", label: "経費" },
];

function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result.split(",")[1]);
    r.onerror = () => reject(new Error("読み込みに失敗しました"));
    r.readAsDataURL(file);
  });
}

// ストレージの安全なフォールバック設定
const safeStorage = {
  get: async (key) => {
    try {
      if (typeof window !== "undefined" && window.storage && window.storage.get) {
        return await window.storage.get(key, true);
      }
      const val = localStorage.getItem(key);
      return val ? { value: val } : null;
    } catch (e) {
      return null;
    }
  },
  set: async (key, val) => {
    try {
      if (typeof window !== "undefined" && window.storage && window.storage.set) {
        return await window.storage.set(key, val, true);
      }
      localStorage.setItem(key, val);
    } catch (e) {
      console.error(e);
    }
  }
};

async function analyzeReceipt(base64, mediaType) {
  const system = `あなたはレシート・伝票を読み取る経理アシスタントです。画像から情報を抽出し、必ず次のJSON形式のみを出力してください。前置き、説明、コードブロックの記号は一切付けないこと。

{
  "date": "YYYY-MM-DD",
  "storeName": "店名または取引先名",
  "amount": 1234,
  "type": "expense または income",
  "category": "次の中から一つ選ぶ: 食材費, 酒類・飲料費, 消耗品費, 水道光熱費, 通信費, 家賃, 人件費, 修繕費, 広告宣伝費, 雑費, 売上, その他",
  "memo": "購入品目や内容の簡潔な要約(20文字以内)",
  "items": [{"name": "商品名", "amount": 700}]
}

日付が読み取れない場合は今日の日付を使うこと。金額は税込の合計金額を数値のみで。売上を記録した伝票の場合は type を income、category を "売上" にすること。
"items" について: 会計伝票・売上伝票で商品ごとの内訳(品名と金額)が読み取れる場合は、すべての商品をitemsに配列で入れること。内訳が読み取れない、または経費のレシートの場合は items は空配列 [] にすること。`;

  const response = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { 
      "Content-Type": "application/json",
      "dangerously-allow-browser": "true"
    },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 1000,
      system,
      messages: [
        {
          role: "user",
          content: [
            { type: "image", source: { type: "base64", media_type: mediaType, data: base64 } },
            { type: "text", text: "このレシート・伝票の内容をJSONで抽出してください。" }
          ]
        }
      ]
    })
  });

  if (!response.ok) throw new Error("API呼び出しに失敗しました");
  const data = await response.json();
  const text = data.content.map(b => b.text || "").join("").trim();
  const clean = text.replace(/```json|```/g, "").trim();
  const parsed = JSON.parse(clean);
  const items = Array.isArray(parsed.items)
    ? parsed.items
        .filter(it => it && (it.name || it.amount))
        .map(it => ({ name: String(it.name || "商品").slice(0, 40), amount: Number(it.amount) || 0 }))
    : [];
  return {
    date: parsed.date || new Date().toISOString().slice(0, 10),
    storeName: parsed.storeName || "不明",
    amount: Number(parsed.amount) || 0,
    type: parsed.type === "income" ? "income" : "expense",
    category: CATEGORY_LIST.includes(parsed.category) ? parsed.category : "その他",
    memo: parsed.memo || "",
    items,
  };
}

function currency(n) {
  return "¥" + Math.round(n).toLocaleString("ja-JP");
}

function todayStr() {
  return new Date().toISOString().slice(0, 10);
}

function periodKey(dateStr, type) {
  const d = dateStr || todayStr();
  if (type === "day") return d.slice(0, 10);
  if (type === "year") return d.slice(0, 4);
  return d.slice(0, 7);
}

function periodLabel(key, type) {
  if (!key) return "";
  if (type === "day") {
    const [y, m, d] = key.split("-");
    return `${y}年${m}月${d}日`;
  }
  if (type === "year") return `${key}年`;
  const [y, m] = key.split("-");
  return `${y}年${m}月`;
}

function shortPeriodLabel(key, type) {
  if (type === "day") return key.slice(5).replace("-", "/");
  if (type === "year") return key;
  return key.slice(5, 7) + "月";
}

const WEEKDAYS = ["日", "月", "火", "水", "木", "金", "土"];
function formatCardDate(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr + "T00:00:00");
  if (isNaN(d.getTime())) return dateStr;
  return `${dateStr} (${WEEKDAYS[d.getDay()]})`;
}

function lastNPeriods(type, n) {
  const arr = [];
  const now = new Date();
  for (let i = n - 1; i >= 0; i--) {
    let dt;
    if (type === "day") dt = new Date(now.getFullYear(), now.getMonth(), now.getDate() - i);
    else if (type === "year") dt = new Date(now.getFullYear() - i, 0, 1);
    else dt = new Date(now.getFullYear(), now.getMonth() - i, 1);
    arr.push(periodKey(dt.toISOString().slice(0, 10), type));
  }
  return arr;
}

const TREND_LENGTH = { day: 14, month: 6, year: 5 };

const inputStyle = {
  width: "100%", padding: "9px 10px", borderRadius: 7,
  border: `1px solid ${COLORS.border}`, background: "#fff", color: COLORS.text,
};

export default function App() {
  const [receipts, setReceipts] = useState([]);
  const [loaded, setLoaded] = useState(false);
  const [processing, setProcessing] = useState([]);
  const [errorMsg, setErrorMsg] = useState("");
  const [periodType, setPeriodType] = useState("month");
  const [selectedPeriod, setSelectedPeriod] = useState("");
  const [typeFilter, setTypeFilter] = useState("all");
  const [filterCategory, setFilterCategory] = useState("すべて");
  const [dragOver, setDragOver] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [editDraft, setEditDraft] = useState(null);
  const fileInputRef = useRef(null);
  const cameraInputRef = useRef(null);

  useEffect(() => {
    (async () => {
      try {
        const res = await safeStorage.get("receipts");
        const list = res ? JSON.parse(res.value) : [];
        setReceipts(list);
        const periods = Array.from(new Set(list.map(r => periodKey(r.date, "month")))).sort().reverse();
        setSelectedPeriod(periods[0] || periodKey(todayStr(), "month"));
      } catch (e) {
        setSelectedPeriod(periodKey(todayStr(), "month"));
      } finally {
        setLoaded(true);
      }
    })();
  }, []);

  const persist = useCallback(async (list) => {
    try {
      await safeStorage.set("receipts", JSON.stringify(list));
    } catch (e) {
      setErrorMsg("保存に失敗しました。もう一度お試しください。");
    }
  }, []);

  const changePeriodType = useCallback((type) => {
    setPeriodType(type);
    const periods = Array.from(new Set(receipts.map(r => periodKey(r.date, type)))).sort().reverse();
    setSelectedPeriod(periods[0] || periodKey(todayStr(), type));
  }, [receipts]);

  const handleFiles = useCallback(async (fileList) => {
    const files = Array.from(fileList || []).filter(f => f.type.startsWith("image/"));
    if (files.length === 0) return;
    setErrorMsg("");

    for (const file of files) {
      const tempId = uid();
      setProcessing(p => [...p, { id: tempId, name: file.name }]);
      try {
        const base64 = await fileToBase64(file);
        const result = await analyzeReceipt(base64, file.type || "image/jpeg");
        const entry = { id: uid(), ...result, addedAt: new Date().toISOString() };
        setReceipts(prev => {
          const next = [entry, ...prev];
          persist(next);
          return next;
        });
        setSelectedPeriod(periodKey(entry.date, periodType));
      } catch (e) {
        setErrorMsg(`「${file.name}」の読み取りに失敗しました。画像がはっきり写っているか確認してください。`);
      } finally {
        setProcessing(p => p.filter(x => x.id !== tempId));
      }
    }
  }, [persist, periodType]);

  const deleteReceipt = useCallback((id) => {
    setReceipts(prev => {
      const next = prev.filter(r => r.id !== id);
      persist(next);
      return next;
    });
  }, [persist]);

  const updateCategory = useCallback((id, category) => {
    setReceipts(prev => {
      const next = prev.map(r => r.id === id ? { ...r, category } : r);
      persist(next);
      return next;
    });
  }, [persist]);

  const startEdit = useCallback((r) => {
    setEditingId(r.id);
    setEditDraft({ ...r, amount: String(r.amount) });
  }, []);

  const cancelEdit = useCallback(() => {
    setEditingId(null);
    setEditDraft(null);
  }, []);

  const saveEdit = useCallback(() => {
    if (!editDraft) return;
    const amountNum = Number(editDraft.amount);
    if (!editDraft.date || !editDraft.storeName.trim() || isNaN(amountNum) || amountNum < 0) {
      setErrorMsg("日付・店名・金額を正しく入力してください。");
      return;
    }
    setReceipts(prev => {
      const next = prev.map(r => r.id === editDraft.id ? { ...editDraft, amount: amountNum } : r);
      persist(next);
      return next;
    });
    setSelectedPeriod(periodKey(editDraft.date, periodType));
    setEditingId(null);
    setEditDraft(null);
  }, [editDraft, persist, periodType]);

  const exportToSpreadsheet = useCallback(() => {
    if (receipts.length === 0) {
      setErrorMsg("エクスポートできるデータがまだありません。");
      return;
    }
    const rows = [...receipts]
      .sort((a, b) => (a.date || "").localeCompare(b.date || ""))
      .map(r => ({
        "日付": r.date,
        "種別": r.type === "income" ? "売上" : "経費",
        "カテゴリ": r.category,
        "店名・取引先": r.storeName,
        "金額": r.amount,
        "メモ": r.memo || "",
      }));
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [{ wch: 12 }, { wch: 8 }, { wch: 14 }, { wch: 24 }, { wch: 12 }, { wch: 30 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "経費・売上明細");
    XLSX.writeFile(wb, `MIYABI_経費売上_${todayStr()}.xlsx`);
  }, [receipts]);

  const periodsForType = Array.from(new Set(receipts.map(r => periodKey(r.date, periodType)))).sort().reverse();
  const periodOptions = periodsForType.includes(selectedPeriod)
    ? periodsForType
    : Array.from(new Set([selectedPeriod, ...periodsForType])).filter(Boolean);

  const currentPeriodList = receipts.filter(r => periodKey(r.date, periodType) === selectedPeriod);
  const periodCount = currentPeriodList.length;

  const categoryOptions = typeFilter === "income"
    ? ["売上"]
    : typeFilter === "expense"
      ? CATEGORY_LIST.filter(c => c !== "売上")
      : CATEGORY_LIST;

  const visibleList = currentPeriodList
    .filter(r => typeFilter === "all" || r.type === typeFilter)
    .filter(r => filterCategory === "すべて" || r.category === filterCategory)
    .sort((a, b) => (b.date || "").localeCompare(a.date || ""));

  const periodExpense = currentPeriodList.filter(r => r.type === "expense").reduce((s, r) => s + r.amount, 0);
  const periodIncome = currentPeriodList.filter(r => r.type === "income").reduce((s, r) => s + r.amount, 0);
  const periodNet = periodIncome - periodExpense;

  const pieData = CATEGORY_LIST
    .map(cat => ({
      name: cat,
      value: currentPeriodList.filter(r => r.category === cat && r.type === "expense").reduce((s, r) => s + r.amount, 0)
    }))
    .filter(d => d.value > 0);

  const rankingMap = new Map();
  currentPeriodList.filter(r => r.type === "income").forEach(r => {
    const items = (r.items && r.items.length) ? r.items : [{ name: r.storeName || "商品", amount: r.amount }];
    items.forEach(it => {
      const key = it.name || "商品";
      const cur = rankingMap.get(key) || { name: key, amount: 0, count: 0 };
      cur.amount += Number(it.amount) || 0;
      cur.count += 1;
      rankingMap.set(key, cur);
    });
  });
  const productRanking = Array.from(rankingMap.values()).sort((a, b) => b.amount - a.amount).slice(0, 10);
  const RANK_COLORS = ["#C9A15A", "#A8A8A8", "#B08A5A"];

  const trendKeys = lastNPeriods(periodType, TREND_LENGTH[periodType]);
  const trendData = trendKeys.map(key => {
    const list = receipts.filter(r => periodKey(r.date, periodType) === key);
    return {
      period: shortPeriodLabel(key, periodType),
      支出: list.filter(r => r.type === "expense").reduce((s, r) => s + r.amount, 0),
      売上: list.filter(r => r.type === "income").reduce((s, r) => s + r.amount, 0),
    };
  });

  const periodPrefix = periodType === "day" ? "本日" : periodType === "year" ? "今年" : "今月";
  const trendTitle = periodType === "day" ? "日別推移(直近14日間)" : periodType === "year" ? "年別推移(直近5年間)" : "月別推移(直近6か月)";

  return (
    <div style={{
      minHeight: "100vh",
      background: COLORS.cream,
      fontFamily: "'Noto Sans JP', sans-serif",
      color: COLORS.text,
      paddingBottom: 48,
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Shippori+Mincho:wght@500;700&family=Noto+Sans+JP:wght@400;500;700&display=swap');
        * { box-sizing: border-box; }
        .mincho { font-family: 'Shippori Mincho', serif; }
        @keyframes scanline {
          0% { transform: translateY(-100%); opacity: 0; }
          10% { opacity: 1; }
          90% { opacity: 1; }
          100% { transform: translateY(2400%); opacity: 0; }
        }
        @keyframes feedIn {
          0% { transform: translateY(-6px); }
          50% { transform: translateY(0px); }
          100% { transform: translateY(-6px); }
        }
        .upload-zone { transition: border-color 0.2s ease, background 0.2s ease; }
        select, button, input { font-family: inherit; }
        button, select, input[type="date"], input[type="number"], input[type="text"] {
          min-height: 40px;
        }
        button:focus-visible, select:focus-visible, input:focus-visible {
          outline: 2px solid ${COLORS.gold};
          outline-offset: 2px;
        }
        button { -webkit-tap-highlight-color: transparent; touch-action: manipulation; }
        input, select { font-size: 16px; }
        .card-row { transition: background 0.15s ease; }
        .tab-btn { transition: background 0.15s ease, color 0.15s ease; }
        @media (max-width: 720px) {
          .grid-2 { grid-template-columns: 1fr !important; }
          .summary-row { grid-template-columns: 1fr 1fr !important; }
          .upload-actions { width: 100%; }
          .upload-actions button { flex: 1; justify-content: center; }
          .upload-zone-inner { width: 100%; }
          .toolbar-select { width: 100%; }
          .fab { display: flex !important; }
          .tab-group { width: 100%; }
          .tab-btn { flex: 1; }
        }
        @media (min-width: 721px) {
          .fab { display: none !important; }
        }
      `}</style>

      {/* Header */}
      <div style={{
        background: `linear-gradient(180deg, ${COLORS.navyDeep} 0%, ${COLORS.navy} 100%)`,
        padding: "28px 20px 22px",
        borderBottom: `3px solid ${COLORS.gold}`,
      }}>
        <div style={{ maxWidth: 1040, margin: "0 auto", display: "flex", alignItems: "center", gap: 14 }}>
          <div style={{
            width: 44, height: 44, borderRadius: "50%", border: `1.5px solid ${COLORS.gold}`,
            display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
          }}>
            <span className="mincho" style={{ color: COLORS.gold, fontSize: 20, fontWeight: 700 }}>雅</span>
          </div>
          <div>
            <div className="mincho" style={{ color: "#fff", fontSize: 21, fontWeight: 700, letterSpacing: 1 }}>
              MIYABI 経費・売上管理
            </div>
            <div style={{ color: "#B8A98A", fontSize: 12, marginTop: 2, letterSpacing: 0.5 }}>
              レシート撮影 → AI読み取り → クラウド保存
            </div>
          </div>
        </div>
      </div>

      <div style={{ maxWidth: 1040, margin: "0 auto", padding: "24px 20px 0" }}>

        {errorMsg && (
          <div style={{
            background: "#FDECEA", border: "1px solid #E4A199", color: "#8A3A32",
            borderRadius: 8, padding: "10px 14px", fontSize: 13.5, marginBottom: 16,
            display: "flex", justifyContent: "space-between", alignItems: "center",
          }}>
            <span>{errorMsg}</span>
            <button onClick={() => setErrorMsg("")} style={{ background: "none", border: "none", cursor: "pointer", color: "#8A3A32" }}>
              <X size={16} />
            </button>
          </div>
        )}

        {/* Upload zone */}
        <div
          className="upload-zone"
          onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={(e) => { e.preventDefault(); setDragOver(false); handleFiles(e.dataTransfer.files); }}
          style={{
            position: "relative",
            overflow: "hidden",
            background: COLORS.navy,
            border: `2px dashed ${dragOver ? COLORS.gold : "#3D4C74"}`,
            borderRadius: 14,
            padding: "28px 24px",
            display: "flex",
            flexWrap: "wrap",
            alignItems: "center",
            justifyContent: "space-between",
            gap: 18,
          }}
        >
          {processing.length > 0 && (
            <div style={{
              position: "absolute", left: 0, right: 0, height: 3,
              background: `linear-gradient(90deg, transparent, ${COLORS.gold}, transparent)`,
              animation: "scanline 1.6s linear infinite",
            }} />
          )}
          <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
            <div style={{
              width: 52, height: 52, borderRadius: 10, background: COLORS.navySoft,
              display: "flex", alignItems: "center", justifyContent: "center",
              animation: processing.length > 0 ? "feedIn 1s ease-in-out infinite" : "none",
              border: `1px solid ${COLORS.gold}`,
            }}>
              <Receipt size={24} color={COLORS.gold} />
            </div>
            <div>
              <div className="mincho" style={{ color: "#fff", fontSize: 16.5, fontWeight: 700 }}>
                レシート・伝票を読み込む
              </div>
              <div style={{ color: "#A9B0C8", fontSize: 12.5, marginTop: 3 }}>
                {processing.length > 0
                  ? `${processing.length}件を読み取り中...`
                  : "写真を撮る、または画像をドラッグ&ドロップ"}
              </div>
            </div>
          </div>

          <div className="upload-actions" style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            <button
              onClick={() => cameraInputRef.current?.click()}
              style={{
                display: "flex", alignItems: "center", justifyContent: "center", gap: 7, padding: "12px 18px",
                background: COLORS.gold, color: COLORS.navyDeep, border: "none",
                borderRadius: 8, fontWeight: 700, fontSize: 14.5, cursor: "pointer",
              }}
            >
              <Camera size={17} /> 撮影する
            </button>
            <button
              onClick={() => fileInputRef.current?.click()}
              style={{
                display: "flex", alignItems: "center", justifyContent: "center", gap: 7, padding: "12px 18px",
                background: "transparent", color: "#fff", border: "1px solid #3D4C74",
                borderRadius: 8, fontWeight: 500, fontSize: 14.5, cursor: "pointer",
              }}
            >
              <Upload size={17} /> 画像を選択
            </button>
          </div>

          <input ref={cameraInputRef} type="file" accept="image/*" capture="environment" multiple
            style={{ display: "none" }} onChange={(e) => { handleFiles(e.target.files); e.target.value = ""; }} />
          <input ref={fileInputRef} type="file" accept="image/*" multiple
            style={{ display: "none" }} onChange={(e) => { handleFiles(e.target.files); e.target.value = ""; }} />
        </div>

        <button
          className="fab"
          onClick={() => cameraInputRef.current?.click()}
          aria-label="レシートを撮影する"
          style={{
            display: "none",
            position: "fixed", right: 18, bottom: 22, zIndex: 20,
            width: 56, height: 56, borderRadius: "50%",
            background: COLORS.gold, border: `2px solid ${COLORS.creamCard}`,
            alignItems: "center", justifyContent: "center", cursor: "pointer",
            boxShadow: "0 6px 16px rgba(22,33,62,0.35)",
          }}
        >
          <Camera size={22} color={COLORS.navyDeep} />
        </button>

        {!loaded ? (
          <div style={{ textAlign: "center", padding: "60px 0", color: COLORS.textMuted }}>
            <Loader2 className="spin" size={22} style={{ animation: "spin 1s linear infinite" }} />
            <div style={{ marginTop: 10, fontSize: 13.5 }}>読み込み中...</div>
          </div>
        ) : (
        <>
        {/* Period tabs + summary */}
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", margin: "26px 0 14px", flexWrap: "wrap", gap: 10 }}>
          <div className="mincho" style={{ fontSize: 18, fontWeight: 700, color: COLORS.navy }}>
            サマリー
          </div>
          <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
            <TabGroup options={PERIOD_TABS} value={periodType} onChange={changePeriodType} />
            <select
              className="toolbar-select"
              value={selectedPeriod}
              onChange={(e) => setSelectedPeriod(e.target.value)}
              style={{
                padding: "10px 12px", borderRadius: 7, border: `1px solid ${COLORS.border}`,
                background: "#fff", fontSize: 14, color: COLORS.text,
              }}
            >
              {periodOptions.map(p => (
                <option key={p} value={p}>{periodLabel(p, periodType)}</option>
              ))}
            </select>
          </div>
        </div>

        {/* Summary cards */}
        <div className="summary-row" style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 14, marginBottom: 22 }}>
          <SummaryCard icon={<TrendingDown size={18} color="#B5654A" />} label={`${periodPrefix}の支出`} value={currency(periodExpense)} color="#B5654A" />
          <SummaryCard icon={<TrendingUp size={18} color="#2E7D5B" />} label={`${periodPrefix}の売上`} value={currency(periodIncome)} color="#2E7D5B" />
          <SummaryCard icon={<Wallet size={18} color={COLORS.goldDeep} />} label="収支" value={(periodNet >= 0 ? "+" : "") + currency(periodNet)} color={periodNet >= 0 ? "#2E7D5B" : "#B5654A"} />
          <SummaryCard icon={<Receipt size={18} color={COLORS.goldDeep} />} label={`${periodPrefix}の件数`} value={`${periodCount}件`} color={COLORS.goldDeep} />
        </div>

        {/* Charts */}
        <div className="grid-2" style={{ display: "grid", gridTemplateColumns: "1.2fr 1fr", gap: 16, marginBottom: 28 }}>
          <ChartCard title={trendTitle}>
            <ResponsiveContainer width="100%" height={220}>
              <BarChart data={trendData} margin={{ top: 6, right: 8, left: 0, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#EEE7D6" />
                <XAxis dataKey="period" tick={{ fontSize: 11.5, fill: COLORS.textMuted }} axisLine={{ stroke: COLORS.border }} tickLine={false} />
                <YAxis tick={{ fontSize: 11, fill: COLORS.textMuted }} axisLine={false} tickLine={false}
                  tickFormatter={(v) => v >= 10000 ? `${Math.round(v / 10000)}万` : v} />
                <Tooltip formatter={(v) => currency(v)} contentStyle={{ fontSize: 12.5, borderRadius: 8, border: `1px solid ${COLORS.border}` }} />
                <Legend wrapperStyle={{ fontSize: 12 }} />
                <Bar dataKey="支出" fill="#B5654A" radius={[4, 4, 0, 0]} />
                <Bar dataKey="売上" fill="#2E7D5B" radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </ChartCard>

          <ChartCard title={`${periodPrefix}の支出内訳`}>
            {pieData.length === 0 ? (
              <div style={{ height: 220, display: "flex", alignItems: "center", justifyContent: "center", color: COLORS.textMuted, fontSize: 13 }}>
                この期間の支出データがありません
              </div>
            ) : (
              <ResponsiveContainer width="100%" height={220}>
                <PieChart>
                  <Pie data={pieData} dataKey="value" nameKey="name" cx="50%" cy="50%" innerRadius={45} outerRadius={75} paddingAngle={2}>
                    {pieData.map((d, i) => <Cell key={i} fill={CATEGORY_COLORS[d.name] || "#999"} />)}
                  </Pie>
                  <Tooltip formatter={(v) => currency(v)} contentStyle={{ fontSize: 12.5, borderRadius: 8, border: `1px solid ${COLORS.border}` }} />
                  <Legend wrapperStyle={{ fontSize: 11 }} layout="vertical" align="right" verticalAlign="middle" />
                </PieChart>
              </ResponsiveContainer>
            )}
          </ChartCard>
        </div>

        {/* Product ranking */}
        <div style={{ marginBottom: 28 }}>
          <div className="mincho" style={{ fontSize: 18, fontWeight: 700, color: COLORS.navy, marginBottom: 12 }}>
            売上商品ランキング({periodPrefix})
          </div>
          <div style={{ background: "#fff", border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: "8px 16px" }}>
            {productRanking.length === 0 ? (
              <div style={{ padding: "24px 4px", textAlign: "center", color: COLORS.textMuted, fontSize: 13.5 }}>
                この期間の売上データがありません
              </div>
            ) : productRanking.map((p, i) => (
              <div key={p.name} style={{
                display: "flex", alignItems: "center", gap: 12, padding: "12px 0",
                borderTop: i === 0 ? "none" : `1px solid ${COLORS.border}`,
              }}>
                <div style={{
                  width: 26, height: 26, borderRadius: "50%", flexShrink: 0,
                  background: i < 3 ? RANK_COLORS[i] : COLORS.cream,
                  color: i < 3 ? "#fff" : COLORS.textMuted,
                  display: "flex", alignItems: "center", justifyContent: "center",
                  fontSize: 12.5, fontWeight: 700,
                }}>
                  {i + 1}
                </div>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 14, fontWeight: 700, color: COLORS.navy, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                    {p.name}
                  </div>
                  <div style={{ fontSize: 11.5, color: COLORS.textMuted }}>{p.count}回</div>
                </div>
                <div className="mincho" style={{ fontSize: 15, fontWeight: 700, color: "#2E7D5B", whiteSpace: "nowrap" }}>
                  {currency(p.amount)}
                </div>
              </div>
            ))}
          </div>
        </div>

        {/* List */}
        <div style={{ display: "flex", flexDirection: "column", gap: 10, marginBottom: 12 }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 10 }}>
            <div className="mincho" style={{ fontSize: 18, fontWeight: 700, color: COLORS.navy }}>明細一覧</div>
            <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
              <button
                onClick={exportToSpreadsheet}
                style={{
                  display: "flex", alignItems: "center", gap: 6, padding: "8px 12px",
                  background: "#fff", color: COLORS.navy, border: `1px solid ${COLORS.border}`,
                  borderRadius: 7, fontWeight: 600, fontSize: 13, cursor: "pointer",
                }}
              >
                <Download size={15} /> スプレッドシート出力
              </button>
              <TabGroup
                options={TYPE_TABS}
                value={typeFilter}
                onChange={(v) => {
                  setTypeFilter(v);
                  const opts = v === "income" ? ["売上"] : v === "expense" ? CATEGORY_LIST.filter(c => c !== "売上") : CATEGORY_LIST;
                  if (filterCategory !== "すべて" && !opts.includes(filterCategory)) setFilterCategory("すべて");
                }}
              />
            </div>
          </div>
          {typeFilter !== "income" && (
            <select
              className="toolbar-select"
              value={filterCategory}
              onChange={(e) => setFilterCategory(e.target.value)}
              style={{ padding: "10px 12px", borderRadius: 7, border: `1px solid ${COLORS.border}`, background: "#fff", fontSize: 14, alignSelf: "flex-start" }}
            >
              <option value="すべて">すべてのカテゴリ</option>
              {categoryOptions.map(c => <option key={c} value={c}>{c}</option>)}
            </select>
          )}
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {visibleList.length === 0 ? (
            <div style={{
              background: "#fff", borderRadius: 12, border: `1px solid ${COLORS.border}`,
              padding: "40px 20px", textAlign: "center", color: COLORS.textMuted, fontSize: 13.5,
            }}>
              この条件に一致する明細はまだありません。上のボタンからレシートを読み込んでください。
            </div>
          ) : visibleList.map(r => {
            const isEditing = editingId === r.id;
            return (
              <div key={r.id} className="card-row" style={{
                background: "#fff", borderRadius: 12, border: `1px solid ${COLORS.border}`,
                borderLeft: `4px solid ${CATEGORY_COLORS[r.category] || COLORS.border}`,
                padding: "14px 16px",
              }}>
                {isEditing ? (
                  <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                      <Field label="日付">
                        <input type="date" value={editDraft.date}
                          onChange={(e) => setEditDraft(d => ({ ...d, date: e.target.value }))}
                          style={inputStyle} />
                      </Field>
                      <Field label="金額">
                        <input type="number" inputMode="numeric" value={editDraft.amount}
                          onChange={(e) => setEditDraft(d => ({ ...d, amount: e.target.value }))}
                          style={inputStyle} />
                      </Field>
                    </div>
                    <Field label="店名・取引先">
                      <input type="text" value={editDraft.storeName}
                        onChange={(e) => setEditDraft(d => ({ ...d, storeName: e.target.value }))}
                        style={inputStyle} />
                    </Field>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                      <Field label="カテゴリ">
                        <select value={editDraft.category}
                          onChange={(e) => setEditDraft(d => ({ ...d, category: e.target.value }))}
                          style={inputStyle}>
                          {CATEGORY_LIST.map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                      </Field>
                      <Field label="種別">
                        <select value={editDraft.type}
                          onChange={(e) => setEditDraft(d => ({ ...d, type: e.target.value }))}
                          style={inputStyle}>
                          <option value="expense">支出</option>
                          <option value="income">売上</option>
                        </select>
                      </Field>
                    </div>
                    <Field label="メモ">
                      <input type="text" value={editDraft.memo}
                        onChange={(e) => setEditDraft(d => ({ ...d, memo: e.target.value }))}
                        style={inputStyle} />
                    </Field>
                    <div style={{ display: "flex", gap: 8, marginTop: 4 }}>
                      <button onClick={saveEdit} style={{
                        flex: 1, padding: "10px 14px", background: COLORS.navy, color: "#fff",
                        border: "none", borderRadius: 8, fontWeight: 700, fontSize: 14, cursor: "pointer",
                      }}>保存</button>
                      <button onClick={cancelEdit} style={{
                        flex: 1, padding: "10px 14px", background: "#fff", color: COLORS.textMuted,
                        border: `1px solid ${COLORS.border}`, borderRadius: 8, fontWeight: 500, fontSize: 14, cursor: "pointer",
                      }}>キャンセル</button>
                    </div>
                  </div>
                ) : (
                  <>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10 }}>
                      <div style={{ minWidth: 0 }}>
                        <div style={{ fontSize: 11.5, color: COLORS.textMuted, marginBottom: 3 }}>{formatCardDate(r.date)}</div>
                        <div style={{ fontSize: 15, fontWeight: 700, color: COLORS.navy, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                          {r.storeName}
                        </div>
                      </div>
                      <div className="mincho" style={{
                        fontSize: 17, fontWeight: 700, whiteSpace: "nowrap",
                        color: r.type === "income" ? "#2E7D5B" : COLORS.text,
                      }}>
                        {r.type === "income" ? "+" : "-"}{currency(r.amount)}
                      </div>
                    </div>

                    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginTop: 10, gap: 10, flexWrap: "wrap" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", minWidth: 0 }}>
                        <select
                          value={r.category}
                          onChange={(e) => updateCategory(r.id, e.target.value)}
                          style={{
                            border: "none", background: (CATEGORY_COLORS[r.category] || "#999") + "22",
                            color: CATEGORY_COLORS[r.category] || COLORS.text,
                            borderRadius: 6, padding: "6px 8px", fontSize: 12.5, fontWeight: 600, cursor: "pointer",
                          }}
                        >
                          {CATEGORY_LIST.map(c => <option key={c} value={c}>{c}</option>)}
                        </select>
                        {r.memo && (
                          <span style={{ color: COLORS.textMuted, fontSize: 12.5, overflow: "hidden", textOverflow: "ellipsis" }}>
                            {r.memo}
                          </span>
                        )}
                      </div>
                      <div style={{ display: "flex", gap: 4, flexShrink: 0 }}>
                        <button
                          onClick={() => startEdit(r)}
                          aria-label="編集"
                          style={{
                            display: "flex", alignItems: "center", justifyContent: "center",
                            width: 36, height: 36, background: "none", border: "none", cursor: "pointer", color: COLORS.goldDeep,
                          }}
                        >
                          <Pencil size={16} />
                        </button>
                        <button
                          onClick={() => deleteReceipt(r.id)}
                          aria-label="削除"
                          style={{
                            display: "flex", alignItems: "center", justifyContent: "center",
                            width: 36, height: 36, background: "none", border: "none", cursor: "pointer", color: "#B0A896",
                          }}
                        >
                          <Trash2 size={16} />
                        </button>
                      </div>
                    </div>
                  </>
                )}
              </div>
            );
          })}
        </div>

        <div style={{ marginTop: 14, fontSize: 11.5, color: COLORS.textMuted, textAlign: "center" }}>
          このデータはブラウザに保存されます。カテゴリはその場で、日付・店名・金額・メモは鉛筆アイコンから編集できます。
        </div>
        </>
        )}
      </div>
    </div>
  );
}

function TabGroup({ options, value, onChange }) {
  return (
    <div className="tab-group" style={{
      display: "flex", background: "#EFE8D6", borderRadius: 8, padding: 3, gap: 2,
    }}>
      {options.map(opt => {
        const active = opt.value === value;
        return (
          <button
            key={opt.value}
            className="tab-btn"
            onClick={() => onChange(opt.value)}
            style={{
              padding: "8px 14px",
              borderRadius: 6,
              border: "none",
              cursor: "pointer",
              fontSize: 13.5,
              fontWeight: 700,
              background: active ? COLORS.navy : "transparent",
              color: active ? "#fff" : COLORS.textMuted,
            }}
          >
            {opt.label}
          </button>
        );
      })}
    </div>
  );
}

function SummaryCard({ icon, label, value, color }) {
  return (
    <div style={{
      background: "#fff", border: `1px solid ${COLORS.border}`, borderRadius: 12,
      padding: "16px 18px", borderLeft: `4px solid ${color}`,
    }}>
      <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 8 }}>
        {icon}
        <span style={{ fontSize: 12.5, color: COLORS.textMuted, fontWeight: 500 }}>{label}</span>
      </div>
      <div className="mincho" style={{ fontSize: 24, fontWeight: 700, color: COLORS.navy }}>{value}</div>
    </div>
  );
}

function ChartCard({ title, children }) {
  return (
    <div style={{ background: "#fff", border: `1px solid ${COLORS.border}`, borderRadius: 12, padding: "16px 16px 6px" }}>
      <div style={{ fontSize: 13.5, fontWeight: 700, color: COLORS.navy, marginBottom: 6 }}>{title}</div>
      {children}
    </div>
  );
}

function Field({ label, children }) {
  return (
    <label style={{ display: "block" }}>
      <div style={{ fontSize: 11.5, color: COLORS.textMuted, marginBottom: 4, fontWeight: 500 }}>{label}</div>
      {children}
    </label>
  );
}
