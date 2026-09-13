import React, { useState, useEffect, useMemo, useRef, useCallback } from "react";
import {
  LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  PieChart, Pie, Cell, BarChart, Bar,
} from "recharts";
import * as XLSX from "xlsx";
import {
  Leaf, Sprout, BookOpen, Layers, Target, Quote, Link2, BarChart3, Upload,
  FileSpreadsheet, Volume2, Star, Search, ChevronRight, ChevronLeft, X, Check,
  RotateCcw, Flame, Trophy, PenLine, Filter, Download, AlertCircle, CheckCircle2,
  XCircle, Clock, TrendingUp, Award, RefreshCw, Home, Sparkles, Menu,
  LayoutDashboard, GraduationCap, Heart, ListChecks, Shuffle, ChevronDown, Play,
  CloudUpload, Zap, Trash2, Plus, LogOut, Mail, Lock, Loader2,
} from "lucide-react";
import { supabase, isSupabaseConfigured } from "./supabaseClient.js";

/* ============================== CONSTANTS ============================== */

const NAV = [
  { id: "dashboard", label: "Trang chủ", icon: LayoutDashboard },
  { id: "myvocab", label: "Từ vựng của tôi", icon: BookOpen },
  { id: "flashcards", label: "Thẻ ghi nhớ", icon: Layers },
  { id: "practice", label: "Luyện tập", icon: Target },
  { id: "idioms", label: "Thành ngữ", icon: Quote },
  { id: "collocations", label: "Cụm từ đi kèm", icon: Link2 },
  { id: "statistics", label: "Thống kê", icon: BarChart3 },
  { id: "import", label: "Nhập từ vựng", icon: Upload },
  { id: "template", label: "Mẫu Excel", icon: FileSpreadsheet },
];

const IMPORT_TYPES = {
  vocab: {
    itemType: "vocab",
    label: "Từ vựng",
    fileName: "VocaBloom_Vocabulary_Template.xlsx",
    columns: ["Vocabulary", "Part of Speech", "Vietnamese Meaning", "English Example"],
    required: ["Vocabulary", "Vietnamese Meaning"],
    colWidths: [18, 16, 36, 46],
    sample: [
      ["resilient", "adjective", "kiên cường, có khả năng phục hồi", "She is a resilient person."],
      ["achieve", "verb", "đạt được", "She worked hard to achieve her goals."],
      ["significant", "adjective", "đáng kể, quan trọng", "The project had a significant impact."],
      ["diligent", "adjective", "siêng năng, chăm chỉ", "He is diligent about his studies every day."],
    ],
    mapRow: (get) => ({ term: get("Vocabulary"), pos: get("Part of Speech") || "—", meaning: get("Vietnamese Meaning"), example: get("English Example") }),
  },
  idiom: {
    itemType: "idiom",
    label: "Thành ngữ",
    fileName: "VocaBloom_Idioms_Template.xlsx",
    columns: ["Idiom", "Category", "Vietnamese Meaning", "English Example"],
    required: ["Idiom", "Vietnamese Meaning"],
    colWidths: [22, 18, 34, 46],
    sample: [
      ["Break the ice", "Social", "phá tan bầu không khí ngượng ngùng", "He told a joke to break the ice at the meeting."],
      ["Piece of cake", "Ease", "việc dễ như ăn bánh", "Don't worry, the exam was a piece of cake."],
      ["Hit the books", "Study", "chăm chỉ học bài", "I need to hit the books tonight."],
    ],
    mapRow: (get) => ({ term: get("Idiom"), topic: get("Category") || "General", meaning: get("Vietnamese Meaning"), example: get("English Example") }),
  },
  collocation: {
    itemType: "collocation",
    label: "Cụm từ đi kèm",
    fileName: "VocaBloom_Collocations_Template.xlsx",
    columns: ["Collocation", "Vietnamese Meaning", "English Example", "Related Collocations"],
    required: ["Collocation", "Vietnamese Meaning"],
    colWidths: [24, 34, 46, 40],
    sample: [
      ["make a decision", "đưa ra quyết định", "She needs more time to make a decision.", "make a mistake, make progress, make an effort"],
      ["take a break", "nghỉ giải lao", "Let's take a break after this lesson.", "take a risk, take action, take responsibility"],
      ["pay attention", "chú ý", "Please pay attention during the lecture.", "pay a visit, pay a compliment, pay the price"],
    ],
    mapRow: (get) => ({ term: get("Collocation"), meaning: get("Vietnamese Meaning"), example: get("English Example"), related: (get("Related Collocations") || "").split(",").map((s) => s.trim()).filter(Boolean) }),
  },
};

function seedVocab() {
  const rows = [
    ["resilient", "adjective", "kiên cường, có khả năng phục hồi", "She is a resilient person who never gives up easily.", "/rɪˈzɪliənt/", "Personality", "medium"],
    ["achieve", "verb", "đạt được", "She worked hard to achieve her goals.", "/əˈtʃiːv/", "Success", "easy"],
    ["significant", "adjective", "đáng kể, quan trọng", "The project had a significant impact on sales.", "/sɪɡˈnɪfɪkənt/", "Business", "medium"],
    ["reluctant", "adjective", "miễn cưỡng, không sẵn lòng", "He was reluctant to accept the new proposal.", "/rɪˈlʌktənt/", "Emotions", "medium"],
    ["enhance", "verb", "nâng cao, cải thiện", "New features enhance the user experience.", "/ɪnˈhæns/", "Technology", "medium"],
    ["abundant", "adjective", "dồi dào, phong phú", "The region has abundant natural resources.", "/əˈbʌndənt/", "Nature", "hard"],
    ["diligent", "adjective", "siêng năng, chăm chỉ", "She is a diligent student who studies every day.", "/ˈdɪlɪdʒənt/", "Personality", "easy"],
    ["inevitable", "adjective", "không thể tránh khỏi", "Change is inevitable in a growing company.", "/ɪnˈevɪtəbl/", "Business", "hard"],
    ["genuine", "adjective", "chân thật, thật lòng", "He gave a genuine apology for the mistake.", "/ˈdʒenjuɪn/", "Emotions", "easy"],
    ["curious", "adjective", "tò mò", "Children are naturally curious about the world.", "/ˈkjʊəriəs/", "Personality", "easy"],
    ["flourish", "verb", "phát triển mạnh mẽ, nở rộ", "Small businesses flourish in a supportive economy.", "/ˈflʌrɪʃ/", "Business", "hard"],
    ["cautious", "adjective", "thận trọng, cẩn thận", "Drivers should be cautious in heavy rain.", "/ˈkɔːʃəs/", "Safety", "medium"],
    ["ambiguous", "adjective", "mơ hồ, không rõ ràng", "The instructions were ambiguous and confusing.", "/æmˈbɪɡjuəs/", "Communication", "hard"],
    ["commence", "verb", "bắt đầu, khởi đầu", "The ceremony will commence at nine o'clock.", "/kəˈmens/", "Formal", "medium"],
  ];
  return rows.map((r) => makeItem("vocab", {
    term: r[0], pos: r[1], meaning: r[2], example: r[3], ipa: r[4], topic: r[5], difficulty: r[6],
  }));
}

function seedIdioms() {
  const rows = [
    ["Break the ice", "phá tan bầu không khí ngượng ngùng, bắt đầu cuộc trò chuyện", "to break the physical ice", "Used to describe an action that relieves tension or awkwardness at the start of a social situation.", "He told a joke to break the ice at the beginning of the meeting.", "Social", "easy"],
    ["Hit the books", "chăm chỉ học bài", "to physically hit books", "Used when someone needs to study hard, especially before a test.", "I have a test tomorrow, so I need to hit the books tonight.", "Study", "easy"],
    ["Under the weather", "cảm thấy không khỏe", "below the weather", "Used to describe feeling slightly ill.", "I'm feeling a bit under the weather today, so I might stay home.", "Health", "easy"],
    ["Bite the bullet", "cắn răng chịu đựng, chấp nhận điều khó khăn", "to bite a bullet", "Used when someone decides to do something difficult they've been avoiding.", "I finally bit the bullet and booked the dentist appointment.", "Decisions", "medium"],
    ["Cost an arm and a leg", "rất đắt đỏ", "cost a person's arm and leg", "Used to say something is extremely expensive.", "That new laptop costs an arm and a leg.", "Money", "medium"],
    ["Piece of cake", "việc dễ như ăn bánh", "a slice of cake", "Used to describe a task that is very easy.", "Don't worry, the exam was a piece of cake.", "Ease", "easy"],
    ["Once in a blue moon", "hiếm khi, thỉnh thoảng mới", "only when the moon appears blue", "Used for something that rarely happens.", "We only eat out once in a blue moon.", "Frequency", "medium"],
    ["On the same page", "có cùng quan điểm, hiểu nhau", "reading the same page of a book", "Used when people agree or understand something the same way.", "Let's make sure we're on the same page before the meeting.", "Communication", "medium"],
  ];
  return rows.map((r) => makeItem("idiom", {
    term: r[0], meaning: r[1], literal: r[2], explanation: r[3], example: r[4], topic: r[5], difficulty: r[6],
  }));
}

function seedCollocations() {
  const rows = [
    ["make a decision", "đưa ra quyết định", "She needs more time to make a decision.", ["make a mistake", "make progress", "make an effort", "make a difference"], "Decisions", "easy"],
    ["take a break", "nghỉ giải lao", "Let's take a break after this lesson.", ["take a risk", "take action", "take responsibility", "take notice"], "Daily life", "easy"],
    ["do homework", "làm bài tập", "He always does his homework before dinner.", ["do business", "do research", "do damage", "do exercise"], "Study", "easy"],
    ["pay attention", "chú ý", "Please pay attention during the lecture.", ["pay a visit", "pay a compliment", "pay the price", "pay respect"], "Study", "medium"],
    ["keep in touch", "giữ liên lạc", "We promised to keep in touch after graduation.", ["keep a secret", "keep calm", "keep a promise", "keep track"], "Relationships", "medium"],
    ["save time", "tiết kiệm thời gian", "This shortcut will save time on the project.", ["save money", "save energy", "save face", "save a life"], "Efficiency", "easy"],
    ["catch a cold", "bị cảm lạnh", "I always catch a cold in winter.", ["catch a bus", "catch attention", "catch a glimpse", "catch fire"], "Health", "medium"],
    ["build confidence", "xây dựng sự tự tin", "Public speaking practice helps build confidence.", ["build a relationship", "build trust", "build a career", "build a reputation"], "Personal growth", "hard"],
  ];
  return rows.map((r) => makeItem("collocation", {
    term: r[0], meaning: r[1], example: r[2], related: r[3], topic: r[4], difficulty: r[5],
  }));
}

let idCounter = 1;
function makeItem(type, fields) {
  return {
    id: `${type}-${idCounter++}-${Math.random().toString(36).slice(2, 7)}`,
    type,
    term: "", pos: "", meaning: "", example: "", ipa: "",
    literal: "", explanation: "", related: [],
    topic: "General", difficulty: "medium", grade: null,
    favorite: false, notes: "",
    status: "new", box: 0, reviews: { correct: 0, incorrect: 0 },
    lastReviewedAt: null, createdAt: Date.now(),
    source: "seed", userSentence: "", sentenceCount: 0,
    ...fields,
  };
}

function withDemoProgress(vocab) {
  const set = (term, patch) => {
    const it = vocab.find((v) => v.term === term);
    if (it) Object.assign(it, patch);
  };
  set("curious", { status: "mastered", box: 4, grade: "easy", reviews: { correct: 6, incorrect: 0 }, lastReviewedAt: Date.now() - 1000 * 60 * 60 * 5 });
  set("genuine", { status: "mastered", box: 5, grade: "easy", reviews: { correct: 5, incorrect: 1 }, lastReviewedAt: Date.now() - 1000 * 60 * 60 * 26 });
  set("diligent", { status: "familiar", box: 2, grade: "good", reviews: { correct: 3, incorrect: 1 }, lastReviewedAt: Date.now() - 1000 * 60 * 60 * 30 });
  set("achieve", { status: "familiar", box: 2, grade: "good", reviews: { correct: 2, incorrect: 0 }, lastReviewedAt: Date.now() - 1000 * 60 * 60 * 50 });
  set("ambiguous", { status: "learning", box: 1, grade: "hard", reviews: { correct: 2, incorrect: 3 }, lastReviewedAt: Date.now() - 1000 * 60 * 60 * 10 });
  set("inevitable", { status: "learning", box: 1, grade: "hard", reviews: { correct: 1, incorrect: 2 }, lastReviewedAt: Date.now() - 1000 * 60 * 60 * 12 });
  set("flourish", { status: "learning", box: 0, grade: "again", reviews: { correct: 1, incorrect: 2 }, lastReviewedAt: Date.now() - 1000 * 60 * 60 * 20 });
  set("reluctant", { favorite: true });
  set("resilient", { favorite: true });
  return vocab;
}

function seedStats() {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  const dailyActivity = {};
  for (let i = 6; i >= 1; i--) {
    const dt = new Date();
    dt.setDate(dt.getDate() - i);
    dailyActivity[fmtDate(dt)] = 4 + Math.floor(Math.random() * 12);
  }
  return {
    todayGoal: 50,
    todayReviewed: 0,
    lastActiveDate: fmtDate(d),
    streak: 6,
    longestStreak: 9,
    totalReviews: 142,
    correct: 116,
    incorrect: 26,
    sentencesWritten: 3,
    dailyActivity,
  };
}

const BADGES = [
  { id: "first10", label: "10 từ đầu tiên", icon: Sprout, check: (items) => items.filter((i) => i.type === "vocab" && i.status !== "new").length >= 10 },
  { id: "words50", label: "Đã học 50 từ", icon: Leaf, check: (items) => items.filter((i) => i.type === "vocab" && i.status !== "new").length >= 50 },
  { id: "words100", label: "Đã học 100 từ", icon: Award, check: (items) => items.filter((i) => i.type === "vocab" && i.status !== "new").length >= 100 },
  { id: "streak7", label: "Chuỗi 7 ngày", icon: Flame, check: (items, stats) => stats.streak >= 7 },
  { id: "acc90", label: "Độ chính xác 90%", icon: Zap, check: (items, stats) => stats.correct + stats.incorrect >= 20 && stats.correct / (stats.correct + stats.incorrect) >= 0.9 },
  { id: "idiomMaster", label: "Cao thủ thành ngữ", icon: Quote, check: (items) => items.filter((i) => i.type === "idiom").length > 0 && items.filter((i) => i.type === "idiom" && i.status === "mastered").length >= 5 },
  { id: "collocMaster", label: "Cao thủ cụm từ đi kèm", icon: Link2, check: (items) => items.filter((i) => i.type === "collocation").length > 0 && items.filter((i) => i.type === "collocation" && i.status === "mastered").length >= 5 },
  { id: "sentenceSmith", label: "Bậc thầy đặt câu", icon: PenLine, check: (items, stats) => (stats.sentencesWritten || 0) >= 10 },
  { id: "selfStarter", label: "Tự thân vận động", icon: Sprout, check: (items) => items.filter((i) => i.source === "manual").length >= 10 },
];

const COLORS = {
  mint: "#7FB596",
  mintLight: "#BFE0CC",
  sage: "#5E7F5A",
  honey: "#E3B04B",
  coral: "#E0897A",
  ink: "#233B2C",
};

/* ============================== UTILITIES ============================== */

function fmtDate(d) {
  return d.toISOString().slice(0, 10);
}
function todayStr() { return fmtDate(new Date()); }
function shuffle(arr) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}
function normalize(s) {
  return String(s || "").trim().toLowerCase().replace(/\s+/g, " ");
}
function cx(...args) { return args.filter(Boolean).join(" "); }

function computeStatus(item) {
  const { correct = 0, incorrect = 0 } = item.reviews || {};
  const total = correct + incorrect;
  if (total === 0) return "new";
  const acc = correct / total;
  if ((item.box || 0) >= 4 && total >= 4 && acc >= 0.85) return "mastered";
  if (total >= 2 && acc >= 0.6) return "familiar";
  return "learning";
}

const STATUS_META = {
  new: { label: "Mới", color: "#8AA694", bg: "#EEF4EC" },
  learning: { label: "Đang học", color: "#C08A34", bg: "#FBF1DC" },
  familiar: { label: "Quen thuộc", color: "#3E7A6A", bg: "#E1F1EC" },
  mastered: { label: "Đã thuộc", color: "#2F6B3F", bg: "#DCF0DE" },
};
const STATUS_ORDER = ["new", "learning", "familiar", "mastered"];

function speak(text, accent = "US") {
  if (!("speechSynthesis" in window)) return;
  try {
    window.speechSynthesis.cancel();
    const utter = new SpeechSynthesisUtterance(text);
    const lang = accent === "US" ? "en-US" : "en-GB";
    const voices = window.speechSynthesis.getVoices();
    const voice = voices.find((v) => v.lang === lang) || voices.find((v) => v.lang && v.lang.startsWith("en"));
    if (voice) utter.voice = voice;
    utter.lang = lang;
    utter.rate = 0.92;
    utter.pitch = 1;
    window.speechSynthesis.speak(utter);
  } catch (e) { /* noop */ }
}

function typeLabel(t) {
  return t === "vocab" ? "Từ vựng" : t === "idiom" ? "Thành ngữ" : "Cụm từ đi kèm";
}
function typeIcon(t) {
  return t === "vocab" ? BookOpen : t === "idiom" ? Quote : Link2;
}

/* ============================== SMALL UI PARTS ============================== */

function ProgressBar({ value, max, colorVar = "var(--vb-mint-dk)", trackVar = "var(--vb-mint-pale)", height = 10 }) {
  const pct = max > 0 ? Math.min(100, Math.round((value / max) * 100)) : 0;
  return (
    <div className="w-full rounded-full overflow-hidden" style={{ height, background: trackVar }}>
      <div
        className="h-full rounded-full transition-all duration-700 ease-out"
        style={{ width: `${pct}%`, background: colorVar }}
      />
    </div>
  );
}

function PronounceButton({ text, size = "md" }) {
  const dim = size === "sm" ? "w-8 h-8" : "w-10 h-10";
  return (
    <div className="inline-flex items-center gap-1">
      <button
        onClick={(e) => { e.stopPropagation(); speak(text, "US"); }}
        title="Listen (US)"
        className={cx(dim, "rounded-full flex items-center justify-center vb-audio-btn")}
      >
        <Volume2 size={size === "sm" ? 15 : 17} />
      </button>
      <button
        onClick={(e) => { e.stopPropagation(); speak(text, "UK"); }}
        className="text-[11px] px-1.5 py-1 rounded-full vb-chip"
        title="Listen (UK)"
      >
        UK
      </button>
    </div>
  );
}

function FavoriteStar({ active, onClick, size = 18 }) {
  return (
    <button
      onClick={(e) => { e.stopPropagation(); onClick(); }}
      className="vb-star-btn"
      title={active ? "Bỏ yêu thích" : "Thêm vào yêu thích"}
    >
      <Star size={size} fill={active ? "#E3B04B" : "none"} color={active ? "#E3B04B" : "#8AA694"} />
    </button>
  );
}

function StatusPill({ status }) {
  const m = STATUS_META[status] || STATUS_META.new;
  return (
    <span className="text-xs font-medium px-2.5 py-1 rounded-full" style={{ color: m.color, background: m.bg }}>
      {m.label}
    </span>
  );
}

const GRADE_META = {
  again: { label: "Lại nào", color: "#B4563F", bg: "#F7E3DD" },
  hard: { label: "Khó nhằn", color: "#C08A34", bg: "#FBF1DC" },
  good: { label: "Ổn áp", color: "#3E7A6A", bg: "#E1F1EC" },
  easy: { label: "Dễ ợt", color: "#2F6B3F", bg: "#DCF0DE" },
};

function GradePill({ grade }) {
  const m = grade && GRADE_META[grade];
  if (!m) {
    return <span className="text-xs font-medium px-2.5 py-1 rounded-full" style={{ color: "#8AA694", background: "#EEF4EC" }}>Chưa chấm</span>;
  }
  return <span className="text-xs font-medium px-2.5 py-1 rounded-full" style={{ color: m.color, background: m.bg }}>{m.label}</span>;
}

function GradePicker({ value, onChange, size = "md" }) {
  const opts = [
    { id: null, label: "—" },
    { id: "again", label: "Lại nào" },
    { id: "hard", label: "Khó nhằn" },
    { id: "good", label: "Ổn áp" },
    { id: "easy", label: "Dễ ợt" },
  ];
  return (
    <select
      value={value || ""}
      onChange={(e) => onChange(e.target.value || null)}
      onClick={(e) => e.stopPropagation()}
      className={cx("vb-select", size === "sm" && "vb-select-sm")}
      title="Tự chấm điểm ôn tập cho từ này"
    >
      {opts.map((o) => <option key={o.label} value={o.id || ""}>{o.label === "—" ? "Chưa chấm" : o.label}</option>)}
    </select>
  );
}

function StatusPicker({ value, onChange, size = "md" }) {
  return (
    <select
      value={value || "new"}
      onChange={(e) => onChange(e.target.value)}
      onClick={(e) => e.stopPropagation()}
      className={cx("vb-select", size === "sm" && "vb-select-sm")}
      title="Tự chỉnh trạng thái học của từ này"
    >
      {STATUS_ORDER.map((s) => <option key={s} value={s}>{STATUS_META[s].label}</option>)}
    </select>
  );
}

function EmptyState({ icon: Icon, title, subtitle, action }) {
  return (
    <div className="flex flex-col items-center justify-center text-center py-16 px-6">
      <div className="w-16 h-16 rounded-full flex items-center justify-center mb-4" style={{ background: "var(--vb-mint-pale)" }}>
        <Icon size={26} color="var(--vb-sage-dk)" />
      </div>
      <h3 className="vb-heading text-lg mb-1" style={{ color: "var(--vb-ink)" }}>{title}</h3>
      <p className="text-sm max-w-sm" style={{ color: "var(--vb-muted)" }}>{subtitle}</p>
      {action}
    </div>
  );
}

function SectionHeader({ title, subtitle, right }) {
  return (
    <div className="flex flex-wrap items-end justify-between gap-4 mb-6">
      <div>
        <h1 className="vb-heading text-2xl sm:text-3xl" style={{ color: "var(--vb-ink)" }}>{title}</h1>
        {subtitle && <p className="text-sm mt-1" style={{ color: "var(--vb-muted)" }}>{subtitle}</p>}
      </div>
      {right}
    </div>
  );
}

function Toast({ toast }) {
  if (!toast) return null;
  const Icon = toast.kind === "error" ? XCircle : toast.kind === "info" ? AlertCircle : CheckCircle2;
  return (
    <div className="fixed bottom-5 right-5 z-50 vb-toast" key={toast.id}>
      <Icon size={18} color={toast.kind === "error" ? "#B4563F" : "#2F6B3F"} />
      <span>{toast.message}</span>
    </div>
  );
}

function ConfirmDialog({ message, confirmLabel, onConfirm, onCancel }) {
  return (
    <div className="fixed inset-0 flex items-center justify-center p-6" style={{ background: "rgba(30,45,35,0.45)", zIndex: 70 }} onClick={onCancel}>
      <div className="vb-card w-full max-w-sm rounded-2xl p-6" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-start gap-3 mb-6">
          <div className="w-10 h-10 rounded-full flex items-center justify-center shrink-0" style={{ background: "#F7E3DD" }}>
            <AlertCircle size={18} color="#B4563F" />
          </div>
          <p className="text-sm font-medium pt-1.5" style={{ color: "var(--vb-ink)" }}>{message}</p>
        </div>
        <div className="flex gap-2.5 justify-end">
          <button onClick={onCancel} className="vb-btn-secondary">Huỷ</button>
          <button onClick={onConfirm} className="vb-btn-primary" style={{ background: "#B4563F" }}>{confirmLabel || "Xác nhận"}</button>
        </div>
      </div>
    </div>
  );
}

/* ============================== APP ROOT ============================== */

export default function App() {
  const [items, setItems] = useState(() => withDemoProgress(seedVocab()).concat(seedIdioms(), seedCollocations()));
  const [stats, setStats] = useState(seedStats);
  const [activeTab, setActiveTab] = useState("dashboard");
  const [mobileNavOpen, setMobileNavOpen] = useState(false);
  const [detailItem, setDetailItem] = useState(null);
  const [toast, setToastState] = useState(null);
  const [quickPool, setQuickPool] = useState(null);
  const [flashSeed, setFlashSeed] = useState(null);
  const readyRef = useRef(false);
  const saveTimer = useRef(null);
  const [earnedBadges, setEarnedBadges] = useState([]);
  const [confirmState, setConfirmState] = useState(null); // { message, confirmLabel, onConfirm }
  const [session, setSession] = useState(null);
  const [authLoading, setAuthLoading] = useState(isSupabaseConfigured);
  const [dataLoading, setDataLoading] = useState(false);
  const [theme, setTheme] = useState(() => {
    try { return localStorage.getItem("vb-theme") || "mint"; } catch (e) { return "mint"; }
  });

  useEffect(() => {
    try { localStorage.setItem("vb-theme", theme); } catch (e) { /* noop */ }
  }, [theme]);

  const showToast = useCallback((message, kind = "success") => {
    setToastState({ message, kind, id: Date.now() });
  }, []);

  const requestConfirm = useCallback((message, onConfirm, confirmLabel = "Xoá") => {
    setConfirmState({ message, confirmLabel, onConfirm });
  }, []);

  useEffect(() => {
    if (!toast) return;
    const t = setTimeout(() => setToastState(null), 2800);
    return () => clearTimeout(t);
  }, [toast]);

  // Track login session (only relevant if Supabase env vars are configured)
  useEffect(() => {
    if (!isSupabaseConfigured) return;
    let mounted = true;
    supabase.auth.getSession().then(({ data }) => {
      if (!mounted) return;
      setSession(data.session || null);
      setAuthLoading(false);
    });
    const { data: listener } = supabase.auth.onAuthStateChange((_event, newSession) => {
      setSession(newSession);
    });
    return () => { mounted = false; listener.subscription.unsubscribe(); };
  }, []);

  // Load persisted data — from Supabase when logged in, otherwise from localStorage
  useEffect(() => {
    if (isSupabaseConfigured) {
      if (!session) { readyRef.current = false; return; }
      setDataLoading(true);
      let cancelled = false;
      (async () => {
        try {
          const { data } = await supabase
            .from("vocabloom_data")
            .select("payload")
            .eq("user_id", session.user.id)
            .maybeSingle();
          if (!cancelled && data && data.payload) {
            const parsed = data.payload;
            if (Array.isArray(parsed.items) && parsed.items.length) setItems(parsed.items);
            if (parsed.stats) setStats(parsed.stats);
            if (Array.isArray(parsed.earnedBadges)) setEarnedBadges(parsed.earnedBadges);
          }
        } catch (e) { /* first login, or offline — keep defaults */ }
        if (!cancelled) { readyRef.current = true; setDataLoading(false); }
      })();
      return () => { cancelled = true; };
    }
    // No login configured: plain browser storage, same as before
    try {
      const raw = localStorage.getItem("vocabloom-data");
      if (raw) {
        const parsed = JSON.parse(raw);
        if (Array.isArray(parsed.items) && parsed.items.length) setItems(parsed.items);
        if (parsed.stats) setStats(parsed.stats);
        if (Array.isArray(parsed.earnedBadges)) setEarnedBadges(parsed.earnedBadges);
      }
    } catch (e) { /* first run - keep seed data */ }
    readyRef.current = true;
  }, [session]);

  // Persist on change (debounced) — to Supabase when logged in, otherwise localStorage
  useEffect(() => {
    if (!readyRef.current) return;
    if (saveTimer.current) clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(() => {
      if (isSupabaseConfigured) {
        if (!session) return;
        supabase
          .from("vocabloom_data")
          .upsert({ user_id: session.user.id, payload: { items, stats, earnedBadges }, updated_at: new Date().toISOString() })
          .then(() => {})
          .catch(() => {});
        return;
      }
      try {
        localStorage.setItem("vocabloom-data", JSON.stringify({ items, stats, earnedBadges }));
      } catch (e) { /* storage full or unavailable */ }
    }, 700);
    return () => clearTimeout(saveTimer.current);
  }, [items, stats, earnedBadges, session]);

  // Badge detection
  useEffect(() => {
    BADGES.forEach((b) => {
      if (!earnedBadges.includes(b.id) && b.check(items, stats)) {
        setEarnedBadges((prev) => (prev.includes(b.id) ? prev : [...prev, b.id]));
        showToast(`🏆 Mở khoá thành tích: ${b.label}`, "success");
      }
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [items, stats]);

  const recordActivity = useCallback((correct) => {
    setStats((prev) => {
      const today = todayStr();
      let { streak, longestStreak, lastActiveDate, todayReviewed, dailyActivity, totalReviews, correct: c, incorrect: i } = prev;
      if (lastActiveDate !== today) {
        const y = new Date(); y.setDate(y.getDate() - 1);
        streak = lastActiveDate === fmtDate(y) ? streak + 1 : 1;
        longestStreak = Math.max(longestStreak, streak);
        todayReviewed = 0;
        lastActiveDate = today;
      }
      todayReviewed += 1;
      totalReviews += 1;
      if (correct) c += 1; else i += 1;
      dailyActivity = { ...dailyActivity, [today]: (dailyActivity[today] || 0) + 1 };
      return { ...prev, streak, longestStreak, lastActiveDate, todayReviewed, dailyActivity, totalReviews, correct: c, incorrect: i };
    });
  }, []);

  const gradeItem = useCallback((id, grade) => {
    setItems((prev) => prev.map((it) => {
      if (it.id !== id) return it;
      let box = it.box || 0;
      const reviews = { ...it.reviews };
      if (grade === "again") { box = 0; reviews.incorrect += 1; }
      else if (grade === "hard") { reviews.correct += 1; }
      else if (grade === "good") { box += 1; reviews.correct += 1; }
      else if (grade === "easy") { box += 2; reviews.correct += 1; }
      const status = computeStatus({ ...it, box, reviews });
      return { ...it, box, reviews, status, grade, lastReviewedAt: Date.now() };
    }));
    recordActivity(grade !== "again");
  }, [recordActivity]);

  const setItemGrade = useCallback((id, grade) => {
    setItems((prev) => prev.map((it) => (it.id === id ? { ...it, grade: grade || null } : it)));
  }, []);

  const setItemStatus = useCallback((id, status) => {
    setItems((prev) => prev.map((it) => (it.id === id ? { ...it, status } : it)));
  }, []);

  const deleteItems = useCallback((ids) => {
    const idSet = new Set(ids);
    setItems((prev) => prev.filter((it) => !idSet.has(it.id)));
  }, []);

  const toggleFavorite = useCallback((id) => {
    setItems((prev) => prev.map((it) => (it.id === id ? { ...it, favorite: !it.favorite } : it)));
  }, []);

  const updateNotes = useCallback((id, notes) => {
    setItems((prev) => prev.map((it) => (it.id === id ? { ...it, notes } : it)));
  }, []);

  const deleteItem = useCallback((id) => {
    setItems((prev) => prev.filter((it) => it.id !== id));
  }, []);

  const addManualItem = useCallback((type, fields) => {
    const item = makeItem(type, { ...fields, source: "manual" });
    setItems((prev) => [...prev, item]);
    showToast(`✅ Đã thêm "${item.term}" vào ${typeLabel(type).toLowerCase()}.`);
  }, [showToast]);

  const recordSentence = useCallback((id, sentence, ok) => {
    setItems((prev) => prev.map((it) => (it.id === id ? { ...it, userSentence: sentence, sentenceCount: (it.sentenceCount || 0) + (ok ? 1 : 0) } : it)));
    if (ok) setStats((prev) => ({ ...prev, sentencesWritten: (prev.sentencesWritten || 0) + 1 }));
  }, []);

  const goto = (tab) => { setActiveTab(tab); setMobileNavOpen(false); };

  const resetData = () => {
    requestConfirm(
      "Thao tác này sẽ xoá toàn bộ từ vựng, tiến độ và thống kê để bắt đầu lại từ đầu. Tiếp tục chứ?",
      () => {
        if (!isSupabaseConfigured) { try { localStorage.removeItem("vocabloom-data"); } catch (e) { /* noop */ } }
        setItems(withDemoProgress(seedVocab()).concat(seedIdioms(), seedCollocations()));
        setStats(seedStats());
        setEarnedBadges([]);
        showToast("Đã đặt lại dữ liệu của bạn.", "info");
      },
      "Đặt lại"
    );
  };

  const handleLogout = () => {
    requestConfirm(
      "Đăng xuất khỏi VocaBloom? Dữ liệu của bạn vẫn được lưu trong tài khoản cho lần sau.",
      async () => {
        readyRef.current = false;
        await supabase.auth.signOut();
        setItems(withDemoProgress(seedVocab()).concat(seedIdioms(), seedCollocations()));
        setStats(seedStats());
        setEarnedBadges([]);
      },
      "Đăng xuất"
    );
  };

  const startWeakPractice = () => {
    const weak = items.filter((i) => (i.reviews.correct + i.reviews.incorrect) > 0)
      .sort((a, b) => (b.reviews.incorrect / (b.reviews.correct + b.reviews.incorrect)) - (a.reviews.incorrect / (a.reviews.correct + a.reviews.incorrect)));
    if (weak.length === 0) { showToast("Chưa ghi nhận lỗi sai nào — luyện tập chút đã nhé!", "info"); return; }
    setQuickPool(weak.slice(0, 20).map((i) => i.id));
    goto("practice");
  };

  const stat = useMemo(() => {
    const vocab = items.filter((i) => i.type === "vocab");
    const idioms = items.filter((i) => i.type === "idiom");
    const collocations = items.filter((i) => i.type === "collocation");
    return {
      vocabTotal: vocab.length,
      idiomTotal: idioms.length,
      collocationTotal: collocations.length,
      vocabLearned: vocab.filter((v) => v.status !== "new").length,
      vocabLearning: vocab.filter((v) => v.status === "learning").length,
      vocabMastered: vocab.filter((v) => v.status === "mastered").length,
    };
  }, [items]);

  if (isSupabaseConfigured && authLoading) {
    return <LoadingScreen message="Đang tải VocaBloom..." />;
  }
  if (isSupabaseConfigured && !session) {
    return <AuthGate />;
  }
  if (isSupabaseConfigured && dataLoading) {
    return <LoadingScreen message="Đang tải kho từ vựng của bạn..." />;
  }

  return (
    <div className={cx("vb-root", `vb-theme-${theme}`, "min-h-screen w-full flex flex-col lg:flex-row")}>
      <style>{VB_STYLES}</style>

      {/* Mobile top bar */}
      <div className="lg:hidden flex items-center justify-between px-4 py-3 sticky top-0 z-40" style={{ background: "var(--vb-sage-dk)" }}>
        <div className="flex items-center gap-2 text-white">
          <Leaf size={20} />
          <span className="vb-heading text-lg">VocaBloom</span>
        </div>
        <button onClick={() => setMobileNavOpen((v) => !v)} className="text-white p-1">
          <Menu size={22} />
        </button>
      </div>

      <Sidebar active={activeTab} goto={goto} stats={stats} open={mobileNavOpen} resetData={resetData} userEmail={isSupabaseConfigured ? session?.user?.email : null} onLogout={isSupabaseConfigured ? handleLogout : null} />

      <main className="flex-1 min-w-0 px-4 py-6 sm:px-8 sm:py-8 lg:px-10 lg:py-9 max-w-6xl mx-auto w-full">
        {activeTab === "dashboard" && (
          <Dashboard items={items} stats={stats} stat={stat} goto={goto} openDetail={setDetailItem} startWeakPractice={startWeakPractice} />
        )}
        {activeTab === "myvocab" && (
          <MyVocabulary items={items} toggleFavorite={toggleFavorite} openDetail={setDetailItem} addManualItem={addManualItem} deleteItem={deleteItem} deleteItems={deleteItems} requestConfirm={requestConfirm} setItemGrade={setItemGrade} setItemStatus={setItemStatus} />
        )}
        {activeTab === "flashcards" && (
          <FlashcardsPage items={items} gradeItem={gradeItem} showToast={showToast} />
        )}
        {activeTab === "practice" && (
          <PracticeView items={items} gradeItem={gradeItem} recordSentence={recordSentence} quickPool={quickPool} clearQuickPool={() => setQuickPool(null)} goto={goto} />
        )}
        {activeTab === "idioms" && (
          <IdiomsCollocationsList type="idiom" items={items} toggleFavorite={toggleFavorite} openDetail={setDetailItem} showToast={showToast} addManualItem={addManualItem} deleteItems={deleteItems} requestConfirm={requestConfirm} />
        )}
        {activeTab === "collocations" && (
          <IdiomsCollocationsList type="collocation" items={items} toggleFavorite={toggleFavorite} openDetail={setDetailItem} showToast={showToast} addManualItem={addManualItem} deleteItems={deleteItems} requestConfirm={requestConfirm} />
        )}
        {activeTab === "statistics" && (
          <StatisticsView items={items} stats={stats} earnedBadges={earnedBadges} startWeakPractice={startWeakPractice} />
        )}
        {activeTab === "import" && (
          <ImportView items={items} setItems={setItems} showToast={showToast} goto={goto} />
        )}
        {activeTab === "template" && <TemplateView />}
      </main>

      {detailItem && (
        <ItemDetailModal
          item={items.find((i) => i.id === detailItem.id) || detailItem}
          onClose={() => setDetailItem(null)}
          toggleFavorite={toggleFavorite}
          updateNotes={updateNotes}
          onDelete={deleteItem}
          requestConfirm={requestConfirm}
          setItemGrade={setItemGrade}
          setItemStatus={setItemStatus}
        />
      )}

      {confirmState && (
        <ConfirmDialog
          message={confirmState.message}
          confirmLabel={confirmState.confirmLabel}
          onCancel={() => setConfirmState(null)}
          onConfirm={() => { confirmState.onConfirm(); setConfirmState(null); }}
        />
      )}

      <Toast toast={toast} />
      <ThemeSwitcher theme={theme} setTheme={setTheme} />
    </div>
  );
}

/* ============================== THEME SWITCHER ============================== */

const THEME_OPTIONS = [
  { id: "mint", label: "Bạc hà", swatch: "#7FB596" },
  { id: "pink", label: "Hồng", swatch: "#F4A6C1" },
  { id: "sky", label: "Xanh da trời", swatch: "#8EC5F0" },
  { id: "butter", label: "Vàng bơ", swatch: "#F6D370" },
  { id: "lilac", label: "Tím", swatch: "#C6A6E8" },
];

function ThemeSwitcher({ theme, setTheme }) {
  const [open, setOpen] = useState(false);
  return (
    <div className="fixed bottom-5 left-5 z-50 flex flex-col items-start gap-2">
      {open && (
        <div className="vb-card vb-sticker rounded-3xl p-3 flex flex-col gap-1.5 mb-1">
          <div className="text-[11px] font-semibold px-1.5 pb-0.5" style={{ color: "var(--vb-muted)" }}>Chọn màu giao diện 🎨</div>
          {THEME_OPTIONS.map((t) => (
            <button
              key={t.id}
              onClick={() => { setTheme(t.id); setOpen(false); }}
              className="flex items-center gap-2.5 px-2.5 py-1.5 rounded-2xl text-xs font-semibold text-left"
              style={{ background: theme === t.id ? "var(--vb-mint-pale)" : "transparent", color: "var(--vb-ink)" }}
            >
              <span className="w-5 h-5 rounded-full inline-block shrink-0" style={{ background: t.swatch, border: theme === t.id ? "2px solid var(--vb-sage-dk)" : "2px solid white", boxShadow: "0 0 0 1px rgba(0,0,0,0.08)" }} />
              {t.label}
              {theme === t.id && <Check size={13} className="ml-auto" />}
            </button>
          ))}
        </div>
      )}
      <button
        onClick={() => setOpen((v) => !v)}
        className="vb-theme-fab"
        title="Đổi màu giao diện"
      >
        <Sparkles size={18} />
      </button>
    </div>
  );
}

/* ============================== SIDEBAR ============================== */

function Sidebar({ active, goto, stats, open, resetData, userEmail, onLogout }) {
  return (
    <aside className={cx(
      "lg:w-64 lg:flex lg:flex-col lg:h-screen lg:sticky lg:top-0 shrink-0",
      open ? "flex flex-col" : "hidden lg:flex"
    )} style={{ background: "var(--vb-sage-dk)" }}>
      <div className="hidden lg:flex items-center gap-2 px-6 pt-7 pb-5">
        <div className="w-9 h-9 rounded-xl flex items-center justify-center" style={{ background: "var(--vb-mint)" }}>
          <Leaf size={19} color="var(--vb-sage-dk)" />
        </div>
        <div>
          <div className="vb-heading text-xl leading-none text-white">VocaBloom</div>
          <div className="text-[11px] tracking-wide" style={{ color: "var(--vb-mint-pale)" }}>Học tiếng Anh, nhẹ nhàng mà chắc</div>
        </div>
      </div>

      {userEmail && (
        <div className="hidden lg:flex items-center gap-2 mx-3 mb-3 px-3.5 py-2.5 rounded-xl" style={{ background: "rgba(255,255,255,0.08)" }}>
          <div className="w-7 h-7 rounded-full flex items-center justify-center shrink-0 text-xs font-semibold" style={{ background: "var(--vb-mint)", color: "var(--vb-sage-dk)" }}>
            {userEmail[0].toUpperCase()}
          </div>
          <span className="text-xs truncate" style={{ color: "var(--vb-mint-pale)" }}>{userEmail}</span>
        </div>
      )}

      <nav className="flex-1 px-3 py-2 flex flex-col gap-1 overflow-y-auto">
        {NAV.map((n) => {
          const Icon = n.icon;
          const isActive = active === n.id;
          return (
            <button
              key={n.id}
              onClick={() => goto(n.id)}
              className="flex items-center gap-3 px-3.5 py-2.5 rounded-xl text-sm text-left transition-colors"
              style={{
                background: isActive ? "rgba(255,255,255,0.14)" : "transparent",
                color: isActive ? "#FFFFFF" : "var(--vb-mint-pale)",
                fontWeight: isActive ? 600 : 500,
              }}
            >
              <Icon size={17} />
              {n.label}
            </button>
          );
        })}
      </nav>

      <div className="px-4 py-4 mx-3 mb-4 rounded-2xl flex items-center gap-3" style={{ background: "rgba(255,255,255,0.1)" }}>
        <Flame size={18} color="#F4C669" />
        <div className="text-white">
          <div className="text-sm font-semibold leading-none">Chuỗi {stats.streak} ngày</div>
          <div className="text-[11px] mt-1" style={{ color: "var(--vb-mint-pale)" }}>Giữ vững nha 🌱</div>
        </div>
      </div>
      <div className="hidden lg:flex items-center justify-between px-4 pb-5">
        <button onClick={resetData} className="text-[11px] text-left" style={{ color: "rgba(255,255,255,0.55)" }}>
          Đặt lại tiến độ
        </button>
        {onLogout && (
          <button onClick={onLogout} className="text-[11px] flex items-center gap-1" style={{ color: "rgba(255,255,255,0.75)" }}>
            <LogOut size={12} /> Đăng xuất
          </button>
        )}
      </div>
    </aside>
  );
}

/* ============================== AUTH (LOGIN / SIGN UP) ============================== */

function LoadingScreen({ message }) {
  return (
    <div className="vb-root min-h-screen w-full flex items-center justify-center">
      <style>{VB_STYLES}</style>
      <div className="flex flex-col items-center gap-3">
        <Loader2 size={28} color="var(--vb-sage-dk)" className="vb-spin" />
        <p className="text-sm" style={{ color: "var(--vb-muted)" }}>{message}</p>
      </div>
    </div>
  );
}

function AuthGate() {
  const [mode, setMode] = useState("signin"); // signin | signup
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [notice, setNotice] = useState("");

  const submit = async (e) => {
    e.preventDefault();
    if (!email.trim() || !password) return;
    setLoading(true);
    setError("");
    setNotice("");
    try {
      if (mode === "signup") {
        const { data, error: err } = await supabase.auth.signUp({ email: email.trim(), password });
        if (err) throw err;
        if (data.session) {
          // Email confirmation is off in the Supabase project — logged in immediately.
        } else {
          setNotice("Đã tạo tài khoản! Kiểm tra email để xác nhận, rồi đăng nhập nhé.");
          setMode("signin");
        }
      } else {
        const { error: err } = await supabase.auth.signInWithPassword({ email: email.trim(), password });
        if (err) throw err;
      }
    } catch (err) {
      setError(err && err.message ? err.message : "Có lỗi xảy ra rồi. Thử lại nhé.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="vb-root min-h-screen w-full flex items-center justify-center p-6">
      <style>{VB_STYLES}</style>
      <div className="vb-card w-full max-w-sm rounded-3xl p-8">
        <div className="flex items-center gap-2 mb-1">
          <div className="w-9 h-9 rounded-xl flex items-center justify-center" style={{ background: "var(--vb-mint-pale)" }}>
            <Leaf size={18} color="var(--vb-sage-dk)" />
          </div>
          <span className="vb-heading text-xl" style={{ color: "var(--vb-ink)" }}>VocaBloom</span>
        </div>
        <p className="text-sm mb-6" style={{ color: "var(--vb-muted)" }}>
          {mode === "signin" ? "Chào mừng quay lại — đăng nhập để vào kho từ vựng của bạn." : "Tạo tài khoản để lưu tiến độ mọi lúc mọi nơi."}
        </p>

        <form onSubmit={submit}>
          <label className="block text-xs font-semibold uppercase tracking-wide mb-1.5" style={{ color: "var(--vb-muted)" }}>Email</label>
          <div className="relative mb-4">
            <Mail size={15} className="absolute left-3 top-1/2 -translate-y-1/2" color="var(--vb-muted)" />
            <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} required className="vb-input w-full pl-9" placeholder="ban@example.com" autoFocus />
          </div>

          <label className="block text-xs font-semibold uppercase tracking-wide mb-1.5" style={{ color: "var(--vb-muted)" }}>Mật khẩu</label>
          <div className="relative mb-5">
            <Lock size={15} className="absolute left-3 top-1/2 -translate-y-1/2" color="var(--vb-muted)" />
            <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} required minLength={6} className="vb-input w-full pl-9" placeholder="Tối thiểu 6 ký tự" />
          </div>

          {error && <p className="text-xs mb-4" style={{ color: "#B4563F" }}>{error}</p>}
          {notice && <p className="text-xs mb-4" style={{ color: "#3E7A6A" }}>{notice}</p>}

          <button type="submit" disabled={loading} className="vb-btn-primary w-full justify-center disabled:opacity-50">
            {loading ? <Loader2 size={15} className="vb-spin" /> : mode === "signin" ? "Đăng nhập" : "Tạo tài khoản"}
          </button>
        </form>

        <p className="text-xs text-center mt-5" style={{ color: "var(--vb-muted)" }}>
          {mode === "signin" ? "Mới đến đây à?" : "Đã có tài khoản rồi?"}{" "}
          <button onClick={() => { setMode(mode === "signin" ? "signup" : "signin"); setError(""); setNotice(""); }} className="font-semibold" style={{ color: "var(--vb-sage-dk)" }}>
            {mode === "signin" ? "Tạo tài khoản" : "Đăng nhập"}
          </button>
        </p>
      </div>
    </div>
  );
}

/* ============================== DASHBOARD ============================== */

const VIBE_QUOTES = [
  "Hôm nay không học thì mai chữ nó bay hết á, đừng có mà cứng đầu 🌱",
  "Đời không cho mình tất cả, nhưng cho mình cái Duolingo miễn phí thì phải học chứ 😤",
  "Không ai đánh thuế ước mơ, nhưng lười thì tự nó phạt mình đó nha 🐌",
  "Học từ mới mỗi ngày, không giàu bằng ai nhưng mà lú tiếng Anh thì đỡ hơn 💅",
  "Cuộc đời là những chuyến đi, còn từ vựng là những chuyến... quên 😭 học lại đi bạn ơi",
  "Chill quá cũng được, nhưng đừng để chữ nó chill mất luôn nha 🧊",
  "Không sao nếu hôm qua lười, hôm nay đứng dậy học tiếp là được, đời mà 🌸",
  "Từ vựng giống người yêu cũ: không ôn là quên liền á 💔",
  "Một ngày không học coi như tặng không kiến thức cho đứa khác 😮‍💨",
  "Cày từ vựng như cày phim bộ thì chắc giỏi tiếng Anh lâu rồi 🍿",
];

const ROAST_QUOTES = [
  "Ê, cái nút Flashcards để làm cảnh hả? Bấm vô học đi 👀",
  "132 từ nằm im re trong kho, tụi nó buồn lắm đó nha 🥲",
  "Học ít thôi mà lười cũng nhiều, vậy chữ nào cũng không chịu ở lại đâu 😂",
  "Streak mà đứt là tự nhiên thấy cuộc đời vô nghĩa liền á, giữ lấy đi!",
  "Đừng nói \"để mai học\", vì \"mai\" của bạn nó dài như deadline vậy đó 📅",
  "Chưa graded từ nào hết vậy? Để mồ côi từ vựng luôn hả 😩",
  "Học 5 phút thôi cũng được, hơn là ngồi đây đọc mấy câu mắng vui của tui nè 😜",
  "Cái nút \"Ôn lại lỗi sai\" nó gọi tên bạn suốt đó, nghe chưa? 📢",
];

function pickDaily(list) {
  const key = todayStr();
  let hash = 0;
  for (let i = 0; i < key.length; i++) hash = (hash * 31 + key.charCodeAt(i)) >>> 0;
  return list[hash % list.length];
}

function Dashboard({ items, stats, stat, goto, openDetail, startWeakPractice }) {
  const recent = items.filter((i) => i.lastReviewedAt).sort((a, b) => b.lastReviewedAt - a.lastReviewedAt).slice(0, 5);
  const vibe = useMemo(() => pickDaily(VIBE_QUOTES), []);
  const roast = useMemo(() => {
    const key = todayStr() + "-roast";
    let hash = 0;
    for (let i = 0; i < key.length; i++) hash = (hash * 31 + key.charCodeAt(i)) >>> 0;
    return ROAST_QUOTES[hash % ROAST_QUOTES.length];
  }, []);

  return (
    <div>
      <div className="vb-hero rounded-3xl px-6 py-8 sm:px-10 sm:py-10 mb-8 relative overflow-hidden">
        <svg className="absolute -right-10 -top-16 opacity-70" width="260" height="260" viewBox="0 0 200 200">
          <path fill="#8FCBA6" d="M45.7,-58.3C58.4,-49.7,67.5,-34.6,71.6,-18.1C75.6,-1.5,74.5,16.6,66.5,30.8C58.5,45,43.6,55.4,27.2,62.1C10.8,68.8,-7.2,71.9,-23.1,67.3C-39,62.7,-52.9,50.4,-61.6,35.1C-70.3,19.7,-73.8,1.2,-70.2,-15.5C-66.6,-32.1,-56,-46.9,-42.1,-55.6C-28.2,-64.2,-14.1,-66.7,1.6,-68.7C17.3,-70.7,34.9,-67,45.7,-58.3Z" transform="translate(100 100)" />
        </svg>
        <div className="relative">
          <p className="vb-heading text-2xl sm:text-3xl" style={{ color: "var(--vb-ink)" }}>{vibe} 🌿</p>
          <p className="text-sm sm:text-base mt-2 max-w-md font-semibold" style={{ color: "var(--vb-sage-dk)" }}>🙃 {roast}</p>

          <div className="flex flex-wrap gap-3 mt-6">
            <HeroStat icon={BookOpen} value={stat.vocabTotal} label="từ" />
            <HeroStat icon={Quote} value={stat.idiomTotal} label="thành ngữ" />
            <HeroStat icon={Link2} value={stat.collocationTotal} label="cụm từ" />
            <HeroStat icon={Flame} value={stats.streak} label="ngày liên tiếp" accent />
          </div>

          <div className="mt-7 max-w-md">
            <div className="flex items-center justify-between mb-1.5 text-sm" style={{ color: "var(--vb-sage-dk)" }}>
              <span className="font-semibold">Mục tiêu hôm nay</span>
              <span>{stats.todayReviewed} / {stats.todayGoal} từ đã ôn</span>
            </div>
            <ProgressBar value={stats.todayReviewed} max={stats.todayGoal} />
          </div>

          <div className="flex flex-wrap gap-2.5 mt-6">
            <button onClick={() => goto("flashcards")} className="vb-btn-primary"><Layers size={15} /> Học thẻ ghi nhớ</button>
            <button onClick={() => goto("practice")} className="vb-btn-secondary"><Target size={15} /> Luyện từ vựng</button>
            <button onClick={startWeakPractice} className="vb-btn-secondary"><RotateCcw size={15} /> Ôn lại lỗi sai</button>
            <button onClick={() => goto("import")} className="vb-btn-secondary"><Upload size={15} /> Nhập từ mới</button>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 sm:gap-4 mb-8">
        <MiniStat label="Đã học" value={stat.vocabLearned} icon={Sprout} />
        <MiniStat label="Đang học" value={stat.vocabLearning} icon={PenLine} />
        <MiniStat label="Đã thuộc" value={stat.vocabMastered} icon={Trophy} />
        <MiniStat label="Độ chính xác" value={`${stats.correct + stats.incorrect > 0 ? Math.round((stats.correct / (stats.correct + stats.incorrect)) * 100) : 0}%`} icon={Target} />
      </div>

      <div className="vb-card rounded-3xl p-6">
        <div className="flex items-center justify-between mb-4">
          <h3 className="vb-heading text-lg" style={{ color: "var(--vb-ink)" }}>Học gần đây</h3>
          <button onClick={() => goto("myvocab")} className="text-sm font-medium flex items-center gap-1" style={{ color: "var(--vb-sage-dk)" }}>
            Xem tất cả <ChevronRight size={15} />
          </button>
        </div>
        {recent.length === 0 ? (
          <EmptyState icon={Sprout} title="Chưa học gì cả" subtitle="Bắt đầu một buổi thẻ ghi nhớ để xem tiến độ gần đây ở đây." />
        ) : (
          <div className="divide-y" style={{ borderColor: "var(--vb-border)" }}>
            {recent.map((it) => {
              const Icon = typeIcon(it.type);
              return (
                <button key={it.id} onClick={() => openDetail(it)} className="w-full flex items-center gap-3 py-3 text-left">
                  <div className="w-9 h-9 rounded-xl flex items-center justify-center shrink-0" style={{ background: "var(--vb-mint-pale)" }}>
                    <Icon size={15} color="var(--vb-sage-dk)" />
                  </div>
                  <div className="min-w-0 flex-1">
                    <div className="text-sm font-medium truncate" style={{ color: "var(--vb-ink)" }}>{it.term}</div>
                    <div className="text-xs truncate" style={{ color: "var(--vb-muted)" }}>{it.meaning}</div>
                  </div>
                  <StatusPill status={it.status} />
                </button>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}

function HeroStat({ icon: Icon, value, label, accent }) {
  return (
    <div className="vb-hero-chip">
      <Icon size={15} color={accent ? "#C97A2B" : "var(--vb-sage-dk)"} />
      <span className="font-semibold" style={{ color: "var(--vb-ink)" }}>{value}</span>
      <span style={{ color: "var(--vb-muted)" }}>{label}</span>
    </div>
  );
}

function MiniStat({ label, value, icon: Icon }) {
  return (
    <div className="vb-card rounded-2xl p-4 flex items-center gap-3">
      <div className="w-10 h-10 rounded-xl flex items-center justify-center shrink-0" style={{ background: "var(--vb-mint-pale)" }}>
        <Icon size={16} color="var(--vb-sage-dk)" />
      </div>
      <div className="min-w-0">
        <div className="text-lg font-semibold leading-none" style={{ color: "var(--vb-ink)" }}>{value}</div>
        <div className="text-xs mt-1 truncate" style={{ color: "var(--vb-muted)" }}>{label}</div>
      </div>
    </div>
  );
}

/* ============================== MY VOCABULARY ============================== */

function MyVocabulary({ items, toggleFavorite, openDetail, addManualItem, deleteItem, deleteItems, requestConfirm, setItemGrade, setItemStatus }) {
  const [q, setQ] = useState("");
  const [pos, setPos] = useState("all");
  const [gradeFilter, setGradeFilter] = useState("all");
  const [topic, setTopic] = useState("all");
  const [status, setStatus] = useState("all");
  const [sort, setSort] = useState("alpha");
  const [favOnly, setFavOnly] = useState(false);
  const [showAdd, setShowAdd] = useState(false);
  const [selectMode, setSelectMode] = useState(false);
  const [selected, setSelected] = useState([]);

  const vocab = items.filter((i) => i.type === "vocab");
  const topics = Array.from(new Set(vocab.map((v) => v.topic))).sort();

  let filtered = vocab.filter((v) => {
    if (q && !(`${v.term} ${v.meaning}`.toLowerCase().includes(q.toLowerCase()))) return false;
    if (pos !== "all" && v.pos !== pos) return false;
    if (gradeFilter !== "all") {
      if (gradeFilter === "none" ? v.grade : v.grade !== gradeFilter) return false;
    }
    if (topic !== "all" && v.topic !== topic) return false;
    if (status !== "all" && v.status !== status) return false;
    if (favOnly && !v.favorite) return false;
    return true;
  });
  filtered = filtered.sort((a, b) => sort === "alpha" ? a.term.localeCompare(b.term) : b.createdAt - a.createdAt);

  const handleDelete = (v) => {
    requestConfirm(`Xoá "${v.term}" khỏi kho từ vựng? Không thể hoàn tác đâu nhé.`, () => deleteItem(v.id));
  };

  const toggleSelected = (id) => setSelected((prev) => prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]);
  const toggleSelectAll = () => setSelected((prev) => prev.length === filtered.length ? [] : filtered.map((v) => v.id));
  const exitSelectMode = () => { setSelectMode(false); setSelected([]); };
  const handleBulkDelete = () => {
    requestConfirm(`Xoá ${selected.length} từ đã chọn khỏi kho từ vựng? Không thể hoàn tác đâu nhé.`, () => { deleteItems(selected); exitSelectMode(); }, "Xoá đã chọn");
  };

  return (
    <div>
      <SectionHeader title="Từ vựng của tôi" subtitle={`${vocab.length} từ trong kho của bạn`} right={
        <div className="flex items-center gap-2">
          <button onClick={() => (selectMode ? exitSelectMode() : setSelectMode(true))} className={cx("vb-btn-secondary text-sm", selectMode && "vb-chip-toggle-active")}>
            <ListChecks size={15} /> {selectMode ? "Xong" : "Chọn nhiều"}
          </button>
          <button onClick={() => setShowAdd(true)} className="vb-btn-primary text-sm"><Plus size={15} /> Thêm từ thủ công</button>
        </div>
      } />

      {selectMode && (
        <div className="vb-bulk-bar mb-4 flex-wrap">
          <label className="flex items-center gap-2 text-sm font-semibold cursor-pointer" style={{ color: "var(--vb-sage-dk)" }}>
            <input type="checkbox" className="vb-checkbox" checked={selected.length > 0 && selected.length === filtered.length} onChange={toggleSelectAll} />
            Chọn tất cả ({filtered.length})
          </label>
          <span className="text-sm" style={{ color: "var(--vb-sage-dk)" }}>Đã chọn {selected.length} từ</span>
          <button onClick={handleBulkDelete} disabled={selected.length === 0} className="vb-btn-primary text-sm ml-auto disabled:opacity-40" style={{ background: "#B4563F", borderColor: "#B4563F" }}>
            <Trash2 size={14} /> Xoá đã chọn
          </button>
        </div>
      )}

      <div className="vb-card rounded-2xl p-4 mb-5 flex flex-wrap gap-3 items-center">
        <div className="relative flex-1 min-w-[180px]">
          <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2" color="var(--vb-muted)" />
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Tìm từ hoặc nghĩa..." className="vb-input pl-9" />
        </div>
        <select className="vb-select" value={pos} onChange={(e) => setPos(e.target.value)}>
          <option value="all">Mọi từ loại</option>
          <option value="noun">Danh từ</option>
          <option value="verb">Động từ</option>
          <option value="adjective">Tính từ</option>
          <option value="adverb">Trạng từ</option>
        </select>
        <select className="vb-select" value={gradeFilter} onChange={(e) => setGradeFilter(e.target.value)}>
          <option value="all">Mọi mức điểm</option>
          <option value="none">Chưa chấm</option>
          <option value="again">Lại nào</option>
          <option value="hard">Khó nhằn</option>
          <option value="good">Ổn áp</option>
          <option value="easy">Dễ ợt</option>
        </select>
        <select className="vb-select" value={topic} onChange={(e) => setTopic(e.target.value)}>
          <option value="all">Mọi chủ đề</option>
          {topics.map((t) => <option key={t} value={t}>{t}</option>)}
        </select>
        <select className="vb-select" value={status} onChange={(e) => setStatus(e.target.value)}>
          <option value="all">Mọi trạng thái</option>
          {STATUS_ORDER.map((s) => <option key={s} value={s}>{STATUS_META[s].label}</option>)}
        </select>
        <select className="vb-select" value={sort} onChange={(e) => setSort(e.target.value)}>
          <option value="alpha">Sắp xếp: A–Z</option>
          <option value="recent">Sắp xếp: Mới thêm gần đây</option>
        </select>
        <button onClick={() => setFavOnly((v) => !v)} className={cx("vb-chip-toggle", favOnly && "vb-chip-toggle-active")}>
          <Star size={13} fill={favOnly ? "#E3B04B" : "none"} color={favOnly ? "#E3B04B" : "var(--vb-sage-dk)"} /> Yêu thích
        </button>
      </div>

      {filtered.length === 0 ? (
        <EmptyState icon={Search} title="Không có từ nào khớp bộ lọc" subtitle="Thử bỏ bớt bộ lọc, hoặc thêm từ đầu tiên của bạn." action={
          <button onClick={() => setShowAdd(true)} className="vb-btn-primary mt-4"><Plus size={15} /> Thêm từ thủ công</button>
        } />
      ) : (
        <div className="vb-card rounded-2xl overflow-hidden">
          <div className={cx("hidden sm:grid gap-3 px-5 py-3 text-xs font-semibold uppercase tracking-wide", selectMode ? "grid-cols-[auto_1.3fr_1.8fr_0.8fr_0.9fr_0.9fr_auto_auto_auto]" : "grid-cols-[1.3fr_1.8fr_0.8fr_0.9fr_0.9fr_auto_auto_auto]")} style={{ color: "var(--vb-muted)", borderBottom: "1px solid var(--vb-border)" }}>
            {selectMode && <span></span>}
            <span>Từ</span><span>Nghĩa</span><span>Từ loại</span><span>Trạng thái</span><span>Điểm</span><span>Âm thanh</span><span>Yêu thích</span><span></span>
          </div>
          <div className="divide-y" style={{ borderColor: "var(--vb-border)" }}>
            {filtered.map((v) => (
              <div key={v.id} onClick={() => (selectMode ? toggleSelected(v.id) : openDetail(v))} className={cx("grid grid-cols-2 gap-3 px-5 py-3.5 items-center cursor-pointer vb-row-hover", selectMode ? "sm:grid-cols-[auto_1.3fr_1.8fr_0.8fr_0.9fr_0.9fr_auto_auto_auto]" : "sm:grid-cols-[1.3fr_1.8fr_0.8fr_0.9fr_0.9fr_auto_auto_auto]")}>
                {selectMode && <input type="checkbox" className="vb-checkbox hidden sm:inline-block" checked={selected.includes(v.id)} onChange={() => toggleSelected(v.id)} onClick={(e) => e.stopPropagation()} />}
                <div className="font-medium" style={{ color: "var(--vb-ink)" }}>{v.term}</div>
                <div className="text-sm truncate" style={{ color: "var(--vb-muted)" }}>{v.meaning}</div>
                <div className="hidden sm:block text-sm capitalize" style={{ color: "var(--vb-muted)" }}>{v.pos}</div>
                <div className="hidden sm:block" onClick={(e) => e.stopPropagation()}>
                  <StatusPicker value={v.status} onChange={(s) => setItemStatus(v.id, s)} size="sm" />
                </div>
                <div className="hidden sm:block" onClick={(e) => e.stopPropagation()}>
                  <GradePicker value={v.grade} onChange={(g) => setItemGrade(v.id, g)} size="sm" />
                </div>
                <div onClick={(e) => e.stopPropagation()}><PronounceButton text={v.term} size="sm" /></div>
                <FavoriteStar active={v.favorite} onClick={() => toggleFavorite(v.id)} />
                <button onClick={(e) => { e.stopPropagation(); handleDelete(v); }} className="vb-icon-btn" title="Xoá từ">
                  <Trash2 size={15} />
                </button>
              </div>
            ))}
          </div>
        </div>
      )}

      {showAdd && <AddItemModal initialType="vocab" onAdd={addManualItem} onClose={() => setShowAdd(false)} />}
    </div>
  );
}

const ADD_ITEM_TYPES = [
  { id: "vocab", label: "Từ vựng" },
  { id: "idiom", label: "Thành ngữ" },
  { id: "collocation", label: "Cụm từ đi kèm" },
];

function AddItemModal({ initialType = "vocab", onAdd, onClose }) {
  const [type, setType] = useState(initialType);
  const [term, setTerm] = useState("");
  const [pos, setPos] = useState("noun");
  const [meaning, setMeaning] = useState("");
  const [example, setExample] = useState("");
  const [literal, setLiteral] = useState("");
  const [explanation, setExplanation] = useState("");
  const [related, setRelated] = useState("");
  const [topic, setTopic] = useState("Từ của tôi");

  const canSubmit = term.trim() && meaning.trim();
  const label = type === "vocab" ? "từ/cụm từ" : type === "idiom" ? "thành ngữ" : "cụm từ đi kèm";

  const submit = () => {
    if (!canSubmit) return;
    const fields = {
      term: term.trim(),
      meaning: meaning.trim(),
      example: example.trim() || `I'm learning to use "${term.trim()}".`,
      topic: topic.trim() || "Từ của tôi",
    };
    if (type === "vocab") fields.pos = pos;
    if (type === "idiom") { fields.literal = literal.trim(); fields.explanation = explanation.trim(); }
    if (type === "collocation") fields.related = related.split(",").map((s) => s.trim()).filter(Boolean);
    onAdd(type, fields);
    onClose();
  };

  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center p-0 sm:p-6" style={{ background: "rgba(30,45,35,0.4)" }} onClick={onClose}>
      <div className="vb-card w-full sm:max-w-md rounded-t-3xl sm:rounded-3xl max-h-[88vh] overflow-y-auto" onClick={(e) => e.stopPropagation()}>
        <div className="p-6 sm:p-7">
          <div className="flex items-center justify-between mb-5">
            <h2 className="vb-heading text-xl" style={{ color: "var(--vb-ink)" }}>Thêm thủ công</h2>
            <button onClick={onClose} className="vb-icon-btn"><X size={18} /></button>
          </div>

          <div className="flex gap-2 mb-5">
            {ADD_ITEM_TYPES.map((t) => (
              <button key={t.id} onClick={() => setType(t.id)} className={cx("vb-pill-option", type === t.id && "vb-pill-option-active")}>{t.label}</button>
            ))}
          </div>

          <label className="block text-xs font-semibold uppercase tracking-wide mb-1.5" style={{ color: "var(--vb-muted)" }}>{label} tiếng Anh *</label>
          <input value={term} onChange={(e) => setTerm(e.target.value)} className="vb-input w-full mb-4" placeholder={type === "vocab" ? "vd: perseverance" : type === "idiom" ? "vd: hit the books" : "vd: take a break"} autoFocus />

          {type === "vocab" && (
            <>
              <label className="block text-xs font-semibold uppercase tracking-wide mb-1.5" style={{ color: "var(--vb-muted)" }}>Từ loại</label>
              <select value={pos} onChange={(e) => setPos(e.target.value)} className="vb-select w-full mb-4">
                <option value="noun">Danh từ</option>
                <option value="verb">Động từ</option>
                <option value="adjective">Tính từ</option>
                <option value="adverb">Trạng từ</option>
                <option value="phrase">Cụm từ</option>
              </select>
            </>
          )}

          <label className="block text-xs font-semibold uppercase tracking-wide mb-1.5" style={{ color: "var(--vb-muted)" }}>Nghĩa tiếng Việt *</label>
          <input value={meaning} onChange={(e) => setMeaning(e.target.value)} className="vb-input w-full mb-4" placeholder="vd: sự kiên trì" />

          {type === "idiom" && (
            <>
              <label className="block text-xs font-semibold uppercase tracking-wide mb-1.5" style={{ color: "var(--vb-muted)" }}>Nghĩa đen (tuỳ chọn)</label>
              <input value={literal} onChange={(e) => setLiteral(e.target.value)} className="vb-input w-full mb-4" placeholder="vd: đánh vào sách theo nghĩa đen" />

              <label className="block text-xs font-semibold uppercase tracking-wide mb-1.5" style={{ color: "var(--vb-muted)" }}>Giải thích (tuỳ chọn)</label>
              <input value={explanation} onChange={(e) => setExplanation(e.target.value)} className="vb-input w-full mb-4" placeholder="Dùng khi nào và như thế nào" />
            </>
          )}

          <label className="block text-xs font-semibold uppercase tracking-wide mb-1.5" style={{ color: "var(--vb-muted)" }}>Ví dụ tiếng Anh (tuỳ chọn)</label>
          <textarea value={example} onChange={(e) => setExample(e.target.value)} className="vb-input w-full mb-4 resize-none" rows={2} placeholder="Một câu ví dụ dùng từ này." />

          {type === "collocation" && (
            <>
              <label className="block text-xs font-semibold uppercase tracking-wide mb-1.5" style={{ color: "var(--vb-muted)" }}>Cụm từ liên quan (tuỳ chọn, cách nhau bằng dấu phẩy)</label>
              <input value={related} onChange={(e) => setRelated(e.target.value)} className="vb-input w-full mb-4" placeholder="vd: take a risk, take action" />
            </>
          )}

          <label className="block text-xs font-semibold uppercase tracking-wide mb-1.5" style={{ color: "var(--vb-muted)" }}>Chủ đề</label>
          <input value={topic} onChange={(e) => setTopic(e.target.value)} className="vb-input w-full mb-6" placeholder="Từ của tôi" />

          <button onClick={submit} disabled={!canSubmit} className="vb-btn-primary w-full justify-center disabled:opacity-40">
            <Plus size={15} /> Thêm {label}
          </button>
        </div>
      </div>
    </div>
  );
}

/* ============================== ITEM DETAIL MODAL ============================== */

function ItemDetailModal({ item, onClose, toggleFavorite, updateNotes, onDelete, requestConfirm, setItemGrade, setItemStatus }) {
  const [notes, setNotes] = useState(item.notes || "");
  useEffect(() => setNotes(item.notes || ""), [item.id]);
  const Icon = typeIcon(item.type);

  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center p-0 sm:p-6" style={{ background: "rgba(30,45,35,0.4)" }} onClick={onClose}>
      <div className="vb-card w-full sm:max-w-lg rounded-t-3xl sm:rounded-3xl max-h-[88vh] overflow-y-auto" onClick={(e) => e.stopPropagation()}>
        <div className="p-6 sm:p-7">
          <div className="flex items-start justify-between mb-4">
            <div className="flex items-center gap-2 text-xs font-medium px-2.5 py-1 rounded-full" style={{ background: "var(--vb-mint-pale)", color: "var(--vb-sage-dk)" }}>
              <Icon size={13} /> {typeLabel(item.type)}
            </div>
            <button onClick={onClose} className="vb-icon-btn"><X size={18} /></button>
          </div>

          <div className="flex items-center gap-3 flex-wrap">
            <h2 className="vb-heading text-3xl" style={{ color: "var(--vb-ink)" }}>{item.term}</h2>
            <FavoriteStar active={item.favorite} onClick={() => toggleFavorite(item.id)} size={22} />
          </div>
          {item.pos && <div className="text-sm italic mt-1" style={{ color: "var(--vb-muted)" }}>{item.pos}</div>}
          {item.ipa && <div className="text-sm mt-1" style={{ color: "var(--vb-sage-dk)" }}>{item.ipa}</div>}

          <div className="mt-3"><PronounceButton text={item.term} /></div>

          <div className="mt-5 flex flex-wrap items-center gap-2">
            <StatusPill status={item.status} />
            <GradePill grade={item.grade} />
            <span className="text-xs font-medium px-2.5 py-1 rounded-full" style={{ background: "var(--vb-mint-pale)", color: "var(--vb-sage-dk)" }}>{item.topic}</span>
          </div>

          <div className="mt-4 grid grid-cols-2 gap-3">
            <div>
              <div className="text-xs font-semibold uppercase tracking-wide mb-1.5" style={{ color: "var(--vb-muted)" }}>Trạng thái học</div>
              <StatusPicker value={item.status} onChange={(s) => setItemStatus(item.id, s)} />
            </div>
            <div>
              <div className="text-xs font-semibold uppercase tracking-wide mb-1.5" style={{ color: "var(--vb-muted)" }}>Điểm ôn tập</div>
              <GradePicker value={item.grade} onChange={(g) => setItemGrade(item.id, g)} />
            </div>
          </div>
          <p className="text-xs mt-1.5" style={{ color: "var(--vb-muted)" }}>Được tự cập nhật khi bạn chấm điểm thẻ ghi nhớ, hoặc bạn tự chỉnh ở đây bất cứ lúc nào.</p>

          <div className="mt-6">
            <div className="text-xs font-semibold uppercase tracking-wide mb-1.5" style={{ color: "var(--vb-muted)" }}>Nghĩa</div>
            <p className="text-base" style={{ color: "var(--vb-ink)" }}>{item.meaning}</p>
          </div>

          {item.literal && (
            <div className="mt-4">
              <div className="text-xs font-semibold uppercase tracking-wide mb-1.5" style={{ color: "var(--vb-muted)" }}>Nghĩa đen</div>
              <p className="text-sm" style={{ color: "var(--vb-muted)" }}>{item.literal}</p>
            </div>
          )}
          {item.explanation && (
            <div className="mt-4">
              <div className="text-xs font-semibold uppercase tracking-wide mb-1.5" style={{ color: "var(--vb-muted)" }}>Giải thích</div>
              <p className="text-sm" style={{ color: "var(--vb-ink)" }}>{item.explanation}</p>
            </div>
          )}

          <div className="mt-4">
            <div className="text-xs font-semibold uppercase tracking-wide mb-1.5" style={{ color: "var(--vb-muted)" }}>Ví dụ</div>
            <p className="text-sm italic vb-example-box">"{item.example}"</p>
          </div>

          {item.related && item.related.length > 0 && (
            <div className="mt-4">
              <div className="text-xs font-semibold uppercase tracking-wide mb-2" style={{ color: "var(--vb-muted)" }}>Cụm từ liên quan</div>
              <div className="flex flex-wrap gap-2">
                {item.related.map((r) => <span key={r} className="vb-chip text-xs px-2.5 py-1">{r}</span>)}
              </div>
            </div>
          )}

          <div className="mt-5">
            <div className="text-xs font-semibold uppercase tracking-wide mb-1.5" style={{ color: "var(--vb-muted)" }}>Ghi chú của bạn</div>
            <textarea
              value={notes}
              onChange={(e) => setNotes(e.target.value)}
              onBlur={() => updateNotes(item.id, notes)}
              placeholder="Thêm ghi chú riêng hoặc mẹo ghi nhớ..."
              className="vb-input w-full resize-none"
              rows={2}
            />
          </div>

          {item.userSentence && (
            <div className="mt-5">
              <div className="text-xs font-semibold uppercase tracking-wide mb-1.5" style={{ color: "var(--vb-muted)" }}>Câu ví dụ của bạn</div>
              <p className="text-sm italic vb-example-box">"{item.userSentence}"</p>
            </div>
          )}

          <button
            onClick={() => requestConfirm(`Xoá "${item.term}" khỏi kho từ vựng? Không thể hoàn tác đâu nhé.`, () => { onDelete(item.id); onClose(); })}
            className="mt-6 text-sm font-medium flex items-center gap-1.5"
            style={{ color: "#B4563F" }}
          >
            <Trash2 size={15} /> Xoá {typeLabel(item.type).toLowerCase()} này
          </button>
        </div>
      </div>
    </div>
  );
}

/* ============================== FLASHCARDS PAGE ============================== */

const FLASHCARD_FILTERS = [
  { id: "all-vocab", label: "Tất cả từ vựng" },
  { id: "new", label: "Từ mới" },
  { id: "learned", label: "Đã học rồi" },
  { id: "favorite", label: "Từ yêu thích" },
  { id: "grade-again", label: "Đã đánh dấu Lại nào" },
  { id: "grade-hard", label: "Đã đánh dấu Khó nhằn" },
  { id: "grade-good", label: "Đã đánh dấu Ổn áp" },
  { id: "grade-easy", label: "Đã đánh dấu Dễ ợt" },
  { id: "idioms", label: "Thành ngữ" },
  { id: "collocations", label: "Cụm từ đi kèm" },
];

function poolForFlashcardFilter(items, filterId) {
  const vocab = items.filter((i) => i.type === "vocab");
  switch (filterId) {
    case "all-vocab": return vocab;
    case "new": return vocab.filter((v) => v.status === "new");
    case "learned": return vocab.filter((v) => v.status !== "new");
    case "favorite": return vocab.filter((v) => v.favorite);
    case "grade-again": return items.filter((i) => i.grade === "again");
    case "grade-hard": return items.filter((i) => i.grade === "hard");
    case "grade-good": return items.filter((i) => i.grade === "good");
    case "grade-easy": return items.filter((i) => i.grade === "easy");
    case "idioms": return items.filter((i) => i.type === "idiom");
    case "collocations": return items.filter((i) => i.type === "collocation");
    default: return vocab;
  }
}

function FlashcardsPage({ items, gradeItem, showToast }) {
  const [filterId, setFilterId] = useState("all-vocab");
  const [deckKey, setDeckKey] = useState(0);
  const [running, setRunning] = useState(false);

  const pool = poolForFlashcardFilter(items, filterId);

  if (running) {
    return (
      <FlashcardDeckRunner
        pool={pool}
        gradeItem={gradeItem}
        onExit={() => setRunning(false)}
        onFinish={() => { setRunning(false); showToast("Xong bộ thẻ — làm tốt lắm! 🌱"); }}
      />
    );
  }

  return (
    <div>
      <SectionHeader title="Thẻ ghi nhớ" subtitle="Lật thẻ, nhớ lại, và tự chấm điểm để ghi nhớ lâu hơn." />
      <div className="vb-card vb-pop rounded-3xl p-6 max-w-xl mx-auto">
        <div className="text-sm font-semibold mb-3" style={{ color: "var(--vb-ink)" }}>Chọn một bộ thẻ 🎴</div>
        <div className="grid grid-cols-2 gap-2.5">
          {FLASHCARD_FILTERS.map((f) => {
            const count = poolForFlashcardFilter(items, f.id).length;
            const active = filterId === f.id;
            return (
              <button
                key={f.id}
                onClick={() => setFilterId(f.id)}
                className={cx("vb-deck-option", active && "vb-deck-option-active")}
              >
                <span>{f.label}</span>
                <span className="text-xs" style={{ color: active ? "var(--vb-mint-pale)" : "var(--vb-muted)" }}>{count} thẻ</span>
              </button>
            );
          })}
        </div>
        <button
          disabled={pool.length === 0}
          onClick={() => { setDeckKey((k) => k + 1); setRunning(true); }}
          className="vb-btn-primary w-full justify-center mt-6 disabled:opacity-40"
        >
          <Play size={15} /> Bắt đầu học thẻ
        </button>
        {pool.length === 0 && <p className="text-xs text-center mt-2" style={{ color: "var(--vb-muted)" }}>Bộ thẻ này chưa có thẻ nào cả.</p>}
      </div>
    </div>
  );
}

function FlashcardDeckRunner({ pool, gradeItem, onExit, onFinish }) {
  const [deck] = useState(() => shuffle(pool).map((i) => i.id));
  const [byId] = useState(() => Object.fromEntries(pool.map((i) => [i.id, i])));
  const [queue, setQueue] = useState(deck);
  const [remaining, setRemaining] = useState(() => new Set(deck));
  const [showBack, setShowBack] = useState(false);

  if (deck.length === 0) {
    return <EmptyState icon={Layers} title="Chưa có thẻ nào" subtitle="Thử chọn bộ thẻ khác xem sao." action={<button onClick={onExit} className="vb-btn-secondary mt-4">Quay lại</button>} />;
  }

  if (queue.length === 0) {
    return (
      <div className="vb-card rounded-3xl p-8 max-w-md mx-auto text-center">
        <Trophy size={36} color="var(--vb-honey)" className="mx-auto mb-3" />
        <h2 className="vb-heading text-2xl mb-2" style={{ color: "var(--vb-ink)" }}>Xong bộ thẻ rồi! 🎉</h2>
        <p className="text-sm mb-6" style={{ color: "var(--vb-muted)" }}>Bạn đã ôn {deck.length} thẻ.</p>
        <div className="flex gap-2.5 justify-center">
          <button onClick={onFinish} className="vb-btn-primary"><Home size={15} /> Về trang chủ</button>
        </div>
      </div>
    );
  }

  const currentId = queue[0];
  const card = byId[currentId];
  const done = deck.length - remaining.size;

  const grade = (g) => {
    gradeItem(currentId, g);
    let nq = queue.slice(1);
    if (g === "again") {
      const pos = Math.min(nq.length, 3);
      nq = [...nq.slice(0, pos), currentId, ...nq.slice(pos)];
    } else {
      setRemaining((r) => { const n = new Set(r); n.delete(currentId); return n; });
    }
    setQueue(nq);
    setShowBack(false);
  };

  return (
    <div>
      <div className="flex items-center justify-between mb-4">
        <button onClick={onExit} className="text-sm font-semibold flex items-center gap-1" style={{ color: "var(--vb-sage-dk)" }}><ChevronLeft size={16} /> Thoát</button>
        <div className="text-sm font-medium" style={{ color: "var(--vb-muted)" }}>Thẻ {Math.min(done + 1, deck.length)} / {deck.length}</div>
      </div>
      <div className="max-w-md mx-auto mb-3"><ProgressBar value={done} max={deck.length} /></div>

      <div className="vb-card rounded-3xl p-8 max-w-md mx-auto min-h-[280px] flex flex-col justify-between">
        <div className="text-center flex-1 flex flex-col items-center justify-center">
          <div className="text-xs font-medium uppercase tracking-wide mb-2" style={{ color: "var(--vb-muted)" }}>{typeLabel(card.type)}</div>
          <div className="vb-heading text-3xl mb-2" style={{ color: "var(--vb-ink)" }}>{card.term}</div>
          {card.ipa && <div className="text-sm mb-2" style={{ color: "var(--vb-sage-dk)" }}>{card.ipa}</div>}
          <PronounceButton text={card.term} />

          {showBack && (
            <div className="mt-5 pt-5 w-full text-left" style={{ borderTop: "2px dashed var(--vb-border)" }}>
              {card.pos && <div className="text-xs italic mb-1" style={{ color: "var(--vb-muted)" }}>{card.pos}</div>}
              <div className="text-base font-medium mb-2" style={{ color: "var(--vb-ink)" }}>{card.meaning}</div>
              <div className="text-sm italic vb-example-box">"{card.example}"</div>
            </div>
          )}
        </div>

        {!showBack ? (
          <button onClick={() => setShowBack(true)} className="vb-btn-primary w-full justify-center mt-6"><ChevronDown size={15} /> Xem nghĩa</button>
        ) : (
          <div className="grid grid-cols-4 gap-2 mt-6">
            <button onClick={() => grade("again")} className="vb-grade-btn" style={{ background: "#F7E3DD", color: "#B4563F" }}>❌ Lại nào</button>
            <button onClick={() => grade("hard")} className="vb-grade-btn" style={{ background: "#FBF1DC", color: "#C08A34" }}>😐 Khó nhằn</button>
            <button onClick={() => grade("good")} className="vb-grade-btn" style={{ background: "#E1F1EC", color: "#3E7A6A" }}>🙂 Ổn áp</button>
            <button onClick={() => grade("easy")} className="vb-grade-btn" style={{ background: "#DCF0DE", color: "#2F6B3F" }}>⭐ Dễ ợt</button>
          </div>
        )}
      </div>
    </div>
  );
}

/* ============================== PRACTICE ============================== */

function PracticeView({ items, gradeItem, recordSentence, quickPool, clearQuickPool, goto }) {
  const [config, setConfig] = useState({ practiceType: "all", questionType: "e2v", gradeFilter: "all", count: 10 });
  const [customSelection, setCustomSelection] = useState([]);
  const [pickerSearch, setPickerSearch] = useState("");
  const [session, setSession] = useState(null);

  useEffect(() => {
    if (quickPool) {
      const pool = items.filter((i) => quickPool.includes(i.id));
      setSession({ pool, questionType: "e2v" });
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [quickPool]);

  if (session) {
    return (
      <PracticeRunner
        pool={session.pool}
        questionType={session.questionType}
        gradeItem={gradeItem}
        recordSentence={recordSentence}
        onExit={() => { setSession(null); clearQuickPool(); }}
      />
    );
  }

  const isCustom = config.practiceType === "custom";

  const buildPool = () => {
    if (isCustom) return items.filter((i) => customSelection.includes(i.id));
    let pool = items.slice();
    if (config.practiceType !== "all") pool = pool.filter((i) => i.type === config.practiceType);
    if (config.gradeFilter !== "all") pool = pool.filter((i) => i.grade === config.gradeFilter);
    return pool;
  };

  const start = () => {
    const pool = isCustom ? buildPool() : shuffle(buildPool()).slice(0, config.count);
    if (pool.length === 0) return;
    setSession({ pool, questionType: config.questionType });
  };

  const poolSize = buildPool().length;
  const pickerList = items.filter((i) => `${i.term} ${i.meaning}`.toLowerCase().includes(pickerSearch.toLowerCase()));
  const toggleCustom = (id) => setCustomSelection((prev) => prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]);

  return (
    <div>
      <SectionHeader title="Luyện tập" subtitle="Tạo một buổi luyện tập riêng theo đúng nhu cầu của bạn." />
      <div className="vb-card rounded-3xl p-6 max-w-xl mx-auto">
        <PracticeConfigGroup label="Bạn muốn luyện gì nào? 🌟" options={[
          { id: "all", label: "Tất cả" }, { id: "vocab", label: "Từ vựng" }, { id: "idiom", label: "Thành ngữ" }, { id: "collocation", label: "Cụm từ đi kèm" }, { id: "custom", label: "Chọn từ cụ thể" },
        ]} value={config.practiceType} onChange={(v) => setConfig((c) => ({ ...c, practiceType: v }))} />

        <PracticeConfigGroup label="Kiểu câu hỏi" options={[
          { id: "flashcard", label: "Thẻ ghi nhớ" }, { id: "e2v", label: "Anh → Việt" }, { id: "v2e", label: "Việt → Anh" }, { id: "typing", label: "Việt → Anh (gõ chữ)" }, { id: "sentence", label: "Tự đặt câu" },
        ]} value={config.questionType} onChange={(v) => setConfig((c) => ({ ...c, questionType: v }))} />

        {isCustom ? (
          <div className="mb-5">
            <div className="text-sm font-semibold mb-2" style={{ color: "var(--vb-ink)" }}>Chọn những từ muốn ôn</div>
            <div className="relative mb-2">
              <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2" color="var(--vb-muted)" />
              <input value={pickerSearch} onChange={(e) => setPickerSearch(e.target.value)} placeholder="Tìm trong kho từ của bạn..." className="vb-input w-full pl-9" />
            </div>
            <div className="rounded-xl overflow-y-auto" style={{ border: "2px solid var(--vb-border)", maxHeight: 240 }}>
              {pickerList.length === 0 ? (
                <p className="text-sm text-center py-6" style={{ color: "var(--vb-muted)" }}>Không tìm thấy từ nào.</p>
              ) : pickerList.map((it) => (
                <label key={it.id} className="flex items-center gap-2.5 px-3.5 py-2.5 cursor-pointer vb-row-hover" style={{ borderBottom: "1px solid var(--vb-border)" }}>
                  <input type="checkbox" className="vb-checkbox" checked={customSelection.includes(it.id)} onChange={() => toggleCustom(it.id)} />
                  <span className="text-sm font-medium shrink-0" style={{ color: "var(--vb-ink)" }}>{it.term}</span>
                  <span className="text-xs truncate" style={{ color: "var(--vb-muted)" }}>{it.meaning}</span>
                </label>
              ))}
            </div>
            <p className="text-xs mt-2" style={{ color: "var(--vb-muted)" }}>Đã chọn {customSelection.length} từ</p>
          </div>
        ) : (
          <>
            <PracticeConfigGroup label="Độ khó" options={[
              { id: "all", label: "Tất cả" }, { id: "hard", label: "Khó nhằn" }, { id: "good", label: "Ổn áp" }, { id: "easy", label: "Dễ ợt" },
            ]} value={config.gradeFilter} onChange={(v) => setConfig((c) => ({ ...c, gradeFilter: v }))} />

            <PracticeConfigGroup label="Số lượng câu hỏi" options={[10, 20, 30, 50].map((n) => ({ id: n, label: String(n) }))} value={config.count} onChange={(v) => setConfig((c) => ({ ...c, count: v }))} />
          </>
        )}

        <button onClick={start} disabled={poolSize === 0} className="vb-btn-primary w-full justify-center mt-3 disabled:opacity-40">
          <Play size={15} /> Bắt đầu luyện tập
        </button>
        <p className="text-xs text-center mt-2" style={{ color: "var(--vb-muted)" }}>
          {isCustom
            ? `Đã chọn ${poolSize} từ cho buổi luyện tập này.`
            : `Có ${poolSize} mục trong kho từ của bạn khớp với lựa chọn này${poolSize < config.count ? " — buổi luyện chỉ dùng số từ đang có, không lấy thêm ngoài kho của bạn" : ""}.`}
        </p>
      </div>
    </div>
  );
}

function PracticeConfigGroup({ label, options, value, onChange }) {
  return (
    <div className="mb-5">
      <div className="text-sm font-medium mb-2" style={{ color: "var(--vb-ink)" }}>{label}</div>
      <div className="flex flex-wrap gap-2">
        {options.map((o) => (
          <button
            key={o.id}
            onClick={() => onChange(o.id)}
            className={cx("vb-pill-option", value === o.id && "vb-pill-option-active")}
          >
            {o.label}
          </button>
        ))}
      </div>
    </div>
  );
}

function PracticeRunner({ pool, questionType, gradeItem, recordSentence, onExit }) {
  if (questionType === "flashcard") {
    return <FlashcardDeckRunner pool={pool} gradeItem={gradeItem} onExit={onExit} onFinish={onExit} />;
  }
  if (questionType === "typing") {
    return <TypingRunner pool={pool} gradeItem={gradeItem} onExit={onExit} />;
  }
  if (questionType === "sentence") {
    return <SentenceRunner pool={pool} gradeItem={gradeItem} recordSentence={recordSentence} onExit={onExit} />;
  }
  return <MCQuizRunner pool={pool} direction={questionType} gradeItem={gradeItem} onExit={onExit} />;
}

function useSessionTimer() {
  const startRef = useRef(Date.now());
  const elapsed = () => Date.now() - startRef.current;
  return elapsed;
}

function PracticeResult({ correct, total, elapsedMs, mistakes, onAgain, onReviewMistakes, onDone }) {
  const secs = Math.round(elapsedMs / 1000);
  const mins = Math.floor(secs / 60);
  const rem = secs % 60;
  const acc = total > 0 ? Math.round((correct / total) * 100) : 0;
  return (
    <div className="vb-card rounded-3xl p-8 max-w-lg mx-auto text-center">
      <Sparkles size={34} color="var(--vb-honey)" className="mx-auto mb-3" />
      <h2 className="vb-heading text-2xl mb-1" style={{ color: "var(--vb-ink)" }}>Luyện tập xong rồi! 🎉</h2>
      <p className="text-lg font-semibold mt-3" style={{ color: "var(--vb-ink)" }}>{correct} / {total} câu đúng</p>
      <p className="text-sm" style={{ color: "var(--vb-muted)" }}>Độ chính xác {acc}% · Thời gian: {mins} phút {rem} giây</p>

      {mistakes.length > 0 && (
        <div className="text-left mt-6">
          <div className="text-xs font-semibold uppercase tracking-wide mb-2" style={{ color: "var(--vb-muted)" }}>Từ cần ôn lại</div>
          <div className="flex flex-wrap gap-2">
            {mistakes.map((m) => <span key={m.id} className="vb-chip text-xs px-2.5 py-1">{m.term}</span>)}
          </div>
        </div>
      )}

      <div className="flex flex-wrap gap-2.5 justify-center mt-7">
        <button onClick={onAgain} className="vb-btn-secondary"><RefreshCw size={15} /> Luyện lại</button>
        {mistakes.length > 0 && <button onClick={onReviewMistakes} className="vb-btn-secondary"><RotateCcw size={15} /> Ôn lại lỗi sai</button>}
        <button onClick={onDone} className="vb-btn-primary"><Home size={15} /> Về trang chủ</button>
      </div>
    </div>
  );
}

function makeMCQuestion(item, pool, direction) {
  const distractorPool = pool.filter((p) => p.id !== item.id);
  const distractors = shuffle(distractorPool).slice(0, 3).map((p) => (direction === "e2v" ? p.meaning : p.term));
  const correctAnswer = direction === "e2v" ? item.meaning : item.term;
  const options = shuffle(Array.from(new Set([correctAnswer, ...distractors])));
  const prompt = direction === "e2v" ? `"${item.term}" nghĩa là gì?` : `"${item.meaning}" nghĩa là:`;
  return { item, prompt, options, correctAnswer };
}

function MCQuizRunner({ pool, direction, gradeItem, onExit }) {
  const [questions] = useState(() => shuffle(pool).map((it) => makeMCQuestion(it, pool, direction)));
  const [idx, setIdx] = useState(0);
  const [selected, setSelected] = useState(null);
  const [correctCount, setCorrectCount] = useState(0);
  const [mistakes, setMistakes] = useState([]);
  const elapsed = useSessionTimer();
  const [finishedMs, setFinishedMs] = useState(null);
  const [restartKey, setRestartKey] = useState(0);

  if (questions.length === 0) {
    return <EmptyState icon={Target} title="Chưa đủ từ để luyện" subtitle="Thêm từ vựng để tạo bài luyện tập." action={<button onClick={onExit} className="vb-btn-secondary mt-4">Quay lại</button>} />;
  }

  if (finishedMs !== null) {
    return (
      <PracticeResult
        correct={correctCount} total={questions.length} elapsedMs={finishedMs} mistakes={mistakes}
        onAgain={() => { setIdx(0); setSelected(null); setCorrectCount(0); setMistakes([]); setFinishedMs(null); setRestartKey((k) => k + 1); }}
        onReviewMistakes={onExit} onDone={onExit}
      />
    );
  }

  const q = questions[idx];

  const choose = (opt) => {
    if (selected) return;
    setSelected(opt);
    const isCorrect = opt === q.correctAnswer;
    gradeItem(q.item.id, isCorrect ? "good" : "again");
    if (isCorrect) setCorrectCount((c) => c + 1); else setMistakes((m) => [...m, q.item]);
  };

  const next = () => {
    if (idx + 1 >= questions.length) { setFinishedMs(elapsed()); return; }
    setIdx((i) => i + 1);
    setSelected(null);
  };

  return (
    <div key={restartKey}>
      <div className="flex items-center justify-between mb-4">
        <button onClick={onExit} className="text-sm font-semibold flex items-center gap-1" style={{ color: "var(--vb-sage-dk)" }}><ChevronLeft size={16} /> Thoát</button>
        <div className="text-sm font-medium" style={{ color: "var(--vb-muted)" }}>Câu {idx + 1} / {questions.length}</div>
      </div>
      <div className="max-w-lg mx-auto mb-3"><ProgressBar value={idx} max={questions.length} /></div>

      <div className="vb-card rounded-3xl p-7 max-w-lg mx-auto">
        <div className="text-lg font-medium mb-5" style={{ color: "var(--vb-ink)" }}>{q.prompt}</div>
        <div className="flex flex-col gap-2.5">
          {q.options.map((opt, i) => {
            const isCorrectOpt = opt === q.correctAnswer;
            const isChosen = opt === selected;
            let style = {};
            if (selected) {
              if (isCorrectOpt) style = { background: "#DCF0DE", borderColor: "#3E7A6A", color: "#2F6B3F" };
              else if (isChosen) style = { background: "#F7E3DD", borderColor: "#B4563F", color: "#B4563F" };
            }
            return (
              <button key={i} onClick={() => choose(opt)} className="vb-option-btn" style={style} disabled={!!selected}>
                <span className="w-5 h-5 rounded-full text-xs flex items-center justify-center shrink-0 font-semibold" style={{ background: "var(--vb-mint-pale)", color: "var(--vb-sage-dk)" }}>{String.fromCharCode(65 + i)}</span>
                <span className="text-left">{opt}</span>
                {selected && isCorrectOpt && <Check size={16} className="ml-auto" />}
                {selected && isChosen && !isCorrectOpt && <X size={16} className="ml-auto" />}
              </button>
            );
          })}
        </div>

        {selected && (
          <div className="mt-5 pt-5" style={{ borderTop: "2px dashed var(--vb-border)" }}>
            <div className="flex items-center gap-2 text-sm font-medium mb-2" style={{ color: selected === q.correctAnswer ? "#2F6B3F" : "#B4563F" }}>
              {selected === q.correctAnswer ? <><CheckCircle2 size={16} /> Chính xác</> : <><XCircle size={16} /> Sai rồi — đáp án đúng: {q.correctAnswer}</>}
            </div>
            <p className="text-sm italic vb-example-box">"{q.item.example}"</p>
            <button onClick={next} className="vb-btn-primary w-full justify-center mt-4">
              {idx + 1 >= questions.length ? "Xem kết quả" : "Câu tiếp theo"} <ChevronRight size={15} />
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

function TypingRunner({ pool, gradeItem, onExit }) {
  const [questions] = useState(() => shuffle(pool));
  const [idx, setIdx] = useState(0);
  const [value, setValue] = useState("");
  const [checked, setChecked] = useState(null);
  const [correctCount, setCorrectCount] = useState(0);
  const [mistakes, setMistakes] = useState([]);
  const elapsed = useSessionTimer();
  const [finishedMs, setFinishedMs] = useState(null);
  const [restartKey, setRestartKey] = useState(0);

  if (questions.length === 0) {
    return <EmptyState icon={PenLine} title="Chưa đủ từ để luyện" subtitle="Thêm từ vựng để luyện gõ chữ." action={<button onClick={onExit} className="vb-btn-secondary mt-4">Quay lại</button>} />;
  }

  if (finishedMs !== null) {
    return (
      <PracticeResult
        correct={correctCount} total={questions.length} elapsedMs={finishedMs} mistakes={mistakes}
        onAgain={() => { setIdx(0); setValue(""); setChecked(null); setCorrectCount(0); setMistakes([]); setFinishedMs(null); setRestartKey((k) => k + 1); }}
        onReviewMistakes={onExit} onDone={onExit}
      />
    );
  }

  const q = questions[idx];

  const check = () => {
    if (checked) return;
    const isCorrect = normalize(value) === normalize(q.term);
    setChecked(isCorrect ? "correct" : "incorrect");
    gradeItem(q.id, isCorrect ? "good" : "again");
    if (isCorrect) setCorrectCount((c) => c + 1); else setMistakes((m) => [...m, q]);
  };

  const next = () => {
    if (idx + 1 >= questions.length) { setFinishedMs(elapsed()); return; }
    setIdx((i) => i + 1);
    setValue("");
    setChecked(null);
  };

  return (
    <div key={restartKey}>
      <div className="flex items-center justify-between mb-4">
        <button onClick={onExit} className="text-sm font-semibold flex items-center gap-1" style={{ color: "var(--vb-sage-dk)" }}><ChevronLeft size={16} /> Thoát</button>
        <div className="text-sm font-medium" style={{ color: "var(--vb-muted)" }}>Câu {idx + 1} / {questions.length}</div>
      </div>
      <div className="max-w-lg mx-auto mb-3"><ProgressBar value={idx} max={questions.length} /></div>

      <div className="vb-card rounded-3xl p-7 max-w-lg mx-auto">
        <div className="text-xs font-semibold uppercase tracking-wide mb-1.5" style={{ color: "var(--vb-muted)" }}>Nghĩa tiếng Việt</div>
        <div className="text-lg font-medium mb-5" style={{ color: "var(--vb-ink)" }}>{q.meaning}</div>

        <input
          value={value}
          onChange={(e) => setValue(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && check()}
          placeholder="Gõ từ hoặc cụm từ tiếng Anh..."
          className="vb-input w-full"
          disabled={!!checked}
          autoFocus
        />

        {!checked ? (
          <button onClick={check} disabled={!value.trim()} className="vb-btn-primary w-full justify-center mt-4 disabled:opacity-40">Kiểm tra</button>
        ) : (
          <div className="mt-5 pt-5" style={{ borderTop: "2px dashed var(--vb-border)" }}>
            <div className="flex items-center gap-2 text-sm font-medium mb-2" style={{ color: checked === "correct" ? "#2F6B3F" : "#B4563F" }}>
              {checked === "correct" ? <><CheckCircle2 size={16} /> Chính xác!</> : <><XCircle size={16} /> Sai rồi — đáp án đúng: {q.term}</>}
            </div>
            <p className="text-sm italic vb-example-box">"{q.example}"</p>
            <button onClick={next} className="vb-btn-primary w-full justify-center mt-4">
              {idx + 1 >= questions.length ? "Xem kết quả" : "Câu tiếp theo"} <ChevronRight size={15} />
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

const SENTENCE_STOPWORDS = new Set(["a", "an", "the", "to", "of", "in", "on", "at", "for", "and", "or", "this", "that", "my", "your", "his", "her", "its", "our", "their"]);
const IRREGULAR_STEMS = {
  make: ["mak", "mad"], take: ["tak", "took"], do: ["do", "did", "don"], break: ["break", "brok"],
  bite: ["bit"], catch: ["catch", "caught"], buy: ["buy", "bought"], bring: ["bring", "brought"],
  think: ["think", "thought"], teach: ["teach", "taught"], find: ["find", "found"], give: ["giv", "gave", "given"],
  go: ["go", "went", "gone"], know: ["know", "knew", "known"], see: ["see", "saw", "seen"],
  write: ["writ", "wrote", "written"], speak: ["speak", "spoke", "spoken"], get: ["get", "got", "gotten"],
  come: ["com", "came"], run: ["run", "ran"], say: ["say", "said"], tell: ["tell", "told"],
  feel: ["feel", "felt"], leave: ["leav", "left"], hold: ["hold", "held"], meet: ["meet", "met"],
  lead: ["lead", "led"], stand: ["stand", "stood"], win: ["win", "won"], lose: ["los", "lost"],
  spend: ["spend", "spent"], build: ["build", "built"], send: ["send", "sent"], lend: ["lend", "lent"],
  pay: ["pa", "paid"], sit: ["sit", "sat"], sleep: ["sleep", "slept"], keep: ["keep", "kept"],
};

function sentenceWordPattern(w) {
  const lw = w.toLowerCase();
  if (SENTENCE_STOPWORDS.has(lw)) return escapeRegExp(w);
  if (IRREGULAR_STEMS[lw]) return `(?:${IRREGULAR_STEMS[lw].map(escapeRegExp).join("|")})\\w*`;
  if (lw.length < 3) return escapeRegExp(w);
  return `${escapeRegExp(w.replace(/e$/i, ""))}\\w*`;
}

// Loose check so common inflections (achieve -> achieved/achieving, make -> made/making)
// still count, while requiring the actual word/phrase to appear — not just similar text.
function sentenceUsesTerm(sentence, term) {
  const words = String(term).trim().split(/\s+/).filter(Boolean);
  if (words.length === 0 || !sentence) return false;
  try {
    const pattern = words.map(sentenceWordPattern).join("\\s+");
    return new RegExp(`\\b${pattern}`, "i").test(sentence);
  } catch (e) {
    return sentence.toLowerCase().includes(String(term).toLowerCase());
  }
}

function SentenceRunner({ pool, gradeItem, recordSentence, onExit }) {
  const [questions] = useState(() => shuffle(pool));
  const [idx, setIdx] = useState(0);
  const [value, setValue] = useState("");
  const [checked, setChecked] = useState(null); // { ok, short }
  const [correctCount, setCorrectCount] = useState(0);
  const [mistakes, setMistakes] = useState([]);
  const elapsed = useSessionTimer();
  const [finishedMs, setFinishedMs] = useState(null);
  const [restartKey, setRestartKey] = useState(0);

  if (questions.length === 0) {
    return <EmptyState icon={PenLine} title="Chưa đủ từ để luyện" subtitle="Thêm từ vựng trước đã nhé." action={<button onClick={onExit} className="vb-btn-secondary mt-4">Quay lại</button>} />;
  }

  if (finishedMs !== null) {
    return (
      <PracticeResult
        correct={correctCount} total={questions.length} elapsedMs={finishedMs} mistakes={mistakes}
        onAgain={() => { setIdx(0); setValue(""); setChecked(null); setCorrectCount(0); setMistakes([]); setFinishedMs(null); setRestartKey((k) => k + 1); }}
        onReviewMistakes={onExit} onDone={onExit}
      />
    );
  }

  const q = questions[idx];

  const check = () => {
    if (checked || !value.trim()) return;
    const ok = sentenceUsesTerm(value, q.term);
    const wc = value.trim().split(/\s+/).filter(Boolean).length;
    setChecked({ ok, short: wc < 4 });
    gradeItem(q.id, ok ? "good" : "again");
    recordSentence(q.id, value.trim(), ok);
    if (ok) setCorrectCount((c) => c + 1); else setMistakes((m) => [...m, q]);
  };

  const next = () => {
    if (idx + 1 >= questions.length) { setFinishedMs(elapsed()); return; }
    setIdx((i) => i + 1);
    setValue("");
    setChecked(null);
  };

  return (
    <div key={restartKey}>
      <div className="flex items-center justify-between mb-4">
        <button onClick={onExit} className="text-sm font-semibold flex items-center gap-1" style={{ color: "var(--vb-sage-dk)" }}><ChevronLeft size={16} /> Thoát</button>
        <div className="text-sm font-medium" style={{ color: "var(--vb-muted)" }}>Câu {idx + 1} / {questions.length}</div>
      </div>
      <div className="max-w-lg mx-auto mb-3"><ProgressBar value={idx} max={questions.length} /></div>

      <div className="vb-card rounded-3xl p-7 max-w-lg mx-auto">
        <div className="text-xs font-semibold uppercase tracking-wide mb-1.5" style={{ color: "var(--vb-muted)" }}>Nghĩa tiếng Việt</div>
        <div className="text-lg font-medium mb-1" style={{ color: "var(--vb-ink)" }}>{q.meaning}</div>
        <div className="text-xs mb-4" style={{ color: "var(--vb-muted)" }}>Viết một câu tiếng Anh dùng đúng {typeLabel(q.type).toLowerCase()} này — từ sẽ được ẩn cho tới khi bạn kiểm tra.</div>

        <textarea
          value={value}
          onChange={(e) => setValue(e.target.value)}
          placeholder="Gõ câu tiếng Anh của bạn ở đây..."
          className="vb-input w-full resize-none"
          rows={3}
          disabled={!!checked}
          autoFocus
        />

        {!checked ? (
          <button onClick={check} disabled={!value.trim()} className="vb-btn-primary w-full justify-center mt-4 disabled:opacity-40">Kiểm tra câu của tôi</button>
        ) : (
          <div className="mt-5 pt-5" style={{ borderTop: "2px dashed var(--vb-border)" }}>
            <div className="flex items-center gap-2 text-sm font-medium mb-2" style={{ color: checked.ok ? "#2F6B3F" : "#B4563F" }}>
              {checked.ok ? <><CheckCircle2 size={16} /> Tốt lắm — dùng đúng rồi!</> : <><XCircle size={16} /> Câu này có vẻ chưa dùng đúng từ đó.</>}
            </div>
            {checked.ok && checked.short && (
              <p className="text-xs mb-2" style={{ color: "var(--vb-muted)" }}>Mẹo: thử viết câu dài hơn một chút lần sau để luyện tốt hơn nhé.</p>
            )}
            <div className="text-xs font-semibold uppercase tracking-wide mb-1" style={{ color: "var(--vb-muted)" }}>Từ/cụm từ đó là</div>
            <div className="font-medium mb-2" style={{ color: "var(--vb-ink)" }}>{q.term}</div>
            <div className="text-xs font-semibold uppercase tracking-wide mb-1" style={{ color: "var(--vb-muted)" }}>Ví dụ tham khảo</div>
            <p className="text-sm italic vb-example-box">"{q.example}"</p>
            <button onClick={next} className="vb-btn-primary w-full justify-center mt-4">
              {idx + 1 >= questions.length ? "Xem kết quả" : "Câu tiếp theo"} <ChevronRight size={15} />
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

/* ============================== IDIOMS / COLLOCATIONS LIST ============================== */

function IdiomsCollocationsList({ type, items, toggleFavorite, openDetail, showToast, addManualItem, deleteItems, requestConfirm }) {
  const [q, setQ] = useState("");
  const [showAdd, setShowAdd] = useState(false);
  const [selectMode, setSelectMode] = useState(false);
  const [selected, setSelected] = useState([]);
  const list = items.filter((i) => i.type === type && `${i.term} ${i.meaning}`.toLowerCase().includes(q.toLowerCase()));
  const title = type === "idiom" ? "Thành ngữ" : "Cụm từ đi kèm";
  const Icon = type === "idiom" ? Quote : Link2;

  const toggleSelected = (id) => setSelected((prev) => prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]);
  const toggleSelectAll = () => setSelected((prev) => prev.length === list.length ? [] : list.map((i) => i.id));
  const exitSelectMode = () => { setSelectMode(false); setSelected([]); };
  const handleBulkDelete = () => {
    requestConfirm(`Xoá ${selected.length} mục đã chọn? Không thể hoàn tác đâu nhé.`, () => { deleteItems(selected); exitSelectMode(); }, "Xoá đã chọn");
  };

  return (
    <div>
      <SectionHeader title={title} subtitle={`${items.filter((i) => i.type === type).length} ${title.toLowerCase()} để khám phá`} right={
        <div className="flex items-center gap-2 flex-wrap justify-end">
          <div className="relative">
            <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2" color="var(--vb-muted)" />
            <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Tìm kiếm..." className="vb-input pl-9" />
          </div>
          <button onClick={() => (selectMode ? exitSelectMode() : setSelectMode(true))} className={cx("vb-btn-secondary text-sm", selectMode && "vb-chip-toggle-active")}>
            <ListChecks size={15} /> {selectMode ? "Xong" : "Chọn nhiều"}
          </button>
          <button onClick={() => setShowAdd(true)} className="vb-btn-primary text-sm"><Plus size={15} /> Thêm thủ công</button>
        </div>
      } />

      {selectMode && (
        <div className="vb-bulk-bar mb-4 flex-wrap">
          <label className="flex items-center gap-2 text-sm font-semibold cursor-pointer" style={{ color: "var(--vb-sage-dk)" }}>
            <input type="checkbox" className="vb-checkbox" checked={selected.length > 0 && selected.length === list.length} onChange={toggleSelectAll} />
            Chọn tất cả ({list.length})
          </label>
          <span className="text-sm" style={{ color: "var(--vb-sage-dk)" }}>Đã chọn {selected.length} mục</span>
          <button onClick={handleBulkDelete} disabled={selected.length === 0} className="vb-btn-primary text-sm ml-auto disabled:opacity-40" style={{ background: "#B4563F", borderColor: "#B4563F" }}>
            <Trash2 size={14} /> Xoá đã chọn
          </button>
        </div>
      )}

      {list.length === 0 ? (
        <EmptyState icon={Icon} title={`Không tìm thấy ${title.toLowerCase()}`} subtitle="Thử từ khoá khác, hoặc tự thêm mục của riêng bạn." action={
          <button onClick={() => setShowAdd(true)} className="vb-btn-primary mt-4"><Plus size={15} /> Thêm thủ công</button>
        } />
      ) : (
        <div className="grid sm:grid-cols-2 gap-4">
          {list.map((it) => (
            <div key={it.id} className="vb-card rounded-2xl p-5 cursor-pointer relative" onClick={() => (selectMode ? toggleSelected(it.id) : openDetail(it))}>
              {selectMode && (
                <input
                  type="checkbox"
                  className="vb-checkbox absolute top-4 left-4"
                  checked={selected.includes(it.id)}
                  onChange={() => toggleSelected(it.id)}
                  onClick={(e) => e.stopPropagation()}
                />
              )}
              <div className={cx("flex items-start justify-between gap-2", selectMode && "pl-7")}>
                <h3 className="vb-heading text-lg" style={{ color: "var(--vb-ink)" }}>{it.term}</h3>
                <FavoriteStar active={it.favorite} onClick={() => toggleFavorite(it.id)} />
              </div>
              <p className="text-sm mt-1.5" style={{ color: "var(--vb-muted)" }}>{it.meaning}</p>
              <p className="text-sm italic mt-3 vb-example-box">"{it.example}"</p>
              <div className="flex items-center justify-between mt-4">
                <div onClick={(e) => e.stopPropagation()}><PronounceButton text={it.term} size="sm" /></div>
                <div className="flex items-center gap-2">
                  <GradePill grade={it.grade} />
                  <button
                    onClick={(e) => { e.stopPropagation(); showToast(`Đã thêm — tìm trong Thẻ ghi nhớ → ${type === "idiom" ? "Thành ngữ" : "Cụm từ đi kèm"}.`); }}
                    className="vb-chip text-xs px-2.5 py-1 flex items-center gap-1"
                  >
                    <Layers size={12} /> Thêm vào Thẻ ghi nhớ
                  </button>
                </div>
              </div>
            </div>
          ))}
        </div>
      )}

      {showAdd && <AddItemModal initialType={type} onAdd={addManualItem} onClose={() => setShowAdd(false)} />}
    </div>
  );
}

/* ============================== STATISTICS ============================== */

function StatisticsView({ items, stats, earnedBadges, startWeakPractice }) {
  const vocab = items.filter((i) => i.type === "vocab");
  const idioms = items.filter((i) => i.type === "idiom");
  const collocations = items.filter((i) => i.type === "collocation");
  const totalReviewed = stats.correct + stats.incorrect;
  const accuracy = totalReviewed > 0 ? Math.round((stats.correct / totalReviewed) * 100) : 0;
  const mastered = items.filter((i) => i.status === "mastered").length;

  const learnedNow = vocab.filter((v) => v.status !== "new").length;
  const progressData = [
    { date: "6 ngày trước", words: Math.max(0, learnedNow - 6) },
    { date: "5 ngày trước", words: Math.max(0, learnedNow - 5) },
    { date: "4 ngày trước", words: Math.max(0, learnedNow - 4) },
    { date: "3 ngày trước", words: Math.max(0, learnedNow - 3) },
    { date: "2 ngày trước", words: Math.max(0, learnedNow - 2) },
    { date: "Hôm qua", words: Math.max(0, learnedNow - 1) },
    { date: "Hôm nay", words: learnedNow },
  ];

  const accuracyData = [
    { name: "Đúng", value: stats.correct || 0 },
    { name: "Sai", value: stats.incorrect || 0 },
  ];

  const typeData = [
    { name: "Từ vựng", value: vocab.length },
    { name: "Thành ngữ", value: idioms.length },
    { name: "Cụm từ", value: collocations.length },
  ];

  const weekLabels = [];
  const weekData = [];
  for (let i = 6; i >= 0; i--) {
    const d = new Date(); d.setDate(d.getDate() - i);
    const key = fmtDate(d);
    weekData.push({ name: d.toLocaleDateString("vi-VN", { weekday: "short" }), value: stats.dailyActivity[key] || 0 });
  }

  const weakest = items
    .filter((i) => (i.reviews.correct + i.reviews.incorrect) > 0)
    .map((i) => ({ ...i, acc: i.reviews.correct / (i.reviews.correct + i.reviews.incorrect) }))
    .sort((a, b) => a.acc - b.acc)
    .slice(0, 6);

  const statusCounts = ["new", "learning", "familiar", "mastered"].map((s) => ({
    name: STATUS_META[s].label, value: items.filter((i) => i.status === s).length, key: s,
  }));

  const topicMap = {};
  items.forEach((i) => {
    if (!topicMap[i.topic]) topicMap[i.topic] = { total: 0, mastered: 0 };
    topicMap[i.topic].total += 1;
    if (i.status === "mastered") topicMap[i.topic].mastered += 1;
  });
  const topicRows = Object.entries(topicMap).sort((a, b) => b[1].total - a[1].total).slice(0, 8);

  const manualCount = items.filter((i) => i.source === "manual").length;

  const typeAccuracy = (type) => {
    const list = items.filter((i) => i.type === type);
    const c = list.reduce((s, i) => s + i.reviews.correct, 0);
    const w = list.reduce((s, i) => s + i.reviews.incorrect, 0);
    return c + w > 0 ? Math.round((c / (c + w)) * 100) : null;
  };
  const idiomAcc = typeAccuracy("idiom");
  const collocAcc = typeAccuracy("collocation");

  return (
    <div>
      <SectionHeader title="Thống kê" subtitle="Xem thành quả học tập của bạn nhé." />

      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 sm:gap-4 mb-6">
        <MiniStat label="Tổng lượt ôn" value={stats.totalReviews} icon={ListChecks} />
        <MiniStat label="Tỷ lệ chính xác" value={`${accuracy}%`} icon={Target} />
        <MiniStat label="Từ đã thuộc" value={mastered} icon={Trophy} />
        <MiniStat label="Chuỗi hiện tại" value={`${stats.streak} ngày`} icon={Flame} />
      </div>
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 sm:gap-4 mb-8">
        <MiniStat label="Câu trả lời đúng" value={stats.correct} icon={CheckCircle2} />
        <MiniStat label="Câu trả lời sai" value={stats.incorrect} icon={XCircle} />
        <MiniStat label="Chuỗi dài nhất" value={`${stats.longestStreak} ngày`} icon={Award} />
        <MiniStat label="Lượt ôn hôm nay" value={stats.todayReviewed} icon={Clock} />
      </div>
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 sm:gap-4 mb-8">
        <MiniStat label="Câu đã viết" value={stats.sentencesWritten || 0} icon={PenLine} />
        <MiniStat label="Từ tự thêm" value={manualCount} icon={Plus} />
        <MiniStat label="Độ chính xác thành ngữ" value={idiomAcc !== null ? `${idiomAcc}%` : "—"} icon={Quote} />
        <MiniStat label="Độ chính xác cụm từ" value={collocAcc !== null ? `${collocAcc}%` : "—"} icon={Link2} />
      </div>

      <div className="grid lg:grid-cols-2 gap-5 mb-6">
        <ChartCard title="Tiến độ học" subtitle="Số từ đã học trong tuần qua">
          <ResponsiveContainer width="100%" height="100%">
            <LineChart data={progressData}>
              <CartesianGrid strokeDasharray="3 3" stroke="var(--vb-border)" vertical={false} />
              <XAxis dataKey="date" tick={{ fontSize: 11, fill: COLORS.sage }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: COLORS.sage }} axisLine={false} tickLine={false} width={28} />
              <Tooltip />
              <Line type="monotone" dataKey="words" stroke={COLORS.mint} strokeWidth={3} dot={{ r: 3, fill: COLORS.mint }} />
            </LineChart>
          </ResponsiveContainer>
        </ChartCard>

        <ChartCard title="Độ chính xác" subtitle="Số câu đúng so với sai">
          {totalReviewed === 0 ? (
            <div className="h-full flex items-center justify-center text-sm" style={{ color: "var(--vb-muted)" }}>Chưa có lượt ôn nào</div>
          ) : (
            <ResponsiveContainer width="100%" height="100%">
              <PieChart>
                <Pie data={accuracyData} dataKey="value" nameKey="name" innerRadius={54} outerRadius={82} paddingAngle={3}>
                  <Cell fill={COLORS.mint} />
                  <Cell fill={COLORS.coral} />
                </Pie>
                <Tooltip />
              </PieChart>
            </ResponsiveContainer>
          )}
        </ChartCard>

        <ChartCard title="Từ vựng theo loại" subtitle="Thành phần kho từ của bạn">
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={typeData}>
              <CartesianGrid strokeDasharray="3 3" stroke="var(--vb-border)" vertical={false} />
              <XAxis dataKey="name" tick={{ fontSize: 11, fill: COLORS.sage }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: COLORS.sage }} axisLine={false} tickLine={false} width={28} />
              <Tooltip />
              <Bar dataKey="value" radius={[8, 8, 0, 0]}>
                <Cell fill={COLORS.mint} /><Cell fill={COLORS.honey} /><Cell fill={COLORS.sage} />
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>

        <ChartCard title="Hoạt động trong tuần" subtitle="Số bài luyện hoàn thành mỗi ngày">
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={weekData}>
              <CartesianGrid strokeDasharray="3 3" stroke="var(--vb-border)" vertical={false} />
              <XAxis dataKey="name" tick={{ fontSize: 11, fill: COLORS.sage }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: COLORS.sage }} axisLine={false} tickLine={false} width={28} />
              <Tooltip />
              <Bar dataKey="value" fill={COLORS.mint} radius={[8, 8, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>

        <ChartCard title="Phân bổ theo cấp độ" subtitle="Mỗi từ, thành ngữ & cụm từ đang ở đâu">
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={statusCounts}>
              <CartesianGrid strokeDasharray="3 3" stroke="var(--vb-border)" vertical={false} />
              <XAxis dataKey="name" tick={{ fontSize: 11, fill: COLORS.sage }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: COLORS.sage }} axisLine={false} tickLine={false} width={28} />
              <Tooltip />
              <Bar dataKey="value" radius={[8, 8, 0, 0]}>
                {statusCounts.map((s) => <Cell key={s.key} fill={STATUS_META[s.key].color} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </ChartCard>
      </div>

      <div className="vb-card rounded-2xl p-6 mb-6">
        <div className="flex items-center justify-between mb-4">
          <h3 className="vb-heading text-lg" style={{ color: "var(--vb-ink)" }}>Từ yếu nhất</h3>
          {weakest.length > 0 && <button onClick={startWeakPractice} className="vb-btn-secondary text-sm"><RotateCcw size={14} /> Luyện từ yếu</button>}
        </div>
        {weakest.length === 0 ? (
          <p className="text-sm" style={{ color: "var(--vb-muted)" }}>Chưa có từ nào yếu cả — duy trì phong độ nhé!</p>
        ) : (
          <div className="grid sm:grid-cols-2 gap-2.5">
            {weakest.map((w) => (
              <div key={w.id} className="flex items-center justify-between px-4 py-2.5 rounded-xl" style={{ background: "var(--vb-mint-pale)" }}>
                <span className="text-sm font-medium" style={{ color: "var(--vb-ink)" }}>{w.term}</span>
                <span className="text-xs" style={{ color: "#B4563F" }}>Độ chính xác {Math.round(w.acc * 100)}%</span>
              </div>
            ))}
          </div>
        )}
      </div>

      <div className="vb-card rounded-2xl p-6 mb-6">
        <h3 className="vb-heading text-lg mb-4" style={{ color: "var(--vb-ink)" }}>Theo chủ đề</h3>
        {topicRows.length === 0 ? (
          <p className="text-sm" style={{ color: "var(--vb-muted)" }}>Chưa có gì để hiển thị.</p>
        ) : (
          <div className="grid sm:grid-cols-2 gap-3">
            {topicRows.map(([name, data]) => (
              <div key={name} className="p-3.5 rounded-xl" style={{ background: "var(--vb-mint-pale)" }}>
                <div className="flex items-center justify-between mb-1.5">
                  <span className="text-sm font-medium" style={{ color: "var(--vb-ink)" }}>{name}</span>
                  <span className="text-xs" style={{ color: "var(--vb-muted)" }}>{data.mastered}/{data.total} đã thuộc</span>
                </div>
                <ProgressBar value={data.mastered} max={data.total} height={7} />
              </div>
            ))}
          </div>
        )}
      </div>

      <div className="vb-card rounded-2xl p-6">
        <h3 className="vb-heading text-lg mb-4" style={{ color: "var(--vb-ink)" }}>Thành tích</h3>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
          {BADGES.map((b) => {
            const earned = earnedBadges.includes(b.id);
            const BIcon = b.icon;
            return (
              <div key={b.id} className="flex flex-col items-center text-center gap-2 p-4 rounded-2xl" style={{ background: earned ? "var(--vb-mint-pale)" : "#F3F5F1", opacity: earned ? 1 : 0.55 }}>
                <BIcon size={22} color={earned ? "var(--vb-sage-dk)" : "#A9B7AB"} />
                <span className="text-xs font-medium" style={{ color: earned ? "var(--vb-ink)" : "#8AA694" }}>{b.label}</span>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

function ChartCard({ title, subtitle, children }) {
  return (
    <div className="vb-card rounded-2xl p-5">
      <div className="mb-3">
        <div className="text-sm font-semibold" style={{ color: "var(--vb-ink)" }}>{title}</div>
        <div className="text-xs" style={{ color: "var(--vb-muted)" }}>{subtitle}</div>
      </div>
      <div style={{ height: 220 }}>{children}</div>
    </div>
  );
}

/* ============================== IMPORT VOCABULARY ============================== */

function ImportView({ items, setItems, showToast, goto }) {
  const [importType, setImportType] = useState("vocab");
  const [dragOver, setDragOver] = useState(false);
  const [error, setError] = useState(null);
  const [preview, setPreview] = useState(null); // { rows: [{...fields, dup}], dupChoice }
  const fileRef = useRef(null);
  const cfg = IMPORT_TYPES[importType];

  const resetForTypeChange = (t) => {
    setImportType(t);
    setError(null);
    setPreview(null);
    if (fileRef.current) fileRef.current.value = "";
  };

  const handleFile = async (file) => {
    setError(null);
    setPreview(null);
    if (!file) return;
    const validExt = /\.(xlsx|xls)$/i.test(file.name);
    if (!validExt) { setError("Định dạng file không hợp lệ. Vui lòng tải lên file .xlsx."); return; }
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
      if (!rows || rows.length < 1) { setError("File này có vẻ trống rỗng."); return; }
      const header = rows[0].map((h) => String(h).trim());
      const findCol = (name) => header.findIndex((h) => h.toLowerCase() === name.toLowerCase());

      const missing = cfg.required.filter((name) => findCol(name) === -1);
      if (missing.length > 0) { setError(`Thiếu cột bắt buộc. Cần có: ${cfg.columns.join(", ")}.`); return; }

      const existingTerms = new Set(items.filter((i) => i.type === cfg.itemType).map((i) => normalize(i.term)));
      const dataRows = rows.slice(1)
        .map((r) => {
          const get = (name) => { const idx = findCol(name); return idx === -1 ? "" : String(r[idx] || "").trim(); };
          return cfg.mapRow(get);
        })
        .filter((r) => r.term);

      if (dataRows.length === 0) { setError(`Không tìm thấy dòng ${cfg.label.toLowerCase()} nào — bảng tính đang trống.`); return; }

      const withDup = dataRows.map((r) => ({ ...r, dup: existingTerms.has(normalize(r.term)) }));
      setPreview({ rows: withDup, dupChoice: "skip" });
    } catch (e) {
      setError("Không đọc được file này. Kiểm tra lại xem có phải file .xlsx hợp lệ không nhé.");
    }
  };

  const onDrop = (e) => {
    e.preventDefault(); setDragOver(false);
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    handleFile(f);
  };

  const confirmImport = () => {
    if (!preview) return;
    const { rows, dupChoice } = preview;
    let toAdd = [];
    let replacedTerms = [];
    let skipped = 0;
    rows.forEach((r) => {
      const { dup, ...fields } = r;
      if (dup) {
        if (dupChoice === "skip") { skipped += 1; return; }
        if (dupChoice === "replace") replacedTerms.push(normalize(r.term));
      }
      toAdd.push(makeItem(cfg.itemType, { ...fields, topic: fields.topic || "Đã nhập", source: "import" }));
    });

    setItems((prev) => {
      let next = prev;
      if (replacedTerms.length) next = next.filter((it) => !(it.type === cfg.itemType && replacedTerms.includes(normalize(it.term))));
      return [...next, ...toAdd];
    });

    showToast(`✅ Đã nhập thành công ${toAdd.length} mục ${cfg.label.toLowerCase()}!${skipped ? ` (bỏ qua ${skipped} mục trùng)` : ""}`);
    setPreview(null);
    if (fileRef.current) fileRef.current.value = "";
  };

  const dupCount = preview ? preview.rows.filter((r) => r.dup).length : 0;

  return (
    <div>
      <SectionHeader title="Nhập từ vựng" subtitle="Tải lên file Excel để thêm hàng loạt từ, thành ngữ, hoặc cụm từ." right={
        <button onClick={() => goto("template")} className="vb-btn-secondary text-sm"><Download size={14} /> Lấy mẫu Excel</button>
      } />

      <div className="flex gap-2 mb-5">
        {Object.entries(IMPORT_TYPES).map(([key, t]) => (
          <button key={key} onClick={() => resetForTypeChange(key)} className={cx("vb-pill-option", importType === key && "vb-pill-option-active")}>{t.label}</button>
        ))}
      </div>

      {!preview && (
        <div
          onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
          onDragLeave={() => setDragOver(false)}
          onDrop={onDrop}
          className="vb-card rounded-3xl p-10 text-center border-2 border-dashed transition-colors"
          style={{ borderColor: dragOver ? "var(--vb-sage-dk)" : "var(--vb-border)", background: dragOver ? "var(--vb-mint-pale)" : "var(--vb-card-bg)" }}
        >
          <CloudUpload size={38} className="mx-auto mb-3" color="var(--vb-sage-dk)" />
          <p className="font-medium mb-1" style={{ color: "var(--vb-ink)" }}>📁 Kéo & thả file Excel {cfg.label} vào đây</p>
          <p className="text-sm mb-4" style={{ color: "var(--vb-muted)" }}>hoặc</p>
          <button onClick={() => fileRef.current && fileRef.current.click()} className="vb-btn-primary mx-auto">Chọn file Excel</button>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => handleFile(e.target.files[0])} />
          <p className="text-xs mt-4" style={{ color: "var(--vb-muted)" }}>Định dạng: .xlsx · Cột bắt buộc: {cfg.required.join(", ")}</p>
        </div>
      )}

      {error && (
        <div className="vb-card rounded-2xl p-4 mt-4 flex items-start gap-3" style={{ background: "#FBEAE4" }}>
          <AlertCircle size={18} color="#B4563F" className="shrink-0 mt-0.5" />
          <p className="text-sm" style={{ color: "#8C3D2A" }}>{error}</p>
        </div>
      )}

      {preview && (
        <div className="vb-card rounded-2xl p-6 mt-2">
          <div className="flex items-center justify-between flex-wrap gap-3 mb-4">
            <div className="font-medium" style={{ color: "var(--vb-ink)" }}>{preview.rows.length} dòng sẵn sàng xem trước</div>
            {dupCount > 0 && (
              <div className="flex items-center gap-2 text-sm">
                <span style={{ color: "var(--vb-muted)" }}>Tìm thấy {dupCount} mục trùng:</span>
                <select className="vb-select" value={preview.dupChoice} onChange={(e) => setPreview((p) => ({ ...p, dupChoice: e.target.value }))}>
                  <option value="skip">Bỏ qua mục trùng</option>
                  <option value="replace">Thay thế mục đã có</option>
                  <option value="import">Vẫn nhập hết</option>
                </select>
              </div>
            )}
          </div>

          <div className="overflow-x-auto rounded-xl" style={{ border: "1px solid var(--vb-border)" }}>
            <table className="w-full text-sm">
              <thead>
                <tr style={{ background: "var(--vb-mint-pale)" }}>
                  <th className="text-left px-4 py-2.5 font-semibold" style={{ color: "var(--vb-sage-dk)" }}>{cfg.columns[0]}</th>
                  <th className="text-left px-4 py-2.5 font-semibold" style={{ color: "var(--vb-sage-dk)" }}>Nghĩa</th>
                  <th className="text-left px-4 py-2.5 font-semibold" style={{ color: "var(--vb-sage-dk)" }}>Ví dụ</th>
                  <th className="text-left px-4 py-2.5 font-semibold" style={{ color: "var(--vb-sage-dk)" }}>Trạng thái</th>
                </tr>
              </thead>
              <tbody className="divide-y" style={{ borderColor: "var(--vb-border)" }}>
                {preview.rows.slice(0, 50).map((r, i) => (
                  <tr key={i}>
                    <td className="px-4 py-2.5 font-medium" style={{ color: "var(--vb-ink)" }}>{r.term}</td>
                    <td className="px-4 py-2.5" style={{ color: "var(--vb-muted)" }}>{r.meaning}</td>
                    <td className="px-4 py-2.5" style={{ color: "var(--vb-muted)" }}>{r.example}</td>
                    <td className="px-4 py-2.5">{r.dup ? <span className="text-xs font-medium" style={{ color: "#C08A34" }}>Trùng</span> : <span className="text-xs font-medium" style={{ color: "#3E7A6A" }}>Mới</span>}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {preview.rows.length > 50 && <p className="text-xs mt-2" style={{ color: "var(--vb-muted)" }}>Đang hiển thị 50/{preview.rows.length} dòng đầu tiên.</p>}

          <div className="flex gap-2.5 justify-end mt-5">
            <button onClick={() => { setPreview(null); if (fileRef.current) fileRef.current.value = ""; }} className="vb-btn-secondary">Huỷ</button>
            <button onClick={confirmImport} className="vb-btn-primary"><Upload size={15} /> Nhập {cfg.label.toLowerCase()}</button>
          </div>
        </div>
      )}
    </div>
  );
}

/* ============================== EXCEL TEMPLATE ============================== */

function downloadTemplate(typeKey) {
  const cfg = IMPORT_TYPES[typeKey];
  const ws = XLSX.utils.aoa_to_sheet([cfg.columns, ...cfg.sample]);
  ws["!cols"] = cfg.colWidths.map((w) => ({ wch: w }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, cfg.label);
  XLSX.writeFile(wb, cfg.fileName);
}

function TemplateView() {
  const [typeKey, setTypeKey] = useState("vocab");
  const cfg = IMPORT_TYPES[typeKey];

  return (
    <div>
      <SectionHeader title="Mẫu Excel" subtitle="Tải mẫu về, điền dữ liệu, rồi nhập vào ứng dụng." />

      <div className="flex gap-2 mb-5">
        {Object.entries(IMPORT_TYPES).map(([key, t]) => (
          <button key={key} onClick={() => setTypeKey(key)} className={cx("vb-pill-option", typeKey === key && "vb-pill-option-active")}>{t.label}</button>
        ))}
      </div>

      <div className="vb-card rounded-3xl p-7 max-w-3xl mx-auto">
        <div className="flex items-start gap-4 flex-wrap">
          <div className="w-14 h-14 rounded-2xl flex items-center justify-center shrink-0" style={{ background: "var(--vb-mint-pale)" }}>
            <FileSpreadsheet size={26} color="var(--vb-sage-dk)" />
          </div>
          <div className="flex-1 min-w-[200px]">
            <h3 className="vb-heading text-lg" style={{ color: "var(--vb-ink)" }}>{cfg.fileName}</h3>
            <p className="text-sm mt-1" style={{ color: "var(--vb-muted)" }}>Gồm đủ các cột cần thiết, đúng thứ tự, kèm vài dòng mẫu để bạn dễ hình dung.</p>
          </div>
          <button onClick={() => downloadTemplate(typeKey)} className="vb-btn-primary shrink-0"><Download size={15} /> Tải mẫu Excel</button>
        </div>

        <div className="mt-6">
          <div className="text-xs font-semibold uppercase tracking-wide mb-2" style={{ color: "var(--vb-muted)" }}>Các cột, đúng theo thứ tự này</div>
          <div className="flex flex-wrap gap-2">
            {cfg.columns.map((h, i) => (
              <span key={h} className="vb-chip text-xs px-3 py-1.5 flex items-center gap-1.5">
                <span className="font-semibold">{i + 1}.</span> {h}{!cfg.required.includes(h) && " (tuỳ chọn)"}
              </span>
            ))}
          </div>
        </div>

        <div className="mt-6 overflow-x-auto rounded-xl" style={{ border: "1px solid var(--vb-border)" }}>
          <table className="w-full text-sm">
            <thead>
              <tr style={{ background: "var(--vb-mint-pale)" }}>
                {cfg.columns.map((h) => <th key={h} className="text-left px-4 py-2.5 font-semibold" style={{ color: "var(--vb-sage-dk)" }}>{h}</th>)}
              </tr>
            </thead>
            <tbody className="divide-y" style={{ borderColor: "var(--vb-border)" }}>
              {cfg.sample.map((row, i) => (
                <tr key={i}>{row.map((cell, j) => <td key={j} className="px-4 py-2.5" style={{ color: j === 0 ? "var(--vb-ink)" : "var(--vb-muted)" }}>{cell}</td>)}</tr>
              ))}
            </tbody>
          </table>
        </div>

        <p className="text-xs mt-4" style={{ color: "var(--vb-muted)" }}>Điền xong rồi thì qua mục <strong>Nhập từ vựng</strong> để tải file lên nhé.</p>
      </div>
    </div>
  );
}

/* ============================== STYLES ============================== */

const VB_STYLES = `
@import url('https://fonts.googleapis.com/css2?family=Baloo+2:wght@500;600;700;800&family=Quicksand:wght@400;500;600;700&display=swap');

.vb-root {
  --vb-bg: #FBF9F1;
  --vb-card-bg: #FFFFFF;
  --vb-sage-dk: #3E5C3B;
  --vb-mint: #7FB596;
  --vb-mint-dk: #5E9B78;
  --vb-mint-pale: #E7F3E9;
  --vb-honey: #E3B04B;
  --vb-ink: #223526;
  --vb-muted: #6C8570;
  --vb-border: #E2EEDD;
  --vb-accent-shadow: rgba(94, 155, 120, 0.28);
  background:
    radial-gradient(circle at 8% 8%, rgba(255,255,255,0.9) 0, transparent 45%),
    var(--vb-bg);
  font-family: 'Quicksand', -apple-system, sans-serif;
}

/* ---------- Pastel theme variants ---------- */
.vb-theme-pink {
  --vb-bg: #FFF3F7;
  --vb-sage-dk: #B5507B;
  --vb-mint: #F4A6C1;
  --vb-mint-dk: #E683A9;
  --vb-mint-pale: #FCE4ED;
  --vb-honey: #F0B65C;
  --vb-ink: #5C2A3D;
  --vb-muted: #A8748C;
  --vb-border: #F8D7E4;
  --vb-accent-shadow: rgba(230, 131, 169, 0.3);
}
.vb-theme-sky {
  --vb-bg: #F0F8FF;
  --vb-sage-dk: #3E6E93;
  --vb-mint: #8EC5F0;
  --vb-mint-dk: #5FA6DD;
  --vb-mint-pale: #E1F0FC;
  --vb-honey: #E3B04B;
  --vb-ink: #223A4E;
  --vb-muted: #6C8CA1;
  --vb-border: #D8ECFB;
  --vb-accent-shadow: rgba(95, 166, 221, 0.3);
}
.vb-theme-butter {
  --vb-bg: #FFFBEF;
  --vb-sage-dk: #937226;
  --vb-mint: #F6D370;
  --vb-mint-dk: #E8BC3F;
  --vb-mint-pale: #FCF2D4;
  --vb-honey: #E68A5C;
  --vb-ink: #4A3B14;
  --vb-muted: #9C8B57;
  --vb-border: #F6E9BE;
  --vb-accent-shadow: rgba(232, 188, 63, 0.32);
}
.vb-theme-lilac {
  --vb-bg: #F8F3FF;
  --vb-sage-dk: #6D4A96;
  --vb-mint: #C6A6E8;
  --vb-mint-dk: #A879D6;
  --vb-mint-pale: #EFE3FA;
  --vb-honey: #E3B04B;
  --vb-ink: #3D2A54;
  --vb-muted: #8B76A3;
  --vb-border: #E9DAF7;
  --vb-accent-shadow: rgba(168, 121, 214, 0.3);
}

.vb-heading { font-family: 'Baloo 2', 'Quicksand', sans-serif; font-weight: 700; }
.vb-card { background: var(--vb-card-bg); border: 2px solid var(--vb-border); border-radius: 22px; box-shadow: 0 6px 0 var(--vb-border), 0 10px 24px rgba(62, 92, 59, 0.06); }
.vb-hero { background: linear-gradient(135deg, var(--vb-mint-pale) 0%, var(--vb-card-bg) 100%); border: 2px solid var(--vb-border); box-shadow: 0 6px 0 var(--vb-border); }
.vb-hero-chip { display: inline-flex; align-items: center; gap: 6px; background: rgba(255,255,255,0.85); padding: 8px 14px; border-radius: 999px; font-size: 13px; font-weight: 600; border: 2px solid var(--vb-border); }
.vb-btn-primary { display: inline-flex; align-items: center; gap: 7px; background: var(--vb-sage-dk); color: white; padding: 11px 20px; border-radius: 999px; font-size: 14px; font-weight: 700; border: 2px solid var(--vb-sage-dk); box-shadow: 0 4px 0 rgba(0,0,0,0.18); transition: transform 0.12s ease, box-shadow 0.12s ease; }
.vb-btn-primary:hover { transform: translateY(-1px); box-shadow: 0 5px 0 rgba(0,0,0,0.18); }
.vb-btn-primary:active { transform: translateY(2px); box-shadow: 0 1px 0 rgba(0,0,0,0.18); }
.vb-btn-secondary { display: inline-flex; align-items: center; gap: 7px; background: white; color: var(--vb-sage-dk); border: 2px solid var(--vb-border); padding: 10px 18px; border-radius: 999px; font-size: 14px; font-weight: 700; box-shadow: 0 3px 0 var(--vb-border); transition: transform 0.12s ease, box-shadow 0.12s ease; }
.vb-btn-secondary:hover { background: var(--vb-mint-pale); transform: translateY(-1px); }
.vb-btn-secondary:active { transform: translateY(2px); box-shadow: 0 1px 0 var(--vb-border); }
.vb-input { border: 2px solid var(--vb-border); border-radius: 14px; padding: 9px 12px; font-size: 14px; background: white; color: var(--vb-ink); outline: none; font-family: 'Quicksand', sans-serif; font-weight: 600; }
.vb-input:focus { border-color: var(--vb-mint-dk); }
.vb-select { border: 2px solid var(--vb-border); border-radius: 14px; padding: 8px 10px; font-size: 13px; background: white; color: var(--vb-ink); outline: none; font-family: 'Quicksand', sans-serif; font-weight: 600; }
.vb-select-sm { padding: 5px 6px; font-size: 12px; border-radius: 11px; max-width: 110px; }
.vb-chip { background: var(--vb-mint-pale); color: var(--vb-sage-dk); border-radius: 999px; font-weight: 700; }
.vb-chip-toggle { display: inline-flex; align-items: center; gap: 5px; border: 2px solid var(--vb-border); padding: 8px 14px; border-radius: 999px; font-size: 13px; font-weight: 700; color: var(--vb-sage-dk); background: white; }
.vb-chip-toggle-active { background: var(--vb-mint-pale); border-color: var(--vb-mint-dk); }
.vb-icon-btn { width: 34px; height: 34px; border-radius: 999px; display: flex; align-items: center; justify-content: center; color: var(--vb-sage-dk); border: 2px solid transparent; }
.vb-icon-btn:hover { background: var(--vb-mint-pale); border-color: var(--vb-border); }
.vb-star-btn { display: flex; align-items: center; justify-content: center; }
.vb-audio-btn { background: var(--vb-mint-pale); color: var(--vb-sage-dk); border: 2px solid var(--vb-border); }
.vb-audio-btn:hover { background: var(--vb-mint-dk); color: white; }
.vb-row-hover:hover { background: var(--vb-mint-pale); }
.vb-example-box { background: var(--vb-mint-pale); color: var(--vb-sage-dk); border-radius: 14px; padding: 10px 14px; border: 2px dashed var(--vb-border); }
.vb-deck-option { display: flex; flex-direction: column; align-items: flex-start; gap: 3px; padding: 13px 15px; border-radius: 18px; border: 2px solid var(--vb-border); background: white; font-size: 13px; font-weight: 700; color: var(--vb-ink); text-align: left; box-shadow: 0 3px 0 var(--vb-border); transition: transform 0.12s ease; }
.vb-deck-option:hover { transform: translateY(-1px); }
.vb-deck-option-active { background: var(--vb-sage-dk); border-color: var(--vb-sage-dk); color: white; box-shadow: 0 3px 0 rgba(0,0,0,0.2); }
.vb-pill-option { padding: 10px 12px; border-radius: 999px; border: 2px solid var(--vb-border); background: white; font-size: 13px; font-weight: 700; color: var(--vb-ink); box-shadow: 0 3px 0 var(--vb-border); transition: transform 0.12s ease; }
.vb-pill-option:hover { transform: translateY(-1px); }
.vb-pill-option-active { background: var(--vb-mint-dk); border-color: var(--vb-mint-dk); color: white; box-shadow: 0 3px 0 rgba(0,0,0,0.18); }
.vb-grade-btn { padding: 11px 4px; border-radius: 16px; font-size: 13px; font-weight: 700; border: 2px solid transparent; }
.vb-option-btn { display: flex; align-items: center; gap: 10px; padding: 12px 14px; border-radius: 16px; border: 2px solid var(--vb-border); background: white; font-size: 14px; font-weight: 600; color: var(--vb-ink); text-align: left; }
.vb-option-btn:not(:disabled):hover { background: var(--vb-mint-pale); }
.vb-toast { display: flex; align-items: center; gap: 8px; background: white; border: 2px solid var(--vb-border); box-shadow: 0 6px 0 var(--vb-border), 0 10px 24px rgba(34,53,38,0.14); padding: 12px 18px; border-radius: 18px; font-size: 14px; color: var(--vb-ink); font-weight: 600; max-width: 340px; }
.vb-sticker { box-shadow: 0 6px 0 var(--vb-border), 0 12px 26px rgba(34,53,38,0.14); }
.vb-theme-fab { width: 46px; height: 46px; border-radius: 999px; display: flex; align-items: center; justify-content: center; background: var(--vb-sage-dk); color: white; border: 3px solid white; box-shadow: 0 6px 16px rgba(0,0,0,0.2); transition: transform 0.15s ease; }
.vb-theme-fab:hover { transform: scale(1.08) rotate(-8deg); }
.vb-bulk-bar { display: flex; align-items: center; gap: 10px; background: var(--vb-mint-pale); border: 2px dashed var(--vb-mint-dk); border-radius: 18px; padding: 10px 16px; }
.vb-checkbox { width: 19px; height: 19px; border-radius: 7px; accent-color: var(--vb-sage-dk); }
.vb-spin { animation: vb-spin 0.9s linear infinite; }
@keyframes vb-spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
@keyframes vb-pop { 0% { transform: scale(0.9); opacity: 0; } 100% { transform: scale(1); opacity: 1; } }
.vb-pop { animation: vb-pop 0.18s ease-out; }
`;
