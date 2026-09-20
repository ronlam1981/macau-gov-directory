/**
 * 澳門博彩監察協調局 (DICJ) 統計資料自動更新系統
 * Google Apps Script — 安裝於 Google Sheets 中
 *
 * 安裝方式：
 *   1. 打開 Google Sheets
 *   2. 點擊「擴充功能」→「Apps Script」
 *   3. 刪除預設代碼，貼上本文件全部內容
 *   4. 點擊「執行」→「initializeSheets」完成初始化
 *   5. 點擊「執行」→「setupMonthlyTrigger」設定自動更新
 */

// ════════════════════════════════════════════════════════
// 設定
// ════════════════════════════════════════════════════════
const CFG = {
  MONTHLY_SHEET:    "每月幸運博彩毛收入",
  QUARTERLY_SHEET:  "每季各博彩項目",
  SUMMARY_SHEET:    "年度彙總",
  ANALYSIS_SHEET:   "分析報告",
  LOG_SHEET:        "更新日誌",
  YEAR_START:       2002,
  BASE_CN: "https://www.dicj.gov.mo/web/cn/information",
  BASE_EN: "https://www.dicj.gov.mo/web/en/information",
};

const MONTH_MAP = {
  "一月":1,"二月":2,"三月":3,"四月":4,"五月":5,"六月":6,
  "七月":7,"八月":8,"九月":9,"十月":10,"十一月":11,"十二月":12,
  "Jan":1,"Feb":2,"Mar":3,"Apr":4,"May":5,"Jun":6,
  "Jul":7,"Aug":8,"Sep":9,"Oct":10,"Nov":11,"Dec":12,
};
const MONTH_ZH = {1:"一月",2:"二月",3:"三月",4:"四月",5:"五月",6:"六月",
                  7:"七月",8:"八月",9:"九月",10:"十月",11:"十一月",12:"十二月"};
const QTR_ZH   = {1:"第一季",2:"第二季",3:"第三季",4:"第四季"};

const GAME_ALIASES = {
  "百家樂貴賓廳":"百家樂貴賓廳","百家樂大眾廳":"百家樂大眾廳",
  "三卡百家樂":"三卡百家樂","廿一點":"廿一點","輪盤":"輪盤",
  "骰寶":"骰寶","牌九":"牌九","其他桌面博彩":"其他桌面博彩",
  "角子機":"角子機","合計":"合計",
  "VIP Baccarat":"百家樂貴賓廳","Mass Market Baccarat":"百家樂大眾廳",
  "Three-card Baccarat":"三卡百家樂","Blackjack":"廿一點",
  "Roulette":"輪盤","Sic-Bo":"骰寶","Sic Bo":"骰寶",
  "Pai Gow":"牌九","Other Table Games":"其他桌面博彩",
  "Slot Machines":"角子機","Total":"合計",
};


// ════════════════════════════════════════════════════════
// 菜單
// ════════════════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📊 DICJ 博彩統計")
    .addItem("🔄 完整更新（所有年份）", "updateAll")
    .addItem("⚡ 快速更新（本年度）", "updateLatest")
    .addSeparator()
    .addItem("⏰ 設置每月自動更新", "setupMonthlyTrigger")
    .addItem("🗑️ 取消自動更新", "removeTriggers")
    .addSeparator()
    .addItem("📈 重新計算分析報告", "calculateAnalysis")
    .addItem("📋 顯示數據摘要", "showSummaryReport")
    .addSeparator()
    .addItem("🏗️ 初始化工作表結構", "initializeSheets")
    .addToUi();
}


// ════════════════════════════════════════════════════════
// 主要更新函數
// ════════════════════════════════════════════════════════
function updateAll() {
  log_("═══════ 開始完整更新 ═══════");
  const t0 = Date.now();
  initializeSheets();

  const year = new Date().getFullYear();
  let monthly = [], quarterly = [];

  for (let y = CFG.YEAR_START; y <= year; y++) {
    const m = scrapeMonthly_(y);
    if (m.length) { monthly = monthly.concat(m); }
    Utilities.sleep(400 + Math.random() * 300);

    const q = scrapeQuarterly_(y);
    if (q.length) { quarterly = quarterly.concat(q); }
    Utilities.sleep(400 + Math.random() * 300);
  }

  const addedM = writeMonthly_(monthly);
  const addedQ = writeQuarterly_(quarterly);
  updateSummary_();
  calculateAnalysis();

  const sec = Math.round((Date.now() - t0) / 1000);
  const msg = `✅ 完整更新完成！月份新增 ${addedM} 筆，季度新增 ${addedQ} 筆，耗時 ${sec} 秒`;
  log_(msg);
  SpreadsheetApp.getUi().alert(msg);
}

function updateLatest() {
  log_("─── 快速更新（本年度）───");
  const year = new Date().getFullYear();

  const monthly    = scrapeMonthly_(year);
  const quarterly  = scrapeQuarterly_(year);
  const addedM     = monthly.length   ? writeMonthly_(monthly)   : 0;
  const addedQ     = quarterly.length ? writeQuarterly_(quarterly) : 0;

  if (addedM + addedQ > 0) {
    updateSummary_();
    calculateAnalysis();
    log_(`✅ 快速更新完成：月份 +${addedM}，季度 +${addedQ}`);
  } else {
    log_("📊 數據已是最新，無需更新");
  }
}

// 每月自動觸發的入口（靜默執行）
function monthlyAutoUpdate_() {
  try {
    updateLatest();
    log_("✅ 每月自動更新完成");
  } catch (e) {
    log_("❌ 自動更新失敗：" + e.message);
  }
}


// ════════════════════════════════════════════════════════
// 爬蟲函數
// ════════════════════════════════════════════════════════
function fetchPage_(url) {
  const opts = {
    method: "get",
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
      "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
      "Accept-Language": "zh-TW,zh;q=0.9,en-US;q=0.8",
      "Referer": "https://www.dicj.gov.mo/web/cn/information/index.html",
      "Cache-Control": "no-cache",
    },
  };
  try {
    const resp = UrlFetchApp.fetch(url, opts);
    if (resp.getResponseCode() === 200) return resp.getContentText("UTF-8");
  } catch (e) {
    // silent
  }
  return null;
}

function scrapeMonthly_(year) {
  const urls = [
    `${CFG.BASE_CN}/DadosEstat_mensal/${year}/index.html`,
    `${CFG.BASE_EN}/DadosEstat_mensal/${year}/index.html`,
  ];
  for (const url of urls) {
    const html = fetchPage_(url);
    if (!html) continue;
    const rows = parseMonthlyHtml_(html, year);
    if (rows.length > 0) {
      log_(`  ✓ ${year} 年每月：${rows.length} 筆`);
      return rows;
    }
  }
  log_(`  ✗ ${year} 年每月：無法獲取`);
  return [];
}

function scrapeQuarterly_(year) {
  const urls = [
    `${CFG.BASE_CN}/DadosEstat/${year}/content.html`,
    `${CFG.BASE_EN}/DadosEstat/${year}/content.html`,
  ];
  for (const url of urls) {
    const html = fetchPage_(url);
    if (!html) continue;
    const rows = parseQuarterlyHtml_(html, year);
    if (rows.length > 0) {
      log_(`  ✓ ${year} 年季度：${rows.length} 筆`);
      return rows;
    }
  }
  log_(`  ✗ ${year} 年季度：無法獲取`);
  return [];
}


// ════════════════════════════════════════════════════════
// HTML 解析
// ════════════════════════════════════════════════════════
function stripTags_(s) {
  return (s || "").replace(/<[^>]+>/g, "").replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").trim();
}

function parseNum_(s) {
  const c = (s || "").replace(/,/g, "").trim();
  if (!c || c === "--" || c === "N/A" || c === "-") return null;
  const n = parseFloat(c);
  return isNaN(n) ? null : n;
}

function getCells_(tr) {
  const cells = [];
  const re = /<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi;
  let m;
  while ((m = re.exec(tr)) !== null) cells.push(stripTags_(m[1]));
  return cells;
}

function parseMonthlyHtml_(html, year) {
  const rows = [];
  const trRe = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  let tr;
  while ((tr = trRe.exec(html)) !== null) {
    const cells = getCells_(tr[0]);
    if (cells.length < 2) continue;

    let month = null;
    const c0 = cells[0];
    for (const [k, v] of Object.entries(MONTH_MAP)) {
      if (c0.includes(k)) { month = v; break; }
    }
    if (!month) continue;

    rows.push({
      year, month,
      rev: parseNum_(cells[1]),
      mom: parseNum_(cells[2]),
      yoy: parseNum_(cells[3]),
      ytd: parseNum_(cells[4]),
    });
  }
  return rows;
}

function parseQuarterlyHtml_(html, year) {
  // Strip scripts/styles
  const clean = html
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<style[\s\S]*?<\/style>/gi, "");

  const tableRe = /<table[^>]*>([\s\S]*?)<\/table>/gi;
  const rows = [];
  let quarter = 1;
  let t;

  while ((t = tableRe.exec(clean)) !== null) {
    const tableHtml = t[0];

    // 從 caption 中解析季度
    const capM = tableHtml.match(/<caption[^>]*>([\s\S]*?)<\/caption>/i);
    if (capM) {
      const cap = stripTags_(capM[1]);
      const qM = cap.match(/第\s*([一二三四1-4])\s*季/);
      if (qM) {
        const qMap = {"一":1,"二":2,"三":3,"四":4,"1":1,"2":2,"3":3,"4":4};
        quarter = qMap[qM[1]] || quarter;
      }
      const qM2 = cap.match(/[Qq]([1-4])/);
      if (qM2) quarter = parseInt(qM2[1]);
    }

    const trRe2 = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
    let tr2;
    let tableHasData = false;

    while ((tr2 = trRe2.exec(tableHtml)) !== null) {
      const cells = getCells_(tr2[0]);
      if (cells.length < 2) continue;

      let game = null;
      const c0 = cells[0];
      for (const [alias, std] of Object.entries(GAME_ALIASES)) {
        if (c0.includes(alias)) { game = std; break; }
      }
      if (!game) continue;

      rows.push({
        year, quarter, game,
        tables: parseNum_(cells[1]),
        rev:    parseNum_(cells[2]),
        pct:    parseNum_(cells[3]),
      });
      tableHasData = true;
    }

    if (tableHasData) quarter = Math.min(quarter + 1, 4);
  }
  return rows;
}


// ════════════════════════════════════════════════════════
// 寫入工作表
// ════════════════════════════════════════════════════════
function writeMonthly_(data) {
  if (!data.length) return 0;
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.MONTHLY_SHEET);
  if (!ws) return 0;

  const existingKeys = getExistingKeys_(ws, [1, 2], 2);
  const newRows = [];

  for (const r of data) {
    const key = `${r.year}|${r.month}`;
    if (existingKeys.has(key)) continue;
    newRows.push([r.year, MONTH_ZH[r.month] || r.month, r.rev, r.mom, r.yoy, r.ytd]);
    existingKeys.add(key);
  }

  if (newRows.length) {
    ws.getRange(ws.getLastRow() + 1, 1, newRows.length, 6).setValues(newRows);
    // 按年份+月份數字排序
    sortMonthlySheet_(ws);
  }
  return newRows.length;
}

function writeQuarterly_(data) {
  if (!data.length) return 0;
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.QUARTERLY_SHEET);
  if (!ws) return 0;

  const existingKeys = getExistingKeys_(ws, [1, 2, 3], 2);
  const newRows = [];

  for (const r of data) {
    const key = `${r.year}|${r.quarter}|${r.game}`;
    if (existingKeys.has(key)) continue;
    newRows.push([r.year, QTR_ZH[r.quarter] || r.quarter, r.game, r.tables, r.rev, r.pct]);
    existingKeys.add(key);
  }

  if (newRows.length) {
    ws.getRange(ws.getLastRow() + 1, 1, newRows.length, 6).setValues(newRows);
  }
  return newRows.length;
}

function getExistingKeys_(ws, cols, startRow) {
  const keys = new Set();
  const lastRow = ws.getLastRow();
  if (lastRow < startRow) return keys;
  const maxCol = Math.max(...cols);
  const data = ws.getRange(startRow, 1, lastRow - startRow + 1, maxCol).getValues();
  for (const row of data) {
    if (row[0]) keys.add(cols.map(c => row[c - 1]).join("|"));
  }
  return keys;
}

function sortMonthlySheet_(ws) {
  const lastRow = ws.getLastRow();
  if (lastRow <= 2) return;
  // 在 col A（年份）和 col B（月份文字→數字）排序
  const range = ws.getRange(2, 1, lastRow - 1, 6);
  const data = range.getValues();
  const monthOrder = {
    "一月":1,"二月":2,"三月":3,"四月":4,"五月":5,"六月":6,
    "七月":7,"八月":8,"九月":9,"十月":10,"十一月":11,"十二月":12
  };
  data.sort((a, b) => {
    if (a[0] !== b[0]) return a[0] - b[0];
    return (monthOrder[a[1]] || 99) - (monthOrder[b[1]] || 99);
  });
  range.setValues(data);
}


// ════════════════════════════════════════════════════════
// 年度彙總
// ════════════════════════════════════════════════════════
function updateSummary_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mws = ss.getSheetByName(CFG.MONTHLY_SHEET);
  const sws = ss.getSheetByName(CFG.SUMMARY_SHEET);
  if (!mws || !sws) return;

  const lastRow = mws.getLastRow();
  if (lastRow <= 1) return;
  const data = mws.getRange(2, 1, lastRow - 1, 6).getValues();

  // 按年份彙總（優先使用累計列的最大值）
  const yearMap = {};
  for (const r of data) {
    const y = r[0], rev = r[2], ytd = r[5];
    if (!y) continue;
    if (!yearMap[y]) yearMap[y] = { sumRev: 0, maxYtd: 0, count: 0 };
    if (rev) { yearMap[y].sumRev += rev; yearMap[y].count++; }
    if (ytd && ytd > yearMap[y].maxYtd) yearMap[y].maxYtd = ytd;
  }

  const years = Object.keys(yearMap).map(Number).sort((a,b) => a - b);
  const rows = [["年份", "全年毛收入（百萬澳門幣）", "較上年增減（%）", "備注"]];

  for (let i = 0; i < years.length; i++) {
    const y = years[i];
    const rev = yearMap[y].maxYtd || yearMap[y].sumRev;
    const prevRev = i > 0 ? (yearMap[years[i-1]].maxYtd || yearMap[years[i-1]].sumRev) : null;
    const yoy = prevRev ? Math.round((rev / prevRev - 1) * 1000) / 10 : null;
    const note = yearMap[y].count < 12 ? `*${yearMap[y].count}月數據` : "";
    rows.push([y, Math.round(rev), yoy, note]);
  }

  sws.clearContents();
  sws.getRange(1, 1, rows.length, 4).setValues(rows);
  sws.getRange(1, 1, 1, 4)
    .setBackground("#833C00").setFontColor("#FFF").setFontWeight("bold");
  sws.getRange(2, 2, rows.length - 1, 1).setNumberFormat("#,##0");
}


// ════════════════════════════════════════════════════════
// 分析報告
// ════════════════════════════════════════════════════════
function calculateAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mws = ss.getSheetByName(CFG.MONTHLY_SHEET);
  const aws = ss.getSheetByName(CFG.ANALYSIS_SHEET);
  if (!mws || !aws) return;

  const lastRow = mws.getLastRow();
  if (lastRow <= 1) return;
  const raw = mws.getRange(2, 1, lastRow - 1, 6).getValues();

  // ── 解析數據 ────────────────────────────────────────────
  const series = [];  // {year, month, rev}
  const monthOrder = {
    "一月":1,"二月":2,"三月":3,"四月":4,"五月":5,"六月":6,
    "七月":7,"八月":8,"九月":9,"十月":10,"十一月":11,"十二月":12
  };

  for (const r of raw) {
    if (!r[0] || r[2] == null) continue;
    const m = monthOrder[r[1]];
    if (!m) continue;
    series.push({ year: r[0], month: m, rev: r[2] });
  }
  series.sort((a, b) => a.year !== b.year ? a.year - b.year : a.month - b.month);

  // ── 季節性指數 ──────────────────────────────────────────
  const monthSum = Array(12).fill(0);
  const monthCnt = Array(12).fill(0);
  let grandSum = 0;
  for (const s of series) {
    monthSum[s.month - 1] += s.rev;
    monthCnt[s.month - 1]++;
    grandSum += s.rev;
  }
  const grandMean = grandSum / series.length;
  const seasonal = monthSum.map((s, i) =>
    monthCnt[i] > 0 ? Math.round((s / monthCnt[i]) / grandMean * 1000) / 1000 : 1
  );

  // ── 線性趨勢回歸（最近 36 個月）────────────────────────
  const recent = series.slice(-36);
  const n = recent.length;
  let sx = 0, sy = 0, sxx = 0, sxy = 0;
  for (let i = 0; i < n; i++) {
    sx += i; sy += recent[i].rev;
    sxx += i * i; sxy += i * recent[i].rev;
  }
  const slope = (n * sxy - sx * sy) / (n * sxx - sx * sx);
  const intercept = (sy - slope * sx) / n;

  // ── 未來 3 個月預測 ─────────────────────────────────────
  const last = recent[n - 1];
  const forecasts = [];
  for (let f = 1; f <= 3; f++) {
    let fM = last.month + f, fY = last.year;
    if (fM > 12) { fM -= 12; fY++; }
    const trend = intercept + slope * (n + f - 1);
    const adj = Math.max(0, trend * (seasonal[fM - 1] || 1));
    forecasts.push({
      label: `${fY}/${MONTH_ZH[fM]}`,
      forecast: Math.round(adj),
      low: Math.round(adj * 0.90),
      high: Math.round(adj * 1.10),
    });
  }

  // ── 最近12個月 YoY ─────────────────────────────────────
  const recent12 = series.slice(-12);
  const yoyRows = recent12.map(s => {
    const prev = series.find(p => p.year === s.year - 1 && p.month === s.month);
    const yoy = prev ? Math.round((s.rev / prev.rev - 1) * 1000) / 10 : null;
    return [`${s.year}/${MONTH_ZH[s.month]}`, s.rev, yoy];
  });

  // ── 寫入分析工作表 ──────────────────────────────────────
  const now = new Date().toLocaleString("zh-TW");
  const output = [
    [`澳門幸運博彩統計 — 分析報告`, "", "", ""],
    [`更新時間：${now}`, "", "", ""],
    ["", "", "", ""],

    ["▌ 一、季節性指數（月份 vs 全年均值）", "", "", ""],
    ["月份", "季節性指數", "偏差", "說明"],
    ...seasonal.map((v, i) => [
      MONTH_ZH[i + 1],
      v,
      `${((v - 1) * 100).toFixed(1)}%`,
      v > 1.1 ? "🔺旺季" : v < 0.9 ? "🔻淡季" : "一般",
    ]),

    ["", "", "", ""],
    ["▌ 二、未來3個月趨勢預測（線性趨勢 × 季節調整）", "", "", ""],
    ["預測月份", "預測毛收入（百萬MOP）", "下限 (−10%)", "上限 (+10%)"],
    ...forecasts.map(f => [f.label, f.forecast, f.low, f.high]),

    ["", "", "", ""],
    ["▌ 三、最近12個月 YoY 表現", "", "", ""],
    ["月份", "毛收入（百萬MOP）", "同比增減（%）", ""],
    ...yoyRows.map(r => [r[0], r[1], r[2], ""]),

    ["", "", "", ""],
    ["▌ 四、回歸模型參數（基於最近36個月）", "", "", ""],
    ["月均趨勢增量（百萬MOP）", Math.round(slope * 10) / 10, "", ""],
    ["基期截距（百萬MOP）",      Math.round(intercept),     "", ""],
    ["訓練月份數",               n,                         "", ""],
  ];

  aws.clearContents();
  aws.getRange(1, 1, output.length, 4).setValues(output);

  // 格式化
  aws.getRange(1,1).setFontSize(14).setFontWeight("bold");
  const hdrBg = "#1F5C5C";
  [4, 8, 12, 16].forEach(r => {
    const row = aws.getRange(r, 1);
    row.setBackground(hdrBg).setFontColor("#FFF").setFontWeight("bold");
  });

  log_("✅ 分析報告已更新");
}


// ════════════════════════════════════════════════════════
// 觸發器
// ════════════════════════════════════════════════════════
function setupMonthlyTrigger() {
  removeTriggers();
  ScriptApp.newTrigger("monthlyAutoUpdate_")
    .timeBased().onMonthDay(5).atHour(9).create();
  log_("✅ 已設定：每月 5 日上午 9:00 自動更新");
  SpreadsheetApp.getUi().alert("✅ 已設定每月自動更新（每月5日 09:00）");
}

function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (["updateLatest","updateAll","monthlyAutoUpdate_"].includes(t.getHandlerFunction())) {
      ScriptApp.deleteTrigger(t);
    }
  });
}


// ════════════════════════════════════════════════════════
// 數據摘要（彈出視窗顯示，不需要郵件權限）
// ════════════════════════════════════════════════════════
function showSummaryReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sws = ss.getSheetByName(CFG.SUMMARY_SHEET);
  if (!sws || sws.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("尚無數據，請先執行「更新最新數據」。");
    return;
  }
  const data = sws.getRange(2, 1, sws.getLastRow() - 1, 4).getValues();
  const latest = data[data.length - 1];
  const prev   = data[data.length - 2];
  const yoy    = latest[2] != null ? (latest[2] >= 0 ? "+" : "") + latest[2] + "%" : "N/A";
  const msg =
    `📊 澳門博彩統計摘要\n\n` +
    `${latest[0]} 年全年毛收入：${Number(latest[1]).toLocaleString()} 百萬 MOP\n` +
    `較 ${prev[0]} 年：${yoy}\n\n` +
    `${prev[0]} 年全年毛收入：${Number(prev[1]).toLocaleString()} 百萬 MOP`;
  SpreadsheetApp.getUi().alert(msg);
}

// 保留舊名稱相容性（空函式，不需要郵件權限）
function sendEmailReport() {
  showSummaryReport();
}


// ════════════════════════════════════════════════════════
// 工作表初始化
// ════════════════════════════════════════════════════════
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const defs = [
    {
      name: CFG.MONTHLY_SHEET,
      hdrs: ["年份","月份","幸運博彩毛收入（百萬澳門幣）","較上月變動（%）","較去年同期變動（%）","本年累計毛收入（百萬澳門幣）"],
      bg: "#1F4E79", widths: [60,70,200,120,160,200],
    },
    {
      name: CFG.QUARTERLY_SHEET,
      hdrs: ["年份","季度","博彩項目","博彩桌/機數","幸運博彩毛收入（百萬澳門幣）","佔總毛收入（%）"],
      bg: "#1F5C5C", widths: [60,70,150,110,200,120],
    },
    {
      name: CFG.SUMMARY_SHEET,
      hdrs: ["年份","全年毛收入（百萬澳門幣）","較上年增減（%）","備注"],
      bg: "#833C00", widths: [60,200,150,120],
    },
    {
      name: CFG.ANALYSIS_SHEET,
      hdrs: null,
      bg: "#1F2D40", widths: [180,180,180,180],
    },
    {
      name: CFG.LOG_SHEET,
      hdrs: ["時間","訊息"],
      bg: "#404040", widths: [160,500],
    },
  ];

  for (const d of defs) {
    let ws = ss.getSheetByName(d.name);
    if (!ws) ws = ss.insertSheet(d.name);

    if (d.hdrs && ws.getLastRow() === 0) {
      const r = ws.getRange(1, 1, 1, d.hdrs.length);
      r.setValues([d.hdrs])
       .setBackground(d.bg)
       .setFontColor("#FFFFFF")
       .setFontWeight("bold")
       .setHorizontalAlignment("center")
       .setWrap(true);
      ws.setFrozenRows(1);
      ws.getRange(1, 1, 1, d.hdrs.length).setRowHeight(40);
    }

    d.widths.forEach((w, i) => ws.setColumnWidth(i + 1, w));
  }

  log_("✅ 工作表結構初始化完成");
}


// ════════════════════════════════════════════════════════
// 日誌
// ════════════════════════════════════════════════════════
function log_(msg) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(CFG.LOG_SHEET);
  if (!ws) return;

  ws.appendRow([new Date().toLocaleString("zh-TW"), msg]);
  console.log(msg);

  // 保留最新 2000 行
  const n = ws.getLastRow();
  if (n > 2001) ws.deleteRows(2, n - 2001);
}
