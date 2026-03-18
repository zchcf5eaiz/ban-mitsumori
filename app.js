/**
 * 盤見積もりアプリ - アプリケーションロジック (app.js)
 *
 * タブ構成:
 *   1. 管理表（盤）       - PDFレイアウト固定、カテゴリ見出し付き、単価編集可
 *   2. 管理表（キュービクル）- キュービクル専用マスタ
 *   3. 見積作成           - 空テーブル＋品目追加方式、区切り線機能
 */

const STORAGE_MASTER         = "ban_master_v6";
const STORAGE_ESTIMATES      = "ban_estimates_v6";
const STORAGE_PRICES         = "ban_prices_v6";
const STORAGE_CUBICLE        = "ban_cubicle_v1";
const STORAGE_CUBICLE_PRICES = "ban_cubicle_prices_v1";
const STORAGE_GROUP_TOTALS   = "ban_group_totals_v1";
const STORAGE_SR1_COMMENTS   = "ban_sr1_comments_v2";

// データバージョン: この値を上げるとlocalStorageのマスタを破棄してデフォルトに戻す
const DATA_VERSION = 6;
const STORAGE_DATA_VERSION = "ban_data_version";

if (parseInt(localStorage.getItem(STORAGE_DATA_VERSION) || "0") < DATA_VERSION) {
  // 全PCで古いデータを破棄してデフォルト値（data.js）を使わせる
  localStorage.removeItem(STORAGE_MASTER);
  localStorage.removeItem(STORAGE_PRICES);
  localStorage.removeItem(STORAGE_CUBICLE_PRICES);
  localStorage.removeItem("ban_option_prices");
  localStorage.removeItem(STORAGE_GROUP_TOTALS);
  localStorage.removeItem(STORAGE_SR1_COMMENTS);
  localStorage.setItem(STORAGE_DATA_VERSION, String(DATA_VERSION));
}

// 割増率チェックボックス: 各マトリクスの基本価格を保持
const matrixBasePrices = {};

// ============================================================
// マスタ管理（盤）
// ============================================================

function loadMaster() {
  try {
    const d = localStorage.getItem(STORAGE_MASTER);
    if (d) return JSON.parse(d);
  } catch {}
  return JSON.parse(JSON.stringify(DEFAULT_MASTER_ITEMS));
}

function saveMaster() {
  localStorage.setItem(STORAGE_MASTER, JSON.stringify(masterItems));
  const prices = {};
  for (const item of masterItems) {
    if (item.basePrice) prices[item.id] = item.basePrice;
  }
  localStorage.setItem(STORAGE_PRICES, JSON.stringify(prices));
}

function resetMaster() {
  if (!confirm("盤マスタを初期状態（PDF初期データ）に戻します。\n保存済み価格は維持されます。よろしいですか？")) return;
  const backup = {};
  for (const item of masterItems) {
    backup[item.id] = item.basePrice;
  }
  masterItems = JSON.parse(JSON.stringify(DEFAULT_MASTER_ITEMS));
  for (const item of masterItems) {
    if (backup[item.id] !== undefined) item.basePrice = backup[item.id];
  }
  saveMaster();
  renderMasterTable();
  showToast("盤マスタを初期状態に戻しました（保存済み価格は維持）");
}

// ============================================================
// マスタ管理（キュービクル）
// ============================================================

/**
 * キュービクル価格管理 — シンプル設計
 * - データ構造は常に DEFAULT_CUBICLE_ITEMS から生成（不変の真実）
 * - 保存するのは価格マップ {id: basePrice} のみ（STORAGE_CUBICLE_PRICES）
 * - 旧形式 STORAGE_CUBICLE からもマイグレーション対応
 */
function loadCubicle() {
  // 1) 保存済み価格マップを読み込む
  let prices = {};
  try {
    const p = localStorage.getItem(STORAGE_CUBICLE_PRICES);
    if (p) prices = JSON.parse(p);
  } catch {}
  // 2) 旧形式（配列全体保存）からの移行: 価格マップになければ旧データから取得
  try {
    const old = localStorage.getItem(STORAGE_CUBICLE);
    if (old) {
      const arr = JSON.parse(old);
      for (const it of arr) {
        if (it.id && it.basePrice && prices[it.id] == null) {
          prices[it.id] = it.basePrice;
        }
      }
      // 旧形式は不要になったので削除
      localStorage.removeItem(STORAGE_CUBICLE);
    }
  } catch {}
  // 3) デフォルトから構築し、保存済み価格を上書き
  const items = JSON.parse(JSON.stringify(DEFAULT_CUBICLE_ITEMS));
  for (const item of items) {
    if (prices[item.id] != null) {
      item.basePrice = prices[item.id];
    }
  }
  return items;
}

function saveCubicle() {
  const prices = {};
  for (const item of cubicleItems) {
    prices[item.id] = item.basePrice;
  }
  try {
    localStorage.setItem(STORAGE_CUBICLE_PRICES, JSON.stringify(prices));
  } catch(e) {
    console.error("saveCubicle failed:", e);
  }
}

function resetCubicle() {
  if (!confirm("キュービクルマスタを初期状態に戻します。\n保存済み価格は維持されます。よろしいですか？")) return;
  // 現在の価格マップを取得（メモリから）
  const prices = {};
  for (const item of cubicleItems) {
    prices[item.id] = item.basePrice;
  }
  // DOM上の合計欄も保存
  document.querySelectorAll(".mg-solar-total-input").forEach(inp => {
    const g = inp.dataset.group;
    if (g) groupTotals[g] = parseFloat(inp.value) || 0;
  });
  saveGroupTotals();
  // デフォルトから再構築して価格を復元
  cubicleItems = JSON.parse(JSON.stringify(DEFAULT_CUBICLE_ITEMS));
  for (const item of cubicleItems) {
    if (prices[item.id] != null) {
      item.basePrice = prices[item.id];
    }
  }
  saveCubicle();
  renderCubicleTable();
  showToast("キュービクルマスタを初期状態に戻しました（保存済み価格は維持）");
}

// ============================================================
// 空調盤SR-1 コメント管理
// ============================================================

function loadSr1Comments() {
  try {
    const d = localStorage.getItem(STORAGE_SR1_COMMENTS);
    if (d) {
      const saved = JSON.parse(d);
      // 保存値が空文字ならデフォルト値で補完
      for (const key of Object.keys(DEFAULT_SR1_COMMENTS)) {
        if (!saved[key] && DEFAULT_SR1_COMMENTS[key]) {
          saved[key] = DEFAULT_SR1_COMMENTS[key];
        }
      }
      return saved;
    }
  } catch {}
  return JSON.parse(JSON.stringify(DEFAULT_SR1_COMMENTS));
}

function saveSr1Comments() {
  localStorage.setItem(STORAGE_SR1_COMMENTS, JSON.stringify(sr1Comments));
}

function onSr1Comment(groupName, value) {
  sr1Comments[groupName] = value;
  saveSr1Comments();
}

// ============================================================
// 価格保存（盤 + キュービクル両方）
// ============================================================

/** 価格を保存（盤 + キュービクル） */
function savePrices() {
  // DOM上のinput値をメモリに反映
  document.querySelectorAll("[data-item-id]").forEach(el => {
    const id = el.dataset.itemId;
    const inp = el.querySelector("input[type='number']");
    if (inp && id) {
      const val = parseFloat(inp.value) || 0;
      const m = getMasterItem(id);
      if (m) m.basePrice = val;
    }
  });
  // 盤
  saveMaster();
  // キュービクル
  saveCubicle();

  // グループ合計欄（K2太陽光等）をDOMから収集して保存
  document.querySelectorAll(".mg-solar-total-input").forEach(inp => {
    const g = inp.dataset.group;
    if (g) groupTotals[g] = parseFloat(inp.value) || 0;
  });
  saveGroupTotals();

  // オプション + 掛率input を一括収集
  const optPrices = {};
  document.querySelectorAll('.pbox-calc input[id^="opt-price-"], .pbox-calc input[id^="opt2-price-"], .pbox-calc input[id^="quick-price-"], .pbox-calc input[id*="-rate-"]').forEach(inp => {
    optPrices[inp.id] = parseFloat(inp.value) || 0;
  });
  localStorage.setItem("ban_option_prices", JSON.stringify(optPrices));

  showToast("価格を保存しました（盤 + キュービクル）");
}

/** 保存済み価格を盤マスタに適用 */
function applySavedPrices() {
  try {
    const d = localStorage.getItem(STORAGE_PRICES);
    if (!d) return;
    const prices = JSON.parse(d);
    for (const item of masterItems) {
      if (prices[item.id] !== undefined) item.basePrice = prices[item.id];
    }
  } catch {}
}

/** 保存済みオプション価格をinputに適用 */
function applySavedOptionPrices() {
  try {
    let d = localStorage.getItem("ban_option_prices");
    if (!d && typeof DEFAULT_OPTION_PRICES !== "undefined") {
      d = JSON.stringify(DEFAULT_OPTION_PRICES);
    }
    if (!d) return;
    const prices = JSON.parse(d);
    for (const id in prices) {
      const el = document.getElementById(id);
      if (el) el.value = prices[id];
    }
  } catch {}
}

// ============================================================
// グループ合計管理（K2太陽光等の合計欄）
// ============================================================

let groupTotals = {};
function loadGroupTotals() {
  try {
    const d = localStorage.getItem(STORAGE_GROUP_TOTALS);
    if (d) { groupTotals = JSON.parse(d); return; }
  } catch {}
  if (typeof DEFAULT_GROUP_TOTALS !== "undefined") {
    groupTotals = JSON.parse(JSON.stringify(DEFAULT_GROUP_TOTALS));
  }
}
function saveGroupTotals() {
  try {
    localStorage.setItem(STORAGE_GROUP_TOTALS, JSON.stringify(groupTotals));
  } catch(e) { console.error("saveGroupTotals failed:", e); }
}
function onGroupTotalChange(groupName, val) {
  groupTotals[groupName] = parseFloat(val) || 0;
  saveGroupTotals();
}

// ============================================================
// マスタ初期化
// ============================================================

let masterItems = loadMaster();
let cubicleItems = loadCubicle();
loadGroupTotals();
let sr1Comments = loadSr1Comments();
saveSr1Comments(); // デフォルト補完した値を即保存

// クリック回数トラッキング（選択色の濃さ管理）
let masterClickCounts = {};

/** 盤 + キュービクル両方から品目を検索 */
function getMasterItem(id) {
  return masterItems.find(m => m.id === id) || cubicleItems.find(m => m.id === id);
}

/** IDがキュービクルアイテムかどうか判定 */
function isCubicleItem(id) {
  return id && id.startsWith("K");
}

/** 保存デバウンス（oninput対策: メモリ更新は即時、localStorage保存は遅延） */
let _saveCubicleTimer = null;
let _saveMasterTimer = null;
function debouncedSaveCubicle() {
  if (_saveCubicleTimer) clearTimeout(_saveCubicleTimer);
  _saveCubicleTimer = setTimeout(() => { saveCubicle(); _saveCubicleTimer = null; }, 300);
}
function debouncedSaveMaster() {
  if (_saveMasterTimer) clearTimeout(_saveMasterTimer);
  _saveMasterTimer = setTimeout(() => { saveMaster(); _saveMasterTimer = null; }, 300);
}

/** 価格変更時の保存先を自動判定 */
function onMasterPriceChange(id, field, val) {
  const m = getMasterItem(id);
  if (m) {
    m[field] = parseFloat(val) || 0;
    if (isCubicleItem(id)) debouncedSaveCubicle(); else debouncedSaveMaster();
    // 太陽光合計積算価格セル更新
    const row = document.querySelector(`tr[data-item-id="${id}"]`);
    if (row) {
      const table = row.closest("table");
      const totalCell = table && table.querySelector(".mg-solar-total");
      if (totalCell) {
        let sum = 0;
        table.querySelectorAll("tr[data-item-id]").forEach(tr => {
          const inp = tr.querySelector("input[type='number']");
          const itm = getMasterItem(tr.dataset.itemId);
          const qty = itm && itm._qty ? parseInt(itm._qty) || 1 : 1;
          sum += Math.round((parseFloat(inp?.value) || 0) * qty);
        });
        const inp2 = totalCell.querySelector("input");
        if (inp2) inp2.value = sum; else totalCell.textContent = sum;
      }
    }
  }
}
function onMasterNoteChange(id, val) {
  const m = getMasterItem(id);
  if (m) {
    m.note = val;
    if (isCubicleItem(id)) saveCubicle(); else saveMaster();
  }
}
function onMasterExtraChange(id, field, val) {
  const m = getMasterItem(id);
  if (m) {
    m[field] = val;
    if (isCubicleItem(id)) saveCubicle(); else saveMaster();
  }
}

// ============================================================
// 見積もりデータ
// ============================================================

function createNewEstimate() {
  return {
    id: genId(),
    name: "",
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    project: {
      projectName: "", customerName: "", location: "",
      date: new Date().toISOString().split("T")[0],
      estimateNo: "", staff: "",
    },
    lines: [],
    listRate: DEFAULT_RATES.listRate,
    netRate: DEFAULT_RATES.netRate,
    notes: "",
  };
}

let currentEstimate = createNewEstimate();
let savedEstimates = loadSavedEstimates();

// ============================================================
// localStorage（見積もり）
// ============================================================

function loadSavedEstimates() {
  try { return JSON.parse(localStorage.getItem(STORAGE_ESTIMATES) || "[]"); }
  catch { return []; }
}
function persistEstimates() {
  localStorage.setItem(STORAGE_ESTIMATES, JSON.stringify(savedEstimates));
}

function saveCurrentEstimate() {
  if (!currentEstimate.name) {
    const n = prompt("見積もり名を入力:", currentEstimate.project.projectName || "新規見積もり");
    if (!n) return;
    currentEstimate.name = n;
  }
  currentEstimate.updatedAt = new Date().toISOString();
  const i = savedEstimates.findIndex(e => e.id === currentEstimate.id);
  const copy = JSON.parse(JSON.stringify(currentEstimate));
  if (i >= 0) savedEstimates[i] = copy; else savedEstimates.push(copy);
  persistEstimates();
  renderEstimateSelector();
  showToast("保存しました: " + currentEstimate.name);
}

function loadEstimate(id) {
  const e = savedEstimates.find(x => x.id === id);
  if (!e) return;
  currentEstimate = JSON.parse(JSON.stringify(e));
  renderEstimateTab();
  showToast("読み込みました: " + currentEstimate.name);
}

function deleteEstimateById(id) {
  const e = savedEstimates.find(x => x.id === id);
  if (!e || !confirm("「" + e.name + "」を削除しますか？")) return;
  savedEstimates = savedEstimates.filter(x => x.id !== id);
  persistEstimates();
  renderEstimateSelector();
  showToast("削除しました");
}

// ============================================================
// タブ切り替え
// ============================================================

function switchTab(name) {
  document.querySelectorAll(".tab-btn").forEach(t => t.classList.toggle("active", t.dataset.tab === name));
  document.querySelectorAll(".tab-content").forEach(t => t.classList.toggle("active", t.id === "tab-" + name));
  if (name === "master") renderMasterTable();
  if (name === "cubicle") renderCubicleTable();
  if (name === "estimate") renderEstimateTab();
}

// ============================================================
// カテゴリ折りたたみ
function toggleCat(headingEl) {
  const block = headingEl.parentElement;
  block.classList.toggle("collapsed");
}

// 管理表（盤）レンダリング — PDFマトリクス形式
// ============================================================

function renderMasterTable() {
  const container = document.getElementById("master-content");

  // CATEGORIESの定義順でグループ化
  const catMap = {};
  for (const item of masterItems) {
    if (!catMap[item.category]) catMap[item.category] = [];
    catMap[item.category].push(item);
  }
  const catGroups = [];
  for (const cat of CATEGORIES) {
    if ((catMap[cat.id] && catMap[cat.id].length > 0) || cat.id === "G" || cat.id === "I") {
      catGroups.push({ id: cat.id, label: cat.name, items: catMap[cat.id] || [] });
    }
  }

  let html = "";
  for (const cg of catGroups) {
    html += `<div class="cat-block">`;
    html += `<div class="cat-heading" onclick="toggleCat(this)"><span class="cat-toggle">▼</span> ${esc(cg.id)}. ${esc(cg.label)}</div>`;

    // カテゴリ内を品名でさらにグループ化
    const nameGroups = [];
    let lastName = null;
    for (const item of cg.items) {
      if (item.name !== lastName) {
        lastName = item.name;
        nameGroups.push({ name: item.name, items: [] });
      }
      nameGroups[nameGroups.length - 1].items.push(item);
    }

    // footerGroupとして使われている名前を収集
    const footerNames = new Set();
    for (const key in MATRIX_GROUPS) {
      const md = MATRIX_GROUPS[key];
      if (md.footerGroup) footerNames.add(md.footerGroup);
      if (md.footerGroups) md.footerGroups.forEach(fg => footerNames.add(fg.group));
    }
    // 通常nameGroupにfooterを付ける設定
    const NAME_GROUP_FOOTERS = {
      "BOX (筐体)": { group: "マスト", label: "マスト" },
    };
    for (const k in NAME_GROUP_FOOTERS) footerNames.add(NAME_GROUP_FOOTERS[k].group);

    const stackCats = ["F", "H"];
    const gridCats = { "B": "cat-tables-grid4", "D": "cat-tables-grid4", "E": "cat-tables-grid3" };
    const gridSpanNames = { "E": { "接地端子盤(100sq,200A)": 2 } };
    const extraClass = stackCats.includes(cg.id) ? " cat-tables-stack" : (gridCats[cg.id] || "");
    const catTablesClass = "cat-tables" + (extraClass ? " " + extraClass : "");

    // サブグループ定義があればチームごとに横並び
    const subGroups = CAT_SUBGROUPS[cg.id];
    if (subGroups) {
      // 明示的に割り当てられた名前を収集（ワイルドカード展開用）
      const assignedNames = new Set();
      for (const sg of subGroups) {
        if (!sg.names.includes("*")) sg.names.forEach(n => assignedNames.add(n));
      }

      html += `<div class="cat-tables cat-tables-grouped">`;
      for (const sg of subGroups) {
        const isWild = sg.names.includes("*");
        const matchName = (name) => isWild ? !assignedNames.has(name) : sg.names.includes(name);

        html += `<div class="cat-subgroup${sg.stack ? " cat-subgroup-stack" : ""}">`;
        if (sg.label) html += `<div class="cat-sub-heading">${esc(sg.label)}</div>`;
        html += `<div class="cat-subgroup-inner">`;
        // stackNames指定があれば、対象を縦積みカラムにまとめる
        if (sg.stackNames && sg.stackNames.length > 0) {
          html += `<div class="cat-stack-col">`;
          for (const ng of nameGroups) {
            if (!sg.stackNames.includes(ng.name)) continue;
            html += renderNameGroup(ng, MATRIX_GROUPS);
          }
          html += `</div>`;
          for (const ng of nameGroups) {
            if (footerNames.has(ng.name)) continue;
            if (!matchName(ng.name)) continue;
            if (sg.stackNames.includes(ng.name)) continue;
            html += renderNameGroup(ng, MATRIX_GROUPS);
          }
        } else {
          for (const ng of nameGroups) {
            if (footerNames.has(ng.name)) continue;
            if (!matchName(ng.name)) continue;
            html += renderNameGroup(ng, MATRIX_GROUPS);
          }
        }
        html += `</div></div>`;
      }
      html += `</div></div>`;
    } else {
      html += `<div class="${catTablesClass}">`;
      // カテゴリIの場合、カスタム入力テーブルを表示
      if (cg.id === "I") {
        html += renderCustomInputTable();
      }
      // カテゴリGの場合、P-BOX計算機を先頭に挿入
      if (cg.id === "G") {
        html += `<div class="cat-tables-grid4">`;
        html += renderPBoxCalculator();
        html += renderDuctCalculator();
        html += renderFrameCalculator();
        html += renderTrayCalculator();
        html += renderOptionTable();
        html += renderOption2Table();
        html += renderQuickItemsTable();
        html += `</div>`;
      }
      const vStacks = VERTICAL_STACKS[cg.id] || [];
      const vStackMap = {};
      for (const vs of vStacks) {
        for (const name of vs) vStackMap[name] = vs;
      }
      const hPairs = HORIZONTAL_PAIRS[cg.id] || [];
      const hPairMap = {};
      for (const hp of hPairs) {
        for (const name of hp) hPairMap[name] = hp;
      }
      const renderedStacks = new Set();
      const renderedHPairs = new Set();
      // 統合マトリクスの描画
      const mergedDefs = MERGED_MATRICES[cg.id] || [];
      const mergedNames = new Set();
      const singleZoneDefs = SINGLE_ZONE_DEFS[cg.id] || [];
      const allSingleNames = new Set();
      for (const sz of singleZoneDefs) {
        const ns = Array.isArray(sz.names) ? sz.names : sz.names;
        for (const n of ns) allSingleNames.add(n);
      }
      const singleZoneMerged = [];
      for (const md of mergedDefs) {
        if (singleZoneDefs.length > 0 && md.singleZone) {
          singleZoneMerged.push(md);
        } else {
          html += renderMergedMatrix(md, nameGroups);
        }
        for (const g of md.groups) mergedNames.add(g.nameGroup);
        if (md.hideNames) for (const n of md.hideNames) mergedNames.add(n);
      }
      // 単品ゾーン描画
      if (singleZoneDefs.length > 0) {
        let singleBlockNum = 0;
        html += `<div class="cubicle-single-zone">`;
        html += `<div class="cubicle-single-zone-heading">単品</div>`;
        for (const sz of singleZoneDefs) {
          const namesArr = Array.isArray(sz.names) ? sz.names : [...sz.names];
          // 配列順序で描画、なければnameGroups順
          const orderedNgs = [];
          for (const name of namesArr) {
            const ng = nameGroups.find(n => n.name === name);
            if (ng && !mergedNames.has(ng.name)) orderedNgs.push(ng);
          }
          // このゾーンに属するsingleZone統合表を収集
          const zoneMerged = singleZoneMerged.filter(md => {
            const namesSet = new Set(namesArr);
            return namesSet.has(md.groups[0].nameGroup);
          });
          if (orderedNgs.length === 0 && zoneMerged.length === 0) continue;
          html += `<div class="cat-single-grid${sz.cols}">`;
          // singleZone統合表をゾーン内に描画
          for (const md of zoneMerged) {
            singleBlockNum++;
            html += renderMergedMatrix(md, nameGroups, singleBlockNum);
          }
          for (const ng of orderedNgs) {
            singleBlockNum++;
            const isWide = sz.wideNames && sz.wideNames.has(ng.name);
            const spanN = sz.spanNames && sz.spanNames[ng.name];
            if (isWide) html += `<div class="cat-single-wide3">`;
            else if (spanN) html += `<div style="grid-column:span ${spanN}">`;
            html += renderNameGroup(ng, MATRIX_GROUPS, singleBlockNum);
            if (isWide || spanN) html += `</div>`;
          }
          html += `</div>`;
        }
        html += `</div>`;
        // 単品ブロックの外に画像を挿入
        if (allSingleNames.has("MCなど(単品)") && cg.id === "C") {
          html += `<div class="cat-img-row3">`;
          html += `<div class="cat-inline-img"><img src="img/tansen-zurei.png" alt="単線接続図例"></div>`;
          html += `<div class="cat-inline-img"><img src="img/seigyoban1.png" alt="キャビネット形式及び単位装置の記号"></div>`;
          html += `<div class="cat-inline-img"><img src="img/seigyoban2.png" alt="単位装置の機能1"></div>`;
          html += `</div>`;
        } else {
          if (allSingleNames.has("MCなど(単品)")) {
            html += `<div class="cat-inline-img"><img src="img/tansen-zurei.png" alt="単線接続図例"></div>`;
          }
          if (allSingleNames.has("フロートスイッチ") && cg.id === "C") {
            html += `<div class="cat-inline-img"><img src="img/seigyoban1.png" alt="キャビネット形式及び単位装置の記号"></div>`;
            html += `<div class="cat-inline-img"><img src="img/seigyoban2.png" alt="単位装置の機能1"></div>`;
          }
        }
      }
      for (const ng of nameGroups) {
        if (footerNames.has(ng.name)) continue;
        if (mergedNames.has(ng.name)) continue;
        if (allSingleNames.has(ng.name)) continue;
        // 空調盤SR-1: 画像と積算表を横並び
        if ((ng.name === "空調盤SR-1" || ng.name === "空調盤SR-1(コインタイマー)") && cg.id === "J") {
          const imgFile = ng.name === "空調盤SR-1" ? "SR-1.png" : "SR-1-coin.png";
          const cmtVal = sr1Comments[ng.name] || "";
          html += `<div class="cat-subgroup"><div class="cat-sub-heading">${esc(ng.name)}</div><div class="cat-subgroup-inner cat-sr1-row">`;
          html += `<div class="cat-sr1-img"><img src="img/${imgFile}" alt="${esc(ng.name)} 図面">`;
          html += `<div style="margin-top:8px;"><label style="font-weight:bold;font-size:13px;">コメント:</label>`;
          const cmtRows = Math.max(2, (cmtVal.match(/\n/g) || []).length + 1);
          html += `<textarea rows="${cmtRows}" style="width:100%;margin-top:4px;resize:vertical;box-sizing:border-box;" onchange="onSr1Comment('${esc(ng.name)}', this.value)">${esc(cmtVal)}</textarea>`;
          html += `</div></div>`;
          html += `<div class="cat-sr1-table">`;
          html += renderNameGroup(ng, MATRIX_GROUPS);
          html += `</div>`;
          html += `</div></div>`;
          continue;
        }
        const hp = hPairMap[ng.name];
        if (hp) {
          const key = hp.join("|");
          if (renderedHPairs.has(key)) continue;
          renderedHPairs.add(key);
          html += `<div class="cat-hpair">`;
          for (const hName of hp) {
            const hng = nameGroups.find(g => g.name === hName);
            if (hng) html += renderNameGroup(hng, MATRIX_GROUPS);
          }
          html += `</div>`;
        } else {
          const vs = vStackMap[ng.name];
          if (vs) {
            const key = vs.join("|");
            if (renderedStacks.has(key)) continue;
            renderedStacks.add(key);
            html += `<div class="cat-stack-col">`;
            for (const vName of vs) {
              const vng = nameGroups.find(g => g.name === vName);
              if (vng) html += renderNameGroup(vng, MATRIX_GROUPS);
            }
            html += `</div>`;
          } else {
            const catSpans = gridSpanNames[cg.id];
            const span = catSpans && catSpans[ng.name];
            if (span) html += `<div style="grid-column:span ${span}">`;
            const ngFooter = NAME_GROUP_FOOTERS[ng.name];
            if (ngFooter) {
              const fng = nameGroups.find(g => g.name === ngFooter.group);
              let ngHtml = renderNameGroup(ng, MATRIX_GROUPS);
              if (fng) {
                let fh = `<table class="mg-table mg-footer-table"><tbody><tr><td class="mg-row-group-header" colspan="2">${esc(ngFooter.label)}</td></tr>`;
                for (const fi of fng.items) {
                  const cnt = masterClickCounts[fi.id] || 0;
                  const bgStyle = cnt > 0 ? `background:rgba(202,138,4,${Math.min(cnt * 0.25, 0.85)})` : "";
                  fh += `<tr class="mg-clickable" style="${bgStyle}" onclick="addFromMaster(event,'${fi.id}')">`;
                  fh += `<td class="mg-row-label">${esc(fi.spec)}</td>`;
                  fh += `<td class="mg-price mg-matrix-cell">`;
                  fh += `<input type="number" min="0" step="0.1" value="${fi.basePrice}"
                        oninput="onMasterPriceChange('${fi.id}','basePrice',this.value)"
                        onchange="onMasterPriceChange('${fi.id}','basePrice',this.value)" onclick="event.stopPropagation()">`;
                  fh += `<span class="mg-cell-overlay" title="クリックで見積に追加"></span>`;
                  fh += `</td></tr>`;
                }
                fh += `</tbody></table>`;
                ngHtml = ngHtml.replace(/<\/div>\s*$/, fh + '</div>');
              }
              html += ngHtml;
            } else {
              html += renderNameGroup(ng, MATRIX_GROUPS);
            }
            if (span) html += `</div>`;
          }
        }
        // 特定の名前グループの後に画像を挿入
        if (ng.name === "MCなど(単品)") {
          html += `<div class="cat-inline-img"><img src="img/tansen-zurei.png" alt="単線接続図例"></div>`;
        }
        if (ng.name === "フロートスイッチ") {
          html += `<div class="cat-inline-img"><img src="img/seigyoban1.png" alt="キャビネット形式及び単位装置の記号"></div>`;
          html += `<div class="cat-inline-img"><img src="img/seigyoban2.png" alt="単位装置の機能1"></div>`;
        }
        if (ng.name === "その他" && cg.id === "F") {
          html += `<div class="cat-inline-img"><img src="img/jakuden-list.png" alt="弱電端子盤リスト" style="max-width:600px"></div>`;
        }
      }
      html += `</div></div>`;
    }
  }
  container.innerHTML = html;
  applySavedOptionPrices();
  updateAddedCount();
}

// ============================================================
// P-BOX 寸法入力型計算機
// ============================================================

function renderPBoxCalculator() {
  return `<div class="name-group pbox-calc">
    <div class="pbox-title">P-BOX 寸法計算</div>
    <div class="pbox-inputs">
      <label>W<input type="number" id="pbox-w" min="0" placeholder="mm" oninput="calcPBox()"></label>
      <label>H<input type="number" id="pbox-h" min="0" placeholder="mm" oninput="calcPBox()"></label>
      <label>D<input type="number" id="pbox-d" min="0" placeholder="mm" oninput="calcPBox()"></label>
    </div>
    <div class="pbox-results" id="pbox-results">
      <div class="pbox-result-cell" id="pbox-r-indoor" onclick="addPBoxToEstimate('屋内')">屋内<br><span>—</span></div>
      <div class="pbox-result-cell" id="pbox-r-outdoor" onclick="addPBoxToEstimate('屋外')">屋外<br><span>—</span></div>
      <div class="pbox-result-cell" id="pbox-r-sus" onclick="addPBoxToEstimate('屋外(SUS)')">屋外(SUS)<br><span>—</span></div>
      <div class="pbox-result-cell" id="pbox-r-indoor-door" onclick="addPBoxToEstimate('屋内(扉付)')">屋内(扉付)<br><span>—</span></div>
      <div class="pbox-result-cell" id="pbox-r-outdoor-door" onclick="addPBoxToEstimate('屋外(扉付)')">屋外(扉付)<br><span>—</span></div>
      <div class="pbox-result-cell" id="pbox-r-sus-door" onclick="addPBoxToEstimate('屋外(SUS,扉付)')">屋外(SUS,扉付)<br><span>—</span></div>
    </div>
    <div class="pbox-rates">
      <div class="pbox-rate-row">
        <label class="pbox-rate-item">≤1000<input type="number" id="pbox-rate-1000" value="25" step="1" oninput="calcPBox()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
        <label class="pbox-rate-item">≤1600<input type="number" id="pbox-rate-1600" value="30" step="1" oninput="calcPBox()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
        <label class="pbox-rate-item">≤2000<input type="number" id="pbox-rate-2000" value="35" step="1" oninput="calcPBox()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
      </div>
      <div class="pbox-rate-row">
        <label class="pbox-rate-item">≤2300<input type="number" id="pbox-rate-2300" value="40" step="1" oninput="calcPBox()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
        <label class="pbox-rate-item">&gt;2300<input type="number" id="pbox-rate-over" value="45" step="1" oninput="calcPBox()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
      </div>
      <div class="pbox-rate-row">
        <label class="pbox-rate-item">屋外<input type="number" id="pbox-rate-outdoor" value="1.2" step="0.1" oninput="calcPBox()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
        <label class="pbox-rate-item">SUS<input type="number" id="pbox-rate-sus" value="3.5" step="0.1" oninput="calcPBox()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
        <label class="pbox-rate-item">扉付<input type="number" id="pbox-rate-door" value="25" step="1" oninput="calcPBox()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
      </div>
    </div>
  </div>`;
}

function getPBoxRates() {
  return {
    r1000: parseFloat(document.getElementById("pbox-rate-1000").value) || 25,
    r1600: parseFloat(document.getElementById("pbox-rate-1600").value) || 30,
    r2000: parseFloat(document.getElementById("pbox-rate-2000").value) || 35,
    r2300: parseFloat(document.getElementById("pbox-rate-2300").value) || 40,
    rOver: parseFloat(document.getElementById("pbox-rate-over").value) || 45,
    outdoor: parseFloat(document.getElementById("pbox-rate-outdoor").value) || 1.2,
    sus: parseFloat(document.getElementById("pbox-rate-sus").value) || 3.5,
    door: parseFloat(document.getElementById("pbox-rate-door").value) || 25,
  };
}

function calcPBoxPrices(sum, rates) {
  let rate;
  if (sum <= 1000) rate = rates.r1000;
  else if (sum <= 1600) rate = rates.r1600;
  else if (sum <= 2000) rate = rates.r2000;
  else if (sum <= 2300) rate = rates.r2300;
  else rate = rates.rOver;

  const indoor = Math.ceil((sum * rate) / 1000);
  const outdoor = Math.ceil(indoor * rates.outdoor);
  const sus = Math.ceil(indoor * rates.sus);
  return {
    "屋内": indoor,
    "屋外": outdoor,
    "屋外(SUS)": sus,
    "屋内(扉付)": indoor + rates.door,
    "屋外(扉付)": outdoor + rates.door,
    "屋外(SUS,扉付)": sus + rates.door,
  };
}

function calcPBox() {
  const w = parseInt(document.getElementById("pbox-w").value) || 0;
  const h = parseInt(document.getElementById("pbox-h").value) || 0;
  const d = parseInt(document.getElementById("pbox-d").value) || 0;
  const sum = w + h + d;

  const allIds = [
    "pbox-r-indoor", "pbox-r-outdoor", "pbox-r-sus",
    "pbox-r-indoor-door", "pbox-r-outdoor-door", "pbox-r-sus-door",
  ];

  if (sum <= 0) {
    for (const id of allIds) {
      const el = document.getElementById(id);
      if (el) el.querySelector("span").textContent = "—";
    }
    return;
  }

  const rates = getPBoxRates();
  const prices = calcPBoxPrices(sum, rates);

  const vals = {
    "pbox-r-indoor": prices["屋内"],
    "pbox-r-outdoor": prices["屋外"],
    "pbox-r-sus": prices["屋外(SUS)"],
    "pbox-r-indoor-door": prices["屋内(扉付)"],
    "pbox-r-outdoor-door": prices["屋外(扉付)"],
    "pbox-r-sus-door": prices["屋外(SUS,扉付)"],
  };

  for (const id in vals) {
    const el = document.getElementById(id);
    if (el) el.querySelector("span").textContent = vals[id];
  }
}

function addPBoxToEstimate(type) {
  const w = parseInt(document.getElementById("pbox-w").value) || 0;
  const h = parseInt(document.getElementById("pbox-h").value) || 0;
  const d = parseInt(document.getElementById("pbox-d").value) || 0;
  const sum = w + h + d;
  if (sum <= 0) return;

  const rates = getPBoxRates();
  const prices = calcPBoxPrices(sum, rates);

  const price = prices[type];
  if (price == null) return;

  const spec = "W" + w + "×H" + h + "×D" + d + " " + type;
  const input = prompt("数量を入力してください\nP-BOX " + spec, "1");
  if (input === null) return;
  const qty = parseInt(input, 10);
  if (isNaN(qty) || qty <= 0) return;

  currentEstimate.lines.push({
    type: "custom",
    lineId: genId(),
    name: "P-BOX",
    spec: spec,
    qty: qty,
    unitPrice: price,
    lineNote: "",
  });

  updateAddedCount();
  updateCubicleAddedCount();
  showToast("追加: P-BOX " + spec + " × " + qty);
}

// ============================================================
// ダクト 寸法入力型計算機
// ============================================================

function renderDuctCalculator() {
  return `<div class="name-group pbox-calc">
    <div class="pbox-title">ダクト 寸法計算</div>
    <div class="pbox-inputs">
      <label>横<input type="number" id="duct-w" min="0" placeholder="mm" oninput="calcDuct()"></label>
      <label>奥行<input type="number" id="duct-h" min="0" placeholder="mm" oninput="calcDuct()"></label>
      <label>高さ<input type="number" id="duct-d" min="0" placeholder="mm" oninput="calcDuct()"></label>
    </div>
    <div class="pbox-results" id="duct-results">
      <div class="pbox-result-cell" id="duct-r-indoor" onclick="addDuctToEstimate('屋内')">屋内<br><span>—</span></div>
      <div class="pbox-result-cell" id="duct-r-outdoor" onclick="addDuctToEstimate('屋外')">屋外<br><span>—</span></div>
      <div class="pbox-result-cell" id="duct-r-sus" onclick="addDuctToEstimate('屋外(SUS)')">屋外(SUS)<br><span>—</span></div>
      <div class="pbox-result-cell" id="duct-r-indoor-door" onclick="addDuctToEstimate('屋内(蓋付)')">屋内(蓋付)<br><span>—</span></div>
      <div class="pbox-result-cell" id="duct-r-outdoor-door" onclick="addDuctToEstimate('屋外(蓋付)')">屋外(蓋付)<br><span>—</span></div>
      <div class="pbox-result-cell" id="duct-r-sus-door" onclick="addDuctToEstimate('屋外(SUS,蓋付)')">屋外(SUS,蓋付)<br><span>—</span></div>
    </div>
    <div class="pbox-rates">
      <div class="pbox-rate-row">
        <label class="pbox-rate-item">≤1000<input type="number" id="duct-rate-1000" value="25" step="1" oninput="calcDuct()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
        <label class="pbox-rate-item">≤1600<input type="number" id="duct-rate-1600" value="30" step="1" oninput="calcDuct()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
        <label class="pbox-rate-item">≤2000<input type="number" id="duct-rate-2000" value="35" step="1" oninput="calcDuct()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
      </div>
      <div class="pbox-rate-row">
        <label class="pbox-rate-item">≤2300<input type="number" id="duct-rate-2300" value="40" step="1" oninput="calcDuct()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
        <label class="pbox-rate-item">&gt;2300<input type="number" id="duct-rate-over" value="45" step="1" oninput="calcDuct()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
      </div>
      <div class="pbox-rate-row">
        <label class="pbox-rate-item">屋外<input type="number" id="duct-rate-outdoor" value="1.2" step="0.1" oninput="calcDuct()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
        <label class="pbox-rate-item">SUS<input type="number" id="duct-rate-sus" value="3.5" step="0.1" oninput="calcDuct()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
        <label class="pbox-rate-item">蓋付<input type="number" id="duct-rate-door" value="25" step="1" oninput="calcDuct()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
      </div>
    </div>
  </div>`;
}

function getDuctRates() {
  return {
    r1000: parseFloat(document.getElementById("duct-rate-1000").value) || 25,
    r1600: parseFloat(document.getElementById("duct-rate-1600").value) || 30,
    r2000: parseFloat(document.getElementById("duct-rate-2000").value) || 35,
    r2300: parseFloat(document.getElementById("duct-rate-2300").value) || 40,
    rOver: parseFloat(document.getElementById("duct-rate-over").value) || 45,
    outdoor: parseFloat(document.getElementById("duct-rate-outdoor").value) || 1.2,
    sus: parseFloat(document.getElementById("duct-rate-sus").value) || 3.5,
    door: parseFloat(document.getElementById("duct-rate-door").value) || 25,
  };
}

function calcDuctPrices(sum, rates) {
  let rate;
  if (sum <= 1000) rate = rates.r1000;
  else if (sum <= 1600) rate = rates.r1600;
  else if (sum <= 2000) rate = rates.r2000;
  else if (sum <= 2300) rate = rates.r2300;
  else rate = rates.rOver;

  const indoor = Math.ceil((sum * rate) / 1000);
  const outdoor = Math.ceil(indoor * rates.outdoor);
  const sus = Math.ceil(indoor * rates.sus);
  return {
    "屋内": indoor,
    "屋外": outdoor,
    "屋外(SUS)": sus,
    "屋内(蓋付)": indoor + rates.door,
    "屋外(蓋付)": outdoor + rates.door,
    "屋外(SUS,蓋付)": sus + rates.door,
  };
}

function calcDuct() {
  const w = parseInt(document.getElementById("duct-w").value) || 0;
  const h = parseInt(document.getElementById("duct-h").value) || 0;
  const d = parseInt(document.getElementById("duct-d").value) || 0;
  const sum = w + h + d;

  const allIds = [
    "duct-r-indoor", "duct-r-outdoor", "duct-r-sus",
    "duct-r-indoor-door", "duct-r-outdoor-door", "duct-r-sus-door",
  ];

  if (sum <= 0) {
    for (const id of allIds) {
      const el = document.getElementById(id);
      if (el) el.querySelector("span").textContent = "—";
    }
    return;
  }

  const rates = getDuctRates();
  const prices = calcDuctPrices(sum, rates);

  const vals = {
    "duct-r-indoor": prices["屋内"],
    "duct-r-outdoor": prices["屋外"],
    "duct-r-sus": prices["屋外(SUS)"],
    "duct-r-indoor-door": prices["屋内(蓋付)"],
    "duct-r-outdoor-door": prices["屋外(蓋付)"],
    "duct-r-sus-door": prices["屋外(SUS,蓋付)"],
  };

  for (const id in vals) {
    const el = document.getElementById(id);
    if (el) el.querySelector("span").textContent = vals[id];
  }
}

function addDuctToEstimate(type) {
  const w = parseInt(document.getElementById("duct-w").value) || 0;
  const h = parseInt(document.getElementById("duct-h").value) || 0;
  const d = parseInt(document.getElementById("duct-d").value) || 0;
  const sum = w + h + d;
  if (sum <= 0) return;

  const rates = getDuctRates();
  const prices = calcDuctPrices(sum, rates);

  const price = prices[type];
  if (price == null) return;

  const spec = "横" + w + "×奥行" + h + "×高さ" + d + " " + type;
  const input = prompt("数量を入力してください\nダクト " + spec, "1");
  if (input === null) return;
  const qty = parseInt(input, 10);
  if (isNaN(qty) || qty <= 0) return;

  currentEstimate.lines.push({
    type: "custom",
    lineId: genId(),
    name: "ダクト",
    spec: spec,
    qty: qty,
    unitPrice: price,
    lineNote: "",
  });

  updateAddedCount();
  updateCubicleAddedCount();
  showToast("追加: ダクト " + spec + " × " + qty);
}

// ============================================================
// 架台・ベース 寸法入力型計算機
// ============================================================

function renderFrameCalculator() {
  return `<div class="name-group pbox-calc">
    <div class="pbox-title">架台・ベース 寸法計算</div>
    <div class="pbox-inputs">
      <label>W<input type="number" id="frame-w" min="0" placeholder="mm" oninput="calcFrame()"></label>
      <label>H<input type="number" id="frame-h" min="0" placeholder="mm" oninput="calcFrame()"></label>
      <label>D<input type="number" id="frame-d" min="0" placeholder="mm" oninput="calcFrame()"></label>
    </div>
    <div class="pbox-results pbox-results-2col" id="frame-results">
      <div class="pbox-result-cell" id="frame-r-l" onclick="addFrameToEstimate('架台(L枠のみ)')">架台(L枠のみ)<br><span>—</span></div>
      <div class="pbox-result-cell" id="frame-r-l-sus" onclick="addFrameToEstimate('架台(L枠のみ,SUS)')">架台(L枠のみ,SUS)<br><span>—</span></div>
      <div class="pbox-result-cell" id="frame-r-lp" onclick="addFrameToEstimate('架台(L枠+プレート)')">架台(L枠+プレート)<br><span>—</span></div>
      <div class="pbox-result-cell" id="frame-r-lp-sus" onclick="addFrameToEstimate('架台(L枠+プレート,SUS)')">架台(L枠+プレート,SUS)<br><span>—</span></div>
      <div class="pbox-result-cell" id="frame-r-b50" onclick="addFrameToEstimate('ベース(H50)')">ベース(H50)<br><span>—</span></div>
      <div class="pbox-result-cell" id="frame-r-b50-sus" onclick="addFrameToEstimate('ベース(H50,SUS)')">ベース(H50,SUS)<br><span>—</span></div>
      <div class="pbox-result-cell" id="frame-r-b100" onclick="addFrameToEstimate('ベース(H100)')">ベース(H100)<br><span>—</span></div>
      <div class="pbox-result-cell" id="frame-r-b100-sus" onclick="addFrameToEstimate('ベース(H100,SUS)')">ベース(H100,SUS)<br><span>—</span></div>
    </div>
    <div class="pbox-rates">
      <div class="pbox-rate-row">
        <label class="pbox-rate-item">L枠<input type="number" id="frame-rate-l" value="60" step="1" oninput="calcFrame()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
        <label class="pbox-rate-item">L+P<input type="number" id="frame-rate-lp" value="120" step="1" oninput="calcFrame()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
        <label class="pbox-rate-item">BH50<input type="number" id="frame-rate-b50" value="50" step="1" oninput="calcFrame()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
      </div>
      <div class="pbox-rate-row">
        <label class="pbox-rate-item">BH100<input type="number" id="frame-rate-b100" value="100" step="1" oninput="calcFrame()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
        <label class="pbox-rate-item">SUS<input type="number" id="frame-rate-sus" value="2" step="0.1" oninput="calcFrame()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
      </div>
    </div>
  </div>`;
}

function getFrameRates() {
  return {
    l: parseFloat(document.getElementById("frame-rate-l").value) || 60,
    lp: parseFloat(document.getElementById("frame-rate-lp").value) || 120,
    b50: parseFloat(document.getElementById("frame-rate-b50").value) || 50,
    b100: parseFloat(document.getElementById("frame-rate-b100").value) || 100,
    sus: parseFloat(document.getElementById("frame-rate-sus").value) || 2,
  };
}

function calcFramePrices(sum, rates) {
  const l = Math.ceil((sum * rates.l) / 1000);
  const lp = Math.ceil((sum * rates.lp) / 1000);
  const b50 = Math.ceil((sum * rates.b50) / 1000);
  const b100 = Math.ceil((sum * rates.b100) / 1000);
  return {
    "架台(L枠のみ)": l,
    "架台(L枠+プレート)": lp,
    "ベース(H50)": b50,
    "ベース(H100)": b100,
    "架台(L枠のみ,SUS)": Math.ceil(l * rates.sus),
    "架台(L枠+プレート,SUS)": Math.ceil(lp * rates.sus),
    "ベース(H50,SUS)": Math.ceil(b50 * rates.sus),
    "ベース(H100,SUS)": Math.ceil(b100 * rates.sus),
  };
}

function calcFrame() {
  const w = parseInt(document.getElementById("frame-w").value) || 0;
  const h = parseInt(document.getElementById("frame-h").value) || 0;
  const d = parseInt(document.getElementById("frame-d").value) || 0;
  const sum = w + h + d;

  const allIds = [
    "frame-r-l", "frame-r-lp", "frame-r-b50", "frame-r-b100",
    "frame-r-l-sus", "frame-r-lp-sus", "frame-r-b50-sus", "frame-r-b100-sus",
  ];

  if (sum <= 0) {
    for (const id of allIds) {
      const el = document.getElementById(id);
      if (el) el.querySelector("span").textContent = "—";
    }
    return;
  }

  const rates = getFrameRates();
  const prices = calcFramePrices(sum, rates);

  const vals = {
    "frame-r-l": prices["架台(L枠のみ)"],
    "frame-r-lp": prices["架台(L枠+プレート)"],
    "frame-r-b50": prices["ベース(H50)"],
    "frame-r-b100": prices["ベース(H100)"],
    "frame-r-l-sus": prices["架台(L枠のみ,SUS)"],
    "frame-r-lp-sus": prices["架台(L枠+プレート,SUS)"],
    "frame-r-b50-sus": prices["ベース(H50,SUS)"],
    "frame-r-b100-sus": prices["ベース(H100,SUS)"],
  };

  for (const id in vals) {
    const el = document.getElementById(id);
    if (el) el.querySelector("span").textContent = vals[id];
  }
}

function addFrameToEstimate(type) {
  const w = parseInt(document.getElementById("frame-w").value) || 0;
  const h = parseInt(document.getElementById("frame-h").value) || 0;
  const d = parseInt(document.getElementById("frame-d").value) || 0;
  const sum = w + h + d;
  if (sum <= 0) return;

  const rates = getFrameRates();
  const prices = calcFramePrices(sum, rates);

  const price = prices[type];
  if (price == null) return;

  const spec = "W" + w + "×H" + h + "×D" + d;
  const input = prompt("数量を入力してください\n" + type + " " + spec, "1");
  if (input === null) return;
  const qty = parseInt(input, 10);
  if (isNaN(qty) || qty <= 0) return;

  currentEstimate.lines.push({
    type: "custom",
    lineId: genId(),
    name: type,
    spec: spec,
    qty: qty,
    unitPrice: price,
    lineNote: "",
  });

  updateAddedCount();
  updateCubicleAddedCount();
  showToast("追加: " + type + " " + spec + " × " + qty);
}

// ============================================================
// 防油トレー 寸法入力型計算機
// ============================================================

function renderTrayCalculator() {
  return `<div class="name-group pbox-calc">
    <div class="pbox-title">防油トレー 寸法計算</div>
    <div class="pbox-inputs">
      <label>W<input type="number" id="tray-w" min="0" placeholder="mm" oninput="calcTray()"></label>
      <label>H<input type="number" id="tray-h" min="0" placeholder="mm" oninput="calcTray()"></label>
      <label>D<input type="number" id="tray-d" min="0" placeholder="mm" oninput="calcTray()"></label>
    </div>
    <div class="pbox-results pbox-results-1col" id="tray-results">
      <div class="pbox-result-cell" id="tray-r-price" onclick="addTrayToEstimate()">防油トレー<br><span>—</span></div>
    </div>
    <div class="pbox-rates">
      <div class="pbox-rate-row">
        <label class="pbox-rate-item">≤1000<input type="number" id="tray-rate-1000" value="25" step="1" oninput="calcTray()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
        <label class="pbox-rate-item">≤1600<input type="number" id="tray-rate-1600" value="30" step="1" oninput="calcTray()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
        <label class="pbox-rate-item">≤2000<input type="number" id="tray-rate-2000" value="35" step="1" oninput="calcTray()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
      </div>
      <div class="pbox-rate-row">
        <label class="pbox-rate-item">≤2300<input type="number" id="tray-rate-2300" value="40" step="1" oninput="calcTray()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
        <label class="pbox-rate-item">&gt;2300<input type="number" id="tray-rate-over" value="45" step="1" oninput="calcTray()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
      </div>
      <div class="pbox-rate-row">
        <label class="pbox-rate-item">加算<input type="number" id="tray-rate-add" value="25" step="1" oninput="calcTray()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
        <label class="pbox-rate-item">倍率<input type="number" id="tray-rate-mul" value="1.2" step="0.1" oninput="calcTray()" onclick="event.stopPropagation()" onfocus="this.select()"></label>
      </div>
    </div>
  </div>`;
}

function getTrayRates() {
  return {
    r1000: parseFloat(document.getElementById("tray-rate-1000").value) || 25,
    r1600: parseFloat(document.getElementById("tray-rate-1600").value) || 30,
    r2000: parseFloat(document.getElementById("tray-rate-2000").value) || 35,
    r2300: parseFloat(document.getElementById("tray-rate-2300").value) || 40,
    rOver: parseFloat(document.getElementById("tray-rate-over").value) || 45,
    add: parseFloat(document.getElementById("tray-rate-add").value) || 25,
    mul: parseFloat(document.getElementById("tray-rate-mul").value) || 1.2,
  };
}

function calcTrayPrice(sum, rates) {
  let rate;
  if (sum <= 1000) rate = rates.r1000;
  else if (sum <= 1600) rate = rates.r1600;
  else if (sum <= 2000) rate = rates.r2000;
  else if (sum <= 2300) rate = rates.r2300;
  else rate = rates.rOver;

  const base = Math.ceil((sum * rate) / 1000);
  return Math.ceil((base + rates.add) * rates.mul);
}

function calcTray() {
  const w = parseInt(document.getElementById("tray-w").value) || 0;
  const h = parseInt(document.getElementById("tray-h").value) || 0;
  const d = parseInt(document.getElementById("tray-d").value) || 0;
  const sum = w + h + d;

  const el = document.getElementById("tray-r-price");
  if (sum <= 0) {
    if (el) el.querySelector("span").textContent = "—";
    return;
  }

  const rates = getTrayRates();
  const price = calcTrayPrice(sum, rates);
  if (el) el.querySelector("span").textContent = price;
}

function addTrayToEstimate() {
  const w = parseInt(document.getElementById("tray-w").value) || 0;
  const h = parseInt(document.getElementById("tray-h").value) || 0;
  const d = parseInt(document.getElementById("tray-d").value) || 0;
  const sum = w + h + d;
  if (sum <= 0) return;

  const rates = getTrayRates();
  const price = calcTrayPrice(sum, rates);

  const spec = "W" + w + "×H" + h + "×D" + d;
  const input = prompt("数量を入力してください\n防油トレー " + spec, "1");
  if (input === null) return;
  const qty = parseInt(input, 10);
  if (isNaN(qty) || qty <= 0) return;

  currentEstimate.lines.push({
    type: "custom",
    lineId: genId(),
    name: "防油トレー",
    spec: spec,
    qty: qty,
    unitPrice: price,
    lineNote: "",
  });

  updateAddedCount();
  updateCubicleAddedCount();
  showToast("追加: 防油トレー " + spec + " × " + qty);
}

// ============================================================
// オプション表（片扉・両扉）
// ============================================================

function renderOptionTable() {
  return `<div class="name-group pbox-calc">
    <div class="pbox-title">オプション</div>
    <table class="mg-table option-table">
      <thead><tr><th>品名</th><th>単価</th></tr></thead>
      <tbody>
        <tr class="mg-row" onclick="addOptionToEstimate(event, '片扉')">
          <td class="mg-spec">片扉</td>
          <td class="mg-price"><input type="number" id="opt-price-single" min="0" step="0.1" value="25"
               onclick="event.stopPropagation()" onfocus="this.select()"></td>
        </tr>
        <tr class="mg-row" onclick="addOptionToEstimate(event, '両扉')">
          <td class="mg-spec">両扉</td>
          <td class="mg-price"><input type="number" id="opt-price-double" min="0" step="0.1" value="25"
               onclick="event.stopPropagation()" onfocus="this.select()"></td>
        </tr>
        <tr class="mg-row" onclick="addOptionToEstimate(event, '窓')">
          <td class="mg-spec">窓</td>
          <td class="mg-price"><input type="number" id="opt-price-window" min="0" step="0.1" value="0"
               onclick="event.stopPropagation()" onfocus="this.select()"></td>
        </tr>
      </tbody>
    </table>
  </div>`;
}

function addOptionToEstimate(event, type) {
  if (event.target.tagName === "INPUT") return;
  const inputId = type === "片扉" ? "opt-price-single" : type === "両扉" ? "opt-price-double" : "opt-price-window";
  const price = parseFloat(document.getElementById(inputId).value) || 0;
  if (price <= 0) return;

  const input = prompt("数量を入力してください\n" + type + " (単価: " + price + ")", "1");
  if (input === null) return;
  const qty = parseInt(input, 10);
  if (isNaN(qty) || qty <= 0) return;

  currentEstimate.lines.push({
    type: "custom",
    lineId: genId(),
    name: type,
    spec: "",
    qty: qty,
    unitPrice: price,
    lineNote: "",
  });

  updateAddedCount();
  updateCubicleAddedCount();
  showToast("追加: " + type + " × " + qty);
}

// ============================================================
// オプション2（ポール取付・コン柱取付・スタンド）
// ============================================================

function renderOption2Table() {
  return `<div class="name-group pbox-calc">
    <div class="pbox-title">オプション2</div>
    <table class="mg-table option-table">
      <thead><tr><th>品名</th><th>単価</th></tr></thead>
      <tbody>
        <tr class="mg-row" onclick="addOption2ToEstimate(event, 'ポール取付')">
          <td class="mg-spec">ポール取付</td>
          <td class="mg-price"><input type="number" id="opt2-price-pole" min="0" step="0.1" value="0"
               onclick="event.stopPropagation()" onfocus="this.select()"></td>
        </tr>
        <tr class="mg-row" onclick="addOption2ToEstimate(event, 'コン柱取付')">
          <td class="mg-spec">コン柱取付</td>
          <td class="mg-price"><input type="number" id="opt2-price-conchu" min="0" step="0.1" value="0"
               onclick="event.stopPropagation()" onfocus="this.select()"></td>
        </tr>
        <tr class="mg-row" onclick="addOption2ToEstimate(event, 'スタンド')">
          <td class="mg-spec">スタンド</td>
          <td class="mg-price"><input type="number" id="opt2-price-stand" min="0" step="0.1" value="0"
               onclick="event.stopPropagation()" onfocus="this.select()"></td>
        </tr>
      </tbody>
    </table>
  </div>`;
}

function addOption2ToEstimate(event, type) {
  if (event.target.tagName === "INPUT") return;
  const idMap = { "ポール取付": "opt2-price-pole", "コン柱取付": "opt2-price-conchu", "スタンド": "opt2-price-stand" };
  const price = parseFloat(document.getElementById(idMap[type]).value) || 0;
  if (price <= 0) return;

  const input = prompt("数量を入力してください\n" + type + " (単価: " + price + ")", "1");
  if (input === null) return;
  const qty = parseInt(input, 10);
  if (isNaN(qty) || qty <= 0) return;

  currentEstimate.lines.push({
    type: "custom",
    lineId: genId(),
    name: type,
    spec: "",
    qty: qty,
    unitPrice: price,
    lineNote: "",
  });

  updateAddedCount();
  updateCubicleAddedCount();
  showToast("追加: " + type + " × " + qty);
}

// ============================================================
// よく出る品目
// ============================================================

function renderQuickItemsTable() {
  return `<div class="name-group pbox-calc">
    <div class="pbox-title">よく出る品目</div>
    <table class="mg-table option-table">
      <thead><tr><th>品名</th><th>単価</th></tr></thead>
      <tbody>
        <tr class="mg-row" onclick="addQuickItemToEstimate(event, '上部ダクト')">
          <td class="mg-spec">上部ダクト</td>
          <td class="mg-price"><input type="number" id="quick-price-duct-top" min="0" step="0.1" value="0"
               onclick="event.stopPropagation()" onfocus="this.select()"></td>
        </tr>
        <tr class="mg-row" onclick="addQuickItemToEstimate(event, '下部ダクト')">
          <td class="mg-spec">下部ダクト</td>
          <td class="mg-price"><input type="number" id="quick-price-duct-btm" min="0" step="0.1" value="0"
               onclick="event.stopPropagation()" onfocus="this.select()"></td>
        </tr>
        <tr class="mg-row" onclick="addQuickItemToEstimate(event, '自立')">
          <td class="mg-spec">自立</td>
          <td class="mg-price"><input type="number" id="quick-price-standalone" min="0" step="0.1" value="0"
               onclick="event.stopPropagation()" onfocus="this.select()"></td>
        </tr>
      </tbody>
    </table>
  </div>`;
}

function addQuickItemToEstimate(event, type) {
  if (event.target.tagName === "INPUT") return;
  const idMap = {
    "上部ダクト": "quick-price-duct-top",
    "下部ダクト": "quick-price-duct-btm",
    "自立": "quick-price-standalone",
  };
  const inputId = idMap[type];
  if (!inputId) return;
  const price = parseFloat(document.getElementById(inputId).value) || 0;
  if (price <= 0) return;

  const input = prompt("数量を入力してください\n" + type + " (単価: " + price + ")", "1");
  if (input === null) return;
  const qty = parseInt(input, 10);
  if (isNaN(qty) || qty <= 0) return;

  currentEstimate.lines.push({
    type: "custom",
    lineId: genId(),
    name: type,
    spec: "",
    qty: qty,
    unitPrice: price,
    lineNote: "",
  });

  updateAddedCount();
  updateCubicleAddedCount();
  showToast("追加: " + type + " × " + qty);
}

// ============================================================
// 管理表（キュービクル）レンダリング
// ============================================================

function renderCubicleTable() {
  const container = document.getElementById("cubicle-content");

  // CUBICLE_CATEGORIESの定義順でグループ化
  const catMap = {};
  for (const item of cubicleItems) {
    if (!catMap[item.category]) catMap[item.category] = [];
    catMap[item.category].push(item);
  }
  const catGroups = [];
  for (const cat of CUBICLE_CATEGORIES) {
    if (catMap[cat.id] && catMap[cat.id].length > 0) {
      catGroups.push({ id: cat.id, label: cat.name, items: catMap[cat.id] });
    } else if (cat.id === "K11") {
      catGroups.push({ id: cat.id, label: cat.name, items: [] });
    }
  }

  let html = "";
  for (const cg of catGroups) {
    html += `<div class="cat-block">`;
    html += `<div class="cat-heading cat-heading-cubicle" onclick="toggleCat(this)"><span class="cat-toggle">▼</span> ${esc(cg.id)}. ${esc(cg.label)}</div>`;

    // カテゴリ内を品名でさらにグループ化
    const nameGroups = [];
    let lastName = null;
    for (const item of cg.items) {
      if (item.name !== lastName) {
        lastName = item.name;
        nameGroups.push({ name: item.name, items: [] });
      }
      nameGroups[nameGroups.length - 1].items.push(item);
    }

    html += `<div class="cat-tables">`;
    // K11カスタム: カスタム入力テーブルを表示
    if (cg.id === "K11") {
      html += renderCubicleCustomInputTable();
      html += `</div></div>`;
      continue;
    }
    // 横並びペア定義: left → { right: [...], leftWith: [...] }
    // leftWith の品名は left と同じ左カラムに縦積みされる
    const CUBICLE_HPAIRS = {
      "受電パターン": { right: ["LBSパターン", "VCT"] },
      "1φTR(単相)": { right: ["エネセーバ回路"], leftWith: ["3φTR(三相)", "TR440V(三相)"] },
      "受電盤・饋電盤(パネル付)": { right: ["低圧配電盤(パネルなし)", "パネル"], splitRight: true },
      "高圧スコットTR": { right: ["低圧スコットTR"] },
    };
    const SINGLE_ZONE_EXCLUDE = new Set(["LBSパターン", "VCT", "エネセーバ回路", "低圧配電盤(パネルなし)", "パネル", "低圧スコットTR"]);
    const SINGLE_ZONE_INCLUDE = new Set(["メーター", "Wh(電力量計)", "銅帯母線用", "VCB(真空遮断器)", "低圧MC-DT(220V)"]);
    // 単品2ゾーン: カテゴリ → 単品2が始まる最初のnameGroup名
    const SINGLE_ZONE2_START = { "K1": "LBS(7.2kV)", "K4": "低圧MC-DT(220V)" };
    const hpairRight = new Set();
    const hpairLeftWith = new Set();
    for (const pair of Object.values(CUBICLE_HPAIRS)) {
      const rights = Array.isArray(pair) ? pair : (pair.right || []);
      rights.forEach(r => hpairRight.add(r));
      if (pair.leftWith) pair.leftWith.forEach(l => hpairLeftWith.add(l));
    }
    // footerGroupとして使われている名前を収集（単品ゾーンから除外）
    const cubFooterNames = new Set();
    for (const key in CUBICLE_MATRIX_GROUPS) {
      const md = CUBICLE_MATRIX_GROUPS[key];
      if (md.footerGroup) cubFooterNames.add(md.footerGroup);
      if (md.footerGroups) md.footerGroups.forEach(fg => cubFooterNames.add(fg.group));
    }
    // マトリクスと単品を分離して、単品は「単品」ゾーンにまとめる
    let inSingleZone = false;
    let singleBlockNum = 0;
    let singleZone2Started = false;
    const zone2Start = SINGLE_ZONE2_START[cg.id];
    for (const ng of nameGroups) {
      // 横並びペアの右側・左グループ・footerGroupは描画済みなのでスキップ
      if (hpairRight.has(ng.name) || hpairLeftWith.has(ng.name) || cubFooterNames.has(ng.name)) continue;
      const isMatrix = (!!CUBICLE_MATRIX_GROUPS[ng.name] || SINGLE_ZONE_EXCLUDE.has(ng.name)) && !SINGLE_ZONE_INCLUDE.has(ng.name);
      if (!isMatrix && !inSingleZone) {
        html += `<div class="cubicle-single-zone">`;
        html += `<div class="cubicle-single-zone-heading">単品</div>`;
        html += `<div class="cubicle-single-zone-inner">`;
        inSingleZone = true;
      } else if (isMatrix && inSingleZone) {
        html += `</div></div>`;
        inSingleZone = false;
      }
      // 単品2ゾーン切り替え
      if (inSingleZone && zone2Start && ng.name === zone2Start && !singleZone2Started) {
        html += `</div></div>`;
        html += `<div class="cubicle-single-zone">`;
        html += `<div class="cubicle-single-zone-heading">単品2</div>`;
        html += `<div class="cubicle-single-zone-inner">`;
        singleZone2Started = true;
        singleBlockNum = 0;
      }
      // 横並びペアの左側の場合、右側を探して横並びレンダリング
      const pairDef = CUBICLE_HPAIRS[ng.name];
      if (pairDef) {
        const rights = Array.isArray(pairDef) ? pairDef : (pairDef.right || []);
        const leftWith = Array.isArray(pairDef) ? [] : (pairDef.leftWith || []);
        html += `<div class="cubicle-hpair">`;
        html += `<div class="cubicle-hpair-col">`;
        html += renderNameGroup(ng, CUBICLE_MATRIX_GROUPS);
        for (const ln of leftWith) {
          const leftNg = nameGroups.find(n => n.name === ln);
          if (leftNg) html += renderNameGroup(leftNg, CUBICLE_MATRIX_GROUPS);
        }
        html += `</div>`;
        if (pairDef.splitRight) {
          for (const rn of rights) {
            const rightNg = nameGroups.find(n => n.name === rn);
            if (rightNg) {
              html += `<div class="cubicle-hpair-col">`;
              html += renderNameGroup(rightNg, CUBICLE_MATRIX_GROUPS);
              html += `</div>`;
            }
          }
        } else if (rights.length > 0) {
          html += `<div class="cubicle-hpair-col">`;
          for (const rn of rights) {
            const rightNg = nameGroups.find(n => n.name === rn);
            if (rightNg) html += renderNameGroup(rightNg, CUBICLE_MATRIX_GROUPS);
          }
          html += `</div>`;
        }
        html += `</div>`;
      } else if (inSingleZone) {
        singleBlockNum++;
        html += renderNameGroup(ng, CUBICLE_MATRIX_GROUPS, singleBlockNum);
      } else {
        html += renderNameGroup(ng, CUBICLE_MATRIX_GROUPS);
      }
    }
    if (inSingleZone) html += `</div></div>`;
    html += `</div></div>`;
  }
  container.innerHTML = html;
  updateCubicleAddedCount();
}

// ============================================================
// K11. キュービクル カスタム入力テーブル
// ============================================================
function renderCubicleCustomInputTable() {
  let h = `<div class="name-group-wide custom-input-table">`;
  h += `<table class="mg-table"><thead><tr>`;
  h += `<th>品名</th><th>仕様</th><th>単価</th>`;
  h += `</tr></thead><tbody>`;
  for (let i = 1; i <= 5; i++) {
    h += `<tr onclick="addCubicleCustomToEstimate(${i})" style="cursor:pointer">`;
    h += `<td><input type="text" id="cub-custom-name-${i}" placeholder="品名"></td>`;
    h += `<td><input type="text" id="cub-custom-spec-${i}" placeholder="仕様"></td>`;
    h += `<td><input type="number" id="cub-custom-price-${i}" value="0"></td>`;
    h += `</tr>`;
  }
  h += `</tbody></table></div>`;
  return h;
}

function addCubicleCustomToEstimate(row) {
  if (event && (event.target.tagName === "INPUT")) return;
  const name = document.getElementById("cub-custom-name-" + row).value.trim();
  const spec = document.getElementById("cub-custom-spec-" + row).value.trim();
  const price = parseFloat(document.getElementById("cub-custom-price-" + row).value) || 0;
  if (!name) { showToast("品名を入力してください"); return; }
  const input = prompt("数量を入力してください\n" + name + (spec ? " " + spec : ""), "1");
  if (input === null) return;
  const qty = parseInt(input, 10);
  if (isNaN(qty) || qty <= 0) return;
  currentEstimate.lines.push({
    type: "custom",
    lineId: genId(),
    name: name,
    spec: spec,
    qty: qty,
    unitPrice: price,
    lineNote: "",
  });
  updateAddedCount();
  updateCubicleAddedCount();
  const tr = document.querySelector(`#cub-custom-name-${row}`).closest("tr");
  if (tr) { tr.classList.add("mg-flash"); setTimeout(() => tr.classList.remove("mg-flash"), 400); }
  showToast("追加: " + name + (spec ? " " + spec : "") + " × " + qty);
}

// ============================================================
// I. カスタム入力テーブル
// ============================================================
function renderCustomInputTable() {
  let h = `<div class="name-group-wide custom-input-table">`;
  h += `<table class="mg-table"><thead><tr>`;
  h += `<th>品名</th><th>仕様</th><th>単価</th>`;
  h += `</tr></thead><tbody>`;
  for (let i = 1; i <= 5; i++) {
    h += `<tr onclick="addCustomToEstimate(${i})" style="cursor:pointer">`;
    h += `<td><input type="text" id="custom-name-${i}" placeholder="品名"></td>`;
    h += `<td><input type="text" id="custom-spec-${i}" placeholder="仕様"></td>`;
    h += `<td><input type="number" id="custom-price-${i}" value="0"></td>`;
    h += `</tr>`;
  }
  h += `</tbody></table></div>`;
  return h;
}

function addCustomToEstimate(row) {
  if (event && (event.target.tagName === "INPUT")) return;
  const name = document.getElementById("custom-name-" + row).value.trim();
  const spec = document.getElementById("custom-spec-" + row).value.trim();
  const price = parseFloat(document.getElementById("custom-price-" + row).value) || 0;
  if (!name) { showToast("品名を入力してください"); return; }
  const input = prompt("数量を入力してください\n" + name + (spec ? " " + spec : ""), "1");
  if (input === null) return;
  const qty = parseInt(input, 10);
  if (isNaN(qty) || qty <= 0) return;
  currentEstimate.lines.push({
    type: "custom",
    lineId: genId(),
    name: name,
    spec: spec,
    qty: qty,
    unitPrice: price,
    lineNote: "",
  });
  updateAddedCount();
  updateCubicleAddedCount();
  const tr = document.querySelector(`#custom-name-${row}`).closest("tr");
  if (tr) { tr.classList.add("mg-flash"); setTimeout(() => tr.classList.remove("mg-flash"), 400); }
  showToast("追加: " + name + (spec ? " " + spec : "") + " × " + qty);
}

// ============================================================
// カテゴリ内 縦積みペア定義（flex内で縦に並べるグループ）
// ============================================================
const VERTICAL_STACKS = {
  "C": [],
};

// カテゴリ内 強制横並びペア定義（overflow:hiddenで幅を制限して横に並べる）
const HORIZONTAL_PAIRS = {
  "C": [],
};

// ============================================================
// カテゴリ内サブグループ定義
// ============================================================
const CAT_SUBGROUPS = {
};

// ============================================================
// マトリクス表示の定義: 横軸(cols) × 縦軸(rows) のグリッド表（盤用）
// ============================================================
const MATRIX_GROUPS = {
  "TB (端子台)": {
    rows: ["50AF(14sq)", "100AF(38sq)", "225AF(100sq)", "400AF(200sq)", "600AF(325sq)", "800AF", "1000AF"],
    cols: ["2P", "3P", "4P"],
    highlightCols: ["3P"],
    accent: true,
  },
  "主幹MCCB": {
    rows: ["50AF", "60AF", "100AF", "125AF", "225AF", "250AF", "400AF", "600AF", "800AF", "1000AF", "1200AF"],
    cols: ["2P", "3P", "4P", "AL/AX", "SHT"],
    highlightCols: ["3P"],
    accent: true,
  },
  "主幹ELCB": {
    rows: ["50AF", "60AF", "100AF", "125AF", "225AF", "250AF", "400AF", "600AF", "800AF", "1000AF", "1200AF"],
    cols: ["2P", "3P", "4P", "AL/AX", "SHT"],
    highlightCols: ["3P"],
    accent: true,
  },
  "分岐TB": {
    rows: ["50AF", "100AF", "225AF", "400AF", "600AF", "800AF", "1000AF"],
    cols: ["2P", "3P", "4P"],
  },
  "分岐MCCB": {
    rows: ["50AF", "60AF", "100AF", "125AF", "225AF", "250AF", "400AF", "600AF", "800AF", "1000AF", "1200AF"],
    cols: ["2P", "3P", "4P", "スペース", "AL/AX", "SHT"],
    highlightCols: ["3P"],
    accent: true,
  },
  "分岐ELCB": {
    rows: ["50AF", "60AF", "100AF", "125AF", "225AF", "250AF", "400AF", "600AF", "800AF", "1000AF", "1200AF"],
    cols: ["2P", "3P", "4P", "スペース", "AL/AX", "SHT"],
    highlightCols: ["3P"],
    accent: true,
  },
  "MCなど(単品)": {
    rows: ["0.2kW", "0.4kW", "0.75kW", "1.5kW", "2.2kW", "3.7kW", "5.5kW", "7.5kW", "11kW", "15kW", "18.5kW", "22kW", "30kW", "37kW", "45kW", "55kW", "75kW", "90kW", "110kW", "132kW"],
    hideRowLabels: true,
    textCols: { "MCAF": "A", "MC200V": "kW", "MC400V": "kW" },
    colGroups: [
      { label: "MC", cols: ["AF", "200V", "400V"] },
      { label: "", cols: ["88"] },
      { label: "", cols: ["88(正逆)"] },
      { label: "Y-△", cols: ["2C", "3C", "M+2C", "CLOS"] },
      { label: "SC", cols: ["220V", "440V"] },
    ],
  },
  "MC-DT": {
    rows: ["30A", "50A", "75A", "100A", "150A", "200A", "300A", "400A", "600A", "800A", "1000A", "1200A", "1600A"],
    cols: ["2P", "3P"],
    highlightCols: ["3P"],
    footerGroup: "操作回路",
    footerLabel: "オプション",
    footerBelow: true,
  },
  "MC": {
    rows: ["20A", "25A", "35A", "50A", "65A", "80A", "100A", "125A", "150A", "180A", "220A", "300A", "400A"],
    cols: ["常時励磁", "ラッチ式"],
    highlightCols: ["常時励磁"],
    accent: true,
  },
  "210V回路セット": {
    rows: ["0.2kW", "0.7kW", "1.5kW", "2.2kW", "3.7kW", "5.5kW", "7.5kW", "11kW", "15kW", "19kW", "22kW", "30kW", "37kW", "45kW", "55kW", "75kW"],
    hideRowLabels: true,
    title: "210V回路セット (MCCB,AM,CT,MC,COSorPBS,RL,GL,(SC))",
    footerGroup: "回路セットオプション",
    accent: true,
    highlightCols: ["直入", "直入(SC付)", "M+2C", "M+2C(SC付)"],
    textCols: { "AF": "AF", "kW": "kW" },
    colGroups: [
      { label: "", cols: ["AF"] },
      { label: "", cols: ["kW"] },
      { label: "", cols: ["直入"] },
      { label: "", cols: ["直入(SC付)"] },
      { label: "", cols: ["正逆"] },
      { label: "", cols: ["正逆(SC付)"] },
      { label: "Y-△", cols: ["2C", "2C(SC付)", "3C", "3C(SC付)", "M+2C", "M+2C(SC付)", "CLOS", "CLOS(SC付)"] },
    ],
  },
  "440V回路セット": {
    rows: ["3.7kW", "5.5kW", "7.5kW", "11kW", "15kW", "19kW", "22kW", "30kW", "37kW", "45kW", "55kW", "75kW", "90kW", "110kW", "150kW"],
    hideRowLabels: true,
    title: "440V回路セット (MCCB,AM,CT,MC,COSorPBS,RL,GL,(SC))",
    footerGroup: "回路セットオプション",
    textCols: { "AF": "AF", "kW": "kW" },
    colGroups: [
      { label: "", cols: ["AF"] },
      { label: "", cols: ["kW"] },
      { label: "", cols: ["直入"] },
      { label: "", cols: ["直入(SC付)"] },
      { label: "", cols: ["正逆"] },
      { label: "", cols: ["正逆(SC付)"] },
      { label: "Y-△", cols: ["2C", "2C(SC付)", "3C", "3C(SC付)", "M+2C", "M+2C(SC付)", "CLOS", "CLOS(SC付)"] },
    ],
  },
  "WHM 電子式": {
    rows: ["1φ2W", "1φ3W 3φ3W"],
    colGroups: [
      { label: "未検定", cols: ["30A", "120A", "250A", "/5A"] },
      { label: "検定付", cols: ["30A", "120A", "250A", "/5A"] },
    ],
  },
  "WHM アナログ": {
    rows: ["1φ2W", "1φ3W 3φ3W", "3φ4W", "1φ2W 発信装置付", "1φ3W 3φ3W 発信装置付", "3φ4W 発信装置付"],
    colGroups: [
      { label: "未検定", cols: ["30A", "120A", "/5A"] },
      { label: "検定付", cols: ["30A", "120A", "/5A"] },
    ],
  },
  "接地端子盤(100sq,200A)": {
    rows: ["E1", "E1,E2,E3", "E1,E2,E3,予備*2"],
    cols: ["屋内型", "屋外型", "屋外型,SUS"],
  },
  "端子スペース": {
    rows: ["10P", "20P", "30P", "40P", "60P", "80P", "100P", "150P", "200P", "250P", "300P", "400P", "600P", "800P", "1000P"],
    cols: ["屋内型", "屋外型", "屋外型,SUS製"],
    showNote: true,
  },
  "電話保安器スペース": {
    rows: ["5P(1個用)", "10P(2個用)", "20P(4個用)", "30P(6個用)", "40P(8個用)", "50P(10個用)", "60P(12個用)", "70P(14個用)", "80P(16個用)"],
    cols: ["屋内型", "屋外型", "屋外型,SUS製"],
    showNote: true,
  },
  "端子スペース(電話保安器スペース付)": {
    rows: ["20P", "30P", "40P", "60P", "80P", "100P", "150P", "200P", "300P", "400P", "500P"],
    cols: ["屋内型", "屋外型", "屋外型,SUS製"],
    showNote: true,
  },
  "TV・情報スペース": {
    rows: ["1台", "2台", "4台", "6台", "8台", "10台", "12台", "16台", "20台"],
    cols: ["屋内型", "屋外型", "屋外型,SUS製"],
    showNote: true,
  },
  "総合盤(前半)": {
    title: "総合盤",
    cols: ["W400", "W500", "W600", "W700", "W800", "W900", "W1000", "W1100", "W1200"],
    rows: [
      "H1800 D300", "H1800 D400", "H1800 D500", "H1800 D600",
      "H1900 D300", "H1900 D400", "H1900 D500", "H1900 D600",
      "H2000 D300", "H2000 D400", "H2000 D500", "H2000 D600",
      "H2100 D300", "H2100 D400", "H2100 D500", "H2100 D600",
      "H2200 D300", "H2200 D400", "H2200 D500", "H2200 D600",
      "H2300 D300", "H2300 D400", "H2300 D500", "H2300 D600",
      "H2400 D300", "H2400 D400", "H2400 D500", "H2400 D600",
      "H2500 D300", "H2500 D400", "H2500 D500", "H2500 D600",
      "H2600 D300", "H2600 D400", "H2600 D500", "H2600 D600",
    ],
    rowGroups: [
      { label: "H1800", rows: ["H1800 D300", "H1800 D400", "H1800 D500", "H1800 D600"] },
      { label: "H1900", rows: ["H1900 D300", "H1900 D400", "H1900 D500", "H1900 D600"] },
      { label: "H2000", rows: ["H2000 D300", "H2000 D400", "H2000 D500", "H2000 D600"] },
      { label: "H2100", rows: ["H2100 D300", "H2100 D400", "H2100 D500", "H2100 D600"] },
      { label: "H2200", rows: ["H2200 D300", "H2200 D400", "H2200 D500", "H2200 D600"] },
      { label: "H2300", rows: ["H2300 D300", "H2300 D400", "H2300 D500", "H2300 D600"] },
      { label: "H2400", rows: ["H2400 D300", "H2400 D400", "H2400 D500", "H2400 D600"] },
      { label: "H2500", rows: ["H2500 D300", "H2500 D400", "H2500 D500", "H2500 D600"] },
      { label: "H2600", rows: ["H2600 D300", "H2600 D400", "H2600 D500", "H2600 D600"] },
    ],
    rowLabelFn: function(row) { return row.split(" ")[1]; },
  },
  "総合盤(後半)": {
    title: "総合盤",
    cols: ["W400", "W500", "W600", "W700", "W800", "W900", "W1000", "W1100", "W1200"],
    rows: [
      "H2700 D300", "H2700 D400", "H2700 D500", "H2700 D600",
      "H2800 D300", "H2800 D400", "H2800 D500", "H2800 D600",
      "H2900 D300", "H2900 D400", "H2900 D500", "H2900 D600",
      "H3000 D300", "H3000 D400", "H3000 D500", "H3000 D600",
      "H3100 D300", "H3100 D400", "H3100 D500", "H3100 D600",
      "H3200 D300", "H3200 D400", "H3200 D500", "H3200 D600",
      "H3300 D300", "H3300 D400", "H3300 D500", "H3300 D600",
      "H3400 D300", "H3400 D400", "H3400 D500", "H3400 D600",
      "H3500 D300", "H3500 D400", "H3500 D500", "H3500 D600",
    ],
    rowGroups: [
      { label: "H2700", rows: ["H2700 D300", "H2700 D400", "H2700 D500", "H2700 D600"] },
      { label: "H2800", rows: ["H2800 D300", "H2800 D400", "H2800 D500", "H2800 D600"] },
      { label: "H2900", rows: ["H2900 D300", "H2900 D400", "H2900 D500", "H2900 D600"] },
      { label: "H3000", rows: ["H3000 D300", "H3000 D400", "H3000 D500", "H3000 D600"] },
      { label: "H3100", rows: ["H3100 D300", "H3100 D400", "H3100 D500", "H3100 D600"] },
      { label: "H3200", rows: ["H3200 D300", "H3200 D400", "H3200 D500", "H3200 D600"] },
      { label: "H3300", rows: ["H3300 D300", "H3300 D400", "H3300 D500", "H3300 D600"] },
      { label: "H3400", rows: ["H3400 D300", "H3400 D400", "H3400 D500", "H3400 D600"] },
      { label: "H3500", rows: ["H3500 D300", "H3500 D400", "H3500 D500", "H3500 D600"] },
    ],
    rowLabelFn: function(row) { return row.split(" ")[1]; },
  },
};

const MERGED_MATRICES = {
  "A": [
    {
      title: "主幹",
      rows: ["50AF","60AF","100AF","125AF","225AF","250AF","400AF","600AF","800AF","1000AF","1200AF"],
      groups: [
        { label: "TB(端子台)", nameGroup: "TB (端子台)", cols: ["2P","3P","4P"],
          rowMap: {"50AF":"50AF(14sq)","100AF":"100AF(38sq)","225AF":"225AF(100sq)",
                   "400AF":"400AF(200sq)","600AF":"600AF(325sq)"} },
        { label: "主幹MCCB", nameGroup: "主幹MCCB", cols: ["2P","3P","4P"] },
        { label: "主幹ELCB", nameGroup: "主幹ELCB", cols: ["2P","3P","4P"] },
        { label: "オプション", nameGroup: "主幹MCCB", cols: ["AL/AX","SHT"] },
      ],
      highlightCols: ["3P"],
      accent: true,
      hideNames: ["オプション"],
    },
    {
      title: "分岐",
      rows: ["50AF","60AF","100AF","125AF","225AF","250AF","400AF","600AF","800AF","1000AF","1200AF"],
      groups: [
        { label: "分岐TB", nameGroup: "分岐TB", cols: ["2P","3P","4P"] },
        { label: "分岐MCCB", nameGroup: "分岐MCCB", cols: ["2P","3P","4P"] },
        { label: "分岐ELCB", nameGroup: "分岐ELCB", cols: ["2P","3P","4P"] },
        { label: "スペース", nameGroup: "分岐ELCB", cols: ["スペース"] },
        { label: "オプション", nameGroup: "分岐MCCB", cols: ["AL/AX","SHT"] },
      ],
      highlightCols: ["3P"],
      accent: true,
    },
    {
      title: "スリム",
      rows: ["2P1E","2P2E"],
      groups: [
        { label: "スリムブレーカー", nameGroup: "スリムブレーカー", cols: ["MCCB","ELCB","スペース"] },
        { label: "プラグインブレーカー", nameGroup: "プラグインブレーカー", cols: ["MCCB","ELCB","スペース"] },
        { label: "漏電表示付ブレーカー(日東)", nameGroup: "漏電表示付ブレーカー(日東)", cols: ["MCCB","ELCB","スペース"] },
      ],
      specReverse: true,
      accent: true,
    },
    {
      title: "SPD (避雷器)",
      rows: ["SPD","SPD（分離器付）","SPD（分離器・接点付）"],
      groups: [
        { label: "SPD (避雷器)", nameGroup: "SPD (避雷器)", cols: ["クラスⅠ","クラスⅡ"],
          specMap: {
            "SPD|クラスⅠ": "クラスⅠ",
            "SPD|クラスⅡ": "クラスⅡ",
            "SPD（分離器付）|クラスⅠ": "クラスⅠ（分離器付）",
            "SPD（分離器付）|クラスⅡ": "クラスⅡ（分離器付）",
            "SPD（分離器・接点付）|クラスⅠ": "クラスⅠ（分離器・接点付）",
            "SPD（分離器・接点付）|クラスⅡ": "クラスⅡ（分離器・接点付）",
          } },
      ],
      accent: true,
      singleZone: true,
    }
  ]
};

const SINGLE_ZONE_DEFS = {
  "A": [
    { names: new Set(["SPD (避雷器)", "その他"]), cols: 3 },
  ],
  "C": [
    { names: new Set(["WHM 電子式", "WHM アナログ"]), cols: 2 },
    { names: ["MC", "MCなど(単品)", "Ry・SW・PLなど", "TM", "エネルギーモニターユニット", "MC-DT", "制御用Tr", "計器", "フロートスイッチ", "ELR(ZCT付)", "コンセント", "警報", "その他"],
      cols: 4, wideNames: new Set(["MCなど(単品)"]), spanNames: { "フロートスイッチ": 2 } },
  ],
  "K10": [
    { names: ["TR1φ(単相)", "TR3φ(三相)", "TR3φ440V"], cols: 3 },
  ],
};

const ACCENT_GROUPS = { "A": new Set(["スリムブレーカー", "その他"]), "B": "*", "C": new Set(["Ry・SW・PLなど"]) };
const HIGHLIGHT_ITEMS = new Set([
  "A054b",
  "B07C", "B07D", "B07E", "B07F", "B08C", "B08D", "B08F",
  "B05o", "B051", "B05c", "B05f", "B05i", "B05j", "B05k", "B05q", "B05m",
  "D061", "D062", "D065", "D068", "D06b", "D06c", "D06d", "D06h", "D06f",
  "F001","F004","F007","F00A","F00D","F00G","F00J","F00M","F00P","F00S","F00V","F00Y","F0a1","F0a4","F0a7",
  "F020","F023","F026","F029","F02C","F02F","F02I","F02L","F02O","F02R","F02U",
  "F030","F033","F036","F039","F03C","F03F","F03I","F03L","F03O",
  "F040","F043","F046","F049","F04C","F04F","F04I","F04L","F04O",
  "K1001","K1002","K1003","K1004","K1007","K1008","K1009","K100A","K100D","K100E","K100F","K1010",
  "K10C0",
  "J020","J021","J022","J023","J024","J025","J026","J027","J028","J029","J02E","J02F","J02G","J02J","J02H",
  // K7 MCCB・ELCB 3P + ELCB ハイライト
  "K9102","K9103","K9104","K9105","K910A","K910B","K910E","K910F",
  "K9112","K9113","K9116","K9117",
  "K911C","K911D","K911E","K911F","K9124","K9125","K9128","K9129",
  "K912C","K912D","K9130","K9131",
  "K9136","K9137","K9138","K9139","K913E","K913F","K9142","K9143",
  "K9146","K9147","K914A","K914B",
  "K9150","K9151","K9152","K9153","K9158","K9159","K915C","K915D",
  "K9160","K9161","K9164","K9165","K9168","K9169",
  "K916C","K916D","K9170","K9171","K9174","K9175",
  "K9178","K9179","K917C","K917D","K9180","K9181",
]);

/** 品名グループを1つのサブテーブルとしてレンダリング（共用） */
function renderNameGroup(ng, matrixDefs, blockNum) {
  // マトリクス定義がある場合はマトリクス表示（rowsがある場合のみ）
  const mdef = matrixDefs && matrixDefs[ng.name];
  if (mdef && mdef.rows) {
    return renderMatrixGroup(ng, mdef, blockNum);
  }
  const single = ng.items.length === 1 && !ng.items[0].spec;
  const cat = ng.items[0] && ng.items[0].category;
  const ag = ACCENT_GROUPS[cat];
  const accentClass = ag && (ag === "*" || ag.has(ng.name)) ? " mg-table-accent" : "";
  const showNote = (ng.name === "フロートスイッチ");
  const showQtyNote = (ng.name === "空調盤SR-1" || ng.name === "空調盤SR-1(コインタイマー)");
  const showTotalCol = ng.items.length > 0 && ng.items[0]._qty != null && !showQtyNote;
  const numPrefix = blockNum ? `<span class="mg-block-num">${blockNum}</span> ` : "";
  let h = `<div class="name-group${accentClass}">`;
  if (mdef && mdef.listComments) h += `<div class="mg-comments-top">${buildCommentHtml(mdef.listComments)}</div>`;
  h += `<table class="mg-table">`;
  h += `<thead><tr>`;
  if (single) {
    h += `<th class="mg-name">${numPrefix}${esc(ng.name)}</th>`;
  } else {
    h += `<th class="mg-name" colspan="2">${numPrefix}${esc(ng.name)}</th>`;
  }
  if (showQtyNote) h += `<th class="mg-qty-col-header">数量</th><th class="mg-note-header">備考</th>`;
  if (showNote) h += `<th class="mg-note-header">備考</th>`;
  if (showTotalCol) h += `<th class="mg-col-header">合計</th>`;
  h += `</tr></thead><tbody>`;

  const colSpan = single ? 1 : 2;
  // 合計用: グループ全体の合計を算出
  let groupTotal = 0;
  if (showTotalCol) {
    for (const item of ng.items) {
      const qty = parseInt(item._qty) || 1;
      groupTotal += Math.round(item.basePrice * qty);
    }
  }
  let totalInserted = false;
  for (const item of ng.items) {
    if (item._divider) {
      h += `<tr><td class="mg-row-group-header" colspan="${colSpan + (showTotalCol ? 1 : 0)}">${esc(item._divider)}</td></tr>`;
    }
    const cnt = masterClickCounts[item.id] || 0;
    const isHL = HIGHLIGHT_ITEMS.has(item.id);
    const bgStyle = cnt > 0 ? `style="background:rgba(202,138,4,${Math.min(cnt * 0.25, 0.85)})"` : isHL ? `style="background:rgba(66,153,225,0.15)"` : "";
    h += `<tr class="mg-row" data-item-id="${item.id}" ${bgStyle} onclick="addFromMaster(event,'${item.id}')">`;
    if (!single) {
      h += `<td class="mg-row-label">${esc(item.spec || "")}</td>`;
    }
    h += `<td class="mg-price"><input type="number" min="0" step="0.1" value="${item.basePrice}"
         oninput="onMasterPriceChange('${item.id}','basePrice',this.value)"
         onchange="onMasterPriceChange('${item.id}','basePrice',this.value)" onclick="event.stopPropagation()"></td>`;
    if (showQtyNote) {
      h += `<td class="mg-qty-col">${esc(item._qty || "")}</td>`;
      h += `<td class="mg-note">${esc(item._memo || "")}</td>`;
    }
    if (showNote) {
      h += `<td class="mg-note">${esc(item.note || "")}</td>`;
    }
    if (showTotalCol && !totalInserted) {
      const ngName = ng.name;
      // 保存済みの合計値があればそちらを使う
      const savedTotal = groupTotals[ngName] != null ? groupTotals[ngName] : groupTotal;
      h += `<td class="mg-price mg-solar-total" data-group="${esc(ngName)}" rowspan="${ng.items.length}" style="vertical-align:middle">`;
      h += `<input type="number" min="0" step="0.1" value="${savedTotal}" class="mg-solar-total-input" data-group="${esc(ngName)}"
             oninput="onGroupTotalChange('${escAttr(ngName)}',this.value)"
             onchange="onGroupTotalChange('${escAttr(ngName)}',this.value)"
             onclick="event.stopPropagation()">`;
      h += `</td>`;
      totalInserted = true;
    }
    h += `</tr>`;
  }
  h += `</tbody></table></div>`;
  return h;
}

/** マトリクス形式でサブテーブルをレンダリング */
function buildCommentHtml(comments) {
  let html = "";
  for (const c of comments) {
    if (typeof c === "string") {
      html += `<div class="mg-comment-line">${esc(c)}</div>`;
    } else {
      html += `<div class="mg-comment-block">`;
      if (c.heading) html += `<div class="mg-comment-heading">${esc(c.heading)}</div>`;
      for (const line of c.lines) html += `<div class="mg-comment-line">${esc(line)}</div>`;
      html += `</div>`;
    }
  }
  return html;
}

function buildSurchargeHtml(surcharges, matrixName) {
  let h = '<div class="mg-surcharges">';
  for (const group of surcharges) {
    h += '<div class="mg-surcharge-group">';
    h += `<div class="mg-surcharge-heading">${esc(group.heading)}</div>`;
    h += '<div class="mg-surcharge-items">';
    for (const item of group.items) {
      if (item.rate === null) {
        h += '<div class="mg-surcharge-item disabled">';
        h += `<span class="mg-surcharge-label">${esc(item.label)} → 見積辞退</span>`;
        h += '</div>';
      } else {
        const key = matrixName + "::" + item.label;
        h += '<div class="mg-surcharge-item">';
        h += `<label><input type="checkbox" data-matrix="${escAttr(matrixName)}" data-rate="${item.rate}" data-key="${escAttr(key)}"
               onchange="onSurchargeChange('${escAttr(matrixName)}')" />`;
        h += ` ${esc(item.label)} → 函体×${item.rate}</label>`;
        h += '</div>';
      }
    }
    h += '</div>'; // mg-surcharge-items
    h += '</div>';
  }
  h += '</div>';
  return h;
}

function onSurchargeChange(matrixName) {
  const checkboxes = document.querySelectorAll(`input[data-matrix="${matrixName}"]`);
  let totalRate = 1;
  for (const cb of checkboxes) {
    if (cb.checked) totalRate *= parseFloat(cb.dataset.rate);
  }
  const basePrices = matrixBasePrices[matrixName];
  if (!basePrices) return;
  for (const [itemId, base] of Object.entries(basePrices)) {
    const newPrice = Math.round(base * totalRate);
    const cell = document.querySelector(`[data-item-id="${itemId}"] input[type="number"]`);
    if (cell) cell.value = newPrice;
    const master = getMasterItem(itemId);
    if (master) master.basePrice = newPrice;
  }
}

function buildRefTablesHtml(refTables) {
  let html = `<div class="mg-ref-tables">`;
  html += `<div class="mg-ref-title">トランスから見たブレーカ遮断容量表</div>`;
  html += `<div class="mg-ref-tables-row">`;
  for (const rt of refTables) {
    html += `<table class="mg-ref-table"><thead>`;
    html += `<tr><th class="mg-ref-heading" colspan="${rt.cols.length + 1}">${esc(rt.heading)}</th></tr>`;
    html += `<tr><th class="mg-ref-corner">ブレーカ＼TR容量</th>`;
    for (const c of rt.cols) html += `<th class="mg-ref-col">${esc(c)}</th>`;
    html += `</tr></thead><tbody>`;
    for (const r of rt.rows) {
      html += `<tr><td class="mg-ref-row-label">${esc(r.label)}</td>`;
      let prev = "";
      for (let vi = 0; vi < r.vals.length; vi++) {
        const v = r.vals[vi];
        let cls = "mg-ref-cell";
        if (v) {
          const c = v.replace(/[()]/g,"").charAt(0);
          cls += c === "C" ? " mg-ref-c" : c === "S" ? " mg-ref-s" : c === "H" ? " mg-ref-h" : " mg-ref-other";
        }
        const label = (v && v !== prev) ? esc(v) : "";
        html += `<td class="${cls}">${label}</td>`;
        prev = v;
      }
      html += `</tr>`;
    }
    html += `</tbody></table>`;
  }
  html += `</div></div>`;
  return html;
}

function renderMergedMatrix(md, nameGroups, blockNum) {
  // 各グループのnameGroupとitemMapを構築
  const groupData = [];
  for (const g of md.groups) {
    const ng = nameGroups.find(n => n.name === g.nameGroup);
    const items = ng ? ng.items : [];
    const itemMap = {};
    for (const item of items) {
      itemMap[item.spec] = item;
    }
    groupData.push({ def: g, ng, itemMap });
  }

  const hlClass = md.accent ? " mg-table-accent" : "";
  let h = `<div class="name-group name-group-wide${hlClass}">`;
  h += `<table class="mg-table mg-matrix">`;

  const numPrefix = blockNum ? `<span class="mg-block-num">${blockNum}</span> ` : "";
  // thead
  const singleGroup = md.groups.length === 1;
  h += `<thead>`;
  if (singleGroup) {
    // 1段ヘッダ: タイトル + 列名
    h += `<tr><th class="mg-name">${numPrefix}${esc(md.title)}</th>`;
    for (const col of md.groups[0].cols) {
      h += `<th class="mg-col-header mg-group-sep">${esc(col)}</th>`;
    }
    h += `</tr>`;
  } else {
    // 2段ヘッダ: 1段目 タイトル + グループ名
    h += `<tr><th class="mg-name" rowspan="2">${numPrefix}${esc(md.title)}</th>`;
    for (let gi = 0; gi < md.groups.length; gi++) {
      const g = md.groups[gi];
      h += `<th class="mg-col-group-header mg-group-sep" colspan="${g.cols.length}">${esc(g.label)}</th>`;
    }
    h += `</tr>`;
    // 2段目: 列名
    h += `<tr>`;
    for (let gi = 0; gi < md.groups.length; gi++) {
      const g = md.groups[gi];
      for (let ci = 0; ci < g.cols.length; ci++) {
        const sep = ci === 0 ? " mg-group-sep" : "";
        h += `<th class="mg-col-header${sep}">${esc(g.cols[ci])}</th>`;
      }
    }
    h += `</tr>`;
  }
  h += `</thead><tbody>`;

  // tbody: 各行
  for (const row of md.rows) {
    h += `<tr>`;
    h += `<td class="mg-row-label">${esc(row)}</td>`;
    for (let gi = 0; gi < md.groups.length; gi++) {
      const g = md.groups[gi];
      const gd = groupData[gi];
      for (let ci = 0; ci < g.cols.length; ci++) {
        const col = g.cols[ci];
        const sep = ci === 0 ? " mg-group-sep" : "";
        // specキーを構築
        const mappedRow = g.rowMap ? (g.rowMap[row] || null) : row;
        if (mappedRow === null) {
          h += `<td class="mg-price mg-empty${sep}">-</td>`;
          continue;
        }
        const key = g.specMap ? g.specMap[mappedRow + "|" + col]
                  : md.specReverse ? (col + " " + mappedRow)
                  : (mappedRow + " " + col);
        const item = gd.itemMap[key];
        if (item) {
          const cnt = masterClickCounts[item.id] || 0;
          const isHL = md.highlightCols && md.highlightCols.includes(col);
          const bgStyle = cnt > 0 ? `background:rgba(202,138,4,${Math.min(cnt * 0.25, 0.85)})` : isHL ? `background:rgba(66,153,225,0.15)` : "";
          h += `<td class="mg-price mg-clickable mg-matrix-cell${sep}" data-item-id="${item.id}" style="${bgStyle}" onclick="addFromMaster(event,'${item.id}')">`;
          h += `<input type="number" min="0" step="0.1" value="${item.basePrice}"
                oninput="onMasterPriceChange('${item.id}','basePrice',this.value)"
                onchange="onMasterPriceChange('${item.id}','basePrice',this.value)" onclick="event.stopPropagation()">`;
          h += `<span class="mg-cell-overlay" title="クリックで見積に追加"></span>`;
          h += `</td>`;
        } else {
          h += `<td class="mg-price mg-empty${sep}">-</td>`;
        }
      }
    }
    h += `</tr>`;
  }

  h += `</tbody></table>`;
  h += `</div>`;
  return h;
}

function renderMatrixGroup(ng, def, blockNum) {
  // specからアイテムを検索するマップ
  const itemMap = {};
  for (const item of ng.items) {
    itemMap[item.spec] = item;
  }

  // colGroups対応: フラットなcols配列とspecキー用プレフィックスを構築
  let flatCols = [];
  let colSpecPrefix = []; // spec検索用: "未検定30A" など
  const hasColGroups = def.colGroups && def.colGroups.length > 0;
  if (hasColGroups) {
    for (const grp of def.colGroups) {
      for (const col of grp.cols) {
        flatCols.push(col);
        colSpecPrefix.push(grp.label + col);
      }
    }
  } else {
    flatCols = def.cols;
    colSpecPrefix = def.cols;
  }

  const showRowLabel = !def.hideRowLabels;
  const hlClass = def.accent ? " mg-table-accent" : "";
  const hasSurcharges = def.commentsTop && def.surcharges && def.surcharges.length > 0;
  const commentsTop = !hasSurcharges && def.commentsTop && def.comments && def.comments.length > 0;
  const hasComments = !commentsTop && !hasSurcharges && def.comments && def.comments.length > 0;
  let h = `<div class="name-group name-group-wide${hlClass}">`;
  const numPrefix = blockNum ? `<span class="mg-block-num">${blockNum}</span> ` : "";
  if (!showRowLabel) h += `<div class="mg-matrix-title">${numPrefix}${esc(def.title || ng.name)}</div>`;
  if (hasSurcharges) h += `<div class="mg-comments-top">${buildSurchargeHtml(def.surcharges, ng.name)}</div>`;
  if (commentsTop) h += `<div class="mg-comments-top">${buildCommentHtml(def.comments)}</div>`;
  if (def.refTables) h += buildRefTablesHtml(def.refTables);
  const hasFooter = def.footerGroup || def.footerGroups;
  if (hasFooter && !def.footerBelow) h += `<div class="mg-matrix-row">`;
  h += `<table class="mg-table mg-matrix">`;

  // tbody行数を計算（コメントrowspan用）
  let totalBodyRows = def.rows.length;
  const rowGroupMap = {};
  if (def.rowGroups) {
    const seenGroups = new Set();
    for (const rg of def.rowGroups) {
      for (const r of rg.rows) rowGroupMap[r] = rg.label;
    }
    for (const row of def.rows) {
      if (rowGroupMap[row] && !seenGroups.has(rowGroupMap[row])) {
        seenGroups.add(rowGroupMap[row]);
        totalBodyRows++;
      }
    }
  }

  const commentCol = hasComments ? 1 : 0;
  const cmtWidthStyle = hasComments && def.commentWidth ? ` style="width:${def.commentWidth}px;min-width:${def.commentWidth}px"` : "";

  const hasRowInfo = def.rowInfoCols && def.rowInfoCols.length > 0;
  const rowInfoColCount = hasRowInfo ? def.rowInfoCols.length : 0;

  h += `<thead>`;
  // 大項目行（colGroupsがある場合）
  if (hasColGroups) {
    h += `<tr>`;
    if (hasComments) h += `<th class="mg-comment-header" rowspan="2"${cmtWidthStyle}></th>`;
    if (showRowLabel) h += `<th class="mg-name" rowspan="2">${numPrefix}${esc(ng.name)}</th>`;
    if (hasRowInfo) {
      for (const ic of def.rowInfoCols) h += `<th class="mg-col-header mg-info-header" rowspan="2">${esc(ic)}</th>`;
    }
    for (const grp of def.colGroups) {
      const dl = grp.displayLabel || grp.label;
      if (dl) {
        h += `<th class="mg-col-group-header" colspan="${grp.cols.length}">${esc(dl)}</th>`;
      } else {
        for (const col of grp.cols) {
          h += `<th class="mg-col-header" rowspan="2">${esc(col)}</th>`;
        }
      }
    }
    h += `</tr><tr>`;
    for (const grp of def.colGroups) {
      const dl = grp.displayLabel || grp.label;
      if (dl) {
        for (const col of grp.cols) {
          h += `<th class="mg-col-header">${esc(col)}</th>`;
        }
      }
    }
    h += `</tr>`;
  } else {
    const colCount = commentCol + (showRowLabel ? 1 : 0) + rowInfoColCount + def.cols.length + (def.showNote ? 1 : 0);
    if (def.title && def.rowGroups) {
      h += `<tr><th class="mg-name" colspan="${colCount}">${esc(def.title)}</th></tr>`;
    } else if (def.title) {
      h += `<tr><th class="mg-name" colspan="${colCount}">${esc(def.title)}</th></tr>`;
      h += `<tr>`;
      if (hasComments) h += `<th class="mg-comment-header"${cmtWidthStyle}></th>`;
      if (showRowLabel) h += `<th class="mg-col-header"></th>`;
      if (hasRowInfo) { for (const ic of def.rowInfoCols) h += `<th class="mg-col-header mg-info-header">${esc(ic)}</th>`; }
      for (const col of def.cols) h += `<th class="mg-col-header">${esc(col)}</th>`;
      if (def.showNote) h += `<th class="mg-note-header">備考</th>`;
      h += `</tr>`;
    } else {
      h += `<tr>`;
      if (hasComments) h += `<th class="mg-comment-header"${cmtWidthStyle}></th>`;
      if (showRowLabel) h += `<th class="mg-name">${numPrefix}${esc(ng.name)}</th>`;
      if (hasRowInfo) { for (const ic of def.rowInfoCols) h += `<th class="mg-col-header mg-info-header">${esc(ic)}</th>`; }
      for (const col of def.cols) h += `<th class="mg-col-header">${esc(col)}</th>`;
      if (def.showNote) h += `<th class="mg-note-header">備考</th>`;
      h += `</tr>`;
    }
  }
  h += `</thead><tbody>`;

  let lastRowGroup = null;
  const totalColSpan = commentCol + (showRowLabel ? 1 : 0) + rowInfoColCount + flatCols.length + (def.showNote ? 1 : 0);
  let commentInserted = false;

  for (const row of def.rows) {
    // 大項目ヘッダ行
    if (def.rowGroups && rowGroupMap[row] && rowGroupMap[row] !== lastRowGroup) {
      lastRowGroup = rowGroupMap[row];
      h += `<tr>`;
      // コメントセル: 最初のtbody行に挿入
      if (hasComments && !commentInserted) {
        h += `<td class="mg-comments-cell" rowspan="${totalBodyRows}"${cmtWidthStyle}>${buildCommentHtml(def.comments)}</td>`;
        commentInserted = true;
      }
      if (showRowLabel) h += `<td class="mg-row-group-header">${esc(lastRowGroup)}</td>`;
      if (hasRowInfo) { for (let i = 0; i < rowInfoColCount; i++) h += `<td class="mg-row-group-col"></td>`; }
      for (let ci = 0; ci < flatCols.length; ci++) {
        h += `<td class="mg-row-group-col">${esc(flatCols[ci])}</td>`;
      }
      if (def.showNote) h += `<td class="mg-row-group-col"></td>`;
      h += `</tr>`;
    }

    const rowLabel = def.rowLabelFn ? def.rowLabelFn(row) : row;
    h += `<tr>`;
    // コメントセル: rowGroupsがない場合は最初のデータ行に挿入
    if (hasComments && !commentInserted) {
      h += `<td class="mg-comments-cell" rowspan="${totalBodyRows}"${cmtWidthStyle}>${buildCommentHtml(def.comments)}</td>`;
      commentInserted = true;
    }
    if (showRowLabel) h += `<td class="mg-row-label">${esc(rowLabel)}</td>`;
    if (hasRowInfo) {
      const info = def.rowInfo && def.rowInfo[row] ? def.rowInfo[row] : [];
      for (let i = 0; i < rowInfoColCount; i++) h += `<td class="mg-row-label mg-info-cell">${esc(info[i] != null ? String(info[i]) : "")}</td>`;
    }
    for (let ci = 0; ci < flatCols.length; ci++) {
      const key = row + " " + colSpecPrefix[ci];
      const item = itemMap[key];
      if (item) {
        if (hasSurcharges) {
          if (!matrixBasePrices[ng.name]) matrixBasePrices[ng.name] = {};
          matrixBasePrices[ng.name][item.id] = item.basePrice;
        }
        const textSuffix = def.textCols && def.textCols[colSpecPrefix[ci]];
        if (textSuffix) {
          const tv = item.basePrice ? item.basePrice + textSuffix : "-";
          h += `<td class="mg-row-label">${tv}</td>`;
        } else {
          const cnt = masterClickCounts[item.id] || 0;
          const isHL = (def.highlightCols && def.highlightCols.includes(flatCols[ci])) || HIGHLIGHT_ITEMS.has(item.id);
          const bgStyle = cnt > 0 ? `background:rgba(202,138,4,${Math.min(cnt * 0.25, 0.85)})` : isHL ? `background:rgba(66,153,225,0.15)` : "";
          h += `<td class="mg-price mg-clickable mg-matrix-cell" data-item-id="${item.id}" style="${bgStyle}" onclick="addFromMaster(event,'${item.id}')">`;
          h += `<input type="number" min="0" step="0.1" value="${item.basePrice}"
                oninput="onMasterPriceChange('${item.id}','basePrice',this.value)"
                onchange="onMasterPriceChange('${item.id}','basePrice',this.value)" onclick="event.stopPropagation()">`;
          h += `<span class="mg-cell-overlay" title="クリックで見積に追加"></span>`;
          h += `</td>`;
        }
      } else {
        h += `<td class="mg-price mg-empty">-</td>`;
      }
    }
    if (def.showNote) {
      const firstKey = row + " " + colSpecPrefix[0];
      const noteItem = itemMap[firstKey];
      h += `<td class="mg-note">${esc(noteItem ? noteItem.note || "" : "")}</td>`;
    }
    h += `</tr>`;
  }

  h += `</tbody></table>`;

  // footerGroup / footerGroups: マトリックスの横or下に別テーブルをくっつける
  const footerList = def.footerGroups
    ? def.footerGroups
    : def.footerGroup ? [{ group: def.footerGroup, label: def.footerLabel || def.footerGroup }] : [];
  if (footerList.length > 0) {
    const allItems = masterItems.concat(cubicleItems);
    if (def.footerBelow && footerList.length > 1) h += `<div style="display:flex;gap:4px;align-items:flex-start">`;
    for (const fg of footerList) {
      const footerItems = allItems.filter(it => it.name === fg.group);
      if (footerItems.length > 0) {
        h += `<table class="mg-table mg-footer-table">`;
        h += `<tbody><tr><td class="mg-row-group-header" colspan="2">${esc(fg.label)}</td></tr>`;
        for (const fi of footerItems) {
          const cnt = masterClickCounts[fi.id] || 0;
          const bgStyle = cnt > 0 ? `background:rgba(202,138,4,${Math.min(cnt * 0.25, 0.85)})` : "";
          h += `<tr class="mg-clickable" style="${bgStyle}" onclick="addFromMaster(event,'${fi.id}')">`;
          h += `<td class="mg-row-label">${esc(fi.spec)}</td>`;
          h += `<td class="mg-price mg-matrix-cell">`;
          h += `<input type="number" min="0" step="0.1" value="${fi.basePrice}"
                oninput="onMasterPriceChange('${fi.id}','basePrice',this.value)"
                onchange="onMasterPriceChange('${fi.id}','basePrice',this.value)" onclick="event.stopPropagation()">`;
          h += `<span class="mg-cell-overlay" title="クリックで見積に追加"></span>`;
          h += `</td></tr>`;
        }
        h += `</tbody></table>`;
      }
    }
    if (def.footerBelow && footerList.length > 1) h += `</div>`;
    if (!def.footerBelow) h += `</div>`;
  }

  h += `</div>`;
  return h;
}

// ============================================================
// 管理表 → 見積追加（クリック即追加）— 盤・キュービクル共用
// ============================================================

/** 行クリックで見積に1行追加 */
function addFromMaster(event, id) {
  // input内のクリックは無視（stopPropagationで来ないはずだが念のため）
  if (event.target.tagName === "INPUT" || event.target.tagName === "SELECT") return;

  const master = getMasterItem(id);
  if (!master) return;

  const label = master.name + (master.spec ? " " + master.spec : "");
  const input = prompt("数量を入力してください\n" + label, "1");
  if (input === null) return; // キャンセル
  const qty = parseInt(input, 10);
  if (isNaN(qty) || qty <= 0) return;

  currentEstimate.lines.push({
    type: "item",
    lineId: genId(),
    masterItemId: id,
    qty: qty,
    unitPrice: master.basePrice,
    lineNote: "",
  });

  // SR-1コメントを備考に反映
  const sr1GroupName = master.name;
  if (sr1Comments[sr1GroupName]) {
    const tag = "【" + sr1GroupName + "】" + sr1Comments[sr1GroupName];
    const notes = currentEstimate.notes || "";
    if (!notes.includes(tag)) {
      currentEstimate.notes = notes ? notes + "\n" + tag : tag;
      const notesEl = document.getElementById("notes-textarea");
      if (notesEl) notesEl.value = currentEstimate.notes;
    }
  }

  // クリック回数を記録
  masterClickCounts[id] = (masterClickCounts[id] || 0) + 1;
  const cnt = masterClickCounts[id];

  // 追加件数バッジ更新
  updateAddedCount();
  updateCubicleAddedCount();

  // 色の更新（クリック回数に応じて濃くなる）
  const alpha = Math.min(cnt * 0.25, 0.85);
  const bgColor = `rgba(202,138,4,${alpha})`;

  // マトリクスセルの場合はtdに直接、通常行の場合はtrに適用
  const cell = event.target.closest("td.mg-matrix-cell");
  const tr = event.target.closest("tr");
  if (cell) {
    cell.style.background = bgColor;
  } else if (tr) {
    tr.style.background = bgColor;
  }

  // 行フラッシュ（視覚フィードバック）
  if (tr) {
    tr.classList.add("mg-flash");
    setTimeout(() => tr.classList.remove("mg-flash"), 400);
  }

  showToast("追加: " + label + " × " + qty);
}

/** 追加済み件数を更新（盤） */
function updateAddedCount() {
  const el = document.getElementById("added-count");
  if (el) el.textContent = currentEstimate.lines.filter(l => l.type === "item" || l.type === "custom").length;
}

/** 追加済み件数を更新（キュービクル） */
function updateCubicleAddedCount() {
  const el = document.getElementById("cubicle-added-count");
  if (el) el.textContent = currentEstimate.lines.filter(l => l.type === "item" || l.type === "custom").length;
}

// ============================================================
// 見積作成 - カスケード選択UI（盤 + キュービクル両方対応）
// ============================================================

/** 全カテゴリ一覧を返す（盤 + キュービクル統合） */
function getAllCategories() {
  const panelCats = CATEGORIES.map(c => ({ ...c, source: "盤" }));
  const cubCats = CUBICLE_CATEGORIES.map(c => ({ ...c, source: "キュービクル" }));
  return [...panelCats, ...cubCats];
}

/** 全品目一覧を返す（盤 + キュービクル統合） */
function getAllItems() {
  return [...masterItems, ...cubicleItems];
}

function renderAddSelectors() {
  const catSel = document.getElementById("add-cat");
  const allCats = getAllCategories();
  catSel.innerHTML = '<option value="">-- カテゴリ --</option>' +
    '<optgroup label="盤">' +
    CATEGORIES.map(c => `<option value="${c.id}">${c.id}. ${esc(c.name)}</option>`).join("") +
    '</optgroup>' +
    '<optgroup label="キュービクル">' +
    CUBICLE_CATEGORIES.map(c => `<option value="${c.id}">${c.id}. ${esc(c.name)}</option>`).join("") +
    '</optgroup>';
  document.getElementById("add-name").innerHTML = '<option value="">-- 品名 --</option>';
  document.getElementById("add-spec").innerHTML = '<option value="">-- 仕様 --</option>';
}

function onAddCatChange(catId) {
  const nameSel = document.getElementById("add-name");
  const specSel = document.getElementById("add-spec");
  specSel.innerHTML = '<option value="">-- 仕様 --</option>';
  if (!catId) {
    nameSel.innerHTML = '<option value="">-- 品名 --</option>';
    return;
  }
  const allItems = getAllItems();
  const names = [...new Set(allItems.filter(m => m.category === catId).map(m => m.name))];
  nameSel.innerHTML = '<option value="">-- 品名 --</option>' +
    names.map(n => `<option value="${escAttr(n)}">${esc(n)}</option>`).join("");
}

function onAddNameChange(name) {
  const catId = document.getElementById("add-cat").value;
  const specSel = document.getElementById("add-spec");
  if (!name || !catId) {
    specSel.innerHTML = '<option value="">-- 仕様 --</option>';
    return;
  }
  const allItems = getAllItems();
  const items = allItems.filter(m => m.category === catId && m.name === name);
  if (items.length === 1 && !items[0].spec) {
    specSel.innerHTML = `<option value="${items[0].id}" selected>(なし)</option>`;
  } else {
    specSel.innerHTML = '<option value="">-- 仕様 --</option>' +
      items.map(m => `<option value="${m.id}">${esc(m.spec || "(なし)")}</option>`).join("");
  }
}

function onAddSpecChange(val) {
  // 仕様選択時の追加処理（将来用）
}

function addSelectedItem() {
  const specSel = document.getElementById("add-spec");
  const masterId = specSel.value;
  if (!masterId) { showToast("品目を選択してください"); return; }
  const master = getMasterItem(masterId);
  if (!master) return;

  currentEstimate.lines.push({
    type: "item",
    lineId: genId(),
    masterItemId: masterId,
    qty: 1,
    unitPrice: master.basePrice,
    lineNote: "",
  });

  renderEstimateLines();
  renderTotals();
  showToast("追加: " + master.name + (master.spec ? " " + master.spec : ""));
}

// ============================================================
// 見積作成 - 明細レンダリング
// ============================================================

function renderEstimateTab() {
  renderProjectInfo();
  renderAddSelectors();
  renderEstimateLines();
  renderTotals();
  renderNotes();
  renderEstimateSelector();
}

function renderProjectInfo() {
  const fields = [
    { key: "projectName", label: "工事名" },
    { key: "estimateNo", label: "見積No." },
    { key: "customerName", label: "お客様名" },
    { key: "date", label: "日付", type: "date" },
    { key: "location", label: "現場" },
    { key: "staff", label: "担当者" },
  ];
  document.getElementById("project-info-grid").innerHTML = fields.map(f => {
    const v = currentEstimate.project[f.key] || "";
    return `<div class="info-row">
      <label>${f.label}</label>
      <input type="${f.type||"text"}" value="${escAttr(v)}"
             onchange="currentEstimate.project['${f.key}']=this.value">
    </div>`;
  }).join("");
}

function renderEstimateLines() {
  const section = document.getElementById("est-section");
  const empty = document.getElementById("est-empty");
  const lines = currentEstimate.lines;

  // 既存のテーブルを削除（est-empty は残す）
  section.querySelectorAll(".est-col-table").forEach(el => el.remove());

  if (lines.length === 0) {
    empty.style.display = "";
    return;
  }
  empty.style.display = "none";

  const rowsHtml = _buildLineRows(lines);
  const theadHtml = _estTheadHtml();

  const html = `<table class="estimate-table est-col-table">${theadHtml}<tbody>${rowsHtml.join("")}</tbody></table>`;
  empty.insertAdjacentHTML("beforebegin", html);
}

/** 行HTML配列を生成（画面・印刷共通） */
function _buildLineRows(lines) {
  let itemNo = 0;
  return lines.map((line, i) => {
    const dragAttrs = `draggable="true" data-line-id="${line.lineId}"
      ondragstart="onLineDragStart(event,'${line.lineId}')"
      ondragover="onLineDragOver(event)"
      ondrop="onLineDrop(event,'${line.lineId}')"
      ondragend="onLineDragEnd(event)"
      ondragleave="onLineDragLeave(event)"`;

    if (line.type === "sep") {
      return `<tr class="sep-row" ${dragAttrs}>
        <td colspan="7">
          <button class="btn-sep-del no-print" onclick="removeLine('${line.lineId}')" title="区切り線を削除">&times;</button>
        </td>
      </tr>`;
    }
    if (line.type === "subtotal") {
      const sub = calcSubtotal(lines, i);
      const result = Math.round(sub * line.rate);
      return `<tr class="subtotal-row" ${dragAttrs}>
        <td class="ec-sep no-print"></td>
        <td class="subtotal-label">小計</td>
        <td class="subtotal-sum" colspan="2">${fmtNum(sub)}</td>
        <td class="subtotal-rate-cell">&times;<input type="number" step="0.01" value="${line.rate}"
          onchange="onSubtotalRate('${line.lineId}',this.value)"></td>
        <td class="subtotal-eq">=</td>
        <td class="subtotal-value">${fmtNum(result)}</td>
        <td class="ec-del no-print">
          <button class="btn btn-danger btn-sm" onclick="removeLine('${line.lineId}')">&times;</button>
        </td>
      </tr>`;
    }
    if (line.type === "comment") {
      return `<tr class="comment-row" ${dragAttrs}>
        <td class="ec-sep no-print">
          <button class="btn-sep" onclick="insertSep('${line.lineId}')" title="下に区切り線を挿入">▶</button>
          <button class="btn-cmt" onclick="insertComment('${line.lineId}')" title="下にコメント行を挿入">💬</button>
        </td>
        <td colspan="6" class="ec-comment-cell">
          <input type="text" class="comment-input" value="${esc(line.text || '')}"
            onchange="onCommentText('${line.lineId}',this.value)"
            placeholder="コメントを入力...">
          <button class="btn btn-danger btn-sm no-print" onclick="removeLine('${line.lineId}')">&times;</button>
        </td>
      </tr>`;
    }
    itemNo++;
    let name, spec, srcBadge;
    if (line.type === "custom") {
      name = line.name || "(カスタム)";
      spec = line.spec || "";
      srcBadge = "";
    } else {
      const m = getMasterItem(line.masterItemId);
      name = m ? m.name : "(不明)";
      spec = m ? (m.spec || "") : "";
      srcBadge = isCubicleItem(line.masterItemId) ? '<span class="src-cubicle">Q</span>' : '';
    }
    const sub = line.qty * line.unitPrice;

    return `<tr ${dragAttrs}>
      <td class="ec-sep no-print">
        <button class="btn-sep" onclick="insertSep('${line.lineId}')" title="下に区切り線を挿入">▶</button>
        <button class="btn-cmt" onclick="insertComment('${line.lineId}')" title="下にコメント行を挿入">💬</button>
        <button class="btn-insert-subtotal" onclick="insertSubtotal('${line.lineId}')" title="小計行を挿入">Σ</button>
      </td>
      <td class="ec-no">${itemNo}</td>
      <td class="ec-name">${srcBadge}${esc(name)}</td>
      <td class="ec-spec">${esc(spec)}</td>
      <td class="ec-qty"><input type="number" min="0" value="${line.qty}"
           onchange="onLineQty('${line.lineId}',this.value)" onfocus="this.select()"></td>
      <td class="ec-price"><input type="number" min="0" step="0.1" value="${line.unitPrice}"
           onchange="onLinePrice('${line.lineId}',this.value)" onfocus="this.select()"></td>
      <td class="ec-subtotal" id="sub-${line.lineId}">${fmtNum(sub)}</td>
      <td class="ec-del no-print">
        <button class="btn btn-danger btn-sm" onclick="removeLine('${line.lineId}')">&times;</button>
      </td>
    </tr>`;
  });
}

function _estTheadHtml() {
  return `<thead><tr>
    <th class="col-sep no-print"></th>
    <th class="col-no">No</th>
    <th class="col-name">品名</th>
    <th class="col-spec">仕様</th>
    <th class="col-qty">数量</th>
    <th class="col-price">単価</th>
    <th class="col-subtotal">金額</th>
    <th class="col-actions no-print"></th>
  </tr></thead>`;
}

// ============================================================
// 印刷用 段組みレイアウト
// ============================================================

const EST_PRINT_ROW_LIMIT = 25; // 1列あたりの最大行数（A4高さ基準）

function renderEstimateLinesForPrint() {
  const section = document.getElementById("est-section");
  const empty = document.getElementById("est-empty");
  const lines = currentEstimate.lines;

  section.querySelectorAll(".est-col-table").forEach(el => el.remove());

  if (lines.length === 0) return;
  empty.style.display = "none";

  const rowsHtml = _buildLineRows(lines);
  const theadHtml = _estTheadHtml();

  // 常に3カラム固定幅。1列目→2列目→3列目の順に埋める。
  // 3列目が溢れたら次ページでまた3カラム。
  const COL_COUNT = 3;
  const perCol = EST_PRINT_ROW_LIMIT;

  section.style.cssText = "display:flex; flex-wrap:wrap; gap:2px; align-items:flex-start; overflow:visible;";

  // 行がある分だけテーブルを作成（最低でも3つ = 1ページ分の3カラム）
  const totalCols = Math.max(COL_COUNT, Math.ceil(rowsHtml.length / perCol));
  // ページ単位で3の倍数にする（空テーブルで幅を揃える）
  const colCount = Math.ceil(totalCols / COL_COUNT) * COL_COUNT;

  for (let c = 0; c < colCount; c++) {
    const chunk = rowsHtml.slice(c * perCol, (c + 1) * perCol);
    const tbl = document.createElement("table");
    tbl.className = "estimate-table est-col-table";
    tbl.style.cssText = "flex:1 1 0; min-width:0; max-width:calc(33.33% - 2px); font-size:9px; border-collapse:collapse; table-layout:fixed;";
    if (chunk.length > 0) {
      tbl.innerHTML = theadHtml + "<tbody>" + chunk.join("") + "</tbody>";
    } else {
      // 空テーブル（幅確保用）— ヘッダーだけ表示
      tbl.innerHTML = theadHtml + "<tbody></tbody>";
    }
    section.insertBefore(tbl, empty);
  }

  // 仕様セルの文字あふれ対策: テキストが収まらなければfont-sizeを縮小
  section.querySelectorAll(".ec-spec").forEach(td => {
    td.classList.add("ec-spec-print");
    const text = td.textContent;
    if (!text) return;
    const origSize = 9;
    let fs = origSize;
    td.style.fontSize = fs + "px";
    while (td.scrollWidth > td.clientWidth && fs > 4) {
      fs -= 0.5;
      td.style.fontSize = fs + "px";
    }
  });
}

function fitPrintToOnePage() {
  // 1ページ縮小を無効化 — 収まらない場合はCSSのpage-breakで次ページへ流す
}

function clearPrintFit() {
  const container = document.querySelector(".main-container");
  if (container.dataset.printScaled) {
    container.style.transform = "";
    container.style.transformOrigin = "";
    container.style.width = "";
    delete container.dataset.printScaled;
  }
  const section = document.getElementById("est-section");
  section.style.cssText = "";
}

window.addEventListener("beforeprint", () => {
  renderEstimateLinesForPrint();
  fitPrintToOnePage();
  // 備考: textareaを非表示にし、全文表示用のdivを挿入
  const ta = document.getElementById("notes-textarea");
  if (ta && ta.value) {
    ta.style.display = "none";
    const div = document.createElement("div");
    div.id = "notes-print-div";
    div.style.cssText = "font-size:8px; white-space:pre-wrap; word-break:break-all; border:1px solid #cbd5e0; border-radius:3px; padding:6px 8px; font-family:inherit; line-height:1.5;";
    div.textContent = ta.value;
    ta.parentNode.insertBefore(div, ta.nextSibling);
  }
});
window.addEventListener("afterprint", () => {
  clearPrintFit();
  renderEstimateLines();
  // 備考: 印刷用divを削除してtextareaを復帰
  const div = document.getElementById("notes-print-div");
  if (div) div.remove();
  const ta = document.getElementById("notes-textarea");
  if (ta) ta.style.display = "";
});

// ============================================================
// 見積作成 - 行操作
// ============================================================

function onLineQty(lineId, val) {
  const l = currentEstimate.lines.find(x => x.lineId === lineId);
  if (!l) return;
  l.qty = Math.max(0, parseInt(val) || 0);
  const el = document.getElementById("sub-" + lineId);
  if (el) el.textContent = fmtNum(l.qty * l.unitPrice);
  renderTotals();
}

function onLinePrice(lineId, val) {
  const l = currentEstimate.lines.find(x => x.lineId === lineId);
  if (!l) return;
  l.unitPrice = parseFloat(val) || 0;
  const el = document.getElementById("sub-" + lineId);
  if (el) el.textContent = fmtNum(l.qty * l.unitPrice);
  renderTotals();
}

function onLineNote(lineId, val) {
  const l = currentEstimate.lines.find(x => x.lineId === lineId);
  if (l) l.lineNote = val;
}

function removeLine(lineId) {
  currentEstimate.lines = currentEstimate.lines.filter(x => x.lineId !== lineId);
  rebuildClickCounts();
  renderMasterTable();
  renderEstimateLines();
  renderTotals();
}

function rebuildClickCounts() {
  masterClickCounts = {};
  for (const line of currentEstimate.lines) {
    if (line.type === "item" && line.masterItemId) {
      masterClickCounts[line.masterItemId] = (masterClickCounts[line.masterItemId] || 0) + 1;
    }
  }
}

function insertSep(afterLineId) {
  const idx = currentEstimate.lines.findIndex(x => x.lineId === afterLineId);
  if (idx < 0) return;
  currentEstimate.lines.splice(idx + 1, 0, { type: "sep", lineId: genId() });
  renderEstimateLines();
}

function insertComment(afterLineId) {
  const idx = currentEstimate.lines.findIndex(x => x.lineId === afterLineId);
  if (idx < 0) return;
  currentEstimate.lines.splice(idx + 1, 0, { type: "comment", lineId: genId(), text: "" });
  renderEstimateLines();
  saveEstimates();
}

function onCommentText(lineId, val) {
  const line = currentEstimate.lines.find(x => x.lineId === lineId);
  if (line) { line.text = val; saveEstimates(); }
}

function insertSubtotal(afterLineId) {
  const idx = currentEstimate.lines.findIndex(x => x.lineId === afterLineId);
  if (idx < 0) return;
  currentEstimate.lines.splice(idx + 1, 0, {
    type: "subtotal", lineId: genId(), rate: 1.0, label: ""
  });
  renderEstimateLines();
  renderTotals();
}

function calcSubtotal(lines, subtotalIndex) {
  let sum = 0;
  for (let i = subtotalIndex - 1; i >= 0; i--) {
    if (lines[i].type === "subtotal") break;
    if (lines[i].type === "item" || lines[i].type === "custom") {
      sum += lines[i].qty * lines[i].unitPrice;
    }
  }
  return sum;
}

function onSubtotalRate(lineId, val) {
  const line = currentEstimate.lines.find(l => l.lineId === lineId);
  if (line) line.rate = parseFloat(val) || 1.0;
  renderEstimateLines();
  renderTotals();
}

// ============================================================
// 見積作成 - ドラッグ&ドロップ並べ替え
// ============================================================

let _dragLineId = null;

function onLineDragStart(e, lineId) {
  _dragLineId = lineId;
  e.dataTransfer.effectAllowed = "move";
  e.dataTransfer.setData("text/plain", lineId);
  requestAnimationFrame(() => {
    const row = e.target.closest("tr");
    if (row) row.classList.add("dragging");
  });
}

function onLineDragOver(e) {
  e.preventDefault();
  e.dataTransfer.dropEffect = "move";
  const row = e.target.closest("tr");
  if (!row || !row.dataset.lineId) return;
  // 挿入位置を上半分/下半分で判定
  const rect = row.getBoundingClientRect();
  const midY = rect.top + rect.height / 2;
  row.classList.remove("drag-over-top", "drag-over-bottom");
  if (e.clientY < midY) {
    row.classList.add("drag-over-top");
  } else {
    row.classList.add("drag-over-bottom");
  }
}

function onLineDragLeave(e) {
  const row = e.target.closest("tr");
  if (row) row.classList.remove("drag-over-top", "drag-over-bottom");
}

function onLineDrop(e, targetLineId) {
  e.preventDefault();
  const srcId = _dragLineId;
  if (!srcId || srcId === targetLineId) {
    _clearDragStyles();
    return;
  }
  const lines = currentEstimate.lines;
  const srcIdx = lines.findIndex(x => x.lineId === srcId);
  const tgtIdx = lines.findIndex(x => x.lineId === targetLineId);
  if (srcIdx < 0 || tgtIdx < 0) { _clearDragStyles(); return; }

  // ドロップ位置（上半分=前に、下半分=後に）
  const row = e.target.closest("tr");
  const rect = row ? row.getBoundingClientRect() : null;
  const insertBefore = rect ? e.clientY < rect.top + rect.height / 2 : false;

  // 配列操作: まず元の位置から取り出し
  const [moved] = lines.splice(srcIdx, 1);
  // 取り出し後のインデックスを再計算
  let newIdx = lines.findIndex(x => x.lineId === targetLineId);
  if (!insertBefore) newIdx++;
  lines.splice(newIdx, 0, moved);

  _clearDragStyles();
  renderEstimateLines();
  renderTotals();
}

function onLineDragEnd(e) {
  const row = e.target.closest("tr");
  if (row) row.classList.remove("dragging");
  _clearDragStyles();
}

function _clearDragStyles() {
  _dragLineId = null;
  document.querySelectorAll(".drag-over-top, .drag-over-bottom, .dragging").forEach(el => {
    el.classList.remove("drag-over-top", "drag-over-bottom", "dragging");
  });
}

// ============================================================
// 合計計算
// ============================================================

function calcGrandTotal() {
  const lines = currentEstimate.lines;
  let total = 0;
  let blockSum = 0;
  for (const line of lines) {
    if (line.type === "item" || line.type === "custom") {
      blockSum += line.qty * line.unitPrice;
    } else if (line.type === "subtotal") {
      total += Math.round(blockSum * line.rate);
      blockSum = 0;
    }
  }
  total += blockSum;
  return total;
}
function calcListPrice() { return calcGrandTotal() * currentEstimate.listRate; }
function calcNetPrice()  { return calcListPrice() * currentEstimate.netRate; }

function renderTotals() {
  document.getElementById("total-grand").textContent = fmtNum(calcGrandTotal());
  const listPrice = Math.round(calcListPrice()) * 1000;
  const netRaw = Math.round(calcNetPrice()) * 1000;
  const netPrice = Math.ceil(netRaw / 10000) * 10000; // 万単位切り上げ
  document.getElementById("total-list").textContent  = fmtNum(listPrice);
  document.getElementById("total-net").textContent    = fmtNum(netPrice);
  document.getElementById("rate-list-input").value = currentEstimate.listRate;
  document.getElementById("rate-net-input").value  = currentEstimate.netRate;
}

function renderNotes() {
  document.getElementById("notes-textarea").value = currentEstimate.notes || "";
}

function renderEstimateSelector() {
  const sel = document.getElementById("estimate-selector");
  sel.innerHTML = '<option value="">-- 保存済み見積もり --</option>' +
    savedEstimates.map(e => {
      const d = new Date(e.updatedAt);
      const ds = d.getFullYear()+"/"+(d.getMonth()+1)+"/"+d.getDate();
      return `<option value="${e.id}" ${e.id===currentEstimate.id?"selected":""}>${esc(e.name)} (${ds})</option>`;
    }).join("");
}

// ============================================================
// ユーザー操作
// ============================================================

function updateListRate(v) { currentEstimate.listRate = parseFloat(v)||DEFAULT_RATES.listRate; renderTotals(); }
function updateNetRate(v)  { currentEstimate.netRate  = parseFloat(v)||DEFAULT_RATES.netRate;  renderTotals(); }
function updateNotes(v) { currentEstimate.notes = v; }

function newEstimate() {
  if (currentEstimate.lines.length > 0 && !confirm("現在の見積もりを破棄して新規作成しますか？")) return;
  currentEstimate = createNewEstimate();
  renderEstimateTab();
  showToast("新規見積もりを作成しました");
}

function onEstimateSelect(v) {
  if (!v) return;
  if (currentEstimate.lines.length > 0 && !confirm("現在の見積もりを破棄して読み込みますか？")) {
    document.getElementById("estimate-selector").value = "";
    return;
  }
  loadEstimate(v);
}

function deleteCurrentEstimate() {
  if (!savedEstimates.find(e => e.id === currentEstimate.id)) { showToast("保存されていません"); return; }
  deleteEstimateById(currentEstimate.id);
  currentEstimate = createNewEstimate();
  renderEstimateTab();
}

// ============================================================
// エクスポート / インポート
// ============================================================

function exportJSON() {
  const blob = new Blob([JSON.stringify(currentEstimate, null, 2)], { type: "application/json" });
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob);
  a.download = (currentEstimate.name || "見積もり") + ".json"; a.click();
  URL.revokeObjectURL(a.href);
  showToast("JSONエクスポートしました");
}

function importJSON() {
  const inp = document.createElement("input"); inp.type = "file"; inp.accept = ".json";
  inp.onchange = e => {
    const f = e.target.files[0]; if (!f) return;
    const r = new FileReader();
    r.onload = ev => {
      try {
        const d = JSON.parse(ev.target.result);
        if (d.lines && d.project) { currentEstimate = d; currentEstimate.id = genId(); renderEstimateTab(); showToast("インポートしました"); }
        else alert("無効なデータです。");
      } catch { alert("JSON読み込み失敗"); }
    };
    r.readAsText(f);
  };
  inp.click();
}

function printEstimate() { window.print(); }

function exportAllData() {
  const data = {};
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i);
    data[key] = localStorage.getItem(key);
  }
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob);
  a.download = "ban_all_data.json"; a.click();
  URL.revokeObjectURL(a.href);
  showToast("全データをエクスポートしました");
}

function importAllData() {
  if (!confirm("現在のデータを上書きします。よろしいですか？")) return;
  const inp = document.createElement("input"); inp.type = "file"; inp.accept = ".json";
  inp.onchange = e => {
    const f = e.target.files[0]; if (!f) return;
    const r = new FileReader();
    r.onload = ev => {
      try {
        const data = JSON.parse(ev.target.result);
        for (const key of Object.keys(data)) {
          localStorage.setItem(key, data[key]);
        }
        showToast("全データを読み込みました。リロードします...");
        setTimeout(() => location.reload(), 1000);
      } catch { alert("データ読み込みに失敗しました"); }
    };
    r.readAsText(f);
  };
  inp.click();
}

// ============================================================
// ユーティリティ
// ============================================================

function genId() { return Date.now().toString(36) + Math.random().toString(36).slice(2,8); }
function fmtNum(n) { if (!n && n !== 0) return ""; return Number(n).toLocaleString("ja-JP"); }
function esc(s) { if (!s) return ""; const d = document.createElement("div"); d.textContent = s; return d.innerHTML; }
function escAttr(s) { return (s||"").replace(/&/g,"&amp;").replace(/"/g,"&quot;").replace(/</g,"&lt;"); }
function showToast(msg) {
  document.querySelectorAll(".toast").forEach(t => t.remove());
  const el = document.createElement("div"); el.className = "toast"; el.textContent = msg;
  document.body.appendChild(el); setTimeout(() => el.remove(), 2600);
}

// ============================================================
// 初期化
// ============================================================
document.addEventListener("DOMContentLoaded", () => {
  renderMasterTable();
});

// ページ離脱時にDOM上の価格をメモリに反映して保存
window.addEventListener("beforeunload", () => {
  document.querySelectorAll("[data-item-id]").forEach(el => {
    const id = el.dataset.itemId;
    const inp = el.querySelector("input[type='number']");
    if (inp && id) {
      const val = parseFloat(inp.value) || 0;
      const m = getMasterItem(id);
      if (m) m.basePrice = val;
    }
  });
  saveMaster();
  saveCubicle();
  // 合計欄も保存
  document.querySelectorAll(".mg-solar-total-input").forEach(inp => {
    const g = inp.dataset.group;
    if (g) groupTotals[g] = parseFloat(inp.value) || 0;
  });
  saveGroupTotals();
});
