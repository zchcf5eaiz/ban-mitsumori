// ============================================================
// 改修見積（添付Excelの考え方を、既存見積の「項目」単位で保持）
// ============================================================

const STORAGE_RENOVATION_PRICES = "ban_renovation_prices_v1";
const STORAGE_RENOVATION_MATERIAL_PRICES = "ban_renovation_material_prices_v1";
const STORAGE_RENOVATION_EQUIPMENT_MASTER = "ban_renovation_equipment_master_v1";
const RENOVATION_NET_FACTOR = 1.2;
const RENOVATION_METHODS = {
  base: { label: "一括係数", description: "TRを含む全体に定価・NETの各係数を使用" },
  transformer: { label: "TR別係数", description: "TRだけ専用の低い係数を使用" },
};
const RENOVATION_QUOTE_METHODS = {
  bara: { label: "実定価ベース（バラ）" },
  govt: { label: "役所3.8" },
  private: { label: "民間2.28" },
  private22: { label: "民間2.2" },
};
const RENOVATION_PRESENTATIONS = {
  lump: { label: "一式", description: "採用した方式定価を1行で表示" },
  detail: { label: "実定価内訳", description: "機器・材料・調査・改修を実定価で表示" },
  allocation: { label: "差額振分", description: "役所3.8／民間2.28／民間2.2の残額を4項目へ振分" },
  matsuoka: { label: "松岡Jr.", description: "松岡市役所向けの社内振分方式" },
};
const RENOVATION_MATSUOKA_RATIOS = Object.freeze({
  material: 0.10,
  survey: 0.05,
  renovation: 0.60,
  design: 0.20,
  transport: 0.02,
  management: 0.02,
  misc: 0.01,
});

let renovationCategoryId = RENOVATION_CATEGORIES[0]?.id || "";
let renovationSearchText = "";
let renovationMaterialGroup = RENOVATION_MATERIAL_GROUPS?.[0] || "電線(KIP)";
let renovationEquipmentMasterCache = null;

function renovationNotesHtml(notes, title = "元表の注意事項") {
  const rows = (notes || []).filter(Boolean);
  if (!rows.length) return "";
  return `<div class="renovation-source-notes"><strong>${esc(title)}</strong><ul>${rows.map(note => `<li>${esc(note)}</li>`).join("")}</ul></div>`;
}

const RENOVATION_AUTO_NOTES_START = "【改修・元表の注意事項】";
const RENOVATION_AUTO_NOTES_END = "【以上】";

function renovationSourceNotesForUnit(unit) {
  const state = getRenovationState(unit);
  if (!state) return [];
  const notes = new Set(RENOVATION_REFERENCE_NOTES?.general || []);
  const addNotes = rows => (rows || []).forEach(note => { if (note) notes.add(note); });
  addNotes(RENOVATION_REFERENCE_NOTES?.patterns?.[state.method]);
  addNotes(RENOVATION_REFERENCE_NOTES?.patterns?.[state.presentationMethod]);
  new Set(state.equipment.map(row => row.category).filter(Boolean)).forEach(category => {
    addNotes(RENOVATION_REFERENCE_NOTES?.categories?.[category]);
  });
  new Set(state.materials.map(row => renovationMaterialDisplayGroup(row)).filter(Boolean)).forEach(group => {
    addNotes(RENOVATION_REFERENCE_NOTES?.materialGroups?.[group]);
  });
  return [...notes];
}

function removeRenovationAutoNotes(value, previousNotes = []) {
  let text = String(value || "");
  // 旧形式の見出し付き自動注記を、既存保存データから取り除く。
  const start = text.indexOf(RENOVATION_AUTO_NOTES_START);
  if (start >= 0) {
    const end = text.indexOf(RENOVATION_AUTO_NOTES_END, start);
    const after = end < 0 ? text.length : end + RENOVATION_AUTO_NOTES_END.length;
    text = `${text.slice(0, start)}\n${text.slice(after)}`;
  }
  // 新形式は見出しを印刷しないため、前回自動生成した行を保存データで識別する。
  const previous = new Set((previousNotes || []).map(note => String(note || "").trim()).filter(Boolean));
  if (previous.size) {
    text = text.split("\n").filter(line => {
      const note = line.trim().replace(/^・/, "").trim();
      return !previous.has(note);
    }).join("\n");
  }
  return text.replace(/\n{3,}/g, "\n\n").trim();
}

/** 手入力の備考を残しつつ、選択内容に関係する元表の注意事項だけを自動更新する。 */
function syncRenovationSourceNotes() {
  const previousNotes = Array.isArray(currentEstimate.renovationAutoNotes)
    ? currentEstimate.renovationAutoNotes
    : [];
  const manualNotes = removeRenovationAutoNotes(currentEstimate.notes, previousNotes);
  const sourceNotes = new Set();
  currentEstimate.units.forEach(unit => {
    renovationSourceNotesForUnit(unit).forEach(note => sourceNotes.add(note));
  });
  currentEstimate.renovationAutoNotes = [...sourceNotes];
  const automaticNotes = [...sourceNotes].map(note => `・${note}`).join("\n");
  currentEstimate.notes = [manualNotes, automaticNotes].filter(Boolean).join("\n\n");
  const notes = document.getElementById("notes-textarea");
  if (notes) notes.value = currentEstimate.notes;
}

function loadRenovationPriceOverrides() {
  try { return JSON.parse(localStorage.getItem(STORAGE_RENOVATION_PRICES) || "{}"); }
  catch { return {}; }
}

function saveRenovationPriceOverrides(overrides) {
  localStorage.setItem(STORAGE_RENOVATION_PRICES, JSON.stringify(overrides));
}

function cloneRenovationMasterRows(value) {
  return JSON.parse(JSON.stringify(value));
}

function loadRenovationEquipmentMaster() {
  if (renovationEquipmentMasterCache) return renovationEquipmentMasterCache;
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_RENOVATION_EQUIPMENT_MASTER) || "null");
    if (saved && Array.isArray(saved.categories) && saved.categories.length && Array.isArray(saved.items)) {
      renovationEquipmentMasterCache = saved;
      return renovationEquipmentMasterCache;
    }
  } catch {}
  renovationEquipmentMasterCache = {
    version: 1,
    categories: cloneRenovationMasterRows(RENOVATION_CATEGORIES),
    items: cloneRenovationMasterRows(RENOVATION_ITEMS),
  };
  return renovationEquipmentMasterCache;
}

function saveRenovationEquipmentMaster(master) {
  master.version = 1;
  master.updatedAt = new Date().toISOString();
  renovationEquipmentMasterCache = master;
  localStorage.setItem(STORAGE_RENOVATION_EQUIPMENT_MASTER, JSON.stringify(master));
}

function renovationEquipmentCategories() {
  return loadRenovationEquipmentMaster().categories;
}

function renovationEquipmentItems() {
  return loadRenovationEquipmentMaster().items;
}

function loadRenovationMaterialPriceOverrides() {
  try { return JSON.parse(localStorage.getItem(STORAGE_RENOVATION_MATERIAL_PRICES) || "{}"); }
  catch { return {}; }
}

function saveRenovationMaterialPriceOverrides(overrides) {
  localStorage.setItem(STORAGE_RENOVATION_MATERIAL_PRICES, JSON.stringify(overrides));
}

function renovationToday() { return new Date().toISOString().slice(0, 10); }
function renovationRoundUp(value, unit = 1000) {
  const n = Number(value) || 0;
  return n <= 0 ? 0 : Math.ceil(n / unit) * unit;
}

function renovationFactor(value, fallback) {
  const n = Number(value);
  return Number.isFinite(n) && n >= 0 ? n : fallback;
}

function getRenovationMasterItem(id) {
  const original = renovationEquipmentItems().find(item => item.id === id);
  if (!original) return null;
  const override = loadRenovationPriceOverrides()[id];
  if (!override?.current) return original;
  return Object.assign({}, original, override.current, {
    history: [...(override.history || []), ...(original.history || [])],
  });
}

function getRenovationMaterialItem(id) {
  const original = RENOVATION_MATERIALS.find(item => item.id === id);
  if (!original) return null;
  const override = loadRenovationMaterialPriceOverrides()[id];
  if (!override?.current) return original;
  return Object.assign({}, original, override.current, {
    history: [...(override.history || []), ...(original.history || [])],
  });
}

function renovationCopperCapacity(spec) {
  const match = String(spec || "").match(/^([\d.]+)\s*[×xX]\s*([\d.]+)$/);
  if (!match) return null;
  const area = Number(match[1]) * Number(match[2]);
  if (!(area > 0)) return null;
  let capacity = 0;
  for (const band of RENOVATION_COPPER_CURRENT_DENSITY) {
    const candidate = area * band.density;
    if (candidate > capacity) capacity = Math.min(band.maxCurrent, candidate);
  }
  return Math.floor(capacity);
}

function renovationMaterialCapacity(item) {
  if (!item) return null;
  if (item.group === "銅バー") {
    const amps = renovationCopperCapacity(item.spec);
    return amps == null ? null : {
      short: `参考 ${fmtNum(amps)}A`,
      detail: `参考許容電流 ${fmtNum(amps)}A（断面積×表1.1.5電流密度、+5%不算入）`,
    };
  }
  const size = Number.parseFloat(String(item.spec || "").replace(/[^\d.].*$/, ""));
  const family = RENOVATION_WIRE_CAPACITY?.[item.name];
  const value = family?.[size];
  if (!value) return item.group === "電線" ? { short: "A資料なし", detail: `${item.name} ${item.spec}の許容電流資料は未登録` } : null;
  const parts = [];
  if (value.breaker != null) parts.push(`目安 ${fmtNum(value.breaker)}A`);
  if (value.allowable != null) parts.push(`許容 ${fmtNum(value.allowable)}A`);
  if (value.parallel) parts.push(value.parallel);
  return { short: parts.join(" ／ "), detail: `${parts.join(" ／ ")}（${family.condition}）` };
}

function renovationMaterialDisplayGroup(item) {
  if (item?.group !== "電線") return item?.group || "その他";
  const family = String(item.name || "").replace(/電線$/, "");
  return `電線(${family || "その他"})`;
}

function renovationEquipmentDisplayCategoryId(item) {
  if (item?.categoryId === "cat-1") {
    if (/^(VT|OPTr)\b/i.test(item.name || "")) return "cat-1-vt";
    const number = Number(String(item.id || "").match(/REN-02-(\d+)/)?.[1] || 0);
    return number >= 34 ? "cat-1-ct-meter" : "cat-1-ct-relay";
  }
  if (item?.categoryId === "cat-2") {
    return /^(VCB\b|遠方手動操作器)/i.test(item.name || "") ? "cat-2-vcb" : "cat-2-contactor";
  }
  return item?.categoryId || "";
}

function renovationEquipmentManufacturer(item) {
  if (String(item?.manufacturer || "").trim()) return String(item.manufacturer).trim();
  if (item?.categoryId !== "cat-4") return "";
  const row = Number(String(item.id || "").match(/^REN-05-(\d+)$/)?.[1] || 0);
  if (row === 5 || (row >= 8 && row <= 21)) return "三菱";
  if ((row >= 24 && row <= 31) || (row >= 35 && row <= 42)) return "オムロン";
  if (row >= 45 && row <= 46) return "光商工";
  if (row === 50 || row === 55) return "音羽電機工業";
  return "";
}

function nullableRenovationPrice(value) {
  const cleaned = String(value ?? "").trim().replace(/,/g, "");
  return cleaned === "" ? null : Math.max(0, Number(cleaned) || 0);
}

function renovationMaterialPrice(item, field) {
  if (item?.[field] != null) return Number(item[field]) || 0;
  if (field === "costPrice") return item?.unitPrice == null ? null : Number(item.unitPrice) || 0;
  return null;
}

function createRenovationState() {
  return {
    enabled: true,
    method: "base",
    quoteMethod: "bara",
    presentationMethod: "detail",
    pricingFactor: RENOVATION_NET_FACTOR,
    listPricingFactor: 1.2,
    transformerFactor: 1.1,
    materialNetFactor: 1.2,
    materialListFactor: 1.2,
    renovationListFactor: 1.5,
    pricingVersion: 4,
    equipment: [],
    materials: [],
    siteSurveyList: 60000,
    siteSurveyNet: 50000,
    siteSurveyCost: 0,
    labor: { hourlyRate: 5000, factor: 1.25, hours: 8, people: 1, cost: 0 },
    matsuokaRatios: { ...RENOVATION_MATSUOKA_RATIOS },
  };
}

function getRenovationState(unit = activeUnit()) {
  const state = unit?.renovation?.enabled ? unit.renovation : null;
  if (!state) return null;
  if (!RENOVATION_QUOTE_METHODS[state.quoteMethod]) state.quoteMethod = "bara";
  // 旧データの「積算パターン」に含まれていた表示方法を、新しい表示方法へ移行する。
  if (state.method === "allocation" || state.method === "matsuoka") {
    state.presentationMethod = state.method;
    state.method = "base";
    state.quoteMethod = "govt";
  }
  if (!RENOVATION_METHODS[state.method]) state.method = "base";
  if (!RENOVATION_PRESENTATIONS[state.presentationMethod]) {
    state.presentationMethod = state.quoteMethod === "bara" ? "detail" : "lump";
  }
  if (state.presentationMethod === "detail") state.quoteMethod = "bara";
  if (state.presentationMethod === "matsuoka") state.quoteMethod = "govt";
  if ((state.presentationMethod === "lump" || state.presentationMethod === "allocation") && state.quoteMethod === "bara") {
    state.quoteMethod = "govt";
  }
  if (state.transformerFactor == null) state.transformerFactor = 1.1;
  if (state.listPricingFactor == null) {
    // 旧版の「標準係数」は実質的にバラ定価の調整値として使われていた。
    state.listPricingFactor = (Number(state.pricingVersion) || 0) < 3
      ? renovationFactor(state.pricingFactor, 1.2)
      : 1.2;
  }
  if (state.materialNetFactor == null) state.materialNetFactor = 1.2;
  if (state.materialListFactor == null) state.materialListFactor = 1.2;
  if (state.renovationListFactor == null) state.renovationListFactor = 1.5;
  if (state.siteSurveyCost == null) state.siteSurveyCost = 0;
  if (state.labor.cost == null) state.labor.cost = 0;
  state.matsuokaRatios = { ...RENOVATION_MATSUOKA_RATIOS };
  if ((Number(state.pricingVersion) || 0) < 2) {
    state.materials.forEach(row => {
      const master = getRenovationMaterialItem(row.materialId);
      const cost = Number(row.costPrice ?? row.unitPrice) || 0;
      if (master?.netPrice == null && Number(row.netPrice) === cost) row.netPrice = null;
    });
    state.pricingVersion = 2;
  }
  if ((Number(state.pricingVersion) || 0) < 3) {
    // 保存済みの調整値はバラ定価側へ残し、NET係数だけ規定値へ戻す。
    state.pricingFactor = RENOVATION_NET_FACTOR;
    state.pricingVersion = 3;
  }
  if ((Number(state.pricingVersion) || 0) < 4) {
    state.pricingFactor = RENOVATION_NET_FACTOR;
    state.pricingVersion = 4;
  }
  return state;
}

function startRenovationUnit() {
  let unit = activeUnit();
  if (unit.lines.length && !unit.renovation?.enabled) {
    const number = currentEstimate.units.filter(u => u.renovation?.enabled).length + 1;
    currentEstimate.units.push({
      unitId: genId(), unitName: `改修工事${number}`, lines: [],
      listRate: DEFAULT_RATES.listRate, netRate: DEFAULT_RATES.netRate,
    });
    activeUnitIndex = currentEstimate.units.length - 1;
    unit = activeUnit();
  } else if (/^項目\d+$/.test(unit.unitName || "")) {
    unit.unitName = "改修工事";
  }
  unit.renovation = createRenovationState();
  syncRenovationQuoteLine(unit);
  renderRenovationTab();
  renderCatalogEstimatePanels();
  showToast("改修見積を開始しました");
}

function calcRenovationPricing(unit) {
  const state = getRenovationState(unit);
  if (!state) return null;
  const sum = (rows, field) => rows.reduce((total, row) => total + (Number(row.qty) || 0) * (Number(row[field]) || 0), 0);
  const trRows = state.equipment.filter(row => row.category === "TR");
  const equipmentList = sum(state.equipment, "listPrice");
  const equipmentNet = sum(state.equipment, "netPrice");
  const equipmentCost = sum(state.equipment, "costPrice");
  const trList = sum(trRows, "listPrice");
  const trNet = sum(trRows, "netPrice");
  const trCost = sum(trRows, "costPrice");
  const materialNetFactor = renovationFactor(state.materialNetFactor, 1.2);
  const materialListFactor = renovationFactor(state.materialListFactor, 1.2);
  const materialCost = state.materials.reduce((total, row) => total + (Number(row.qty) || 0) * (Number(row.costPrice ?? row.unitPrice) || 0), 0);
  const materialNetBase = state.materials.reduce((total, row) => {
    const qty = Number(row.qty) || 0;
    const costPrice = Number(row.costPrice ?? row.unitPrice) || 0;
    const unitNet = row.netPrice == null ? costPrice * materialNetFactor : Number(row.netPrice) || 0;
    return total + qty * unitNet;
  }, 0);
  const materialNet = renovationRoundUp(materialNetBase, 1000);
  const labor = renovationRoundUp(
    (Number(state.labor.hourlyRate) || 0) * (Number(state.labor.factor) || 0) *
    (Number(state.labor.hours) || 0) * (Number(state.labor.people) || 0), 1000
  );
  const laborCost = Number(state.labor.cost) || 0;
  const hasExplicitMaterialList = state.materials.some(row => row.listPrice != null);
  const materialListBase = hasExplicitMaterialList ? state.materials.reduce((total, row) => {
    const qty = Number(row.qty) || 0;
    const costPrice = Number(row.costPrice ?? row.unitPrice) || 0;
    const unitNet = row.netPrice == null ? costPrice * materialNetFactor : Number(row.netPrice) || 0;
    const unit = row.listPrice == null ? unitNet * materialListFactor : Number(row.listPrice) || 0;
    return total + qty * unit;
  }, 0) : materialNet * materialListFactor;
  const materialList = renovationRoundUp(materialListBase, 1000);
  const renovationList = renovationRoundUp(labor * renovationFactor(state.renovationListFactor, 1.5), 1000);
  const directList = equipmentList + materialList + (Number(state.siteSurveyList) || 0) + renovationList;
  const directNet = equipmentNet + materialNet + (Number(state.siteSurveyNet) || 0) + labor;
  const cost = equipmentCost + materialCost + (Number(state.siteSurveyCost) || 0) + laborCost;
  const netFactor = RENOVATION_NET_FACTOR;
  const listFactor = renovationFactor(state.listPricingFactor, 1.2);
  const trFactor = renovationFactor(state.transformerFactor, 1.1);
  const splitTr = state.method === "transformer";
  const normalListBase = directList - trList;
  const normalNetBase = directNet - trNet;
  const buildCandidate = (kind) => {
    let normalList, normalNet, trPartList = 0, trPartNet = 0;
    if (kind === "bara") {
      normalList = renovationRoundUp(normalListBase * listFactor, 1000);
      normalNet = renovationRoundUp(normalNetBase * netFactor * 1.15, 10000);
      if (splitTr) {
        trPartList = renovationRoundUp(trList * trFactor, 1000);
        trPartNet = renovationRoundUp(trNet * trFactor * 1.15, 10000);
      }
    } else {
      const multiplier = kind === "govt" ? 3.8 : kind === "private22" ? 2.2 : 2.28;
      const netRate = kind === "govt" ? 0.33 : 0.55;
      normalList = renovationRoundUp(normalNetBase * netFactor * multiplier, 1000);
      normalNet = renovationRoundUp(normalList * netRate, 10000);
      if (splitTr) {
        trPartList = renovationRoundUp(trNet * trFactor * multiplier, 1000);
        trPartNet = renovationRoundUp(trPartList * netRate, 10000);
      }
    }
    // TR別計算を使わない場合はTRも標準側に含まれる。
    if (!splitTr) {
      if (kind === "bara") {
        normalList = renovationRoundUp(directList * listFactor, 1000);
        normalNet = renovationRoundUp(directNet * netFactor * 1.15, 10000);
      } else {
        const multiplier = kind === "govt" ? 3.8 : kind === "private22" ? 2.2 : 2.28;
        const netRate = kind === "govt" ? 0.33 : 0.55;
        normalList = renovationRoundUp(directNet * netFactor * multiplier, 1000);
        normalNet = renovationRoundUp(normalList * netRate, 10000);
      }
    }
    const list = normalList + trPartList;
    const net = normalNet + trPartNet;
    const costRate = net > 0 ? cost / net * 100 : 0;
    const equipmentCostRate = net > 0 ? (equipmentCost + materialCost) / net * 100 : 0;
    const grossMargin = net > 0 ? (net - cost) / net * 100 : 0;
    return {
      list, net, overhead: Math.max(0, list - directList), costRate, equipmentCostRate, grossMargin,
      netToListRate: list > 0 ? net / list * 100 : 0,
    };
  };
  const candidates = {
    bara: buildCandidate("bara"),
    govt: buildCandidate("govt"),
    private: buildCandidate("private"),
    private22: buildCandidate("private22"),
  };
  if (!RENOVATION_QUOTE_METHODS[state.quoteMethod]) state.quoteMethod = "bara";
  const selected = candidates[state.quoteMethod];
  const formulaList = selected.list;
  const formulaNet = selected.net;
  const distributionDifference = formulaList - equipmentList;
  const distributionFeasible = distributionDifference >= 0;
  const distributionOverrun = Math.max(0, -distributionDifference);

  const allocations = {};
  if (state.presentationMethod === "allocation" && distributionFeasible) {
    const difference = distributionDifference;
    const bases = {
      material: materialList,
      survey: Number(state.siteSurveyList) || 0,
      renovation: renovationList,
      overhead: candidates.bara.overhead,
    };
    const denominator = Object.values(bases).reduce((total, value) => total + value, 0) || 1;
    const entries = Object.entries(bases);
    let allocated = 0;
    entries.forEach(([key, value], index) => {
      const remaining = Math.max(0, difference - allocated);
      const amount = index === entries.length - 1
        ? remaining
        : Math.min(remaining, renovationRoundUp(difference * value / denominator, 1000));
      allocations[key] = amount;
      allocated += amount;
    });
  } else if (state.presentationMethod === "matsuoka" && distributionFeasible) {
    const entries = Object.entries(state.matsuokaRatios || {});
    let allocated = 0;
    entries.forEach(([key, ratio], index) => {
      const difference = distributionDifference;
      const remaining = Math.max(0, difference - allocated);
      const value = index === entries.length - 1
        ? remaining
        : Math.min(remaining, renovationRoundUp(difference * (Number(ratio) || 0), 1000));
      allocations[key] = value;
      allocated += value;
    });
  }
  const detailBelowActual = state.presentationMethod === "detail" && formulaList < directList;
  const distributionFallback = (state.presentationMethod === "allocation" || state.presentationMethod === "matsuoka") && !distributionFeasible;
  const list = detailBelowActual || distributionFallback ? directList : formulaList;
  const net = detailBelowActual || distributionFallback ? directNet : formulaNet;
  const overhead = Math.max(0, list - directList);
  return {
    raw: cost, list, net, overhead, allocations, candidates,
    equipmentList, equipmentNet, equipmentCost,
    materialCost, materialNet, materialList, labor, laborCost, renovationList, directList, directNet,
    trList, trNet, trCost,
    distributionDifference, distributionFeasible, distributionOverrun,
    calculationMethodLabel: RENOVATION_METHODS[state.method]?.label || "改修",
    quoteMethodLabel: RENOVATION_QUOTE_METHODS[state.quoteMethod]?.label || "実定価ベース（バラ）",
    presentationMethodLabel: RENOVATION_PRESENTATIONS[state.presentationMethod]?.label || "一式",
    rateLabel: `${RENOVATION_METHODS[state.method]?.label || "改修"}・${RENOVATION_QUOTE_METHODS[state.quoteMethod].label}・${RENOVATION_PRESENTATIONS[state.presentationMethod]?.label || "一式"}`,
  };
}

/** 改修1項目分の社内計算情報。画面表示と印刷の隣接列で共用する。 */
function renovationInternalCardHtml(unit) {
  const state = getRenovationState(unit);
  const pricing = calcRenovationPricing(unit);
  if (!state || !pricing) return "";
  const percent = (value, base) => Number(base) > 0 ? `${(Number(value) / Number(base) * 100).toFixed(1)}%` : "—";
  const totalCost = Number(pricing.raw) || 0;
  const profit = (Number(pricing.net) || 0) - totalCost;
  const method = RENOVATION_METHODS[state.method]?.label || "改修";
  const quote = RENOVATION_QUOTE_METHODS[state.quoteMethod]?.label || "バラ方式";
  const presentation = RENOVATION_PRESENTATIONS[state.presentationMethod]?.label || "一式";
  const row = (label, list, net, cost, className = "") => `<tr${className ? ` class="${className}"` : ""}>` +
    `<td>${esc(label)}</td><td>${fmtNum(list)}</td><td>${fmtNum(net)}</td><td>${cost == null ? "—" : fmtNum(cost)}</td></tr>`;
  const detailRow = (item, list, net, cost) => `<tr>` +
    `<td><strong>${item.manufacturer ? `${esc(item.manufacturer)}｜` : ""}${esc(item.name || "未名称")}</strong>${item.spec ? `<small>${esc(item.spec)}</small>` : ""}</td>` +
    `<td>${fmtNum(item.qty)}</td><td>${fmtNum(list)}</td><td>${fmtNum(net)}</td><td>${fmtNum(cost)}</td></tr>`;
  const equipmentRows = state.equipment.map(item => detailRow(
    item,
    Number(item.listPrice) || 0,
    Number(item.netPrice) || 0,
    Number(item.costPrice) || 0
  )).join("") || `<tr><td colspan="5" class="renovation-internal-empty">選択なし</td></tr>`;
  const materialRows = state.materials.map(item => {
    const cost = Number(item.costPrice ?? item.unitPrice) || 0;
    const net = item.netPrice == null ? cost * state.materialNetFactor : Number(item.netPrice) || 0;
    const list = item.listPrice == null ? net * state.materialListFactor : Number(item.listPrice) || 0;
    return detailRow(item, list, net, cost);
  }).join("") || `<tr><td colspan="5" class="renovation-internal-empty">選択なし</td></tr>`;
  const detailSectionRow = (title, note, rows) => `<tr class="renovation-internal-detail-row"><td colspan="4">` +
    `<div class="renovation-internal-details"><h4>${esc(title)} <small>${esc(note)}</small></h4>` +
    `<table><thead><tr><th>品名・仕様</th><th>数量</th><th>定価</th><th>NET</th><th>仕入</th></tr></thead>` +
    `<tbody>${rows}</tbody></table></div></td></tr>`;
  const conditionGroup = (number, label, options, currentKey) => `<div class="renovation-condition-group">` +
    `<strong><b>${number}</b>${esc(label)}</strong><div>${Object.entries(options).map(([key, option]) => {
      const current = key === currentKey;
      return `<span class="renovation-condition-option${current ? " current" : ""}">${esc(option.label)}` +
        `${current ? `<em>選択中</em>` : ""}</span>`;
    }).join("")}</div></div>`;
  const candidateCards = Object.entries(RENOVATION_QUOTE_METHODS).map(([key, option]) => {
    const candidate = pricing.candidates[key];
    if (!candidate) return "";
    const current = key === state.quoteMethod;
    return `<article class="renovation-pattern-card${current ? " current" : ""}">` +
      `<div class="renovation-pattern-card-heading"><strong>${esc(option.label)}</strong>` +
      `${current ? `<span class="renovation-pattern-badge">選択中</span>` : ""}</div>` +
      `<div class="renovation-pattern-prices"><span>定価<b>${fmtNum(candidate.list)}</b></span><span>NET<b>${fmtNum(candidate.net)}</b></span></div>` +
      `<div class="renovation-pattern-metrics">` +
      `<span>NET／定価<b>${candidate.netToListRate.toFixed(1)}%</b></span>` +
      `<span>機器原価率<b>${candidate.equipmentCostRate.toFixed(1)}%</b></span>` +
      `<span>総原価率<b>${candidate.costRate.toFixed(1)}%</b></span>` +
      `<span>粗利率<b>${candidate.grossMargin.toFixed(1)}%</b></span>` +
      `</div></article>`;
  }).join("");

  return `<section class="renovation-internal-card">` +
    `<div class="renovation-internal-card-heading"><strong>${esc(unit.unitName || "改修工事")}</strong>` +
    `<span>${esc(method)}・${esc(quote)}・${esc(presentation)}</span></div>` +
    `<div class="renovation-internal-conditions"><h4>選択した見積条件 <small>全候補と現在の選択</small></h4>` +
    conditionGroup("①", "TRの扱い", RENOVATION_METHODS, state.method) +
    conditionGroup("②", "定価・NET算出方式", RENOVATION_QUOTE_METHODS, state.quoteMethod) +
    conditionGroup("③", "見積明細の出し方", RENOVATION_PRESENTATIONS, state.presentationMethod) +
    `</div>` +
    `<table class="renovation-internal-total-table"><thead><tr><th>計算項目</th><th>定価</th><th>NET</th><th>仕入・社内原価</th></tr></thead><tbody>` +
    row("機器", pricing.equipmentList, pricing.equipmentNet, pricing.equipmentCost) +
    detailSectionRow("機器内訳", "各単価", equipmentRows) +
    row("材料費", pricing.materialList, pricing.materialNet, pricing.materialCost) +
    detailSectionRow("材料内訳", "各単価／空欄登録は係数計算", materialRows) +
    row("現地調査費", state.siteSurveyList, state.siteSurveyNet, state.siteSurveyCost) +
    row("改修費", pricing.renovationList, pricing.labor, pricing.laborCost) +
    row("方式差額・諸経費", pricing.list - pricing.directList, pricing.net - pricing.directNet, null) +
    row("採用合計", pricing.list, pricing.net, totalCost, "renovation-internal-total") +
    `</tbody></table>` +
    `<div class="renovation-internal-metrics">` +
    `<span>NET利益額<b>${fmtNum(profit)}</b></span>` +
    `<span>粗利率<b>${percent(profit, pricing.net)}</b></span>` +
    `<span>総原価率<b>${percent(totalCost, pricing.net)}</b></span>` +
    `<span>NET／定価<b>${percent(pricing.net, pricing.list)}</b></span>` +
    `</div>` +
    `<div class="renovation-internal-patterns">` +
    `<h4>定価・NET算出方式 4パターン <small>提出NET比較用</small></h4>` +
    `<div class="renovation-pattern-card-grid">${candidateCards}</div>` +
    `<p>現在の見積には「選択中」の方式を反映。ほか3方式は提出NET・原価率・粗利率の比較用です。</p>` +
    `</div></section>`;
}

/** 印刷用の社内計算情報を、A4相当の高さごとの列HTMLへ分割する。 */
function renovationInternalPrintColumns(unit, rowLimit = 35) {
  const state = getRenovationState(unit);
  const pricing = calcRenovationPricing(unit);
  if (!state || !pricing) return [];
  const percent = (value, base) => Number(base) > 0 ? `${(Number(value) / Number(base) * 100).toFixed(1)}%` : "—";
  const totalCost = Number(pricing.raw) || 0;
  const profit = (Number(pricing.net) || 0) - totalCost;
  const method = RENOVATION_METHODS[state.method]?.label || "改修";
  const quote = RENOVATION_QUOTE_METHODS[state.quoteMethod]?.label || "バラ方式";
  const presentation = RENOVATION_PRESENTATIONS[state.presentationMethod]?.label || "一式";
  const columns = [[]];
  const weights = [0];
  const nextColumn = () => { columns.push([]); weights.push(0); };
  const addBlock = (html, weight) => {
    const ci = columns.length - 1;
    if (weights[ci] > 0 && weights[ci] + weight > rowLimit) nextColumn();
    columns[columns.length - 1].push(html);
    weights[weights.length - 1] += weight;
  };
  const totalRow = (label, list, net, cost, className = "") => `<tr${className ? ` class="${className}"` : ""}>` +
    `<td>${esc(label)}</td><td>${fmtNum(list)}</td><td>${fmtNum(net)}</td><td>${cost == null ? "—" : fmtNum(cost)}</td></tr>`;
  const totalTable = rows => `<table class="renovation-internal-total-table"><thead><tr><th>計算項目</th><th>定価</th><th>NET</th><th>仕入・社内原価</th></tr></thead><tbody>${rows}</tbody></table>`;
  const detailRow = (item, list, net, cost) => `<tr><td><strong>${item.manufacturer ? `${esc(item.manufacturer)}｜` : ""}${esc(item.name || "未名称")}</strong>` +
    `${item.spec ? `<small>${esc(item.spec)}</small>` : ""}</td><td>${fmtNum(item.qty)}</td>` +
    `<td>${fmtNum(list)}</td><td>${fmtNum(net)}</td><td>${fmtNum(cost)}</td></tr>`;
  const detailTable = (title, note, rows, continued) => `<div class="renovation-internal-details">` +
    `<h4>${esc(title)}${continued ? "（続き）" : ""} <small>${esc(note)}</small></h4>` +
    `<table><thead><tr><th>品名・仕様</th><th>数量</th><th>定価</th><th>NET</th><th>仕入</th></tr></thead>` +
    `<tbody>${rows.join("") || `<tr><td colspan="5" class="renovation-internal-empty">選択なし</td></tr>`}</tbody></table></div>`;
  const addDetails = (title, note, rows) => {
    const sourceRows = rows.length ? rows : [""];
    let index = 0;
    let continued = false;
    while (index < sourceRows.length) {
      let ci = columns.length - 1;
      if (weights[ci] > rowLimit - 3) { nextColumn(); ci++; }
      const capacity = Math.max(1, Math.floor(rowLimit - weights[ci] - 2));
      const chunk = sourceRows.slice(index, index + capacity);
      addBlock(detailTable(title, note, chunk[0] === "" ? [] : chunk, continued), 2 + chunk.length);
      index += chunk.length;
      continued = true;
    }
  };
  const conditionGroup = (number, label, options, currentKey) => `<div class="renovation-condition-group">` +
    `<strong><b>${number}</b>${esc(label)}</strong><div>${Object.entries(options).map(([key, option]) => {
      const current = key === currentKey;
      return `<span class="renovation-condition-option${current ? " current" : ""}">${esc(option.label)}` +
        `${current ? `<em>選択中</em>` : ""}</span>`;
    }).join("")}</div></div>`;
  addBlock(`<div class="renovation-internal-conditions"><h4>選択した見積条件 <small>全候補と現在の選択</small></h4>` +
    conditionGroup("①", "TRの扱い", RENOVATION_METHODS, state.method) +
    conditionGroup("②", "定価・NET算出方式", RENOVATION_QUOTE_METHODS, state.quoteMethod) +
    conditionGroup("③", "見積明細の出し方", RENOVATION_PRESENTATIONS, state.presentationMethod) + `</div>`, 5);

  addBlock(totalTable(totalRow("機器", pricing.equipmentList, pricing.equipmentNet, pricing.equipmentCost)), 2);
  addDetails("機器内訳", "各単価", state.equipment.map(item => detailRow(
    item, Number(item.listPrice) || 0, Number(item.netPrice) || 0, Number(item.costPrice) || 0
  )));
  addBlock(totalTable(totalRow("材料費", pricing.materialList, pricing.materialNet, pricing.materialCost)), 2);
  addDetails("材料内訳", "各単価／空欄登録は係数計算", state.materials.map(item => {
    const cost = Number(item.costPrice ?? item.unitPrice) || 0;
    const net = item.netPrice == null ? cost * state.materialNetFactor : Number(item.netPrice) || 0;
    const list = item.listPrice == null ? net * state.materialListFactor : Number(item.listPrice) || 0;
    return detailRow(item, list, net, cost);
  }));
  addBlock(totalTable(
    totalRow("現地調査費", state.siteSurveyList, state.siteSurveyNet, state.siteSurveyCost) +
    totalRow("改修費", pricing.renovationList, pricing.labor, pricing.laborCost) +
    totalRow("方式差額・諸経費", pricing.list - pricing.directList, pricing.net - pricing.directNet, null) +
    totalRow("採用合計", pricing.list, pricing.net, totalCost, "renovation-internal-total")
  ), 5);
  addBlock(`<div class="renovation-internal-metrics">` +
    `<span>NET利益額<b>${fmtNum(profit)}</b></span><span>粗利率<b>${percent(profit, pricing.net)}</b></span>` +
    `<span>総原価率<b>${percent(totalCost, pricing.net)}</b></span><span>NET／定価<b>${percent(pricing.net, pricing.list)}</b></span></div>`, 2);
  const candidateCards = Object.entries(RENOVATION_QUOTE_METHODS).map(([key, option]) => {
    const candidate = pricing.candidates[key];
    const current = key === state.quoteMethod;
    return `<article class="renovation-pattern-card${current ? " current" : ""}"><div class="renovation-pattern-card-heading">` +
      `<strong>${esc(option.label)}</strong>${current ? `<span class="renovation-pattern-badge">選択中</span>` : ""}</div>` +
      `<div class="renovation-pattern-prices"><span>定価<b>${fmtNum(candidate.list)}</b></span><span>NET<b>${fmtNum(candidate.net)}</b></span></div>` +
      `<div class="renovation-pattern-metrics"><span>NET／定価<b>${candidate.netToListRate.toFixed(1)}%</b></span>` +
      `<span>機器原価率<b>${candidate.equipmentCostRate.toFixed(1)}%</b></span><span>総原価率<b>${candidate.costRate.toFixed(1)}%</b></span>` +
      `<span>粗利率<b>${candidate.grossMargin.toFixed(1)}%</b></span></div></article>`;
  }).join("");
  addBlock(`<div class="renovation-internal-patterns"><h4>定価・NET算出方式 4パターン <small>提出NET比較用</small></h4>` +
    `<div class="renovation-pattern-card-grid">${candidateCards}</div>` +
    `<p>現在の見積には「選択中」の方式を反映。ほか3方式は提出NET・原価率・粗利率の比較用です。</p></div>`, 8);

  return columns.map((blocks, index) => `<section class="renovation-internal-card renovation-internal-print-flow">` +
    `<div class="renovation-internal-card-heading"><strong>${esc(unit.unitName || "改修工事")}${index ? "（続き）" : ""}</strong>` +
    `<span>${esc(method)}・${esc(quote)}・${esc(presentation)}</span></div>${blocks.join("")}</section>`);
}

function syncRenovationQuoteLine(unit = activeUnit()) {
  const state = getRenovationState(unit);
  if (!state) return;
  const pricing = calcRenovationPricing(unit);
  const previousIds = new Map(unit.lines.filter(row => row.renovationGenerated).map(row => [row.renovationKey || "summary", row.lineId]));
  const normalLines = unit.lines.filter(row => !row.renovationGenerated);
  const generated = [];
  const add = (key, name, spec, qty, unitPrice) => {
    if (!(Number(qty) > 0) && !(Number(unitPrice) > 0)) return;
    generated.push({
      lineId: previousIds.get(key) || genId(), type: "custom", renovationGenerated: true,
      renovationKey: key, name, spec, qty: Number(qty) || 0, unitPrice: Number(unitPrice) || 0,
    });
  };
  const methodLabel = RENOVATION_METHODS[state.method]?.label || "改修";
  const quoteLabel = RENOVATION_QUOTE_METHODS[state.quoteMethod]?.label || "バラ方式";
  const presentation = state.presentationMethod || "detail";
  const splitRows = presentation !== "lump";

  if (!splitRows) {
    const spec = `${methodLabel}・${quoteLabel}／バラ定価×${state.listPricingFactor}・NET×${RENOVATION_NET_FACTOR}` +
      `${state.method === "transformer" ? `・TR×${state.transformerFactor}` : ""}` +
      `／材料NET×${state.materialNetFactor}・材料定価×${state.materialListFactor}・改修×${state.renovationListFactor}`;
    add("summary", unit.unitName || "改修工事", spec, 1, pricing.list);
  } else {
    state.equipment.forEach((row, index) => add(
      `equipment:${row.lineId || index}`,
      row.name || "機器",
      row.spec || row.category || "",
      row.qty,
      row.listPrice
    ));

    if ((presentation === "allocation" || presentation === "matsuoka") && pricing.distributionFeasible) {
      const labels = {
        material: "材料費", survey: "現地調査費", renovation: "改修費", overhead: "諸経費",
        design: "設計費", transport: "運搬費", management: "管理費", misc: "雑費",
      };
      const allocationLabel = presentation === "matsuoka"
        ? "松岡Jr.社内比率"
        : `${quoteLabel}の差額按分`;
      Object.entries(pricing.allocations).forEach(([key, value]) =>
        add(`allocation:${key}`, labels[key] || key, allocationLabel, 1, value));
    } else {
      add("material", "材料費", `登録価格優先／未登録は原価×${state.materialNetFactor}×${state.materialListFactor}`, 1, pricing.materialList);
      add("survey", "現地調査費", `入力NET ${fmtNum(state.siteSurveyNet)}`, 1, state.siteSurveyList);
      add("renovation", "改修費", `${fmtNum(state.labor.hourlyRate)}円 × ${state.labor.factor} × ${state.labor.hours}時間 × ${state.labor.people}人`, 1, pricing.renovationList);
      if (presentation === "allocation" || presentation === "matsuoka") {
        add("distribution-warning", "【振分不可】実定価内訳へ自動切替", `方式定価を機器実定価が${fmtNum(pricing.distributionOverrun)}円超過`, 1, 0);
      } else {
        add("overhead", "諸経費", `${methodLabel}・${quoteLabel}の差額`, 1, pricing.overhead);
      }
    }
  }
  if (Array.isArray(state.lineOrder) && state.lineOrder.length) {
    const order = new Map(state.lineOrder.map((key, index) => [key, index]));
    generated.sort((a, b) => {
      const ai = order.has(a.renovationKey) ? order.get(a.renovationKey) : Number.MAX_SAFE_INTEGER;
      const bi = order.has(b.renovationKey) ? order.get(b.renovationKey) : Number.MAX_SAFE_INTEGER;
      return ai - bi;
    });
  }
  state.lineOrder = generated.map(row => row.renovationKey);
  unit.lines = [...normalLines, ...generated];
}

/** 見積作成で並べ替えた改修明細の順番を、次回再計算後も維持する。 */
function rememberRenovationLineOrder(unit) {
  const state = getRenovationState(unit);
  if (!state) return;
  state.lineOrder = unit.lines.filter(row => row.renovationGenerated).map(row => row.renovationKey);
}

function syncAllRenovationQuoteLines() {
  currentEstimate.units.forEach(unit => {
    if (unit?.renovation?.enabled) syncRenovationQuoteLine(unit);
  });
  syncRenovationSourceNotes();
}

function captureRenovationViewPosition() {
  const root = document.getElementById("renovation-root");
  return {
    windowX: window.scrollX,
    windowY: window.scrollY,
    equipmentTop: root?.querySelector("#renovation-catalog-list")?.scrollTop || 0,
    materialTop: root?.querySelector(".renovation-material-catalog-scroll")?.scrollTop || 0,
    selectionTop: root?.querySelector(".renovation-work .renovation-table-wrap")?.scrollTop || 0,
  };
}

function restoreRenovationViewPosition(position) {
  if (!position) return;
  const root = document.getElementById("renovation-root");
  const equipment = root?.querySelector("#renovation-catalog-list");
  const material = root?.querySelector(".renovation-material-catalog-scroll");
  const selection = root?.querySelector(".renovation-work .renovation-table-wrap");
  if (equipment) equipment.scrollTop = position.equipmentTop;
  if (material) material.scrollTop = position.materialTop;
  if (selection) selection.scrollTop = position.selectionTop;
  window.scrollTo(position.windowX, position.windowY);
}

function refreshRenovation() {
  const viewPosition = captureRenovationViewPosition();
  const unit = activeUnit();
  syncRenovationQuoteLine(unit);
  syncRenovationSourceNotes();
  renderRenovationTab();
  renderCatalogEstimatePanels();
  restoreRenovationViewPosition(viewPosition);
  requestAnimationFrame(() => restoreRenovationViewPosition(viewPosition));
}

function setRenovationMethod(method) {
  const state = getRenovationState();
  if (!state || !RENOVATION_METHODS[method]) return;
  state.method = method;
  refreshRenovation();
}

function setRenovationQuoteMethod(method) {
  const state = getRenovationState();
  if (!state || !RENOVATION_QUOTE_METHODS[method]) return;
  if (state.presentationMethod === "detail" && method !== "bara") {
    showToast("実定価内訳は実定価ベース（バラ）で固定されます");
    return;
  }
  if (state.presentationMethod === "matsuoka" && method !== "govt") {
    showToast("松岡Jr.は役所3.8方式で固定されます");
    return;
  }
  if ((state.presentationMethod === "lump" || state.presentationMethod === "allocation") && method === "bara") {
    showToast(`${RENOVATION_PRESENTATIONS[state.presentationMethod].label}は役所3.8方式、民間2.28方式または民間2.2方式を選択してください`);
    return;
  }
  state.quoteMethod = method;
  refreshRenovation();
}

function setRenovationPresentation(method) {
  const state = getRenovationState();
  if (!state || !RENOVATION_PRESENTATIONS[method]) return;
  state.presentationMethod = method;
  if (method === "detail") state.quoteMethod = "bara";
  if (method === "matsuoka") state.quoteMethod = "govt";
  if ((method === "lump" || method === "allocation") && state.quoteMethod === "bara") state.quoteMethod = "govt";
  refreshRenovation();
  const pricing = calcRenovationPricing();
  if ((method === "allocation" || method === "matsuoka") && !pricing.distributionFeasible) {
    showToast(`振分不可：機器実定価が方式定価を${fmtNum(pricing.distributionOverrun)}円超過しています`);
  }
}

function updateRenovationNumber(path, value) {
  const state = getRenovationState();
  if (!state) return;
  const number = Math.max(0, Number(value) || 0);
  if (path.startsWith("labor.")) state.labor[path.split(".")[1]] = number;
  else state[path] = number;
  refreshRenovation();
}

function addRenovationEquipment(id) {
  const state = getRenovationState();
  const item = getRenovationMasterItem(id);
  if (!state || !item) return;
  const found = state.equipment.find(row => row.itemId === id && row.listPrice === item.listPrice && row.costPrice === item.costPrice);
  if (found) found.qty += 1;
  else state.equipment.push({
    lineId: genId(), itemId: id, name: item.name, spec: item.spec || "", category: item.category,
    manufacturer: renovationEquipmentManufacturer(item),
    qty: item.defaultQty || 1, listPrice: item.listPrice || 0, netPrice: item.netPrice || 0,
    costPrice: item.costPrice || 0, note: item.note || "", priceLabel: item.priceLabel || "", effectiveDate: item.effectiveDate || null,
  });
  refreshRenovation();
}

function updateRenovationEquipment(lineId, field, value) {
  const row = getRenovationState()?.equipment.find(item => item.lineId === lineId);
  if (!row) return;
  row[field] = field === "note" ? String(value || "") : Math.max(0, Number(value) || 0);
  refreshRenovation();
}

function removeRenovationEquipment(lineId) {
  const state = getRenovationState();
  if (!state) return;
  state.equipment = state.equipment.filter(row => row.lineId !== lineId);
  refreshRenovation();
}

function addRenovationMaterial(id) {
  const state = getRenovationState();
  const item = getRenovationMaterialItem(id);
  if (!state || !item) return;
  const found = state.materials.find(row => row.materialId === id);
  if (found) found.qty += 1;
  else state.materials.push({
    lineId: genId(), materialId: id, group: item.group || "その他", name: item.name, spec: item.spec, unit: item.unit, qty: 1,
    unitPrice: renovationMaterialPrice(item, "costPrice") || 0,
    listPrice: item.listPrice ?? null,
    netPrice: renovationMaterialPrice(item, "netPrice"),
    costPrice: renovationMaterialPrice(item, "costPrice") || 0,
    note: item.note || "",
  });
  refreshRenovation();
}

function updateRenovationMaterial(lineId, field, value) {
  const row = getRenovationState()?.materials.find(item => item.lineId === lineId);
  if (!row) return;
  if (field === "note") row[field] = String(value || "");
  else if ((field === "listPrice" || field === "netPrice") && String(value).trim() === "") row[field] = null;
  else row[field] = Math.max(0, Number(value) || 0);
  refreshRenovation();
}

function removeRenovationMaterial(lineId) {
  const state = getRenovationState();
  if (!state) return;
  state.materials = state.materials.filter(row => row.lineId !== lineId);
  refreshRenovation();
}

function setRenovationCategory(id) {
  renovationCategoryId = id;
  renovationSearchText = "";
  const search = document.querySelector(".renovation-search");
  if (search) search.value = "";
  updateRenovationCategoryHighlight();
  renderRenovationCatalog();
}
function setRenovationSearch(value) {
  renovationSearchText = value.trim().toLowerCase();
  updateRenovationCategoryHighlight();
  renderRenovationCatalog();
}

function renovationMasterSafeId(prefix) {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

function refreshRenovationMasterView() {
  const viewPosition = captureRenovationViewPosition();
  const categories = renovationEquipmentCategories();
  if (!categories.some(category => category.id === renovationCategoryId)) renovationCategoryId = categories[0]?.id || "";
  renderRenovationTab();
  restoreRenovationViewPosition(viewPosition);
  requestAnimationFrame(() => restoreRenovationViewPosition(viewPosition));
}

function showRenovationCategoryManager() {
  document.getElementById("renovation-category-manager")?.remove();
  const categories = renovationEquipmentCategories();
  const items = renovationEquipmentItems();
  const overlay = document.createElement("div");
  overlay.id = "renovation-category-manager";
  overlay.className = "renovation-modal-overlay";
  overlay.onclick = event => { if (event.target === overlay) overlay.remove(); };
  overlay.innerHTML = `<div class="renovation-modal renovation-master-modal"><button class="renovation-modal-close" onclick="this.closest('.renovation-modal-overlay').remove()">×</button>` +
    `<h3>機器カテゴリ管理</h3><div class="renovation-master-toolbar"><button onclick="addRenovationCategory()">＋ カテゴリ追加</button>` +
    `<button onclick="exportRenovationEquipmentMaster()">マスタJSON出力</button><button onclick="importRenovationEquipmentMaster()">マスタJSON読込</button></div>` +
    `<table><thead><tr><th>順番</th><th>カテゴリ名</th><th>項目数</th><th>操作</th></tr></thead><tbody>` +
    categories.map((category, index) => `<tr><td><button onclick="moveRenovationCategory('${category.id}',-1)"${index === 0 ? " disabled" : ""}>↑</button>` +
      `<button onclick="moveRenovationCategory('${category.id}',1)"${index === categories.length - 1 ? " disabled" : ""}>↓</button></td>` +
      `<td>${esc(category.label)}</td><td>${items.filter(item => renovationEquipmentDisplayCategoryId(item) === category.id).length}</td>` +
      `<td><button onclick="editRenovationCategory('${category.id}')">編集</button> <button class="renovation-master-delete" onclick="deleteRenovationCategory('${category.id}')">削除</button></td></tr>`).join("") +
    `</tbody></table><p>JSON出力したファイルを共有いただければ、その内容をアプリの初期データへ反映できます。</p></div>`;
  document.body.appendChild(overlay);
}

function addRenovationCategory() {
  const label = prompt("追加するカテゴリ名", "新しいカテゴリ");
  if (label == null || !label.trim()) return;
  const master = loadRenovationEquipmentMaster();
  const category = { id: renovationMasterSafeId("cat-custom"), label: label.trim(), sourceLabel: "ユーザー追加" };
  master.categories.push(category);
  saveRenovationEquipmentMaster(master);
  renovationCategoryId = category.id;
  showRenovationCategoryManager();
  refreshRenovationMasterView();
}

function editRenovationCategory(id) {
  const master = loadRenovationEquipmentMaster();
  const category = master.categories.find(row => row.id === id);
  if (!category) return;
  const label = prompt("カテゴリ名", category.label || "");
  if (label == null || !label.trim()) return;
  category.label = label.trim();
  master.items.forEach(item => {
    if (renovationEquipmentDisplayCategoryId(item) === id) item.category = category.label;
  });
  saveRenovationEquipmentMaster(master);
  showRenovationCategoryManager();
  refreshRenovationMasterView();
}

function deleteRenovationCategory(id) {
  const master = loadRenovationEquipmentMaster();
  const category = master.categories.find(row => row.id === id);
  if (!category || master.categories.length <= 1) {
    showToast("カテゴリは最低1件必要です");
    return;
  }
  const itemCount = master.items.filter(item => renovationEquipmentDisplayCategoryId(item) === id).length;
  if (!confirm(`「${category.label}」を削除します。${itemCount ? `カテゴリ内の${itemCount}項目も削除されます。` : ""}\n既存見積に追加済みの明細は残ります。`)) return;
  master.categories = master.categories.filter(row => row.id !== id);
  master.items = master.items.filter(item => renovationEquipmentDisplayCategoryId(item) !== id);
  saveRenovationEquipmentMaster(master);
  renovationCategoryId = master.categories[0]?.id || "";
  showRenovationCategoryManager();
  refreshRenovationMasterView();
}

function moveRenovationCategory(id, direction) {
  const master = loadRenovationEquipmentMaster();
  const index = master.categories.findIndex(row => row.id === id);
  const next = index + Number(direction);
  if (index < 0 || next < 0 || next >= master.categories.length) return;
  [master.categories[index], master.categories[next]] = [master.categories[next], master.categories[index]];
  saveRenovationEquipmentMaster(master);
  showRenovationCategoryManager();
  refreshRenovationMasterView();
}

function showRenovationEquipmentEditor(id = "") {
  const categories = renovationEquipmentCategories();
  const current = id ? getRenovationMasterItem(id) : null;
  const selectedCategory = current ? renovationEquipmentDisplayCategoryId(current) : renovationCategoryId;
  const overlay = document.createElement("div");
  overlay.className = "renovation-modal-overlay";
  overlay.onclick = event => { if (event.target === overlay) overlay.remove(); };
  overlay.innerHTML = `<div class="renovation-modal renovation-master-modal"><button class="renovation-modal-close" type="button" onclick="this.closest('.renovation-modal-overlay').remove()">×</button>` +
    `<h3>${current ? "機器項目を編集" : "機器項目を追加"}</h3><form class="renovation-master-form" data-item-id="${escAttr(id)}" onsubmit="saveRenovationEquipmentEditor(event)">` +
    `<label>カテゴリ<select name="categoryId" required>${categories.map(category => `<option value="${escAttr(category.id)}"${category.id === selectedCategory ? " selected" : ""}>${esc(category.label)}</option>`).join("")}</select></label>` +
    `<label>項目名<input name="name" required value="${escAttr(current?.name || "")}"></label>` +
    `<label>メーカー<input name="manufacturer" value="${escAttr(renovationEquipmentManufacturer(current) || "")}"></label>` +
    `<label>型式・仕様<input name="spec" value="${escAttr(current?.spec || "")}"></label>` +
    `<label class="wide">備考<textarea name="note">${esc(current?.note || "")}</textarea></label>` +
    `<label>定価<input name="listPrice" inputmode="numeric" value="${current?.listPrice == null ? "" : escAttr(String(current.listPrice))}" placeholder="未登録は空欄"></label>` +
    `<label>NET<input name="netPrice" inputmode="numeric" value="${current?.netPrice == null ? "" : escAttr(String(current.netPrice))}" placeholder="未登録は空欄"></label>` +
    `<label>仕入<input name="costPrice" inputmode="numeric" value="${current?.costPrice == null ? "" : escAttr(String(current.costPrice))}" placeholder="未登録は空欄"></label>` +
    `<label>追加時数量<input name="defaultQty" type="number" min="1" step="1" value="${escAttr(String(current?.defaultQty || 1))}"></label>` +
    `<div class="renovation-master-form-actions"><button type="button" onclick="this.closest('.renovation-modal-overlay').remove()">キャンセル</button><button type="submit">保存</button></div>` +
    `</form></div>`;
  document.body.appendChild(overlay);
}

function saveRenovationEquipmentEditor(event) {
  event.preventDefault();
  const form = event.currentTarget;
  const data = new FormData(form);
  const master = loadRenovationEquipmentMaster();
  const id = form.dataset.itemId || renovationMasterSafeId("ren-custom");
  const index = master.items.findIndex(item => item.id === id);
  const previous = index >= 0 ? getRenovationMasterItem(id) : null;
  const categoryId = String(data.get("categoryId") || "");
  const category = master.categories.find(row => row.id === categoryId);
  if (!category) return;
  const next = {
    ...(index >= 0 ? master.items[index] : {}),
    id, categoryId, category: category.label,
    name: String(data.get("name") || "").trim(),
    manufacturer: String(data.get("manufacturer") || "").trim(),
    spec: String(data.get("spec") || "").trim(),
    note: String(data.get("note") || "").trim(),
    defaultQty: Math.max(1, Number(data.get("defaultQty")) || 1),
    listPrice: nullableRenovationPrice(data.get("listPrice")),
    netPrice: nullableRenovationPrice(data.get("netPrice")),
    costPrice: nullableRenovationPrice(data.get("costPrice")),
    priceLabel: index >= 0 ? "アプリで編集" : "ユーザー追加",
    effectiveDate: renovationToday(),
  };
  if (previous) {
    next.history = [{
      label: previous.priceLabel || "編集前", effectiveDate: previous.effectiveDate || null,
      changedAt: new Date().toISOString(), listPrice: previous.listPrice, netPrice: previous.netPrice,
      costPrice: previous.costPrice, note: previous.note || "",
    }, ...(previous.history || [])];
  } else next.history = [];
  if (index >= 0) master.items[index] = next;
  else master.items.push(next);
  const overrides = loadRenovationPriceOverrides();
  delete overrides[id];
  saveRenovationPriceOverrides(overrides);
  saveRenovationEquipmentMaster(master);
  renovationCategoryId = categoryId;
  form.closest(".renovation-modal-overlay")?.remove();
  refreshRenovationMasterView();
  showToast(index >= 0 ? "機器項目を更新しました" : "機器項目を追加しました");
}

function deleteRenovationEquipmentMasterItem(id) {
  const item = getRenovationMasterItem(id);
  if (!item || !confirm(`「${item.name}」を機器マスタから削除します。\n既存見積に追加済みの明細は残ります。`)) return;
  const master = loadRenovationEquipmentMaster();
  master.items = master.items.filter(row => row.id !== id);
  saveRenovationEquipmentMaster(master);
  const overrides = loadRenovationPriceOverrides();
  delete overrides[id];
  saveRenovationPriceOverrides(overrides);
  renderRenovationCatalog();
}

function moveRenovationEquipmentMasterItem(id, direction) {
  const master = loadRenovationEquipmentMaster();
  const current = master.items.find(item => item.id === id);
  if (!current) return;
  const categoryId = renovationEquipmentDisplayCategoryId(current);
  const siblings = master.items.filter(item => renovationEquipmentDisplayCategoryId(item) === categoryId);
  const siblingIndex = siblings.findIndex(item => item.id === id);
  const target = siblings[siblingIndex + Number(direction)];
  if (!target) return;
  const index = master.items.findIndex(item => item.id === id);
  const targetIndex = master.items.findIndex(item => item.id === target.id);
  [master.items[index], master.items[targetIndex]] = [master.items[targetIndex], master.items[index]];
  saveRenovationEquipmentMaster(master);
  renderRenovationCatalog();
}

function exportRenovationEquipmentMaster() {
  const source = loadRenovationEquipmentMaster();
  const master = cloneRenovationMasterRows(source);
  master.items = source.items.map(item => cloneRenovationMasterRows(getRenovationMasterItem(item.id) || item));
  master.exportedAt = new Date().toISOString();
  const blob = new Blob([JSON.stringify(master, null, 2)], { type: "application/json" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `改修機器マスタ_${renovationToday()}.json`;
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}

function importRenovationEquipmentMaster() {
  const input = document.createElement("input");
  input.type = "file";
  input.accept = ".json,application/json";
  input.onchange = () => {
    const file = input.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const master = JSON.parse(reader.result);
        if (!Array.isArray(master.categories) || !master.categories.length || !Array.isArray(master.items)) throw new Error("形式不正");
        const ids = new Set();
        master.categories.forEach(category => {
          if (!/^[A-Za-z0-9_-]+$/.test(category.id || "") || !String(category.label || "").trim() || ids.has(category.id)) throw new Error("カテゴリ不正");
          ids.add(category.id);
        });
        const itemIds = new Set();
        master.items.forEach(item => {
          if (!/^[A-Za-z0-9_-]+$/.test(item.id || "") || !String(item.name || "").trim() || itemIds.has(item.id)) throw new Error("項目不正");
          if (!ids.has(renovationEquipmentDisplayCategoryId(item))) throw new Error("カテゴリ参照不正");
          itemIds.add(item.id);
        });
        if (!confirm(`カテゴリ${master.categories.length}件、機器${master.items.length}件を読み込み、現在の機器マスタを置き換えますか？`)) return;
        saveRenovationEquipmentMaster({ version: 1, categories: master.categories, items: master.items });
        saveRenovationPriceOverrides({});
        renovationCategoryId = master.categories[0].id;
        document.getElementById("renovation-category-manager")?.remove();
        refreshRenovationMasterView();
        showToast("改修機器マスタを読み込みました");
      } catch {
        alert("改修機器マスタとして読み込めないJSONです。");
      }
    };
    reader.readAsText(file);
  };
  input.click();
}

function updateRenovationCategoryHighlight() {
  document.querySelectorAll(".renovation-categories button[data-category-id]").forEach(button => {
    const isActive = !renovationSearchText && button.dataset.categoryId === renovationCategoryId;
    button.classList.toggle("active", isActive);
    button.setAttribute("aria-pressed", String(isActive));
  });
}
function setRenovationMaterialGroup(group) {
  if (!RENOVATION_MATERIAL_GROUPS.includes(group)) return;
  renovationMaterialGroup = group;
  renderRenovationTab();
}

function editRenovationMaterialPrice(id) {
  const item = getRenovationMaterialItem(id);
  if (!item) return;
  const list = prompt("新しい定価（空欄は材料定価係数で計算）", item.listPrice ?? "");
  if (list === null) return;
  const net = prompt("新しいNET（空欄は仕入×材料NET係数で計算）", item.netPrice ?? "");
  if (net === null) return;
  const cost = prompt("新しい仕入（空欄は未登録）", item.costPrice ?? item.unitPrice ?? "");
  if (cost === null) return;
  const note = prompt("価格メモ・備考", item.note ?? item.priceLabel ?? "");
  if (note === null) return;
  const overrides = loadRenovationMaterialPriceOverrides();
  const previous = {
    effectiveDate: item.effectiveDate || null,
    changedAt: new Date().toISOString(),
    listPrice: item.listPrice ?? null,
    netPrice: item.netPrice ?? null,
    costPrice: item.costPrice ?? item.unitPrice ?? null,
    note: item.note || "",
    label: item.priceLabel || "変更前",
  };
  overrides[id] = {
    current: {
      listPrice: nullableRenovationPrice(list),
      netPrice: nullableRenovationPrice(net),
      costPrice: nullableRenovationPrice(cost),
      unitPrice: nullableRenovationPrice(cost),
      note: note.trim(),
      priceLabel: "アプリで変更",
      effectiveDate: renovationToday(),
    },
    history: [previous, ...(overrides[id]?.history || [])],
  };
  saveRenovationMaterialPriceOverrides(overrides);
  renderRenovationTab();
  showToast("材料の3価格・備考・変更履歴を保存しました（選択済み材料の価格は維持されます）");
}

function showRenovationMaterialPriceHistory(id) {
  const item = getRenovationMaterialItem(id);
  if (!item) return;
  const rows = [
    { effectiveDate: item.effectiveDate, listPrice: item.listPrice, netPrice: item.netPrice, costPrice: item.costPrice ?? item.unitPrice, note: item.note || "", label: item.priceLabel || "現行" },
    ...(item.history || []),
  ];
  const format = value => value == null ? "未登録" : fmtNum(value);
  const overlay = document.createElement("div");
  overlay.className = "renovation-modal-overlay";
  overlay.onclick = event => { if (event.target === overlay) overlay.remove(); };
  overlay.innerHTML = `<div class="renovation-modal"><button class="renovation-modal-close" onclick="this.closest('.renovation-modal-overlay').remove()">×</button>` +
    `<h3>${esc(item.name)} ${esc(item.spec)} の価格履歴</h3><table><thead><tr><th>時点</th><th>定価</th><th>NET</th><th>仕入</th><th>備考</th></tr></thead><tbody>` +
    rows.map(row => `<tr><td>${esc(row.effectiveDate || row.label || "時期不明")}</td><td>${format(row.listPrice)}</td><td>${format(row.netPrice)}</td><td>${format(row.costPrice ?? row.unitPrice)}</td><td>${esc(row.note || "")}</td></tr>`).join("") +
    `</tbody></table><p>登録価格の変更は、変更後に追加する材料へ反映されます。選択済み材料は見積ごとの単価を維持します。</p></div>`;
  document.body.appendChild(overlay);
}

function editRenovationMasterPrice(id) {
  const item = getRenovationMasterItem(id);
  if (!item) return;
  const list = prompt("新しい定価（空欄は未登録）", item.listPrice ?? "");
  if (list === null) return;
  const net = prompt("新しいNET（空欄は未登録）", item.netPrice ?? "");
  if (net === null) return;
  const cost = prompt("新しい仕入（空欄は未登録）", item.costPrice ?? "");
  if (cost === null) return;
  const note = prompt("価格メモ・備考", item.note ?? item.priceLabel ?? "");
  if (note === null) return;
  const overrides = loadRenovationPriceOverrides();
  const previous = {
    label: item.priceLabel || "変更前", effectiveDate: item.effectiveDate || null,
    changedAt: new Date().toISOString(), listPrice: item.listPrice, netPrice: item.netPrice, costPrice: item.costPrice, note: item.note || "",
  };
  const oldHistory = overrides[id]?.history || [];
  overrides[id] = {
    current: {
      listPrice: list.trim() ? Math.max(0, Number(list.replace(/,/g, "")) || 0) : null,
      netPrice: net.trim() ? Math.max(0, Number(net.replace(/,/g, "")) || 0) : null,
      costPrice: cost.trim() ? Math.max(0, Number(cost.replace(/,/g, "")) || 0) : null,
      note: note.trim(),
      priceLabel: "アプリで変更", effectiveDate: renovationToday(),
    },
    history: [previous, ...oldHistory],
  };
  saveRenovationPriceOverrides(overrides);
  renderRenovationCatalog();
  showToast("価格と変更履歴を保存しました（既存見積の単価は維持されます）");
}

function showRenovationPriceHistory(id) {
  const item = getRenovationMasterItem(id);
  if (!item) return;
  const format = value => value == null ? "未登録" : fmtNum(value);
  const rows = [
    { label: item.priceLabel || "現行", effectiveDate: item.effectiveDate, listPrice: item.listPrice, netPrice: item.netPrice, costPrice: item.costPrice, note: item.note || "" },
    ...(item.history || []),
  ];
  const overlay = document.createElement("div");
  overlay.className = "renovation-modal-overlay";
  overlay.onclick = event => { if (event.target === overlay) overlay.remove(); };
  overlay.innerHTML = `<div class="renovation-modal"><button class="renovation-modal-close" onclick="this.closest('.renovation-modal-overlay').remove()">×</button>` +
    `<h3>${esc(item.name)} の価格履歴</h3><table><thead><tr><th>時点</th><th>定価</th><th>NET</th><th>仕入</th><th>備考</th></tr></thead><tbody>` +
    rows.map(row => `<tr><td>${esc(row.effectiveDate || row.label || "時期不明")}</td><td>${format(row.listPrice)}</td><td>${format(row.netPrice)}</td><td>${format(row.costPrice)}</td><td>${esc(row.note || "")}</td></tr>`).join("") +
    `</tbody></table><p>時期不明の履歴は、元Excelに日付の記載がない価格です。</p></div>`;
  document.body.appendChild(overlay);
}

function renovationUnitSwitcherHtml() {
  return `<div class="renovation-unit-switcher">${currentEstimate.units.map((unit, index) =>
    `<button class="${index === activeUnitIndex ? "active" : ""}" onclick="activateUnit(${index});renderRenovationTab()">${esc(unit.unitName)}</button>`
  ).join("")}<button onclick="addRenovationUnit()">＋ 改修項目</button>` +
    `<button class="renovation-confirm-button" onclick="switchTab('estimate')">見積作成で確認</button></div>`;
}

function addRenovationUnit() {
  const number = currentEstimate.units.filter(unit => unit.renovation?.enabled).length + 1;
  currentEstimate.units.push({
    unitId: genId(), unitName: `改修工事${number}`, lines: [],
    listRate: DEFAULT_RATES.listRate, netRate: DEFAULT_RATES.netRate, renovation: createRenovationState(),
  });
  activeUnitIndex = currentEstimate.units.length - 1;
  syncRenovationQuoteLine();
  renderRenovationTab();
}

function renderRenovationCatalog() {
  const container = document.getElementById("renovation-catalog-list");
  if (!container) return;
  const masterItems = renovationEquipmentItems();
  const selectedQuantities = new Map();
  (getRenovationState()?.equipment || []).forEach(row => {
    selectedQuantities.set(row.itemId, (selectedQuantities.get(row.itemId) || 0) + (Number(row.qty) || 0));
  });
  let items = masterItems.filter(item => renovationEquipmentDisplayCategoryId(item) === renovationCategoryId);
  if (renovationSearchText) {
    items = masterItems.filter(item => `${renovationEquipmentManufacturer(item)} ${item.name} ${item.spec} ${item.note || ""} ${item.category}`.toLowerCase().includes(renovationSearchText));
  }
  const category = renovationEquipmentCategories().find(row => row.id === renovationCategoryId);
  const categoryNotes = renovationSearchText ? [] : (RENOVATION_REFERENCE_NOTES?.categories?.[category?.label] || []);
  container.innerHTML = items.map((source, index) => {
    const item = getRenovationMasterItem(source.id);
    const manufacturer = renovationEquipmentManufacturer(item);
    const unavailable = item.listPrice == null && item.netPrice == null && item.costPrice == null;
    const selectedQty = selectedQuantities.get(item.id) || 0;
    return `<div class="renovation-catalog-item${unavailable ? " price-missing" : ""}${selectedQty ? " selected" : ""}">` +
      `<button class="renovation-catalog-select" aria-pressed="${selectedQty > 0}" onclick="addRenovationEquipment('${item.id}')" title="この機器を追加"><strong>＋ ${manufacturer ? `${esc(manufacturer)}｜` : ""}${esc(item.name)}</strong>${selectedQty ? `<em class="renovation-selected-badge">選択済 ×${selectedQty}</em>` : ""}${item.spec ? `<small>${esc(item.spec)}</small>` : ""}` +
      `<span>定価 ${item.listPrice == null ? "未登録" : fmtNum(item.listPrice)} ／ NET ${item.netPrice == null ? "未登録" : fmtNum(item.netPrice)} ／ 仕入 ${item.costPrice == null ? "未登録" : fmtNum(item.costPrice)}</span>` +
      `${item.note ? `<small class="renovation-item-note">備考：${esc(item.note)}</small>` : ""}</button>` +
      `<div class="renovation-catalog-actions"><button onclick="moveRenovationEquipmentMasterItem('${item.id}',-1)" title="上へ"${renovationSearchText || index === 0 ? " disabled" : ""}>↑</button>` +
      `<button onclick="moveRenovationEquipmentMasterItem('${item.id}',1)" title="下へ"${renovationSearchText || index === items.length - 1 ? " disabled" : ""}>↓</button>` +
      `<button onclick="showRenovationPriceHistory('${item.id}')">履歴</button>` +
      `<button onclick="showRenovationEquipmentEditor('${item.id}')">編集</button>` +
      `<button class="renovation-master-delete" onclick="deleteRenovationEquipmentMasterItem('${item.id}')">削除</button></div></div>`;
  }).join("") || `<div class="renovation-empty">該当する機器はありません。</div>`;
  if (categoryNotes.length) container.insertAdjacentHTML("beforeend", renovationNotesHtml(categoryNotes, `${category.label}の元表メモ`));
}

function renovationEstimatePreviewHtml(unit, pricing) {
  let itemNo = 0;
  const rows = unit.lines.map(line => {
    if (line.type === "sep") return `<tr class="sep-row"><td colspan="6"></td></tr>`;
    if (line.type === "comment") return `<tr class="comment-row"><td colspan="6">${esc(line.text || "")}</td></tr>`;
    if (line.type === "subtotal") return "";
    itemNo++;
    let name = line.name || "(カスタム)";
    let spec = line.spec || "";
    if (line.type === "item") {
      const display = getEstimateLineDisplay(line);
      name = display.name;
      spec = display.spec;
    }
    return `<tr><td>${itemNo}</td><td>${esc(name)}</td><td>${esc(spec)}</td>` +
      `<td>${fmtNum(line.unitPrice)}</td><td>${fmtNum(line.qty)}</td><td>${fmtNum((Number(line.unitPrice) || 0) * (Number(line.qty) || 0))}</td></tr>`;
  }).filter(Boolean).join("");
  const sourceNotes = renovationSourceNotesForUnit(unit);
  return `<div class="renovation-estimate-preview">` +
    `<div class="renovation-estimate-preview-head"><strong>見積作成（自動連動）</strong><button onclick="switchTab('estimate')">見積作成を開く</button></div>` +
    `<div class="renovation-estimate-preview-table"><table><thead><tr><th>No</th><th>品名</th><th>仕様</th><th>単価</th><th>数量</th><th>金額</th></tr></thead>` +
    `<tbody>${rows || `<tr><td colspan="6" class="renovation-empty">明細がありません。</td></tr>`}</tbody></table></div>` +
    `<div class="renovation-estimate-preview-totals"><span>定価 <b>${fmtNum(pricing.list)}</b></span><span>NET <b>${fmtNum(pricing.net)}</b></span></div>` +
    (sourceNotes.length ? `<div class="renovation-estimate-preview-notes"><strong>備考・特記事項へ自動記載</strong>${sourceNotes.map(note => `<span>・${esc(note)}</span>`).join("")}</div>` : "") +
    `</div>`;
}

function renderRenovationTab() {
  const root = document.getElementById("renovation-root");
  if (!root) return;
  const unit = activeUnit();
  const state = getRenovationState(unit);
  if (!state) {
    startRenovationUnit();
    return;
  }
  syncRenovationQuoteLine(unit);
  syncRenovationSourceNotes();
  const pricing = calcRenovationPricing(unit);
  const equipmentRows = state.equipment.map(row => `<tr><td><strong>${row.manufacturer ? `${esc(row.manufacturer)}｜` : ""}${esc(row.name)}</strong><small>${esc(row.spec || row.category)}</small></td>` +
    `<td><input type="number" min="0" step="1" value="${row.qty}" onchange="updateRenovationEquipment('${row.lineId}','qty',this.value)"></td>` +
    `<td><input type="number" min="0" step="1" value="${row.listPrice}" onchange="updateRenovationEquipment('${row.lineId}','listPrice',this.value)"></td>` +
    `<td><input type="number" min="0" step="1" value="${row.netPrice}" onchange="updateRenovationEquipment('${row.lineId}','netPrice',this.value)"></td>` +
    `<td><input type="number" min="0" step="1" value="${row.costPrice}" onchange="updateRenovationEquipment('${row.lineId}','costPrice',this.value)"></td>` +
    `<td><input class="renovation-note-input" type="text" value="${escAttr(row.note || "")}" onchange="updateRenovationEquipment('${row.lineId}','note',this.value)"></td>` +
    `<td>${fmtNum(row.qty * row.costPrice)}</td>` +
    `<td><button class="renovation-delete" onclick="removeRenovationEquipment('${row.lineId}')">×</button></td></tr>`).join("");
  const materialRows = state.materials.map(row => {
    const capacity = renovationMaterialCapacity(row);
    return `<tr><td>${esc(row.name)}<small>${esc(row.spec)}${row.unit ? ` ／ ${esc(row.unit)}` : ""}</small>${capacity ? `<small class="renovation-capacity">${esc(capacity.detail)}</small>` : ""}${row.note ? `<small class="renovation-item-note">備考：${esc(row.note)}</small>` : ""}</td>` +
    `<td><input type="number" min="0" step="0.1" value="${row.qty}" onchange="updateRenovationMaterial('${row.lineId}','qty',this.value)"></td>` +
    `<td><input type="number" min="0" step="1" value="${row.listPrice ?? ""}" placeholder="係数" onchange="updateRenovationMaterial('${row.lineId}','listPrice',this.value)"></td>` +
    `<td><input type="number" min="0" step="1" value="${row.netPrice ?? ""}" placeholder="係数" onchange="updateRenovationMaterial('${row.lineId}','netPrice',this.value)"></td>` +
    `<td><input type="number" min="0" step="1" value="${row.costPrice ?? row.unitPrice ?? 0}" onchange="updateRenovationMaterial('${row.lineId}','costPrice',this.value)"></td>` +
    `<td><input class="renovation-note-input" type="text" value="${escAttr(row.note || "")}" onchange="updateRenovationMaterial('${row.lineId}','note',this.value)"></td>` +
    `<td>${fmtNum(row.qty * (row.costPrice ?? row.unitPrice ?? 0))}</td><td><button class="renovation-delete" onclick="removeRenovationMaterial('${row.lineId}')">×</button></td></tr>`;
  }).join("");
  const isDistribution = state.presentationMethod === "allocation" || state.presentationMethod === "matsuoka";
  const candidateCards = Object.entries(RENOVATION_QUOTE_METHODS).map(([key, method]) => {
    const result = pricing.candidates[key];
    const locked = (state.presentationMethod === "detail" && key !== "bara") ||
      (state.presentationMethod === "matsuoka" && key !== "govt") ||
      ((state.presentationMethod === "lump" || state.presentationMethod === "allocation") && key === "bara");
    return `<button class="renovation-price-card${state.quoteMethod === key ? " active" : ""}" ${locked ? "disabled title=\"この表示方法では選択できません\"" : `onclick="setRenovationQuoteMethod('${key}')"`}>` +
      `<span class="renovation-price-card-title">${esc(method.label)}${state.quoteMethod === key ? "（見積採用中）" : ""}</span>` +
      `<span><small>方式定価</small><strong>${fmtNum(result.list)}</strong></span>` +
      `<span><small>NET</small><strong>${fmtNum(result.net)}</strong></span>` +
      `<span class="renovation-rate-row"><small>NET／定価 ${result.netToListRate.toFixed(1)}%</small><small>機器原価率 ${result.equipmentCostRate.toFixed(1)}%</small><small>総原価率 ${result.costRate.toFixed(1)}%</small><small>粗利率 ${result.grossMargin.toFixed(1)}%</small></span>` +
    `</button>`;
  }).join("");
  const overheadNet = Math.max(0, pricing.net - pricing.directNet);
  const overheadLabel = isDistribution ? "振分対象差額" : "諸経費（定価差額）";
  const visibleMaterials = RENOVATION_MATERIALS
    .filter(item => renovationMaterialDisplayGroup(item) === renovationMaterialGroup)
    .map(item => getRenovationMaterialItem(item.id) || item);
  const selectedMaterialQuantities = new Map();
  state.materials.forEach(row => {
    selectedMaterialQuantities.set(row.materialId, (selectedMaterialQuantities.get(row.materialId) || 0) + (Number(row.qty) || 0));
  });
  const materialNotes = RENOVATION_REFERENCE_NOTES?.materialGroups?.[renovationMaterialGroup] || [];
  const equipmentCategories = renovationEquipmentCategories();
  if (!equipmentCategories.some(category => category.id === renovationCategoryId)) renovationCategoryId = equipmentCategories[0]?.id || "";
  root.innerHTML = `<div class="renovation-sticky-controls">` + renovationUnitSwitcherHtml() + `</div>` +
    `<div class="renovation-layout"><section class="renovation-catalog"><div class="renovation-section-title">1. 機器を選択</div>` +
      `<input class="renovation-search" type="search" placeholder="全価格表から検索" value="${escAttr(renovationSearchText)}" oninput="setRenovationSearch(this.value)">` +
      `<div class="renovation-categories">${equipmentCategories.map(cat => `<button data-category-id="${escAttr(cat.id)}" class="${cat.id === renovationCategoryId && !renovationSearchText ? "active" : ""}" aria-pressed="${cat.id === renovationCategoryId && !renovationSearchText}" onclick="setRenovationCategory('${cat.id}')">${esc(cat.label)}</button>`).join("")}</div>` +
      `<div class="renovation-master-toolbar"><button onclick="showRenovationEquipmentEditor()">＋ 機器項目</button><button onclick="showRenovationCategoryManager()">カテゴリ管理</button>` +
      `<button onclick="exportRenovationEquipmentMaster()" title="他のPCへの移行や初期データ化の共有に使用">マスタ出力</button><button onclick="importRenovationEquipmentMaster()">マスタ読込</button></div>` +
      `<div id="renovation-catalog-list" class="renovation-catalog-list"></div></section>` +
    `<section class="renovation-catalog renovation-material-catalog"><div class="renovation-section-title">2. 材料を選択</div>` +
      `<div class="renovation-material-catalog-scroll">` +
      `<div class="renovation-material-groups">${RENOVATION_MATERIAL_GROUPS.map(group => `<button class="${group === renovationMaterialGroup ? "active" : ""}" onclick="setRenovationMaterialGroup('${escAttr(group)}')">${esc(group)}</button>`).join("")}</div>` +
      `<div class="renovation-material-buttons">${visibleMaterials.map(item => {
        const capacity = renovationMaterialCapacity(item);
        const materialCost = renovationMaterialPrice(item, "costPrice");
        const materialNet = renovationMaterialPrice(item, "netPrice");
        const selectedQty = selectedMaterialQuantities.get(item.id) || 0;
        return `<div class="renovation-material-card${selectedQty ? " selected" : ""}"><button class="renovation-material-add ${materialCost == null ? "price-missing" : ""}" aria-pressed="${selectedQty > 0}" onclick="addRenovationMaterial('${item.id}')">` +
          `<span>＋ ${esc(item.name)} ${esc(item.spec)}</span>${selectedQty ? `<em class="renovation-selected-badge">選択済 ×${selectedQty}</em>` : ""}<small>定価 ${item.listPrice == null ? "係数計算" : fmtNum(item.listPrice)} ／ NET ${materialNet == null ? "係数計算" : fmtNum(materialNet)} ／ 仕入 ${materialCost == null ? "未登録" : fmtNum(materialCost)}${capacity ? ` ／ ${esc(capacity.short)}` : ""}</small>` +
          `${item.note ? `<small class="renovation-item-note">備考：${esc(item.note)}</small>` : ""}</button>` +
          `<div class="renovation-catalog-actions renovation-material-card-actions"><button onclick="showRenovationMaterialPriceHistory('${item.id}')" title="価格履歴">履歴</button><button onclick="editRenovationMaterialPrice('${item.id}')" title="登録価格と備考を変更">価格・備考</button></div></div>`;
      }).join("")}</div>` +
      renovationNotesHtml(materialNotes, `${renovationMaterialGroup}の元表メモ`) +
      `</div></section>` +
    `<div class="renovation-work"><section><div class="renovation-section-title">3. 選択内容・費用調整</div>` +
      `<div class="renovation-selection-subtitle">選択した機器</div>` +
      `<div class="renovation-table-wrap"><table><thead><tr><th>機器</th><th>数量</th><th>定価</th><th>NET</th><th>仕入</th><th>備考</th><th>仕入計</th><th></th></tr></thead><tbody>${equipmentRows || `<tr><td colspan="8" class="renovation-empty">左の機器一覧から追加してください。</td></tr>`}</tbody></table></div>` +
      `<div class="renovation-selection-subtitle">選択した材料</div>` +
      `<div class="renovation-table-wrap"><table><thead><tr><th>材料</th><th>数量</th><th>定価</th><th>NET</th><th>仕入</th><th>備考</th><th>仕入計</th><th></th></tr></thead><tbody>${materialRows || `<tr><td colspan="8" class="renovation-empty">材料を追加してください。</td></tr>`}</tbody></table></div></section>` +
    `<section><div class="renovation-section-title">4. 現地調査費</div><div class="renovation-cost-inputs renovation-cost-inputs-small">` +
      `<label class="renovation-important-input">現地調査 定価<input type="number" min="0" value="${state.siteSurveyList}" onchange="updateRenovationNumber('siteSurveyList',this.value)"></label>` +
      `<label class="renovation-important-input">現地調査 NET<input type="number" min="0" value="${state.siteSurveyNet}" onchange="updateRenovationNumber('siteSurveyNet',this.value)"></label>` +
      `<label>現地調査 社内原価<input type="number" min="0" value="${state.siteSurveyCost}" onchange="updateRenovationNumber('siteSurveyCost',this.value)"><small>粗利率の計算に使用</small></label>` +
      `</div></section>` +
    `<section><div class="renovation-section-title">5. 改修費</div><div class="renovation-cost-inputs">` +
      `<label>時間単価<input type="number" min="0" value="${state.labor.hourlyRate}" onchange="updateRenovationNumber('labor.hourlyRate',this.value)"></label>` +
      `<label>工数係数<input type="number" min="0" step="0.05" value="${state.labor.factor}" onchange="updateRenovationNumber('labor.factor',this.value)"></label>` +
      `<label>作業時間<input type="number" min="0" step="0.5" value="${state.labor.hours}" onchange="updateRenovationNumber('labor.hours',this.value)"></label>` +
      `<label class="renovation-important-input">人数<input type="number" min="0" step="1" value="${state.labor.people}" onchange="updateRenovationNumber('labor.people',this.value)"></label>` +
      `<label>改修費 社内原価<input type="number" min="0" value="${state.labor.cost}" onchange="updateRenovationNumber('labor.cost',this.value)"><small>粗利率の計算に使用</small></label>` +
      `</div><div class="renovation-calculated-cost renovation-calculated-cost-three"><span>改修費 定価</span><b>${fmtNum(pricing.renovationList)}</b><span>改修費 NET</span><b>${fmtNum(pricing.labor)}</b><span>社内原価</span><b>${fmtNum(pricing.laborCost)}</b></div></section>` +
    `<section><div class="renovation-section-title">6. 諸経費・計算係数</div><div class="renovation-cost-inputs renovation-factor-inputs">` +
      `<label class="renovation-factor">バラ定価係数<input type="number" min="0" step="0.01" value="${state.listPricingFactor}" onchange="updateRenovationNumber('listPricingFactor',this.value)"><small>基本1.20。諸経費とのバランスに応じて自由に変更</small></label>` +
      `<label class="renovation-factor">NET計算係数<input type="number" value="${RENOVATION_NET_FACTOR}" disabled><small>常に1.20で固定</small></label>` +
      `<label class="renovation-factor">TR別計算係数<input type="number" min="0" step="0.01" value="${state.transformerFactor}" onchange="updateRenovationNumber('transformerFactor',this.value)"><small>「TR別計算」を選んだ時だけ使用</small></label>` +
      `<label>材料 NET係数<input type="number" min="0" step="0.01" value="${state.materialNetFactor}" onchange="updateRenovationNumber('materialNetFactor',this.value)"><small>未登録NETは仕入×この係数</small></label>` +
      `<label>材料 定価係数<input type="number" min="0" step="0.01" value="${state.materialListFactor}" onchange="updateRenovationNumber('materialListFactor',this.value)"></label>` +
      `<label>改修費 定価係数<input type="number" min="0" step="0.01" value="${state.renovationListFactor}" onchange="updateRenovationNumber('renovationListFactor',this.value)"></label>` +
      `</div><div class="renovation-calculated-cost renovation-overhead-box"><span>${overheadLabel}（採用定価－直接費）</span><b>${fmtNum(pricing.overhead)}</b><span>NET差額（採用NET－直接NET）</span><b>${fmtNum(overheadNet)}</b></div></section></div>` +
    `<div class="renovation-side"><div class="renovation-pattern-panel"><div class="renovation-pattern-heading"><strong>見積条件の選択</strong><span>上から順に選びます</span></div>` +
      `<div class="renovation-pattern-subtitle"><b>①</b> TRの扱い</div>` +
      `<div class="renovation-methods">${Object.entries(RENOVATION_METHODS).map(([key, method]) =>
        `<button class="${state.method === key ? "active" : ""}" aria-pressed="${state.method === key}" onclick="setRenovationMethod('${key}')"><strong>${method.label}</strong><span>${method.description}</span></button>`
      ).join("")}</div><div class="renovation-pattern-subtitle"><b>②</b> 定価・NET算出方式</div>` +
      `<div class="renovation-price-cards">${candidateCards}</div>` +
      `<div class="renovation-pattern-subtitle"><b>③</b> 見積明細の出し方</div>` +
      `<div class="renovation-methods">${Object.entries(RENOVATION_PRESENTATIONS).map(([key, method]) =>
        `<button class="${state.presentationMethod === key ? "active" : ""}" aria-pressed="${state.presentationMethod === key}" onclick="setRenovationPresentation('${key}')"><strong>${method.label}</strong><span>${method.description}</span></button>`
      ).join("")}</div></div>` +
    `<aside class="renovation-summary"><div class="renovation-section-title">計算結果・見積作成</div>` +
      renovationEstimatePreviewHtml(unit, pricing) +
    `</aside></div></div>`;
  renderRenovationCatalog();
  const stickyControls = root.querySelector(".renovation-sticky-controls");
  if (typeof observeStickyOffsetElement === "function") observeStickyOffsetElement(stickyControls);
  requestAnimationFrame(() => {
    if (typeof updateStickyOffsets === "function") updateStickyOffsets();
  });
}
