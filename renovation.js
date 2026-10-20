// ============================================================
// 改修見積（添付Excelの考え方を、既存見積の「項目」単位で保持）
// ============================================================

const STORAGE_RENOVATION_PRICES = "ban_renovation_prices_v1";
const STORAGE_RENOVATION_MATERIAL_PRICES = "ban_renovation_material_prices_v1";
const RENOVATION_METHODS = {
  base: { label: "一括係数", description: "TRを含む全体に標準係数を使用" },
  transformer: { label: "TR別係数", description: "TRだけ専用の低い係数を使用" },
};
const RENOVATION_QUOTE_METHODS = {
  bara: { label: "実定価ベース（バラ）" },
  govt: { label: "役所3.8" },
  private: { label: "民間2.28" },
};
const RENOVATION_PRESENTATIONS = {
  lump: { label: "一式", description: "採用した方式定価を1行で表示" },
  detail: { label: "実定価内訳", description: "機器・材料・調査・改修を実定価で表示" },
  allocation: { label: "差額振分", description: "役所3.8／民間2.28の残額を4項目へ振分" },
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

function removeRenovationAutoNotes(value) {
  const text = String(value || "");
  const start = text.indexOf(RENOVATION_AUTO_NOTES_START);
  if (start < 0) return text.trim();
  const end = text.indexOf(RENOVATION_AUTO_NOTES_END, start);
  const after = end < 0 ? text.length : end + RENOVATION_AUTO_NOTES_END.length;
  return `${text.slice(0, start)}\n${text.slice(after)}`.replace(/\n{3,}/g, "\n\n").trim();
}

/** 手入力の備考を残しつつ、選択内容に関係する元表の注意事項だけを自動更新する。 */
function syncRenovationSourceNotes() {
  const manualNotes = removeRenovationAutoNotes(currentEstimate.notes);
  const sourceNotes = new Set();
  currentEstimate.units.forEach(unit => {
    renovationSourceNotesForUnit(unit).forEach(note => sourceNotes.add(note));
  });
  const automaticNotes = sourceNotes.size
    ? `${RENOVATION_AUTO_NOTES_START}\n${[...sourceNotes].map(note => `・${note}`).join("\n")}\n${RENOVATION_AUTO_NOTES_END}`
    : "";
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
  const original = RENOVATION_ITEMS.find(item => item.id === id);
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
    pricingFactor: 1.2,
    transformerFactor: 1.1,
    materialNetFactor: 1.2,
    materialListFactor: 1.2,
    renovationListFactor: 1.5,
    pricingVersion: 2,
    equipment: [],
    materials: [],
    siteSurveyList: 60000,
    siteSurveyNet: 50000,
    siteSurveyCost: 0,
    labor: { hourlyRate: 5000, factor: 1.25, hours: 8, people: 1 },
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
  if (state.materialNetFactor == null) state.materialNetFactor = 1.2;
  if (state.materialListFactor == null) state.materialListFactor = 1.2;
  if (state.renovationListFactor == null) state.renovationListFactor = 1.5;
  if (state.siteSurveyCost == null) state.siteSurveyCost = 0;
  state.matsuokaRatios = { ...RENOVATION_MATSUOKA_RATIOS };
  if ((Number(state.pricingVersion) || 0) < 2) {
    state.materials.forEach(row => {
      const master = getRenovationMaterialItem(row.materialId);
      const cost = Number(row.costPrice ?? row.unitPrice) || 0;
      if (master?.netPrice == null && Number(row.netPrice) === cost) row.netPrice = null;
    });
    state.pricingVersion = 2;
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
  const cost = equipmentCost + materialCost + (Number(state.siteSurveyCost) || 0) + labor;
  const factor = renovationFactor(state.pricingFactor, 1.2);
  const trFactor = renovationFactor(state.transformerFactor, 1.1);
  const splitTr = state.method === "transformer";
  const normalListBase = directList - trList;
  const normalNetBase = directNet - trNet;
  const buildCandidate = (kind) => {
    let normalList, normalNet, trPartList = 0, trPartNet = 0;
    if (kind === "bara") {
      normalList = renovationRoundUp(normalListBase * factor, 10000);
      normalNet = renovationRoundUp(normalNetBase * factor * 1.15, 10000);
      if (splitTr) {
        trPartList = renovationRoundUp(trList * trFactor, 10000);
        trPartNet = renovationRoundUp(trNet * trFactor * 1.15, 10000);
      }
    } else {
      const multiplier = kind === "govt" ? 3.8 : 2.28;
      const netRate = kind === "govt" ? 0.33 : 0.55;
      normalList = renovationRoundUp(normalNetBase * factor * multiplier, 10000);
      normalNet = renovationRoundUp(normalList * netRate, 10000);
      if (splitTr) {
        trPartList = renovationRoundUp(trNet * trFactor * multiplier, 10000);
        trPartNet = renovationRoundUp(trPartList * netRate, 10000);
      }
    }
    // TR別計算を使わない場合はTRも標準側に含まれる。
    if (!splitTr) {
      if (kind === "bara") {
        normalList = renovationRoundUp(directList * factor, 10000);
        normalNet = renovationRoundUp(directNet * factor * 1.15, 10000);
      } else {
        const multiplier = kind === "govt" ? 3.8 : 2.28;
        const netRate = kind === "govt" ? 0.33 : 0.55;
        normalList = renovationRoundUp(directNet * factor * multiplier, 10000);
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
    materialCost, materialNet, materialList, labor, renovationList, directList, directNet,
    trList, trNet, trCost,
    distributionDifference, distributionFeasible, distributionOverrun,
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
    `<td><strong>${esc(item.name || "未名称")}</strong>${item.spec ? `<small>${esc(item.spec)}</small>` : ""}</td>` +
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

  return `<section class="renovation-internal-card">` +
    `<div class="renovation-internal-card-heading"><strong>${esc(unit.unitName || "改修工事")}</strong>` +
    `<span>${esc(method)}・${esc(quote)}・${esc(presentation)}</span></div>` +
    `<table class="renovation-internal-total-table"><thead><tr><th>計算項目</th><th>定価</th><th>NET</th><th>仕入・社内原価</th></tr></thead><tbody>` +
    row("機器", pricing.equipmentList, pricing.equipmentNet, pricing.equipmentCost) +
    row("材料費", pricing.materialList, pricing.materialNet, pricing.materialCost) +
    row("現地調査費", state.siteSurveyList, state.siteSurveyNet, state.siteSurveyCost) +
    row("改修費", pricing.renovationList, pricing.labor, pricing.labor) +
    row("方式差額・諸経費", pricing.list - pricing.directList, pricing.net - pricing.directNet, null) +
    row("採用合計", pricing.list, pricing.net, totalCost, "renovation-internal-total") +
    `</tbody></table>` +
    `<div class="renovation-internal-metrics">` +
    `<span>NET利益額<b>${fmtNum(profit)}</b></span>` +
    `<span>粗利率<b>${percent(profit, pricing.net)}</b></span>` +
    `<span>総原価率<b>${percent(totalCost, pricing.net)}</b></span>` +
    `<span>NET／定価<b>${percent(pricing.net, pricing.list)}</b></span>` +
    `</div>` +
    `<div class="renovation-internal-details">` +
    `<h4>機器内訳 <small>各単価</small></h4>` +
    `<table><thead><tr><th>品名・仕様</th><th>数量</th><th>定価</th><th>NET</th><th>仕入</th></tr></thead><tbody>${equipmentRows}</tbody></table>` +
    `<h4>材料内訳 <small>各単価／空欄登録は係数計算</small></h4>` +
    `<table><thead><tr><th>品名・仕様</th><th>数量</th><th>定価</th><th>NET</th><th>仕入</th></tr></thead><tbody>${materialRows}</tbody></table>` +
    `</div></section>`;
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
    const spec = `${methodLabel}・${quoteLabel}／標準×${state.pricingFactor}` +
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
  unit.lines = [...normalLines, ...generated];
}

function syncAllRenovationQuoteLines() {
  currentEstimate.units.forEach(unit => {
    if (unit?.renovation?.enabled) syncRenovationQuoteLine(unit);
  });
  syncRenovationSourceNotes();
}

function refreshRenovation() {
  const unit = activeUnit();
  syncRenovationQuoteLine(unit);
  syncRenovationSourceNotes();
  renderRenovationTab();
  renderCatalogEstimatePanels();
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
    showToast(`${RENOVATION_PRESENTATIONS[state.presentationMethod].label}は役所3.8方式または民間2.28方式を選択してください`);
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
  let items = RENOVATION_ITEMS.filter(item => renovationEquipmentDisplayCategoryId(item) === renovationCategoryId);
  if (renovationSearchText) {
    items = RENOVATION_ITEMS.filter(item => `${renovationEquipmentManufacturer(item)} ${item.name} ${item.spec} ${item.category}`.toLowerCase().includes(renovationSearchText));
  }
  const category = RENOVATION_CATEGORIES.find(row => row.id === renovationCategoryId);
  const categoryNotes = renovationSearchText ? [] : (RENOVATION_REFERENCE_NOTES?.categories?.[category?.label] || []);
  container.innerHTML = items.map(source => {
    const item = getRenovationMasterItem(source.id);
    const manufacturer = renovationEquipmentManufacturer(item);
    const unavailable = item.listPrice == null && item.netPrice == null && item.costPrice == null;
    return `<div class="renovation-catalog-item${unavailable ? " price-missing" : ""}">` +
      `<button class="renovation-catalog-select" onclick="addRenovationEquipment('${item.id}')" title="この機器を追加"><strong>＋ ${manufacturer ? `${esc(manufacturer)}｜` : ""}${esc(item.name)}</strong>${item.spec ? `<small>${esc(item.spec)}</small>` : ""}` +
      `<span>定価 ${item.listPrice == null ? "未登録" : fmtNum(item.listPrice)} ／ NET ${item.netPrice == null ? "未登録" : fmtNum(item.netPrice)} ／ 仕入 ${item.costPrice == null ? "未登録" : fmtNum(item.costPrice)}</span>` +
      `${item.note ? `<small class="renovation-item-note">備考：${esc(item.note)}</small>` : ""}</button>` +
      `<div class="renovation-catalog-actions"><button onclick="showRenovationPriceHistory('${item.id}')">履歴</button>` +
      `<button onclick="editRenovationMasterPrice('${item.id}')">価格・備考</button></div></div>`;
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
  const selected = pricing.candidates[state.quoteMethod];
  const equipmentRows = state.equipment.map(row => `<tr><td><strong>${esc(row.name)}</strong><small>${esc(row.spec || row.category)}</small></td>` +
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
  const allocation = isDistribution
    ? (pricing.distributionFeasible
      ? `<div class="renovation-allocation"><strong>${state.presentationMethod === "matsuoka" ? "松岡Jr.社内比率" : RENOVATION_QUOTE_METHODS[state.quoteMethod].label} 振分額</strong>` +
        Object.entries(pricing.allocations).map(([key, value]) => `<span>${esc({material:"材料",survey:"現地調査",renovation:"改修",overhead:"諸経費",design:"設計",transport:"運搬",management:"管理",misc:"雑費"}[key] || key)} <b>${fmtNum(value)}</b></span>`).join("") + `</div>`
      : `<div class="renovation-distribution-warning"><strong>振分不可</strong><span>機器実定価が方式定価を <b>${fmtNum(pricing.distributionOverrun)}円</b> 上回っています。</span><small>見積作成には、金額を隠さず実定価内訳（${fmtNum(pricing.directList)}円）を自動反映します。方式か選択機器を見直してください。</small></div>`)
    : "";
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
  const priceComparison = `<div class="renovation-price-comparison">` +
    `<span>機器実定価<b>${fmtNum(pricing.equipmentList)}</b></span>` +
    `<span>実定価合計<b>${fmtNum(pricing.directList)}</b></span>` +
    `<span>選択方式定価<b>${fmtNum(selected.list)}</b></span>` +
    `<span class="adopted">見積作成への反映額<b>${fmtNum(pricing.list)}</b></span></div>`;
  const visibleMaterials = RENOVATION_MATERIALS
    .filter(item => renovationMaterialDisplayGroup(item) === renovationMaterialGroup)
    .map(item => getRenovationMaterialItem(item.id) || item);
  const materialNotes = RENOVATION_REFERENCE_NOTES?.materialGroups?.[renovationMaterialGroup] || [];
  const estimateNotes = [
    ...(RENOVATION_REFERENCE_NOTES?.general || []),
    ...(RENOVATION_REFERENCE_NOTES?.patterns?.[state.method] || []),
    ...(RENOVATION_REFERENCE_NOTES?.patterns?.[state.presentationMethod] || []),
  ];

  root.innerHTML = `<div class="renovation-sticky-controls">` + renovationUnitSwitcherHtml() + `</div>` +
    `<div class="renovation-layout"><section class="renovation-catalog"><div class="renovation-section-title">1. 機器を選択</div>` +
      `<input class="renovation-search" type="search" placeholder="全価格表から検索" value="${escAttr(renovationSearchText)}" oninput="setRenovationSearch(this.value)">` +
      `<div class="renovation-categories">${RENOVATION_CATEGORIES.map(cat => `<button data-category-id="${escAttr(cat.id)}" class="${cat.id === renovationCategoryId && !renovationSearchText ? "active" : ""}" aria-pressed="${cat.id === renovationCategoryId && !renovationSearchText}" onclick="setRenovationCategory('${cat.id}')">${esc(cat.label)}</button>`).join("")}</div>` +
      `<div id="renovation-catalog-list" class="renovation-catalog-list"></div></section>` +
    `<section class="renovation-catalog renovation-material-catalog"><div class="renovation-section-title">2. 材料を選択</div>` +
      `<div class="renovation-material-catalog-scroll">` +
      `<div class="renovation-material-groups">${RENOVATION_MATERIAL_GROUPS.map(group => `<button class="${group === renovationMaterialGroup ? "active" : ""}" onclick="setRenovationMaterialGroup('${escAttr(group)}')">${esc(group)}</button>`).join("")}</div>` +
      `<div class="renovation-material-buttons">${visibleMaterials.map(item => {
        const capacity = renovationMaterialCapacity(item);
        const materialCost = renovationMaterialPrice(item, "costPrice");
        const materialNet = renovationMaterialPrice(item, "netPrice");
        return `<div class="renovation-material-card"><button class="renovation-material-add ${materialCost == null ? "price-missing" : ""}" onclick="addRenovationMaterial('${item.id}')">` +
          `<span>＋ ${esc(item.name)} ${esc(item.spec)}</span><small>定価 ${item.listPrice == null ? "係数計算" : fmtNum(item.listPrice)} ／ NET ${materialNet == null ? "係数計算" : fmtNum(materialNet)} ／ 仕入 ${materialCost == null ? "未登録" : fmtNum(materialCost)}${capacity ? ` ／ ${esc(capacity.short)}` : ""}</small>` +
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
      `<label>現地調査 定価<input type="number" min="0" value="${state.siteSurveyList}" onchange="updateRenovationNumber('siteSurveyList',this.value)"></label>` +
      `<label>現地調査 NET<input type="number" min="0" value="${state.siteSurveyNet}" onchange="updateRenovationNumber('siteSurveyNet',this.value)"></label>` +
      `<label>現地調査 社内原価<input type="number" min="0" value="${state.siteSurveyCost}" onchange="updateRenovationNumber('siteSurveyCost',this.value)"><small>粗利率の計算に使用</small></label>` +
      `</div></section>` +
    `<section><div class="renovation-section-title">5. 改修費</div><div class="renovation-cost-inputs">` +
      `<label>時間単価<input type="number" min="0" value="${state.labor.hourlyRate}" onchange="updateRenovationNumber('labor.hourlyRate',this.value)"></label>` +
      `<label>工数係数<input type="number" min="0" step="0.05" value="${state.labor.factor}" onchange="updateRenovationNumber('labor.factor',this.value)"></label>` +
      `<label>作業時間<input type="number" min="0" step="0.5" value="${state.labor.hours}" onchange="updateRenovationNumber('labor.hours',this.value)"></label>` +
      `<label>人数<input type="number" min="0" step="1" value="${state.labor.people}" onchange="updateRenovationNumber('labor.people',this.value)"></label>` +
      `</div><div class="renovation-calculated-cost"><span>改修費 定価</span><b>${fmtNum(pricing.renovationList)}</b><span>改修費 NET</span><b>${fmtNum(pricing.labor)}</b></div></section>` +
    `<section><div class="renovation-section-title">6. 諸経費・計算係数</div><div class="renovation-cost-inputs renovation-factor-inputs">` +
      `<label class="renovation-factor">標準係数<input type="number" min="0" step="0.01" value="${state.pricingFactor}" onchange="updateRenovationNumber('pricingFactor',this.value)"><small>通常1.20。高額時は1.10、1.05などへ変更</small></label>` +
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
      priceComparison +
      `<div class="renovation-result-breakdown"><table><thead><tr><th>項目</th><th>定価</th><th>NET</th><th>原価</th></tr></thead><tbody>` +
      `<tr><td>機器</td><td>${fmtNum(pricing.equipmentList)}</td><td>${fmtNum(pricing.equipmentNet)}</td><td>${fmtNum(pricing.equipmentCost)}</td></tr>` +
      `<tr><td>材料費</td><td>${fmtNum(pricing.materialList)}</td><td>${fmtNum(pricing.materialNet)}</td><td>${fmtNum(pricing.materialCost)}</td></tr>` +
      `<tr><td>現地調査費</td><td>${fmtNum(state.siteSurveyList)}</td><td>${fmtNum(state.siteSurveyNet)}</td><td>${fmtNum(state.siteSurveyCost)}</td></tr>` +
      `<tr><td>改修費</td><td>${fmtNum(pricing.renovationList)}</td><td>${fmtNum(pricing.labor)}</td><td>${fmtNum(pricing.labor)}</td></tr>` +
      `<tr class="overhead"><td>${overheadLabel}</td><td>${fmtNum(pricing.overhead)}</td><td>${fmtNum(overheadNet)}</td><td>—</td></tr>` +
      `</tbody></table></div>${allocation}` +
      renovationNotesHtml(estimateNotes) +
      `<p>機器原価率は元Excelと同じく機器仕入＋材料原価で算出。総原価率・粗利率は現地調査の社内原価と改修費も含めて算出しています。</p>` +
    `</aside></div></div>`;
  renderRenovationCatalog();
  const stickyControls = root.querySelector(".renovation-sticky-controls");
  if (typeof observeStickyOffsetElement === "function") observeStickyOffsetElement(stickyControls);
  requestAnimationFrame(() => {
    if (typeof updateStickyOffsets === "function") updateStickyOffsets();
  });
}
