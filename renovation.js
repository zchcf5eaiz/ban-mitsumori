// ============================================================
// 改修見積（添付Excelの考え方を、既存見積の「項目」単位で保持）
// ============================================================

const STORAGE_RENOVATION_PRICES = "ban_renovation_prices_v1";
const STORAGE_RENOVATION_MATERIAL_PRICES = "ban_renovation_material_prices_v1";
const RENOVATION_METHODS = {
  base: { label: "標準", description: "TRを含め、すべて同じ係数で計算" },
  transformer: { label: "TR別計算", description: "TR部分だけ低い係数で計算" },
  allocation: { label: "3.8振分", description: "役所3.8定価の差額を4項目へ振分" },
  matsuoka: { label: "松岡Jr.", description: "市役所指定の見積方式" },
};
const RENOVATION_QUOTE_METHODS = {
  bara: { label: "バラ方式" },
  govt: { label: "役所3.8方式" },
  private: { label: "民間2.28方式" },
};

let renovationCategoryId = RENOVATION_CATEGORIES[0]?.id || "";
let renovationSearchText = "";
let renovationMaterialGroup = RENOVATION_MATERIAL_GROUPS?.[0] || "電線(KIP)";

function renovationNotesHtml(notes, title = "元表の注意事項") {
  const rows = (notes || []).filter(Boolean);
  if (!rows.length) return "";
  return `<div class="renovation-source-notes"><strong>${esc(title)}</strong><ul>${rows.map(note => `<li>${esc(note)}</li>`).join("")}</ul></div>`;
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
  return item?.unitPrice == null ? null : Number(item.unitPrice) || 0;
}

function createRenovationState() {
  return {
    enabled: true,
    method: "base",
    quoteMethod: "bara",
    pricingFactor: 1.2,
    transformerFactor: 1.1,
    materialListFactor: 1.2,
    renovationListFactor: 1.5,
    equipment: [],
    materials: [],
    siteSurveyList: 60000,
    siteSurveyNet: 50000,
    labor: { hourlyRate: 5000, factor: 1.25, hours: 8, people: 1 },
    matsuokaRatios: {
      material: 0.10, survey: 0.05, renovation: 0.58, overhead: 0.21,
      design: 0.02, transport: 0.02, management: 0.01, misc: 0.01,
    },
  };
}

function getRenovationState(unit = activeUnit()) {
  const state = unit?.renovation?.enabled ? unit.renovation : null;
  if (!state) return null;
  if (!RENOVATION_QUOTE_METHODS[state.quoteMethod]) state.quoteMethod = "bara";
  if (state.transformerFactor == null) state.transformerFactor = 1.1;
  if (state.materialListFactor == null) state.materialListFactor = 1.2;
  if (state.renovationListFactor == null) state.renovationListFactor = 1.5;
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
  const materialCost = state.materials.reduce((total, row) => total + (Number(row.qty) || 0) * (Number(row.costPrice ?? row.unitPrice) || 0), 0);
  const materialNet = state.materials.reduce((total, row) => total + (Number(row.qty) || 0) * (Number(row.netPrice ?? row.costPrice ?? row.unitPrice) || 0), 0);
  const labor = renovationRoundUp(
    (Number(state.labor.hourlyRate) || 0) * (Number(state.labor.factor) || 0) *
    (Number(state.labor.hours) || 0) * (Number(state.labor.people) || 0), 1000
  );
  const materialListBase = state.materials.reduce((total, row) => {
    const qty = Number(row.qty) || 0;
    const explicitList = row.listPrice;
    const unit = explicitList == null
      ? (Number(row.costPrice ?? row.unitPrice) || 0) * (Number(state.materialListFactor) || 1.2)
      : Number(explicitList) || 0;
    return total + qty * unit;
  }, 0);
  const materialList = renovationRoundUp(materialListBase, 1000);
  const renovationList = renovationRoundUp(labor * (Number(state.renovationListFactor) || 1.5), 1000);
  const directList = equipmentList + materialList + (Number(state.siteSurveyList) || 0) + renovationList;
  const directNet = equipmentNet + materialNet + (Number(state.siteSurveyNet) || 0) + labor;
  const cost = equipmentCost + materialCost + (Number(state.siteSurveyNet) || 0) + labor;
  const factor = Math.max(0, Number(state.pricingFactor) || 1.2);
  const trFactor = Math.max(0, Number(state.transformerFactor) || 1.1);
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
  const list = selected.list;
  const net = selected.net;
  const overhead = selected.overhead;

  const allocations = {};
  if (state.method === "allocation") {
    const difference = Math.max(0, candidates.govt.list - equipmentList);
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
  } else if (state.method === "matsuoka") {
    const entries = Object.entries(state.matsuokaRatios || {});
    let allocated = 0;
    entries.forEach(([key, ratio], index) => {
      const difference = Math.max(0, candidates.govt.list - equipmentList);
      const remaining = Math.max(0, difference - allocated);
      const value = index === entries.length - 1
        ? remaining
        : Math.min(remaining, renovationRoundUp(difference * (Number(ratio) || 0), 1000));
      allocations[key] = value;
      allocated += value;
    });
  }
  return {
    raw: cost, list, net, overhead, allocations, candidates,
    equipmentList, equipmentNet, equipmentCost,
    materialCost, materialNet, materialList, labor, renovationList, directList, directNet,
    trList, trNet, trCost,
    rateLabel: `${RENOVATION_METHODS[state.method]?.label || "改修"}・${RENOVATION_QUOTE_METHODS[state.quoteMethod].label}`,
  };
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
  const splitRows = state.quoteMethod === "bara" || state.method === "allocation" || state.method === "matsuoka";

  if (!splitRows) {
    const spec = `${methodLabel}・${quoteLabel}／標準×${state.pricingFactor}` +
      `${state.method === "transformer" ? `・TR×${state.transformerFactor}` : ""}` +
      `／材料×${state.materialListFactor}・改修×${state.renovationListFactor}`;
    add("summary", unit.unitName || "改修工事", spec, 1, pricing.list);
  } else {
    state.equipment.forEach((row, index) => add(
      `equipment:${row.lineId || index}`,
      row.name || "機器",
      row.spec || row.category || "",
      row.qty,
      row.listPrice
    ));

    if (state.method === "allocation" || state.method === "matsuoka") {
      const labels = {
        material: "材料費", survey: "現地調査費", renovation: "改修費", overhead: "諸経費",
        design: "設計費", transport: "運搬費", management: "管理費", misc: "雑費",
      };
      const allocationLabel = state.method === "matsuoka" ? "松岡Jr.指定比率" : "役所3.8定価の差額按分";
      Object.entries(pricing.allocations).forEach(([key, value]) =>
        add(`allocation:${key}`, labels[key] || key, allocationLabel, 1, value));
    } else {
      add("material", "材料費", `材料原価 ${fmtNum(pricing.materialCost)} × ${state.materialListFactor}`, 1, pricing.materialList);
      add("survey", "現地調査費", `入力NET ${fmtNum(state.siteSurveyNet)}`, 1, state.siteSurveyList);
      add("renovation", "改修費", `${fmtNum(state.labor.hourlyRate)}円 × ${state.labor.factor} × ${state.labor.hours}時間 × ${state.labor.people}人`, 1, pricing.renovationList);
      add("overhead", "諸経費", `${methodLabel}・${quoteLabel}の差額`, 1, pricing.overhead);
    }
  }
  unit.lines = [...normalLines, ...generated];
}

function syncAllRenovationQuoteLines() {
  currentEstimate.units.forEach(unit => {
    if (unit?.renovation?.enabled) syncRenovationQuoteLine(unit);
  });
}

function refreshRenovation() {
  const unit = activeUnit();
  syncRenovationQuoteLine(unit);
  renderRenovationTab();
  renderCatalogEstimatePanels();
}

function setRenovationMethod(method) {
  const state = getRenovationState();
  if (!state || !RENOVATION_METHODS[method]) return;
  state.method = method;
  if (method === "allocation" || method === "matsuoka") state.quoteMethod = "govt";
  refreshRenovation();
}

function setRenovationQuoteMethod(method) {
  const state = getRenovationState();
  if (!state || !RENOVATION_QUOTE_METHODS[method]) return;
  state.quoteMethod = method;
  refreshRenovation();
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
    netPrice: renovationMaterialPrice(item, "netPrice") || 0,
    costPrice: renovationMaterialPrice(item, "costPrice") || 0,
    note: item.note || "",
  });
  refreshRenovation();
}

function updateRenovationMaterial(lineId, field, value) {
  const row = getRenovationState()?.materials.find(item => item.lineId === lineId);
  if (!row) return;
  if (field === "note") row[field] = String(value || "");
  else if (field === "listPrice" && String(value).trim() === "") row[field] = null;
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
  const net = prompt("新しいNET（空欄は仕入価格を使用）", item.netPrice ?? item.unitPrice ?? "");
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
    netPrice: item.netPrice ?? item.unitPrice ?? null,
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
    { effectiveDate: item.effectiveDate, listPrice: item.listPrice, netPrice: item.netPrice ?? item.unitPrice, costPrice: item.costPrice ?? item.unitPrice, note: item.note || "", label: item.priceLabel || "現行" },
    ...(item.history || []),
  ];
  const format = value => value == null ? "未登録" : fmtNum(value);
  const overlay = document.createElement("div");
  overlay.className = "renovation-modal-overlay";
  overlay.onclick = event => { if (event.target === overlay) overlay.remove(); };
  overlay.innerHTML = `<div class="renovation-modal"><button class="renovation-modal-close" onclick="this.closest('.renovation-modal-overlay').remove()">×</button>` +
    `<h3>${esc(item.name)} ${esc(item.spec)} の価格履歴</h3><table><thead><tr><th>時点</th><th>定価</th><th>NET</th><th>仕入</th><th>備考</th></tr></thead><tbody>` +
    rows.map(row => `<tr><td>${esc(row.effectiveDate || row.label || "時期不明")}</td><td>${format(row.listPrice)}</td><td>${format(row.netPrice ?? row.unitPrice)}</td><td>${format(row.costPrice ?? row.unitPrice)}</td><td>${esc(row.note || "")}</td></tr>`).join("") +
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
  ).join("")}<button onclick="addRenovationUnit()">＋ 改修項目</button></div>`;
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
      `<div><strong>${manufacturer ? `${esc(manufacturer)}｜` : ""}${esc(item.name)}</strong>${item.spec ? `<small>${esc(item.spec)}</small>` : ""}` +
      `<span>定価 ${item.listPrice == null ? "未登録" : fmtNum(item.listPrice)} ／ NET ${item.netPrice == null ? "未登録" : fmtNum(item.netPrice)} ／ 仕入 ${item.costPrice == null ? "未登録" : fmtNum(item.costPrice)}</span>` +
      `${item.note ? `<small class="renovation-item-note">備考：${esc(item.note)}</small>` : ""}</div>` +
      `<div class="renovation-catalog-actions"><button onclick="showRenovationPriceHistory('${item.id}')">履歴</button>` +
      `<button onclick="editRenovationMasterPrice('${item.id}')">価格・備考</button>` +
      `<button class="primary" onclick="addRenovationEquipment('${item.id}')">追加</button></div></div>`;
  }).join("") || `<div class="renovation-empty">該当する機器はありません。</div>`;
  if (categoryNotes.length) container.insertAdjacentHTML("beforeend", renovationNotesHtml(categoryNotes, `${category.label}の元表メモ`));
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
    `<td><input type="number" min="0" step="1" value="${row.netPrice ?? row.unitPrice ?? 0}" onchange="updateRenovationMaterial('${row.lineId}','netPrice',this.value)"></td>` +
    `<td><input type="number" min="0" step="1" value="${row.costPrice ?? row.unitPrice ?? 0}" onchange="updateRenovationMaterial('${row.lineId}','costPrice',this.value)"></td>` +
    `<td><input class="renovation-note-input" type="text" value="${escAttr(row.note || "")}" onchange="updateRenovationMaterial('${row.lineId}','note',this.value)"></td>` +
    `<td>${fmtNum(row.qty * (row.costPrice ?? row.unitPrice ?? 0))}</td><td><button class="renovation-delete" onclick="removeRenovationMaterial('${row.lineId}')">×</button></td></tr>`;
  }).join("");
  const allocation = (state.method === "allocation" || state.method === "matsuoka") ? `<div class="renovation-allocation"><strong>${state.method === "matsuoka" ? "松岡Jr." : "3.8"} 振分額</strong>` +
    Object.entries(pricing.allocations).map(([key, value]) => `<span>${esc({material:"材料",survey:"現地調査",renovation:"改修",overhead:"諸経費",design:"設計",transport:"運搬",management:"管理",misc:"雑費"}[key] || key)} <b>${fmtNum(value)}</b></span>`).join("") + `</div>` : "";
  const candidateCards = Object.entries(RENOVATION_QUOTE_METHODS).map(([key, method]) => {
    const result = pricing.candidates[key];
    return `<button class="renovation-price-card${state.quoteMethod === key ? " active" : ""}" onclick="setRenovationQuoteMethod('${key}')">` +
      `<span class="renovation-price-card-title">${esc(method.label)}${state.quoteMethod === key ? "（見積採用中）" : ""}</span>` +
      `<span><small>定価</small><strong>${fmtNum(result.list)}</strong></span>` +
      `<span><small>NET</small><strong>${fmtNum(result.net)}</strong></span>` +
      `<span class="renovation-rate-row"><small>NET／定価 ${result.netToListRate.toFixed(1)}%</small><small>機器原価率 ${result.equipmentCostRate.toFixed(1)}%</small><small>総原価率 ${result.costRate.toFixed(1)}%</small><small>粗利率 ${result.grossMargin.toFixed(1)}%</small></span>` +
    `</button>`;
  }).join("");
  const overheadNet = Math.max(0, selected.net - pricing.directNet);
  const visibleMaterials = RENOVATION_MATERIALS
    .filter(item => renovationMaterialDisplayGroup(item) === renovationMaterialGroup)
    .map(item => getRenovationMaterialItem(item.id) || item);
  const materialNotes = RENOVATION_REFERENCE_NOTES?.materialGroups?.[renovationMaterialGroup] || [];
  const estimateNotes = [
    ...(RENOVATION_REFERENCE_NOTES?.general || []),
    ...(RENOVATION_REFERENCE_NOTES?.patterns?.[state.method] || []),
  ];

  root.innerHTML = `<div class="renovation-sticky-controls">` + renovationUnitSwitcherHtml() +
    `<div class="renovation-heading"><div><h2>改修見積 — ${esc(unit.unitName)}</h2><p>計算内容は同じ項目の「見積作成」に自動反映されます。</p></div>` +
    `<button class="btn btn-success" onclick="switchTab('estimate')">見積作成で確認</button></div></div>` +
    `<div class="renovation-layout"><section class="renovation-catalog"><div class="renovation-section-title">1. 機器を選択</div>` +
      `<input class="renovation-search" type="search" placeholder="全価格表から検索" value="${escAttr(renovationSearchText)}" oninput="setRenovationSearch(this.value)">` +
      `<div class="renovation-categories">${RENOVATION_CATEGORIES.map(cat => `<button data-category-id="${escAttr(cat.id)}" class="${cat.id === renovationCategoryId && !renovationSearchText ? "active" : ""}" aria-pressed="${cat.id === renovationCategoryId && !renovationSearchText}" onclick="setRenovationCategory('${cat.id}')">${esc(cat.label)}</button>`).join("")}</div>` +
      `<div id="renovation-catalog-list" class="renovation-catalog-list"></div></section>` +
    `<div class="renovation-work"><section><div class="renovation-section-title">2. 選択した機器</div>` +
      `<div class="renovation-table-wrap"><table><thead><tr><th>機器</th><th>数量</th><th>定価</th><th>NET</th><th>仕入</th><th>備考</th><th>仕入計</th><th></th></tr></thead><tbody>${equipmentRows || `<tr><td colspan="8" class="renovation-empty">左の価格表から追加してください。</td></tr>`}</tbody></table></div></section>` +
    `<section><div class="renovation-section-title">3. 材料費</div>` +
      `<div class="renovation-material-groups">${RENOVATION_MATERIAL_GROUPS.map(group => `<button class="${group === renovationMaterialGroup ? "active" : ""}" onclick="setRenovationMaterialGroup('${escAttr(group)}')">${esc(group)}</button>`).join("")}</div>` +
      `<div class="renovation-material-buttons">${visibleMaterials.map(item => {
        const capacity = renovationMaterialCapacity(item);
        const materialCost = renovationMaterialPrice(item, "costPrice");
        const materialNet = renovationMaterialPrice(item, "netPrice");
        return `<div class="renovation-material-card"><button class="renovation-material-add ${materialCost == null ? "price-missing" : ""}" onclick="addRenovationMaterial('${item.id}')">` +
          `<span>＋ ${esc(item.name)} ${esc(item.spec)}</span><small>定価 ${item.listPrice == null ? "係数計算" : fmtNum(item.listPrice)} ／ NET ${materialNet == null ? "未登録" : fmtNum(materialNet)} ／ 仕入 ${materialCost == null ? "未登録" : fmtNum(materialCost)}${capacity ? ` ／ ${esc(capacity.short)}` : ""}</small>` +
          `${item.note ? `<small class="renovation-item-note">備考：${esc(item.note)}</small>` : ""}</button>` +
          `<div class="renovation-material-card-actions"><button onclick="showRenovationMaterialPriceHistory('${item.id}')" title="価格履歴">履歴</button><button onclick="editRenovationMaterialPrice('${item.id}')" title="登録価格と備考を変更">価格・備考</button></div></div>`;
      }).join("")}</div>` +
      renovationNotesHtml(materialNotes, `${renovationMaterialGroup}の元表メモ`) +
      `<div class="renovation-table-wrap"><table><thead><tr><th>材料</th><th>数量</th><th>定価</th><th>NET</th><th>仕入</th><th>備考</th><th>仕入計</th><th></th></tr></thead><tbody>${materialRows || `<tr><td colspan="8" class="renovation-empty">材料を追加してください。</td></tr>`}</tbody></table></div></section>` +
    `<section><div class="renovation-section-title">4. 現地調査費</div><div class="renovation-cost-inputs renovation-cost-inputs-small">` +
      `<label>現地調査 定価<input type="number" min="0" value="${state.siteSurveyList}" onchange="updateRenovationNumber('siteSurveyList',this.value)"></label>` +
      `<label>現地調査 NET<input type="number" min="0" value="${state.siteSurveyNet}" onchange="updateRenovationNumber('siteSurveyNet',this.value)"></label>` +
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
      `<label>材料 定価係数<input type="number" min="0" step="0.01" value="${state.materialListFactor}" onchange="updateRenovationNumber('materialListFactor',this.value)"></label>` +
      `<label>改修費 定価係数<input type="number" min="0" step="0.01" value="${state.renovationListFactor}" onchange="updateRenovationNumber('renovationListFactor',this.value)"></label>` +
      `</div><div class="renovation-calculated-cost renovation-overhead-box"><span>採用方式の諸経費 定価</span><b>${fmtNum(selected.overhead)}</b><span>採用方式の諸経費 NET</span><b>${fmtNum(overheadNet)}</b></div></section></div>` +
    `<div class="renovation-side"><div class="renovation-pattern-panel"><div class="renovation-pattern-heading"><strong>積算パターンを選択</strong><span>お客様に合わせて選択</span></div>` +
      `<div class="renovation-methods">${Object.entries(RENOVATION_METHODS).map(([key, method]) =>
        `<button class="${state.method === key ? "active" : ""}" aria-pressed="${state.method === key}" onclick="setRenovationMethod('${key}')"><strong>${method.label}</strong><span>${method.description}</span></button>`
      ).join("")}</div></div>` +
    `<aside class="renovation-summary"><div class="renovation-section-title">計算結果</div>` +
      `<div class="renovation-result-breakdown"><table><thead><tr><th>項目</th><th>定価</th><th>NET</th><th>原価</th></tr></thead><tbody>` +
      `<tr><td>機器</td><td>${fmtNum(pricing.equipmentList)}</td><td>${fmtNum(pricing.equipmentNet)}</td><td>${fmtNum(pricing.equipmentCost)}</td></tr>` +
      `<tr><td>材料費</td><td>${fmtNum(pricing.materialList)}</td><td>${fmtNum(pricing.materialNet)}</td><td>${fmtNum(pricing.materialCost)}</td></tr>` +
      `<tr><td>現地調査費</td><td>${fmtNum(state.siteSurveyList)}</td><td>${fmtNum(state.siteSurveyNet)}</td><td>${fmtNum(state.siteSurveyNet)}</td></tr>` +
      `<tr><td>改修費</td><td>${fmtNum(pricing.renovationList)}</td><td>${fmtNum(pricing.labor)}</td><td>${fmtNum(pricing.labor)}</td></tr>` +
      `<tr class="overhead"><td>諸経費</td><td>${fmtNum(selected.overhead)}</td><td>${fmtNum(overheadNet)}</td><td>—</td></tr>` +
      `</tbody></table></div><div class="renovation-price-cards">${candidateCards}</div>${allocation}` +
      renovationNotesHtml(estimateNotes) +
      `<p>機器原価率は元Excelと同じく機器仕入＋材料原価で算出。総原価率・粗利率は現地調査と改修費も含めて算出しています。</p>` +
    `</aside></div></div>`;
  renderRenovationCatalog();
  const stickyControls = root.querySelector(".renovation-sticky-controls");
  if (typeof observeStickyOffsetElement === "function") observeStickyOffsetElement(stickyControls);
  requestAnimationFrame(() => {
    if (typeof updateStickyOffsets === "function") updateStickyOffsets();
  });
}
