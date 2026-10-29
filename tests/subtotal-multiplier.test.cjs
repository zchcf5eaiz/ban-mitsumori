const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function createApp() {
  const storage = new Map();
  const context = vm.createContext({
    localStorage: {
      getItem: key => storage.get(key) ?? null,
      setItem: (key, value) => storage.set(key, value),
      removeItem: key => storage.delete(key),
    },
    document: {
      addEventListener() {},
      createElement() {
        return { textContent: '', get innerHTML() {
          return String(this.textContent).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        } };
      },
    },
    window: { addEventListener() {} },
    setTimeout() {}, clearTimeout() {},
  });
  for (const file of ['data.js', 'data-cubicle.js', 'data-user-defaults.js', 'app.js']) {
    vm.runInContext(fs.readFileSync(path.join(__dirname, '..', file), 'utf8'), context, { filename: file });
  }
  vm.runInContext('refreshSubtotalViews = () => {};', context);
  return code => JSON.parse(JSON.stringify(vm.runInContext(code, context)));
}

const item = (price, qty = 1) => ({type: 'custom', lineId: 'item', name: '試験品', unitPrice: price, qty});
const subtotal = (rate, secondRate) => ({type: 'subtotal', lineId: 'subtotal', rate, ...(secondRate === undefined ? {} : {secondRate})});

test('従来の小計は同じ計算で、2段目は一度目の結果に掛ける', () => {
  const run = createApp();
  assert.equal(run('calcSubtotalResult(100, {rate:1.2})'), 120);
  assert.equal(run('calcSubtotalResult(100, {rate:1.2, secondRate:0.8})'), 96);
  // 一括乗算なら85になるケース。各段階で切り上げて86を得る。
  assert.equal(run('calcSubtotalResult(101, {rate:1.05, secondRate:0.8})'), 86);
  assert.equal(run('calcSubtotalResult(101, {rate:1.05, secondRate:0})'), 0);
});

test('積算合計は2段目の最終金額だけを数え、別小計と末尾の明細を保持する', () => {
  const run = createApp();
  const lines = [item(50, 2), {type:'comment',text:'備考'}, subtotal(1.2, 0.8), item(30), subtotal(2), item(10)];
  assert.equal(run(`calcLinesGrandTotal(${JSON.stringify(lines)})`), 166);
  assert.equal(run(`calcSubtotal(${JSON.stringify(lines)}, 4)`), 30);
  assert.equal(run('calcLinesGrandTotal([{type:"subtotal",rate:2,secondRate:3}])'), 0);
});

test('盤・キュービクル双方の項目の定価とNETへ最終金額が伝わる', () => {
  const run = createApp();
  for (const sourceCatalog of ['master', 'cubicle']) {
    const unit = {lines:[{...item(50,2),type:'item',sourceCatalog},subtotal(1.2,0.8)],listRate:1.5,netRate:0.33};
    const pricing = run(`calcUnitPricing(${JSON.stringify(unit)})`);
    assert.equal(pricing.raw, 96);
    assert.equal(pricing.list, 144000);
    assert.equal(pricing.net, 50000);
  }
});

test('追加・編集・解除で既存の小計倍率を保持し、別項目へ影響しない', () => {
  const run = createApp();
  const result = run(`(() => {
    currentEstimate.units = [
      {lines:[{type:'subtotal',lineId:'a',rate:1.2}]},
      {lines:[{type:'subtotal',lineId:'b',rate:2}]}
    ];
    addSubtotalMultiplier('a');
    onSubtotalRate('a', '0.8', 'secondRate');
    addSubtotalMultiplier('a');
    const added = JSON.parse(JSON.stringify(currentEstimate.units));
    removeSubtotalMultiplier('a');
    return {added, removed:currentEstimate.units};
  })()`);
  assert.deepEqual(result.added[0].lines[0], {type:'subtotal',lineId:'a',rate:1.2,secondRate:0.8});
  assert.deepEqual(result.removed[0].lines[0], {type:'subtotal',lineId:'a',rate:1.2});
  assert.deepEqual(result.added[1], result.removed[1]);
  assert.equal(result.removed[1].lines[0].rate, 2);
});

test('空欄・無効値を安全に扱い、0倍は有効にする', () => {
  const run = createApp();
  for (const value of ['', 'abc', '-1', 'Infinity', null]) {
    assert.equal(run(`subtotalRate(${JSON.stringify(value)})`), 1);
  }
  assert.equal(run('subtotalRate("0")'), 0);
  assert.equal(run('subtotalRate("0.25")'), 0.25);
});

test('保存・JSON読込時の移行で2段目の倍率と金額を保持する', () => {
  const run = createApp();
  const estimate = {units:[{unitId:'u',unitName:'試験',lines:[item(100),subtotal(1.2,0.8)],listRate:1.5,netRate:0.33}]};
  const restored = run(`migrateEstimate(JSON.parse(JSON.stringify(${JSON.stringify(estimate)})))`);
  assert.deepEqual(restored, estimate);
  assert.equal(run(`calcLinesGrandTotal(${JSON.stringify(restored.units[0].lines)})`), 96);
});

test('倍率の理由を各段ごとに保存し、管理表・見積・印刷へ安全に表示する', () => {
  const run = createApp();
  const result = run(`(() => {
    currentEstimate.units = [{lines:[
      {type:'custom',lineId:'item',qty:1,unitPrice:100},
      {type:'subtotal',lineId:'sub',rate:1.2,secondRate:0.8}
    ]}];
    document.querySelectorAll = () => [];
    onSubtotalReason('sub', '特殊仕様 "割増" <確認>', 'rateReason');
    onSubtotalReason('sub', '調整値引き', 'secondRateReason');
    const restored = migrateEstimate(JSON.parse(JSON.stringify(currentEstimate)));
    const unit = restored.units[0];
    return {line:unit.lines[1], total:calcLinesGrandTotal(unit.lines),
      estimate:_buildLineRows(unit.lines)[1], catalog:buildCatalogEstimateRows(unit)};
  })()`);
  assert.equal(result.line.rateReason, '特殊仕様 "割増" <確認>');
  assert.equal(result.line.secondRateReason, '調整値引き');
  assert.equal(result.total, 96);
  for (const html of [result.estimate, result.catalog]) {
    assert.match(html, /&quot;割増&quot;/);
    assert.match(html, /&lt;確認&gt;/);
    assert.match(html, /subtotal-reason-text has-reason">調整値引き</);
  }
  assert.equal(run(`estimatePrintRowWeight(${JSON.stringify(result.estimate)})`), 4);
});

test('見積・管理表・印刷の共通行に中間金額と最終金額を表示する', () => {
  const run = createApp();
  const lines = [item(100), subtotal(1.2, 0.8)];
  const estimateRows = run(`_buildLineRows(${JSON.stringify(lines)})`);
  const catalogRows = run(`buildCatalogEstimateRows({lines:${JSON.stringify(lines)}})`);
  for (const html of [estimateRows[1], catalogRows]) {
    assert.match(html, /subtotal-second-row/);
    assert.match(html, /subtotal-sum">120</);
    assert.match(html, /subtotal-value">96</);
    assert.match(html, /subtotal-rate-text[^>]*>&times; 0.8</);
  }
  assert.equal(run(`estimatePrintRowWeight(${JSON.stringify(estimateRows[1])})`), 2);
  const oldRows = run(`_buildLineRows(${JSON.stringify([item(100),subtotal(1.2)])})`);
  assert.match(oldRows[1], /さらに×/);
  assert.doesNotMatch(oldRows[1], /subtotal-second-row/);
  assert.equal(run(`estimatePrintRowWeight(${JSON.stringify(oldRows[1])})`), 1);
});

test('印刷の段組みで小計と再計算を分断せず、2行分で改列する', () => {
  const run = createApp();
  const lines = [item(100), subtotal(1.2, 0.8)];
  const chunks = run(`renovationDetailPrintChunks({lines:${JSON.stringify(lines)}}, 3)`);
  assert.equal(chunks.length, 2);
  assert.equal(chunks[1].length, 1);
  assert.equal((chunks[1][0].match(/<tr /g) || []).length, 2);
});

test('数量・単価変更後は入力欄を作り直さず両段の表示金額を更新する', () => {
  const run = createApp();
  const result = run(`(() => {
    currentEstimate.units = [{lines:[
      {type:'custom',lineId:'item',qty:2,unitPrice:100},
      {type:'subtotal',lineId:'sub',rate:1.2,secondRate:0.8}
    ]}];
    const rows = ['first','second'].map(stage => ({
      dataset:{subtotalId:'sub',subtotalStage:stage},
      cells:{'.subtotal-sum':{},'.subtotal-value':{}},
      querySelector(selector) { return this.cells[selector]; }
    }));
    document.querySelectorAll = () => rows;
    refreshSubtotalAmounts();
    const before = rows.map(row => Object.values(row.cells).map(cell => cell.textContent));
    currentEstimate.units[0].lines[0].unitPrice = 50;
    refreshSubtotalAmounts();
    return {before,after:rows.map(row => Object.values(row.cells).map(cell => cell.textContent))};
  })()`);
  assert.deepEqual(result.before, [['200','240'],['240','192']]);
  assert.deepEqual(result.after, [['100','120'],['120','96']]);
});
