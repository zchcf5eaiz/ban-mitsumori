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
    document: { addEventListener() {} },
    window: { addEventListener() {} },
    setTimeout() {}, clearTimeout() {},
    prompt: () => '2',
  });
  for (const file of ['data.js', 'data-cubicle.js', 'data-user-defaults.js', 'app.js']) {
    vm.runInContext(fs.readFileSync(path.join(__dirname, '..', file), 'utf8'), context, { filename: file });
  }
  // DOMの描画だけを省略し、実際の追加・移行・見積表示の処理を実行する。
  vm.runInContext(`
    updateAddedCount = () => {};
    updateCubicleAddedCount = () => {};
    refreshCatalogSelectionStyles = () => {};
    showToast = () => {};
  `, context);
  return code => JSON.parse(JSON.stringify(vm.runInContext(code, context)));
}

test('管理表の容量で見積表示する（MC・正逆・Y-△・SC両電圧）', () => {
  const run = createApp();
  const expected = {
    Bm091: '88 / 25A / 200V 5.5kW / 400V 11kW',
    Bm102: '88 / 35A / 200V 7.5kW / 400V 15kW',
    Bm113: '88 / 50A / 200V 11kW / 400V 22kW',
    Bm114: '88(正逆) / 50A / 200V 11kW / 400V 22kW',
    Bm115: 'Y-△2C / 50A / 200V 11kW / 400V 22kW',
    Bm116: 'Y-△3C / 50A / 200V 11kW / 400V 22kW',
    Bm117: 'Y-△M+2C / 50A / 200V 11kW / 400V 22kW',
    Bm118: 'Y-△CLOS / 50A / 200V 11kW / 400V 22kW',
    Bm097: 'SC220V / 200V 5.5kW',
    Bm098: 'SC440V / 400V 11kW',
    Bm212: '88 / 400A / 200V 110kW / 400V 220kW',
  };
  for (const [id, spec] of Object.entries(expected)) {
    assert.equal(run(`getMasterEstimateDisplay(getMasterItem('${id}')).spec`), spec, id);
  }
});

test('容量が空欄のとき内部行キーや0kWを表示しない', () => {
  const run = createApp();
  assert.equal(run("getMasterEstimateDisplay(getMasterItem('Bm003')).spec"), '88 / 200V 0.2kW');
  assert.equal(run("getMasterEstimateDisplay(getMasterItem('Bm058')).spec"), '88 / 400V 5.5kW');
  assert.equal(run("getMasterEstimateDisplay(getMasterItem('Bm064')).spec"), 'SC220V / 容量記載なし');
});

test('全160価格セルの追加で単価・数量・商品IDを保持し、表示名を保存する', () => {
  const run = createApp();
  const results = run(`masterItems.filter(m => m.name === 'MCなど(単品)' &&
    !/ MC(?:AF|200V|400V)$/.test(m.spec)).map(master => {
    addFromMaster({target: {tagName: 'SPAN', closest: () => null}}, master.id);
    const line = activeUnit().lines.at(-1);
    return {line, master, display: getMasterEstimateDisplay(master)};
  })`);
  assert.equal(results.length, 160);
  for (const { line, master, display } of results) {
    assert.equal(line.masterItemId, master.id);
    assert.equal(line.qty, 2);
    assert.equal(line.unitPrice, master.basePrice);
    assert.equal(line.sourceSpec, display.spec);
    assert.equal(line.sourceName, display.name);
  }
});

test('保存済みの旧仕様を補正し、編集済み単価・数量は変更しない', () => {
  const run = createApp();
  const result = run(`migrateEstimateLines([
    {type:'item', masterItemId:'Bm113', sourceName:'MCなど(単品)', sourceSpec:'18.5kW 88', qty:3, unitPrice:999},
    {type:'item', masterItemId:'Bm097', qty:4, unitPrice:123}
  ])`);
  assert.equal(result[0].sourceSpec, '88 / 50A / 200V 11kW / 400V 22kW');
  assert.equal(result[0].unitPrice, 999);
  assert.equal(result[0].qty, 3);
  assert.equal(result[1].sourceSpec, 'SC220V / 200V 5.5kW');
  assert.equal(result[1].unitPrice, 123);
  assert.equal(result[1].qty, 4);
});

test('追加時の仕様を保持し、後日の容量マスタ変更や再移行で書き換えない', () => {
  const run = createApp();
  const result = run(`(() => {
    addFromMaster({target:{tagName:'SPAN',closest:()=>null}}, 'Bm113');
    getMasterItem('Bm111').basePrice = 123;
    const line = migrateEstimateLines(activeUnit().lines)[0];
    return getEstimateLineDisplay(line);
  })()`);
  assert.equal(result.spec, '88 / 50A / 200V 11kW / 400V 22kW');
});

test('個別編集された仕様・名称と通常品目の表示を保持する', () => {
  const run = createApp();
  const lines = [
    {type:'item', masterItemId:'Bm113', sourceName:'指定機器', sourceSpec:'18.5kW 88', qty:1, unitPrice:65},
    {type:'item', masterItemId:'Bm113', sourceName:'MCなど(単品)', sourceSpec:'客先指定仕様', qty:1, unitPrice:65},
    {type:'item', masterItemId:'A001', sourceName:'主幹MCCB', sourceSpec:'50AF 2P', qty:1, unitPrice:61},
  ];
  assert.deepEqual(run(`migrateEstimateLines(${JSON.stringify(lines)})`), lines);
  assert.equal(run("getMasterEstimateDisplay(getMasterItem('A001')).spec"), '50AF 2P');
  assert.equal(run("getEstimateLineDisplay({masterItemId:'missing',sourceName:'保存品目',sourceSpec:'保存仕様'}).spec"), '保存仕様');
});

test('従来の別商品ID移行も引き続き適用される', () => {
  const run = createApp();
  assert.equal(run("migrateEstimateLines([{type:'item',masterItemId:'B060',unitPrice:35}])[0].masterItemId"), 'B06u');
});
