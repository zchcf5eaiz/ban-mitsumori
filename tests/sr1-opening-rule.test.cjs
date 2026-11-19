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
        return { textContent: '', get innerHTML() { return String(this.textContent); } };
      },
    },
    window: { addEventListener() {} },
    setTimeout() {}, clearTimeout() {},
  });
  for (const file of ['data.js', 'data-cubicle.js', 'data-renovation.js', 'data-user-defaults.js', 'app.js', 'renovation.js']) {
    vm.runInContext(fs.readFileSync(path.join(__dirname, '..', file), 'utf8'), context, { filename: file });
  }
  return code => vm.runInContext(code, context);
}

test('SR-1の抜き加工は基本300・上段25・中下段50をそのまま合計する', () => {
  const run = createApp();
  assert.equal(run('sr1OpeningAmount(7, true, true)'), 575);
  assert.equal(run('sr1OpeningAmount(3, true, true)'), 475);
  assert.equal(run('sr1OpeningAmount(8, true, true)'), 600);
  assert.equal(run('sr1OpeningAmount(6, true, true)'), 550);
  assert.equal(run('sr1OpeningAmount(1, false, false)'), 325);
});
