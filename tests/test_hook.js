// 在模拟 Frida 环境中验证完整性检查与析构，不连接 QQ 音乐。
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
let readSize = 2, destructed = 0, writes = 0, closed = 0, failWrite = false;
const context = {
  Process: {arch: 'ia32', findModuleByName: () => ({findExportByName: name => name})},
  NativeFunction: function(name) {
    if (name.startsWith('??0')) return () => {};
    if (name.startsWith('??1')) return () => {destructed++;};
    if (name.startsWith('?Open')) return () => true;
    if (name.startsWith('?GetSize')) return () => 4;
    if (name.startsWith('?Read')) return () => readSize;
    throw new Error('Unexpected symbol: ' + name);
  },
  Memory: {alloc: () => ({readByteArray: size => new Uint8Array(size)}), allocUtf16String: s => s},
  File: function() {
    this.write = () => {writes++; if (failWrite) throw new Error('write failed');};
    this.close = () => {closed++;};
  },
  rpc: {exports: {}},
};
vm.runInNewContext(fs.readFileSync(path.join(__dirname, '../src/hook_qqmusic_v17.js'), 'utf8'), context);
assert.throws(() => context.rpc.exports.decrypt('input', 'output'), /读取不完整/);
assert.equal(destructed, 1);
assert.equal(writes, 0);
readSize = 4;
assert.equal(context.rpc.exports.decrypt('input', 'output').size, 4);
assert.equal(destructed, 2);
assert.equal(closed, 1);
failWrite = true;
assert.throws(() => context.rpc.exports.decrypt('input', 'output'), /write failed/);
assert.equal(destructed, 3);
assert.equal(closed, 2);
console.log('Hook integrity and cleanup checks passed.');
