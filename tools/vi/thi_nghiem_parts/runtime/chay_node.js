'use strict';
// node chay_node.js <khung.js> <mo-hinh.js> <mo-hinh.json> bang-kiem|luoi   (luoi đọc danh sách tham số từ stdin)
var fs = require('fs');
var vm = require('vm');

function main() {
  var khung = fs.readFileSync(process.argv[2], 'utf8');
  var moHinh = fs.readFileSync(process.argv[3], 'utf8');
  var khaiBao = fs.readFileSync(process.argv[4], 'utf8');
  var cheDo = process.argv[5];
  var vao = cheDo === 'luoi' ? fs.readFileSync(0, 'utf8') : '[]';
  var hop = vm.createContext({ KHAI_BAO_JSON: khaiBao, VAO_JSON: vao });
  var gioiHan = { timeout: 10000 };
  vm.runInContext(khung, hop, gioiHan);
  vm.runInContext(moHinh, hop, gioiHan);
  var lenh = cheDo === 'luoi'
    ? 'JSON.stringify(JSON.parse(VAO_JSON).map(function (v) {' +
      ' return THI_NGHIEM_MO_HINH.tinh(THI_NGHIEM_KHUNG.gopThamSo(JSON.parse(KHAI_BAO_JSON), v)); }))'
    : 'JSON.stringify(THI_NGHIEM_KHUNG.tuKiem(JSON.parse(KHAI_BAO_JSON), THI_NGHIEM_MO_HINH))';
  process.stdout.write(vm.runInContext(lenh, hop, gioiHan) + '\n');
}

try { main(); } catch (loi) { process.stdout.write(JSON.stringify({ loi: String(loi) }) + '\n'); process.exit(1); }
