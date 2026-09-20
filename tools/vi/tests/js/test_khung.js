'use strict';
var test = require('node:test');
var assert = require('node:assert');
var path = require('node:path');

require(path.join(__dirname, '..', '..', 'thi_nghiem_parts', 'runtime', 'khung.js'));
var K = globalThis.THI_NGHIEM_KHUNG;

var KHAI_BAO = {
  thamSo: [{ ma: 'x', kieu: 'so', macDinh: 2 }, { ma: 'kieu', kieu: 'chon', macDinh: 'a' }],
  daiLuongDo: [{ ma: 'y', saiSo: 0.5, chuSo: 2 }, { ma: 'z', saiSo: 0, chuSo: 0 }],
  bangKiem: [
    { vao: {}, ra: { y: 4 }, saiSo: 0.001 },
    { vao: { x: 3 }, ra: { y: 9 }, saiSo: 0.001 },
    { vao: { x: 0 }, ra: { y: 0, z: null }, saiSo: 0.001 }
  ]
};
var DUNG = { tinh: function (p) { return { y: p.x * p.x, z: null }; } };

function cauHinh(nguoi) {
  return { nguoiThaoTac: nguoi, duDoan: { dapAn: 'B' }, quanSat: { soLanDo: 3 } };
}

test('bo sinh so co hat giong cho cung day so', function () {
  var a = K.taoNgauNhien(1), b = K.taoNgauNhien(1), c = K.taoNgauNhien(2);
  var dayA = [a(), a(), a()], dayB = [b(), b(), b()];
  assert.deepStrictEqual(dayA, dayB);
  assert.notStrictEqual(dayA[0], c());
  dayA.forEach(function (x) { assert.ok(x >= 0 && x < 1); });
  assert.strictEqual(K.taoNgauNhien(1)(), 0.6270739405881613);
});

test('nhieu chuan co trung binh 0 va do lech 1', function () {
  var rng = K.taoNgauNhien(42), n = 20000, tong = 0, binhPhuong = 0;
  for (var i = 0; i < n; i += 1) { var x = K.nhieuChuan(rng); tong += x; binhPhuong += x * x; }
  assert.ok(Math.abs(tong / n) < 0.03);
  assert.ok(Math.abs(Math.sqrt(binhPhuong / n) - 1) < 0.03);
});

test('tu kiem dat khi mo hinh dung, ke ca gia tri null', function () {
  assert.deepStrictEqual(K.tuKiem(KHAI_BAO, DUNG), { tong: 3, dat: 3, truot: [] });
});

test('tu kiem bao dung dong truot khi mo hinh sai', function () {
  var sai = { tinh: function (p) { return { y: p.x * p.x + (p.x === 3 ? 1 : 0), z: null }; } };
  var kq = K.tuKiem(KHAI_BAO, sai);
  assert.strictEqual(kq.dat, 2);
  assert.deepStrictEqual(kq.truot, [{ dong: 2, ma: 'y', mong: 9, duoc: 10 }]);
});

test('tu kiem khong nem loi khi mo hinh nem loi hoac tra ve thieu', function () {
  var nem = { tinh: function () { throw new Error('hong'); } };
  assert.strictEqual(K.tuKiem(KHAI_BAO, nem).dat, 0);
  var thieu = { tinh: function () { return {}; } };
  assert.strictEqual(K.tuKiem(KHAI_BAO, thieu).dat, 0);
  var nan = { tinh: function () { return { y: NaN, z: null }; } };
  assert.strictEqual(K.tuKiem(KHAI_BAO, nan).dat, 0);
});

test('ap dung phep bien doi va tra null khi khong xac dinh', function () {
  assert.strictEqual(K.apDungPhep('khong', 3), 3);
  assert.strictEqual(K.apDungPhep('binh-phuong', 3), 9);
  assert.strictEqual(K.apDungPhep('nghich-dao', 4), 0.25);
  assert.strictEqual(K.apDungPhep('can', 9), 3);
  assert.ok(Math.abs(K.apDungPhep('ln', Math.E) - 1) < 1e-12);
  assert.strictEqual(K.apDungPhep('nghich-dao', 0), null);
  assert.strictEqual(K.apDungPhep('ln', 0), null);
  assert.strictEqual(K.apDungPhep('can', -1), null);
  assert.strictEqual(K.apDungPhep('khong', null), null);
});

test('khop tuyen tinh tim dung duong thang', function () {
  var kq = K.khopTuyenTinh([[0, 1], [1, 3], [2, 5], [3, 7]]);
  assert.ok(Math.abs(kq.heSoGoc - 2) < 1e-12);
  assert.ok(Math.abs(kq.tungDoGoc - 1) < 1e-12);
  assert.ok(Math.abs(kq.tuongQuan - 1) < 1e-12);
  assert.strictEqual(K.khopTuyenTinh([[1, 1]]), null);
  assert.strictEqual(K.khopTuyenTinh([[1, 1], [1, 2]]), null);
});

test('dinh dang dung dau phay, khong in am khong, gach ngang khi thieu', function () {
  assert.strictEqual(K.dinhDang(2.0071, 3), '2,007');
  assert.strictEqual(K.dinhDang(-0.0001, 2), '0,00');
  assert.strictEqual(K.dinhDang(-0, 1), '0,0');
  assert.strictEqual(K.dinhDang(null, 2), '—');
  assert.strictEqual(K.dinhDang(NaN, 2), '—');
  assert.strictEqual(K.dinhDang(12, 0), '12');
});

test('danh dau thoat HTML truoc roi moi doi chi so', function () {
  assert.strictEqual(K.danhDau('H~2~SO~4~ <b> m/s^2^ **dam**'),
    'H<sub>2</sub>SO<sub>4</sub> &lt;b&gt; m/s<sup>2</sup> <b>dam</b>');
});

test('mau do chi cong nhieu khi bat va khi dai luong co sai so', function () {
  var thuc = { y: 4, z: 7 };
  assert.deepStrictEqual(K.mauDo(KHAI_BAO, thuc, K.taoNgauNhien(1), false), { y: 4, z: 7 });
  var nhieu = K.mauDo(KHAI_BAO, thuc, K.taoNgauNhien(1), true);
  assert.notStrictEqual(nhieu.y, 4);
  assert.ok(Math.abs(nhieu.y - 4) < 3);
  assert.strictEqual(nhieu.z, 7);
  assert.deepStrictEqual(K.mauDo(KHAI_BAO, { y: null }, K.taoNgauNhien(1), true), { y: null, z: null });
});

test('nhiem vu khoa thao tac cho toi khi du doan', function () {
  var nv = K.taoNhiemVu(cauHinh('nhom'));
  assert.strictEqual(nv.duocThaoTac(), false);
  assert.strictEqual(nv.ghiLanDo({ x: 1 }), false);
  assert.strictEqual(nv.chonDuDoan('  '), false);
  assert.strictEqual(nv.chonDuDoan('A'), true);
  assert.strictEqual(nv.buoc, 'quan-sat');
  assert.strictEqual(nv.chonDuDoan('B'), false);
  assert.strictEqual(nv.duDoan, 'A');
  assert.strictEqual(nv.duDoanDung(), null);
});

test('nhiem vu mo giai thich khi du so lan do va cham du doan luc do', function () {
  var nv = K.taoNhiemVu(cauHinh('nhom'));
  nv.chonDuDoan('B');
  nv.ghiLanDo({ x: 1 }); nv.ghiLanDo({ x: 2 });
  assert.strictEqual(nv.buoc, 'quan-sat');
  nv.ghiLanDo({ x: 3 });
  assert.strictEqual(nv.buoc, 'giai-thich');
  assert.strictEqual(nv.duDoanDung(), true);
  assert.strictEqual(nv.xoaLanDo(0), true);
  assert.strictEqual(nv.xoaLanDo(9), false);
  assert.strictEqual(nv.lanDo.length, 2);
  assert.strictEqual(nv.buoc, 'giai-thich');
});

test('che do giao vien bo khoa nhung khong tu chuyen buoc', function () {
  var nv = K.taoNhiemVu(cauHinh('giao-vien'));
  assert.strictEqual(nv.duocThaoTac(), true);
  assert.strictEqual(nv.ghiLanDo({ x: 1 }), true);
  assert.strictEqual(nv.buoc, 'du-doan');
  var nhom = K.taoNhiemVu(cauHinh('nhom'));
  nhom.batGiaoVien();
  assert.strictEqual(nhom.duocThaoTac(), true);
});

test('du doan cau mo khong co dap an thi khong cham', function () {
  var ch = cauHinh('nhom'); ch.duDoan.dapAn = null;
  var nv = K.taoNhiemVu(ch);
  nv.chonDuDoan('em nghi la tang');
  nv.ghiLanDo({}); nv.ghiLanDo({}); nv.ghiLanDo({});
  assert.strictEqual(nv.duDoanDung(), null);
});

test('diem do thi bo lan do khong xac dinh', function () {
  var doThi = { tung: { ma: 'y', phep: 'binh-phuong' }, hoanh: { ma: 'x', phep: 'nghich-dao' } };
  assert.deepStrictEqual(K.diemDoThi(doThi, [{ x: 2, y: 3 }, { x: 0, y: 1 }, { x: 4, y: null }]), [[0.5, 9]]);
});
