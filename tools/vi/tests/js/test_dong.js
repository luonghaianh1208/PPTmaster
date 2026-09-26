'use strict';
var test = require('node:test');
var assert = require('node:assert');
var path = require('node:path');

var RT = path.join(__dirname, '..', '..', 'video_ma_parts', 'runtime');
require(path.join(RT, 'dong.js'));
require(path.join(RT, 'khung-video.js'));
require(path.join(RT, 'nhan.js'));
var D = globalThis.THI_DONG;
var V = globalThis.THI_VIDEO;
var N = globalThis.THI_NHAN;

function gan(a, b, sai) { return Math.abs(a - b) < (sai || 1e-9); }
function hienThi(html) {
  return html.replace(/<span class="an">[^<]*<\/span>/g, '').replace(/<[^>]+>/g, '');
}
// Thẻ mở/đóng phải cân: mỗi thẻ đóng khớp thẻ mở gần nhất.
function canThe(html) {
  var ngan = [];
  var re = /<(\/?)([a-z]+)[^>]*>/g;
  var m;
  while ((m = re.exec(html))) {
    if (m[1]) {
      if (ngan.pop() !== m[2]) { return false; }
    } else {
      ngan.push(m[2]);
    }
  }
  return ngan.length === 0;
}

test('ham lam muot: 0 tai 0, 1 tai 1, kep ngoai 0..1, xac dinh', function () {
  [D.easeOutBack, D.easeInOut, D.lo_xo].forEach(function (f) {
    assert.ok(gan(f(0), 0), f.name + '(0)');
    assert.ok(gan(f(1), 1), f.name + '(1)');
    assert.ok(gan(f(-3), 0) && gan(f(7), 1), f.name + ' kep');
    assert.strictEqual(f(0.37), f(0.37));
  });
  assert.ok(D.easeOutBack(0.7) > 1, 'easeOutBack vuot qua 1 roi ve');
  assert.ok(gan(D.easeInOut(0.5), 0.5));
  var vuot = false;
  for (var x = 0; x <= 1; x += 0.01) { if (D.lo_xo(x) > 1.01) { vuot = true; } }
  assert.ok(vuot, 'lo xo nay qua dich');
  assert.ok(gan(D.lo_xo(1, 20, 0.2), 1));
});

test('soChay: 0 tai batDau, giua chua xong, sau khi xong la gia tri kieu Viet', function () {
  assert.strictEqual(D.soChay(2, 2, 0.8, 1500, 0), '0');
  assert.strictEqual(D.soChay(1, 2, 0.8, 1500, 0), '0');
  var giua = D.soChay(2.4, 2, 0.8, 1500, 0);
  assert.ok(Number(giua) > 0 && Number(giua) < 1500, giua);
  assert.strictEqual(D.soChay(2.8, 2, 0.8, 1500, 0), '1500');
  assert.strictEqual(D.soChay(9, 2, 0.8, 1500.5, 1), '1500,5');
  assert.strictEqual(D.soChay(2, 2, 0.8, 1500.5, 1), '0,0');
  assert.strictEqual(D.soChay(9, 0, 0.8, -2.25, 2), '-2,25');
  assert.strictEqual(D.soChay(0, 0, 0.8, -2.25, 2), '0,00');
  assert.ok(/^\d+,\d$/.test(D.soChay(2.3, 2, 0.8, 1500.5, 1)));
});

test('nayChu: chu i bat dau lech 0,04 s, cuoi cung dung yen', function () {
  var a = D.nayChu(0, 10, 1.0, 1.0);
  assert.strictEqual(a.a, 0);
  var b = D.nayChu(3, 10, 1.0 + 3 * 0.04, 1.0);
  assert.strictEqual(b.a, 0, 'chu 3 chua hien khi toi luot');
  var sau = D.nayChu(3, 10, 1.0 + 3 * 0.04 + 0.1, 1.0);
  assert.ok(sau.a > 0);
  var het = D.nayChu(9, 10, 1.0 + D.thoiGianNay(10), 1.0);
  assert.deepStrictEqual(het, { s: 1, y: 0, a: 1 });
  assert.deepStrictEqual(D.nayChu(2, 10, 1.3, 1.0), D.nayChu(2, 10, 1.3, 1.0));
  assert.ok(gan(D.thoiGianNay(10), 0.04 * 9 + D.NAY));
});

test('khoa: chu thuong, bo dau cau, giu dau thanh', function () {
  assert.strictEqual(N.khoa('Chu kì,'), 'chu kì');
  assert.strictEqual(N.khoa('“Đơn”.'), 'đơn');
  assert.notStrictEqual(N.khoa('kỳ'), N.khoa('kì'));
});

test('tachCum: kieu, noi dung hien thi va vi tri tinh tren chu hien thi', function () {
  var ds = N.tachCum('Ta có ==chu kì== và ((tần số)) cùng __H~2~O__.');
  assert.deepStrictEqual(ds.map(function (c) { return c.kieu; }), ['to', 'khoanh', 'gach']);
  assert.deepStrictEqual(ds.map(function (c) { return c.noiDung; }), ['chu kì', 'tần số', 'H2O']);
  assert.strictEqual(ds[0].viTri, 'Ta có '.length);
  assert.strictEqual(ds[1].viTri, 'Ta có chu kì và '.length);
  assert.strictEqual(ds[2].viTri, 'Ta có chu kì và tần số cùng '.length);
  assert.deepStrictEqual(N.tachCum('không có gì'), []);
});

var TU = [
  { t: 1.0, chu: 'Chu' }, { t: 1.2, chu: 'kì,' }, { t: 1.5, chu: 'là' },
  { t: 2.0, chu: 'thời' }, { t: 2.2, chu: 'gian.' }, { t: 3.0, chu: 'Chu' }, { t: 3.2, chu: 'kì' }
];

test('thoiDiemNhan: lan dau sau batDauMuc; khong phan biet hoa thuong, bo dau cau, giu dau thanh', function () {
  assert.strictEqual(N.thoiDiemNhan({ noiDung: 'chu kì' }, TU, 0.5, 5), 1.0);
  assert.strictEqual(N.thoiDiemNhan({ noiDung: 'Chu kì' }, TU, 1.1, 5), 3.0, 'lan sau batDauMuc');
  assert.strictEqual(N.thoiDiemNhan({ noiDung: 'thời gian' }, TU, 0, 5), 2.0);
  assert.strictEqual(N.thoiDiemNhan({ noiDung: 'chu kỳ' }, TU, 0, 5), 5.3, 'khac dau thanh thi khong khop');
  assert.strictEqual(N.thoiDiemNhan({ noiDung: 'tần số' }, TU, 0, 4), 4.3);
  assert.strictEqual(N.thoiDiemNhan({ noiDung: 'chu kì' }, [], 0, 2), 2.3);
  assert.strictEqual(N.thoiDiemNhan({ noiDung: 'kì là' }, TU, 0, 5), 1.2, 'cum cat ngang cau van khop');
  var coKhoa = [{ t: 1, chu: 'Chu', khoa: 'chu' }, { t: 1.4, chu: 'kì', khoa: 'kì' }];
  assert.strictEqual(N.thoiDiemNhan({ noiDung: 'chu kì' }, coKhoa, 0, 5), 1);
});

test('catDanhDau: cum va so boc dung the, chu hien thi dung, danh dau khong dem', function () {
  var chu = 'Ta có ==chu kì== và ((tần số)) và __biên độ__ bằng {{1500.5}} m';
  var tong = V.demKyTu(chu);
  assert.strictEqual(tong, 'Ta có chu kì và tần số và biên độ bằng 1500,5 m'.length);
  var du = V.catDanhDau(chu, tong);
  assert.strictEqual(hienThi(du), 'Ta có chu kì và tần số và biên độ bằng 1500,5 m');
  assert.ok(du.indexOf('<span class="cum to" data-cum="0">chu kì</span>') >= 0, du);
  assert.ok(du.indexOf('<span class="cum khoanh" data-cum="1">tần số</span>') >= 0, du);
  assert.ok(du.indexOf('<span class="cum gach" data-cum="2">biên độ</span>') >= 0, du);
  assert.ok(du.indexOf('<span class="so" data-so="0">1500,5</span>') >= 0, du);
  assert.ok(canThe(du));
});

test('catDanhDau: cat giua cum hay giua the long trong cum van can the, phan con lai an', function () {
  var chu = 'A ==x **đậm** <b>== z {{12}}';
  var tong = V.demKyTu(chu);
  var chuThuan = 'A x đậm <b> z 12';
  assert.strictEqual(tong, chuThuan.length);
  for (var n = 0; n <= tong; n++) {
    var html = V.catDanhDau(chu, n);
    assert.ok(canThe(html), 'n=' + n + ' ' + html);
    assert.ok((html.match(/<b>/g) || []).length <= 1, 'chu `<b>` phai duoc thoat: ' + html);
    assert.strictEqual(html.split('data-cum="0"').length, 2, 'cum luon co mat de bo cuc khong nhay: n=' + n);
    var so = (html.match(/<span class="ngoi"><\/span>/g) || []).length;
    assert.strictEqual(so, n > 0 && n < tong ? 1 : 0, 'n=' + n);
  }
  var giua = V.catDanhDau(chu, 6);
  assert.strictEqual(giua, 'A <span class="cum to" data-cum="0">x <b>đậ<span class="ngoi"></span><span class="an">m</span></b>' +
    '<span class="an"> &lt;b&gt;</span></span><span class="an"> z </span><span class="so" data-so="0"><span class="an">12</span></span>');
  assert.strictEqual(V.catDanhDau('a ((b)) ~2~', 0).indexOf('<span class="cum khoanh" data-cum="0"><span class="an">b</span></span>') >= 0, true);
});

test('catDanhDau: so dang chay giu chu so cuoi de giu be rong', function () {
  var html = V.catDanhDau('được {{1500}} m', 99, ['37']);
  assert.ok(html.indexOf('<span class="so" data-so="0"><span class="so-cuoi">1500</span><span class="so-chay">37</span></span>') >= 0, html);
  assert.strictEqual(V.catDanhDau('được {{1500}} m', 99, ['1500']), V.catDanhDau('được {{1500}} m', 99));
  assert.strictEqual(V.catDanhDau('{{abc}} {{1,5}}', 99), '{{abc}} {{1,5}}', 'khong phai so thi giu nguyen chu');
});

test('catDanhDau: cu phap cu giu nguyen', function () {
  assert.strictEqual(V.catDanhDau('**ab**c', 1), '<b>a<span class="ngoi"></span><span class="an">b</span></b><span class="an">c</span>');
  assert.strictEqual(V.catDanhDau('H~2~O', 99), 'H<sub>2</sub>O');
  assert.strictEqual(V.catDanhDau('a < b', 99), 'a &lt; b');
  var doc = V.catDanhDau('==</script><img src=x onerror=alert(1)>== {{1}}', 99);
  assert.ok(doc.indexOf('<img') === -1 && doc.indexOf('</script') === -1, doc);
});

test('catDanhDau che do nay: moi chu mot span, tu khong bi ngat, chu hien thi du', function () {
  var chu = 'Con ==lắc== đơn';
  var tong = V.demKyTu(chu);
  var html = V.catDanhDau(chu, tong, null, function (i) { return { s: 1, y: 0, a: 1 }; });
  assert.strictEqual(html.replace(/<[^>]+>/g, ''), 'Con lắc đơn');
  assert.strictEqual((html.match(/class="nay"/g) || []).length, 'Conlắcđơn'.length);
  assert.ok(html.indexOf('data-cum="0"') >= 0);
  assert.ok(canThe(html));
  var dau = V.catDanhDau(chu, tong, null, function (i) { return i === 0 ? { s: 0.5, y: -10, a: 0.25 } : { s: 1, y: 0, a: 0 }; });
  assert.ok(dau.indexOf('opacity:0.25') >= 0, dau);
});

test('duongKhoanh: vong kin quanh hop, xac dinh, nam trong khung va tren y 620', function () {
  var d = N.duongKhoanh({ x: 100, y: 200, w: 180, h: 40 }, 3);
  assert.strictEqual(d, N.duongKhoanh({ x: 100, y: 200, w: 180, h: 40 }, 3));
  var so = d.match(/-?\d+(\.\d+)?/g).map(Number);
  var xs = so.filter(function (_, i) { return i % 2 === 0; });
  var ys = so.filter(function (_, i) { return i % 2 === 1; });
  assert.ok(Math.min.apply(null, xs) < 100 && Math.max.apply(null, xs) > 280);
  assert.ok(Math.min.apply(null, ys) < 200 && Math.max.apply(null, ys) > 240);
  var sat = N.duongKhoanh({ x: 1150, y: 585, w: 125, h: 30 }, 1).match(/-?\d+(\.\d+)?/g).map(Number);
  sat.forEach(function (v, i) {
    if (i % 2 === 0) { assert.ok(v >= 0 && v <= 1280, 'x ' + v); } else { assert.ok(v >= 0 && v <= 620, 'y ' + v); }
  });
  var g = N.duongGach([{ x: 10, y: 600, w: 100, h: 30 }], 2).match(/-?\d+(\.\d+)?/g).map(Number);
  g.forEach(function (v, i) { if (i % 2 === 1) { assert.ok(v <= 620, 'gach y ' + v); } });
});
