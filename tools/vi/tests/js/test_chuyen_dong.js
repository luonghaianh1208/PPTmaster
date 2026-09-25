'use strict';
var test = require('node:test');
var assert = require('node:assert');
var path = require('node:path');

var RT = path.join(__dirname, '..', '..', 'video_ma_parts', 'runtime');
require(path.join(RT, 'ban-tay.js'));
require(path.join(RT, 'may-quay.js'));
var T = globalThis.THI_BAN_TAY;
var Q = globalThis.THI_MAY_QUAY;
var NGHI = { x: 1220, y: 760 };

// Ngòi giả: mỗi mục có đoạn thẳng riêng; điểm ngòi chạy từ đầu tới cuối đoạn theo p.
var DUONG = { a: [[100, 100], [300, 100]], b: [[100, 300], [500, 300]], c: [[700, 200], [900, 400]] };
function ngoi(m, p) {
  var d = DUONG[m.id];
  return { x: d[0][0] + (d[1][0] - d[0][0]) * p, y: d[0][1] + (d[1][1] - d[0][1]) * p };
}
function muc(id, batDau, thoiLuong, tuy) {
  var m = { id: id, kieu: 'chu', batDau: batDau, thoiLuong: thoiLuong };
  Object.keys(tuy || {}).forEach(function (k) { m[k] = tuy[k]; });
  return m;
}
function gan(a, b, sai) { return Math.abs(a - b) < (sai || 1e-9); }

test('ban tay: xac dinh theo t', function () {
  var ds = [muc('a', 1, 1), muc('b', 2.4, 1), muc('c', 5, 1)];
  for (var t = 0; t < 8; t += 0.137) {
    assert.deepStrictEqual(T.viTri(ds, t, ngoi, NGHI), T.viTri(ds, t, ngoi, NGHI));
  }
});

test('ban tay: dang ve thi o dung ngoi cua muc do', function () {
  var ds = [muc('a', 1, 1), muc('b', 2.4, 1)];
  var v = T.viTri(ds, 1.5, ngoi, NGHI);
  assert.strictEqual(v.hien, true);
  assert.strictEqual(v.kieu, 'but');
  assert.ok(gan(v.x, 200) && gan(v.y, 100), JSON.stringify(v));
});

test('ban tay: khoang trong 0,4 s thi noi suy tuyen tinh tu cuoi muc truoc toi dau muc sau', function () {
  var ds = [muc('a', 1, 1), muc('b', 2.4, 1)];
  var v = T.viTri(ds, 2.2, ngoi, NGHI);
  assert.strictEqual(v.hien, true);
  assert.ok(gan(v.x, (300 + 100) / 2) && gan(v.y, (100 + 300) / 2), JSON.stringify(v));
  var dau = T.viTri(ds, 2.0, ngoi, NGHI);
  assert.ok(gan(dau.x, 300) && gan(dau.y, 100));
});

test('ban tay va may quay: hai muc cung bat dau thi uu tien hinh, roi chu, roi net', function () {
  var ds = [muc('a', 1, 0.3, { kieu: 'net' }), muc('b', 1, 2.5), muc('c', 1, 1.4, { kieu: 'hinh' })];
  var v = T.viTri(ds, 1.2, ngoi, NGHI);
  assert.ok(gan(v.x, ngoi(ds[2], 0.2 / 1.4).x), 've hinh truoc: ' + JSON.stringify(v));
  var sau = T.viTri(ds, 2.6, ngoi, NGHI);
  assert.ok(gan(sau.x, ngoi(ds[1], 1.6 / 2.5).x), 'hinh xong thi viet tiep chu: ' + JSON.stringify(sau));
  var ds2 = [muc('a', 1, 0.3, { kieu: 'net' }), muc('b', 1, 1.5)];
  assert.ok(gan(T.viTri(ds2, 1.2, ngoi, NGHI).x, ngoi(ds2[1], 0.2 / 1.5).x), 'chu truoc net');
  var hop = { a: HOP_TAM, b: HOP_TAM, c: HOP_TAM };
  assert.strictEqual(Q.mucTieu(ds, hop, 1.2), 'c');
  assert.strictEqual(Q.mucTieu(ds, hop, 2.6), 'b');
  assert.strictEqual(Q.mucTieu(ds2, hop, 1.2), 'b');
});
var HOP_TAM = { x: 100, y: 100, w: 100, h: 100 };

test('ban tay va may quay: gach chan (quay: false) bat dau muon hon van nhuong tay cho muc may quay dang nhin', function () {
  // Cảnh minh-hoa thật: tiêu đề viết 0,55–1,4 s, gạch chân 1,4–1,8 s, hình đầu tiên vẽ 1,0–2,8 s.
  var ds = [muc('a', 0.55, 0.85), muc('b', 1.4, 0.4, { kieu: 'net', quay: false }), muc('c', 1.0, 1.8, { kieu: 'hinh' })];
  var hop = { a: HOP_TAM, b: HOP_TAM, c: HOP_TAM };
  var v = T.viTri(ds, 1.6, ngoi, NGHI);
  assert.strictEqual(Q.mucTieu(ds, hop, 1.6), 'c');
  assert.ok(gan(v.x, ngoi(ds[2], 0.6 / 1.8).x) && gan(v.y, ngoi(ds[2], 0.6 / 1.8).y), 'tay o hinh may quay dang nhin: ' + JSON.stringify(v));
  // Chỉ còn gạch chân đang vẽ thì tay vẫn vẽ gạch chân.
  var ds2 = [muc('a', 0.55, 0.85), muc('b', 1.4, 0.4, { kieu: 'net', quay: false })];
  var v2 = T.viTri(ds2, 1.6, ngoi, NGHI);
  assert.ok(gan(v2.x, ngoi(ds2[1], 0.5).x), 'chi con gach chan: ' + JSON.stringify(v2));
  // Mọi t: mục máy quay nhắm, nếu đang vẽ, chính là mục tay đang vẽ.
  for (var t = 0; t < 3; t += 1 / 30) {
    var id = Q.mucTieu(ds, hop, t);
    var m = ds.filter(function (x) { return x.id === id; })[0];
    if (!m || t < m.batDau || t >= m.batDau + m.thoiLuong) { continue; }
    var p = (t - m.batDau) / m.thoiLuong;
    var tay = T.viTri(ds, t, ngoi, NGHI);
    assert.ok(gan(tay.x, ngoi(m, p).x) && gan(tay.y, ngoi(m, p).y), 't = ' + t + ': ' + JSON.stringify(tay));
  }
});

test('ban tay: bo qua muc dong va muc anh', function () {
  var ds = [muc('a', 1, 1), muc('b', 1.2, 1, { dong: true }), muc('c', 1.3, 0.4, { kieu: 'anh' })];
  var v = T.viTri(ds, 1.5, ngoi, NGHI);
  assert.ok(gan(v.x, 200) && gan(v.y, 100), JSON.stringify(v));
});

test('ban tay: ranh 2 s thi nghi o goc va an', function () {
  var ds = [muc('a', 1, 1), muc('c', 5, 1)];
  var v = T.viTri(ds, 3.5, ngoi, NGHI);
  assert.strictEqual(v.hien, false);
  assert.ok(gan(v.x, NGHI.x) && gan(v.y, NGHI.y));
  var luot = T.viTri(ds, 2.2, ngoi, NGHI);
  assert.strictEqual(luot.hien, true, 'dang luot ve goc nghi');
  assert.ok(luot.x > 300 && luot.x < NGHI.x);
  assert.strictEqual(T.viTri(ds, 2.4, ngoi, NGHI).hien, false, 'luot ve nghi xong trong 0,4 s');
});

test('ban tay: truoc muc dau tien tay luot vao tu goc nghi, bat dau khong som hon t = 0', function () {
  var ds = [muc('a', 0.2, 1)];
  var v0 = T.viTri(ds, 0, ngoi, NGHI);
  assert.ok(gan(v0.x, NGHI.x) && gan(v0.y, NGHI.y), JSON.stringify(v0));
  var v1 = T.viTri(ds, 0.1, ngoi, NGHI);
  assert.ok(gan(v1.x, (NGHI.x + 100) / 2) && gan(v1.y, (NGHI.y + 100) / 2), JSON.stringify(v1));
  var ds2 = [muc('a', 3, 1)];
  assert.strictEqual(T.viTri(ds2, 1, ngoi, NGHI).hien, false);
  var v2 = T.viTri(ds2, 2.8, ngoi, NGHI);
  assert.ok(v2.hien && gan(v2.x, NGHI.x + (100 - NGHI.x) * 0.5));
});

test('ban tay: lau bang dung gie, x tang dan theo t, y = 380', function () {
  var ds = [muc('a', 0.55, 1)];
  var truoc = -Infinity;
  for (var i = 0; i <= 10; i++) {
    var t = 0.5 * i / 10;
    var v = T.viTri(ds, t, ngoi, NGHI, { lauBang: true });
    assert.strictEqual(v.kieu, 'gie');
    assert.strictEqual(v.hien, true);
    assert.strictEqual(v.y, 380);
    assert.ok(gan(v.x, -120 + 1400 * t / 0.5));
    assert.ok(v.x > truoc);
    truoc = v.x;
  }
  assert.strictEqual(T.viTri(ds, 0.3, ngoi, NGHI).kieu, 'but', 'khong lau bang thi khong co gie');
});

test('ban tay: do dai lau bang theo tuy.giayLau (tu du.giayLauBang)', function () {
  var ds = [muc('a', 1.05, 1)];
  var v = T.viTri(ds, 0.8, ngoi, NGHI, { lauBang: true, giayLau: 1.0 });
  assert.strictEqual(v.kieu, 'gie');
  assert.ok(gan(v.x, -120 + 1400 * 0.8 / 1.0));
  assert.strictEqual(T.viTri(ds, 0.8, ngoi, NGHI, { lauBang: true }).kieu, 'but', 'mac dinh 0,5 giay');
});

test('ban tay: o khung cuoi canh ngan tay da nghi', function () {
  var gh = 2.5667;
  var ds = [muc('a', 0.55, gh - 0.2 - 0.55)];
  var v = T.viTri(ds, gh - 1 / 30, ngoi, NGHI, { lauBang: true, gh: gh });
  assert.strictEqual(v.hien, false);
  assert.strictEqual(T.viTri(ds, gh - 0.2, ngoi, NGHI, { gh: gh }).hien, false);
});

// ---------- máy quay ----------

var HOP = {
  a: { x: 60, y: 30, w: 1160, h: 116 },
  b: { x: 900, y: 240, w: 300, h: 300 },
  c: { x: 124, y: 400, w: 400, h: 40 },
  d: { x: 480, y: 560, w: 200, h: 58 }
};
var DS = [muc('a', 0.55, 1.2), muc('b', 1.0, 1.5), muc('c', 3, 1), muc('d', 3.3, 0.8), muc('e', 4, 1, { dong: true })];
var GH = 7;

function bienDoi(s, h) { return { x: s.z * h.x + s.tx, y: s.z * h.y + s.ty, w: s.z * h.w, h: s.z * h.h }; }
function phuKin(s) {
  return s.tx <= 1e-6 && s.ty <= 1e-6 && s.tx + 1280 * s.z >= 1280 - 1e-6 && s.ty + 720 * s.z >= 720 - 1e-6;
}

test('may quay: xac dinh theo t', function () {
  for (var t = 0; t < GH; t += 0.173) {
    assert.deepStrictEqual(Q.tinh(DS, HOP, t, GH, {}), Q.tinh(DS, HOP, t, GH, {}));
  }
});

test('may quay: z trong [1; 1,35] tai 200 diem va ve {1,0,0} tai cuoi canh', function () {
  for (var i = 0; i < 200; i++) {
    var s = Q.tinh(DS, HOP, GH * i / 199, GH, {});
    assert.ok(s.z >= 1 - 1e-9 && s.z <= 1.35 + 1e-9, 'z = ' + s.z);
  }
  assert.deepStrictEqual(Q.tinh(DS, HOP, GH, GH, {}), { z: 1, tx: 0, ty: 0 });
  assert.deepStrictEqual(Q.tinh(DS, HOP, GH - 0.2, GH, {}), { z: 1, tx: 0, ty: 0 });
});

test('may quay: moi diem lop bang phu kin khung va hop muc tieu nam trong khung, day <= 620', function () {
  var coPhong = false;
  for (var i = 0; i < 400; i++) {
    var t = GH * i / 399;
    var s = Q.tinh(DS, HOP, t, GH, {});
    assert.ok(phuKin(s), 't=' + t + ' ' + JSON.stringify(s));
    var id = Q.mucTieu(DS, HOP, t);
    if (!id) { continue; }
    var b = bienDoi(s, HOP[id]);
    assert.ok(b.y + b.h <= 620 + 1e-6, 't=' + t + ' ' + id + ' day ' + (b.y + b.h));
    assert.ok(b.x >= -1e-6 && b.y >= -1e-6 && b.x + b.w <= 1280 + 1e-6, 't=' + t + ' ' + id + ' ' + JSON.stringify(b));
    if (s.z > 1.05) { coPhong = true; }
  }
  assert.ok(coPhong, 'may quay phai co luc phong vao muc nho');
});

test('may quay: muc tieu la muc dang ve, khong co thi muc vua xong; bo qua muc dong', function () {
  assert.strictEqual(Q.mucTieu(DS, HOP, 0.2), null);
  assert.strictEqual(Q.mucTieu(DS, HOP, 0.8), 'a');
  assert.strictEqual(Q.mucTieu(DS, HOP, 1.2), 'b');
  assert.strictEqual(Q.mucTieu(DS, HOP, 2.8), 'b');
  assert.strictEqual(Q.mucTieu(DS, HOP, 3.5), 'd');
  assert.strictEqual(Q.mucTieu(DS, HOP, 4.5), 'd');
  assert.strictEqual(Q.mucTieu(DS, HOP, 4.05), 'd', 'c xong truoc, d con dang ve');
});

test('may quay: hop sat vach phu de van duoc nang len tren 620', function () {
  var hop = { y: { x: 124, y: 556, w: 700, h: 72 } };
  var ds = [muc('y', 1, 1)];
  var s = Q.tinh(ds, hop, 3, 8, {});
  var b = bienDoi(s, hop.y);
  assert.ok(b.y + b.h <= 620 + 1e-6, JSON.stringify(s));
  assert.ok(phuKin(s));
});

test('may quay: canh thi nghiem chi day cham tu 1 len 1,06, giu tam', function () {
  var s0 = Q.tinh(DS, HOP, 0, GH, { day: true });
  var s1 = Q.tinh(DS, HOP, GH, GH, { day: true });
  assert.ok(gan(s0.z, 1));
  assert.ok(gan(s1.z, 1.06, 1e-6));
  assert.ok(gan(s1.tx, 640 * (1 - s1.z), 1e-6) && gan(s1.ty, 360 * (1 - s1.z), 1e-6));
  assert.ok(phuKin(s1));
});

test('may quay: tat may quay thi luon {1,0,0}', function () {
  [0, 1.2, 3.5, GH].forEach(function (t) {
    assert.deepStrictEqual(Q.tinh(DS, HOP, t, GH, { mayQuay: false }), { z: 1, tx: 0, ty: 0 });
  });
});

test('may quay: chuyen muc tieu muot, khong nhay qua 0,1 giua hai khung 1/30 s ngoai luc bat buoc', function () {
  var truoc = Q.tinh(DS, HOP, 0, GH, {});
  for (var i = 1; i <= GH * 30; i++) {
    var s = Q.tinh(DS, HOP, i / 30, GH, {});
    assert.ok(Math.abs(s.z - truoc.z) < 0.1, 't=' + i / 30 + ' dz=' + (s.z - truoc.z));
    truoc = s;
  }
});

test('may quay: canh ngan 2,5667 s co hinh va o chu day 628 van giu rang buoc, ve dung {1,0,0}', function () {
  var gh = 2.5667;
  var hop = { hinh: { x: 910, y: 240, w: 300, h: 300 }, y: { x: 124, y: 556, w: 700, h: 72 } };
  var ds = [muc('hinh', 0.55, 1.2, { kieu: 'hinh' }), muc('y', 1.0, gh - 0.2 - 1.0)];
  var veNha = gh - 1.2;
  for (var i = 0; i <= 240; i++) {
    var t = gh * i / 240;
    var s = Q.tinh(ds, hop, t, gh, {});
    assert.ok(phuKin(s), 't=' + t + ' ' + JSON.stringify(s));
    assert.ok(s.z >= 1 - 1e-9 && s.z <= 1.35 + 1e-9, 't=' + t + ' z=' + s.z);
    var id = Q.mucTieu(ds, hop, t);
    if (id && t < veNha) {
      var b = bienDoi(s, hop[id]);
      assert.ok(b.y + b.h <= 620 + 1e-6, 't=' + t + ' ' + id + ' day ' + (b.y + b.h));
    }
  }
  assert.deepStrictEqual(Q.tinh(ds, hop, gh - 1 / 30, gh, {}), { z: 1, tx: 0, ty: 0 });
  assert.deepStrictEqual(Q.tinh(ds, hop, gh - 0.2, gh, {}), { z: 1, tx: 0, ty: 0 });
});

test('ban tay va may quay: cung ranh gioi bat dau (t = batDau la dang ve)', function () {
  var ds = [muc('a', 1, 1)];
  var v = T.viTri(ds, 1, ngoi, NGHI);
  assert.ok(v.hien && gan(v.x, 100) && gan(v.y, 100), JSON.stringify(v));
  assert.strictEqual(T.dangVe(ds, 1), ds[0]);
  assert.strictEqual(Q.mucTieu(ds, { a: HOP_TAM }, 1), 'a');
});
