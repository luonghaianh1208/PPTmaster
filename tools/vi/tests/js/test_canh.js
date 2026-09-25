'use strict';
var test = require('node:test');
var assert = require('node:assert');
var path = require('node:path');
var fs = require('node:fs');

var RT = path.join(__dirname, '..', '..', 'video_ma_parts', 'runtime');
var TN = path.join(__dirname, '..', '..', 'thi_nghiem_parts');
require(path.join(TN, 'runtime', 'khung.js'));
require(path.join(TN, 'mo_hinh', 'li-con-lac-don.js'));
require(path.join(RT, 'khung-video.js'));
var LOAI = ['tieu-de', 'khai-niem', 'cong-thuc', 'y-tung-y', 'quy-trinh', 'so-sanh', 'do-thi', 'thi-nghiem'];
LOAI.forEach(function (l) { require(path.join(RT, 'canh', l + '.js')); });
var V = globalThis.THI_VIDEO;
var C = globalThis.THI_CANH;
var KHAI_BAO = JSON.parse(fs.readFileSync(path.join(TN, 'mo_hinh', 'li-con-lac-don.json'), 'utf8'));

function du(loai, thoiLuong, moc, truong, them) {
  var d = { so: 1, loai: loai, thoiLuong: thoiLuong, danDau: 0.7, moc: moc, truong: truong };
  Object.keys(them || {}).forEach(function (k) { d[k] = them[k]; });
  return d;
}
function tatCa(gh) {
  var s = gh === 2.5;
  return {
    'tieu-de': du('tieu-de', gh, [], { chu: ['Chuẩn độ acid – base'], phu: ['Hoá học 11'] }),
    'khai-niem': du('khai-niem', gh, [], { 'thuat-ngu': ['Chuẩn độ'], 'dinh-nghia': ['Xác định nồng độ một dung dịch.'] }),
    'cong-thuc': du('cong-thuc', gh, s ? [0.7, 1.0] : [2.5, 4.5], { 'bieu-thuc': ['C~M~ = n/V'], 'giai-thich': ['n là số mol', 'V là thể tích'] }),
    'y-tung-y': du('y-tung-y', gh, s ? [0.7, 1.0, 1.3] : [0.7, 3, 5.5], { 'tieu-de': ['Ba bước'], y: ['a', 'b', 'c'] }),
    'quy-trinh': du('quy-trinh', gh, s ? [0.7, 1.0, 1.3] : [0.7, 3.5, 6], { 'tieu-de': ['Quy trình'], buoc: ['Đo', 'Nhỏ', 'Dừng'] }),
    'so-sanh': du('so-sanh', gh, s ? [0.7, 0.9, 1.1, 1.3] : [0.7, 2.5, 4.5, 6.5],
      { 'tieu-de': ['So sánh'], trai: ['Acid'], phai: ['Base'], 'y-trai': ['a', 'b'], 'y-phai': ['c', 'd'] }),
    'do-thi': du('do-thi', gh, s ? [0.7, 1.0, 1.3] : [0.7, 3, 5.5], { 'tieu-de': ['Đồ thị'], 'truc-ngang': ['t'], 'truc-doc': ['v'], diem: ['0, 0', '1, 2', '2, 4'] },
      { diem: [[0, 0], [1, 2], [2, 4]] }),
    'thi-nghiem': du('thi-nghiem', gh, [], { mau: ['li-con-lac-don'] },
      { khaiBao: KHAI_BAO, thamSo: { 'chieu-dai': [[0, 0.4], [6, 1.6]] }, do: ['chu-ki'] })
  };
}
function hienThi(html) {
  return html.replace(/<span class="an">.*?<\/span>/g, '').replace(/<[^>]+>/g, '');
}

test('catDanhDau: che chu, giu the dong mo va thoat ky tu dac biet', function () {
  assert.strictEqual(V.catDanhDau('**ab**c', 1), '<b>a<span class="an">b</span></b><span class="an">c</span>');
  assert.strictEqual(V.catDanhDau('H~2~O', 99), 'H<sub>2</sub>O');
  assert.strictEqual(V.catDanhDau('m/s^2^', 99), 'm/s<sup>2</sup>');
  assert.strictEqual(V.catDanhDau('a < b & "c"', 99), 'a &lt; b &amp; &quot;c&quot;');
  var doc = V.catDanhDau('</script><img src=x onerror=alert(1)>', 99);
  assert.ok(doc.indexOf('<img') === -1 && doc.indexOf('</script') === -1, doc);
  assert.strictEqual(hienThi(V.catDanhDau('**ab**c', 2)), 'ab');
  assert.strictEqual(hienThi(V.catDanhDau('Nhờ ướt', 3)), 'Nhờ');
  assert.strictEqual(V.catDanhDau('abc', 0), '<span class="an">abc</span>');
});

test('demKyTu bo dau danh dau; tienDo kep 0..1', function () {
  assert.strictEqual(V.demKyTu('H~2~SO~4~ **x**'), 'H2SO4 x'.length);
  assert.strictEqual(V.tienDo(0, 1, 2), 0);
  assert.strictEqual(V.tienDo(2, 1, 2), 0.5);
  assert.strictEqual(V.tienDo(9, 1, 2), 1);
  assert.strictEqual(V.tienDo(1, 1, 0), 1);
});

test('duong ve la xac dinh theo hat giong', function () {
  assert.strictEqual(V.duongQua([[0, 0], [100, 0]], 5), V.duongQua([[0, 0], [100, 0]], 5));
  assert.notStrictEqual(V.duongQua([[0, 0], [100, 0]], 5), V.duongQua([[0, 0], [100, 0]], 6));
});

test('moi loai canh: muc xac dinh, t=0 chua hien gi, cuoi canh hien du', function () {
  [12, 2.5].forEach(function (gh) {
    var bo = tatCa(gh);
    LOAI.forEach(function (l) {
      var a = C[l].muc(bo[l]);
      var b = C[l].muc(bo[l]);
      assert.deepStrictEqual(a, b, l);
      assert.ok(a.length > 0, l);
      a.forEach(function (m) {
        if (m.dong) { return; }
        assert.ok(m.batDau > 0, l + ':' + m.id + ' phai bat dau sau t=0');
        assert.strictEqual(V.tienDo(0, m.batDau, m.thoiLuong), 0, l + ':' + m.id);
        assert.ok(m.batDau + m.thoiLuong <= gh - 0.2 + 1e-9, l + ':' + m.id + ' phai xong truoc cuoi canh ' + gh);
        assert.strictEqual(V.tienDo(gh, m.batDau, m.thoiLuong), 1, l + ':' + m.id);
      });
      var ids = a.map(function (m) { return m.id; });
      assert.strictEqual(new Set(ids).size, ids.length, l + ' co id trung');
    });
  });
});

test('moi loai canh chua trong day 90 px cho phu de, ke ca khi du so dong toi da', function () {
  var DAY = 630;
  var bo = tatCa(12);
  var moc = function (n) { return Array.apply(null, Array(n)).map(function (_, k) { return 0.7 + k; }); };
  var nhieu = function (n, chu) { return Array.apply(null, Array(n)).map(function () { return chu; }); };
  bo['y-tung-y'] = du('y-tung-y', 12, moc(6), { 'tieu-de': ['Sáu ý'], y: nhieu(6, 'ý') });
  bo['cong-thuc'] = du('cong-thuc', 12, moc(4), { 'bieu-thuc': ['T = 2π√(l/g)'], 'giai-thich': nhieu(4, 'g') });
  bo['so-sanh'] = du('so-sanh', 12, moc(8), { 'tieu-de': ['S'], trai: ['A'], phai: ['B'], 'y-trai': nhieu(4, 'a'), 'y-phai': nhieu(4, 'b') });
  bo['quy-trinh'] = du('quy-trinh', 12, moc(5), { 'tieu-de': ['Q'], buoc: nhieu(5, 'b') });
  bo['thi-nghiem'].thamSo = { 'chieu-dai': [[0, 0.4]], g: [[0, 9.8]], 'khoi-luong': [[0, 0.2]] };
  bo['thi-nghiem'].do = ['chu-ki', 'thoi-gian-10-dao-dong'];
  LOAI.forEach(function (l) {
    C[l].muc(bo[l]).forEach(function (m) {
      var day = m.kieu === 'net'
        ? Math.max.apply(null, m.d.match(/-?\d+(\.\d+)?/g).map(Number).filter(function (_, i) { return i % 2 === 1; }))
        : m.y + m.cao;
      assert.ok(day <= DAY, l + ':' + m.id + ' xuong toi ' + day);
    });
  });
});

test('y k hien dung tai moc cau k', function () {
  var bo = tatCa(12);
  var ids = function (l, tienTo) {
    var muc = C[l].muc(bo[l]);
    return [0, 1, 2].map(function (k) { return muc.filter(function (m) { return m.id === tienTo + k; })[0]; });
  };
  ids('y-tung-y', 'y-').forEach(function (m, k) { assert.strictEqual(m.batDau, bo['y-tung-y'].moc[k]); });
  ids('quy-trinh', 'buoc-').forEach(function (m, k) { assert.strictEqual(m.batDau, bo['quy-trinh'].moc[k]); });
  ids('do-thi', 'diem-').forEach(function (m, k) { assert.strictEqual(m.batDau, bo['do-thi'].moc[k]); });
  var ss = C['so-sanh'].muc(bo['so-sanh']);
  var at = function (id) { return ss.filter(function (m) { return m.id === id; })[0].batDau; };
  assert.strictEqual(at('y-trai-0'), 0.7);
  assert.strictEqual(at('y-trai-1'), 2.5);
  assert.strictEqual(at('y-phai-0'), 4.5);
  assert.strictEqual(at('y-phai-1'), 6.5);
  var ct = C['cong-thuc'].muc(bo['cong-thuc']);
  var bt = ct.filter(function (m) { return m.id === 'bieu-thuc'; })[0];
  ct.filter(function (m) { return /^giai-thich-/.test(m.id); }).forEach(function (m, k) {
    assert.ok(m.batDau >= bo['cong-thuc'].moc[k]);
    assert.ok(m.batDau >= bt.batDau + bt.thoiLuong);
  });
});

test('do thi: diem nam trong vung ve, cung y thi vao giua', function () {
  var muc = C['do-thi'].muc(tatCa(12)['do-thi']);
  var vong = muc.filter(function (m) { return /^diem-/.test(m.id); });
  assert.strictEqual(vong.length, 3);
  vong.forEach(function (m) {
    var so = m.d.match(/-?\d+(\.\d+)?/g).map(Number);
    var cx = so[0] + 9;
    assert.ok(cx >= 190 - 1 && cx <= 1090 + 1, 'x ngoai vung: ' + cx);
  });
  var bang = tatCa(12)['do-thi'];
  bang.diem = [[0, 5], [1, 5]];
  var muc2 = C['do-thi'].muc(bang).filter(function (m) { return /^diem-/.test(m.id); });
  var y = muc2.map(function (m) { return Number(m.d.match(/-?\d+(\.\d+)?/g)[1]); });
  assert.strictEqual(y[0], y[1]);
  assert.ok(y[0] > 270 && y[0] < 560);
});

test('thi nghiem: noi suy tham so theo moc thoi gian', function () {
  var d = tatCa(12)['thi-nghiem'];
  var T = C['thi-nghiem'];
  assert.strictEqual(T.thamSoTai(d, 0)['chieu-dai'], 0.4);
  assert.ok(Math.abs(T.thamSoTai(d, 3)['chieu-dai'] - 1.0) < 1e-9);
  assert.strictEqual(T.thamSoTai(d, 6)['chieu-dai'], 1.6);
  assert.strictEqual(T.thamSoTai(d, 100)['chieu-dai'], 1.6);
  assert.strictEqual(T.thamSoTai(d, 3)['g'], 9.8);
  d.thamSo = { 'chieu-dai': [[2, 0.4]] };
  assert.strictEqual(T.thamSoTai(d, 1)['chieu-dai'], 1.0);
  assert.strictEqual(T.thamSoTai(d, 2)['chieu-dai'], 0.4);
});

test('thi nghiem: so do la so cua tinh() cua mo hinh, o moi thoi diem', function () {
  var d = tatCa(12)['thi-nghiem'];
  var T = C['thi-nghiem'];
  var M = globalThis.THI_NGHIEM_MO_HINH;
  [0, 1.5, 3, 5, 6, 11].forEach(function (t) {
    var g = T.giaTri(d, t);
    assert.strictEqual(g.dai['chu-ki'], M.tinh(g.tham)['chu-ki']);
  });
  assert.ok(Math.abs(T.giaTri(d, 3).dai['chu-ki'] - 2.007089923) < 5e-4);
});

test('thi nghiem: mo hinh mot-lan chay lai tu moc tham-so gan nhat, mo hinh khac tinh tu danDau', function () {
  var T = C['thi-nghiem'];
  var d = { danDau: 0.7, khaiBao: { hoatHinh: 'mot-lan' },
    thamSo: { 'goc-nem': [[0, 30], [8, 60]], 'van-toc': [[12, 15]] } };
  assert.strictEqual(T.mocGanNhat(d, 5), 0.7);
  assert.strictEqual(T.mocGanNhat(d, 9), 8);
  assert.strictEqual(T.mocGanNhat(d, 13), 12);
  assert.ok(Math.abs(T.thoiGianMoHinh(d, 5) - 4.3) < 1e-9);
  assert.strictEqual(T.thoiGianMoHinh(d, 8), 0);
  assert.strictEqual(T.thoiGianMoHinh(d, 10), 2);
  assert.strictEqual(T.thoiGianMoHinh(d, 14), 2);
  assert.strictEqual(T.thoiGianMoHinh(d, 0.2), 0);
  d.khaiBao.hoatHinh = 'lap';
  assert.ok(Math.abs(T.thoiGianMoHinh(d, 14) - 13.3) < 1e-9);
});

test('thi nghiem: khong co dong chu nao bat dau tai t=0 ngoai dong so do', function () {
  var muc = C['thi-nghiem'].muc(tatCa(12)['thi-nghiem']);
  var dong = muc.filter(function (m) { return m.dong; });
  assert.ok(dong.length >= 2);
});
