# Thí nghiệm ảo cho Toán, Vật lí, Hoá học — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Thêm loại việc thứ 9 cho PPT Master bản Việt: từ `thi-nghiem.md`, công cụ `tools/vi/thi_nghiem.py` ghép ra một file HTML thí nghiệm ảo chạy không cần mạng, một phiếu học tập Word và `can-soat.md`, với 8 mô hình Toán – Lí – Hoá đã kiểm bằng số.

**Architecture:** Máy dựng từ linh kiện. Khung chạy chung (`runtime/khung.js`, `khung.css`) lo giao diện, ba bước Dự đoán – Quan sát – Giải thích, bảng số liệu, đồ thị, nhiễu đo và tự kiểm. Mỗi mô hình là hai file: `<mã>.json` (tham số, đại lượng đo, công thức, điều kiện, bảng số kiểm) và `<mã>.js` (hàm `tinh`, `ve`). Python đọc `thi-nghiem.md`, kiểm nó với khai báo JSON, chạy bảng số kiểm qua Node khi máy có, rồi nhúng tất cả vào một file HTML. Một bản tính lại độc lập bằng Python (`tham_chieu.py`) dùng cho test đối chiếu và cho bảng số liệu lí tưởng của phiếu.

**Tech Stack:** Python 3.10+ thư viện chuẩn, `python-docx` (đã có trong `tools/vi/requirements-vi.txt`), `tools/vi/word_parts/` sẵn có; JavaScript ES5 thuần không thư viện; Node chỉ dùng để kiểm số và chạy test (không bắt buộc trên máy thầy cô); `unittest`.

**Spec:** `docs/vi/phat-trien/2026-09-20-thi-nghiem-ao-design.md`

## Global Constraints

- Không sửa gì trong `skills/`, `LICENSE`, `SPONSORS*.md`, `requirements.txt` gốc, `skills/ppt-master/requirements.txt`; không đụng `skills/ppt-master/scripts/attribution_guard.py`.
- Không commit gì trong `projects/`; không tạo hay commit `.env`; không force push.
- Python của công cụ chỉ dùng thư viện chuẩn, cộng `python-docx` cho phiếu (import muộn). Không thêm dòng nào vào file requirements.
- JavaScript thuần, không thư viện ngoài. File HTML ghép ra không được chứa `http://`, `https://`, `//cdn`, `@import`, `url(`, `<link`, `src=`.
- Font: `"Segoe UI", Arial, sans-serif`. Quy ước chữ: `H~2~SO~4~`, `m/s^2^`, `**in đậm**`.
- Mọi công cụ dòng lệnh in **đúng một dòng JSON** ra stdout ở mọi nhánh, mã thoát 0 hoặc 1; tiến trình đi ra stderr.
- Tài liệu trong `docs/vi/` và chuỗi hiển thị viết tiếng Việt; tên hàm Python tiếng Anh, tên trong JavaScript và khoá JSON tiếng Việt không dấu theo đúng mã trong plan. File trong `docs/vi/tro-ly/` không có link Markdown.
- Chạy lệnh từ thư mục gốc repo bằng `venv\Scripts\python.exe`. Test JavaScript cần Node; máy không có Node thì các test đó tự bỏ qua, nhưng máy thi công plan này **phải có Node** (`node --version` ≥ 18).
- Mã trong plan đã được dựng thử và chạy qua test trước khi viết plan. Chép nguyên văn; nếu một test trượt, tìm lỗi chép trước khi sửa mã.
- Commit theo Conventional Commits, tiếng Anh, kết thúc bằng dòng `Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>`.
- `gh` luôn kèm `--repo luonghaianh1208/PPTmaster`.

## File Structure

```
tools/vi/thi_nghiem.py                         CLI: kiểm → ghép HTML → phiếu → can-soat.md → một dòng JSON
tools/vi/thi_nghiem_parts/__init__.py          (rỗng)
tools/vi/thi_nghiem_parts/tham_chieu.py        8 hàm tính lại bằng Python + mulberry32
tools/vi/thi_nghiem_parts/thu_vien.py          Model, load(), check_declaration(), check_js(), list_models()
tools/vi/thi_nghiem_parts/kiem_so.py           find_node(), bang_kiem(), luoi(), CheckError
tools/vi/thi_nghiem_parts/parse.py             read_meta(), parse_experiment(), Experiment.config(), ParseError
tools/vi/thi_nghiem_parts/build_html.py        build(), write(), BuildError
tools/vi/thi_nghiem_parts/phieu.py             build(), ideal_table()
tools/vi/thi_nghiem_parts/runtime/khung.js     logic thuần + giao diện
tools/vi/thi_nghiem_parts/runtime/khung.css
tools/vi/thi_nghiem_parts/runtime/chay_node.js chạy mô hình trong vm của Node
tools/vi/thi_nghiem_parts/mo_hinh/<mã>.json|js 8 mô hình
tools/vi/tests/js/test_khung.js                test Node cho logic khung
tools/vi/tests/thi_nghiem_mau.py               dữ liệu mẫu dùng chung cho test
tools/vi/tests/test_thi_nghiem_khung.py        chạy test Node
tools/vi/tests/test_thi_nghiem_tham_chieu.py   định luật
tools/vi/tests/test_thi_nghiem_mo_hinh.py      khuôn mô hình, đối chiếu JS–Python
tools/vi/tests/test_thi_nghiem_parse.py
tools/vi/tests/test_thi_nghiem_html.py
tools/vi/tests/test_thi_nghiem_phieu.py
tools/vi/tests/test_thi_nghiem_cong_cu.py
docs/vi/tro-ly/thi-nghiem-ao.md                loại việc thứ 9 (cho AI)
docs/vi/tro-ly/mo-hinh-thi-nghiem.md           danh mục mẫu + khuôn mô hình mới (cho AI)
docs/vi/thi-nghiem-ao.md                       tài liệu cho thầy cô
```

Giao ước giữa JavaScript và Python (mọi task dùng đúng các tên này):

- Toàn cục JS: `THI_NGHIEM_KHUNG` (khung), `THI_NGHIEM_MO_HINH` (mô hình đang nạp).
- Mô hình JS: `tinh(p) -> {mã_đại_lượng: số|null}`, `ve(ctx, p, t, kt, d)`, tuỳ chọn `thoiLuong(p, d)`.
- Khai báo JSON: `ma`, `ten`, `mon`, `hoatHinh` ∈ `khong|mot-lan|lap`, `thamSo[]`, `daiLuongDo[]`, `congThuc{bieuThuc,dieuKien,nguon}`, `bangKiem[]{vao,ra,saiSo}`.
- Cấu hình nhúng vào HTML (`Experiment.config()`): `tieuDe, mon, lop, nguoiThaoTac, saiSo, thamSo{mã: {kieu: truot|chon|co-dinh, ...}}, duDoan{cau, luaChon[{ma, noiDung}], dapAn}, quanSat{soLanDo, cot[], doThi{tung{ma, phep}, hoanh{ma, phep}}|null}, giaiThich{cau, goiY}, ketLuan`.
- Phép biến đổi trục đồ thị: `khong`, `binh-phuong`, `nghich-dao`, `ln`, `can`.

---

### Task 1: Logic của khung chạy và bộ chạy Node

Phần logic thuần của khung: bộ sinh số có hạt giống, nhiễu chuẩn, tự kiểm mô hình, phép biến đổi, khớp tuyến tính, định dạng số kiểu Việt, đánh dấu chỉ số, lấy mẫu đo, máy trạng thái ba bước. Chạy được trong Node nên test được. Giao diện sẽ thêm ở Task 8.

**Files:**
- Create: `tools/vi/thi_nghiem_parts/__init__.py` (rỗng)
- Create: `tools/vi/thi_nghiem_parts/runtime/khung.js`
- Create: `tools/vi/thi_nghiem_parts/runtime/chay_node.js`
- Create: `tools/vi/tests/js/test_khung.js`
- Create: `tools/vi/tests/test_thi_nghiem_khung.py`

**Interfaces:**
- Produces: `THI_NGHIEM_KHUNG.{taoNgauNhien(hatGiong) -> () => [0,1), nhieuChuan(rng), thamSoMacDinh(khaiBao), gopThamSo(khaiBao, vao), tuKiem(khaiBao, moHinh) -> {tong, dat, truot[{dong, ma, mong, duoc}]}, apDungPhep(phep, x) -> số|null, khopTuyenTinh(diem[[x,y]]) -> {heSoGoc, tungDoGoc, tuongQuan}|null, dinhDang(x, chuSo) -> chuỗi, danhDau(chu) -> HTML, mauDo(khaiBao, ketQua, rng, batSaiSo), taoNhiemVu(cauHinh), diemDoThi(doThi, lanDo)}`.
- Produces: `node chay_node.js <khung.js> <mo-hinh.js> <mo-hinh.json> bang-kiem|luoi` in một dòng JSON; chế độ `luoi` đọc danh sách tham số từ stdin; lỗi thì in `{"loi": "..."}` và thoát 1.

- [ ] **Step 1: Viết test Node**

`tools/vi/tests/js/test_khung.js`:

```js
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
```

`tools/vi/tests/test_thi_nghiem_khung.py`:

```python
"""Chạy bộ test Node của phần logic trong khung chạy thí nghiệm ảo."""

import shutil
import subprocess
import unittest
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
HAS_NODE = shutil.which("node") is not None
NEED_NODE = "máy không có Node nên không chạy được JavaScript"


@unittest.skipUnless(HAS_NODE, NEED_NODE)
class RuntimeLogicTest(unittest.TestCase):
    def test_runtime_logic_passes_its_node_tests(self):
        proc = subprocess.run([shutil.which("node"), "--test", str(TOOLS_VI / "tests" / "js" / "test_khung.js")],
                              capture_output=True, text=True, encoding="utf-8", errors="replace")
        self.assertEqual(proc.returncode, 0, proc.stdout[-2000:])


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Chạy để thấy trượt**

Run: `venv\Scripts\python.exe -m unittest tools.vi.tests.test_thi_nghiem_khung -v`
Expected: FAIL (Node báo không tìm thấy `khung.js`).

- [ ] **Step 3: Viết logic khung và bộ chạy Node**

Tạo file rỗng `tools/vi/thi_nghiem_parts/__init__.py`.

`tools/vi/thi_nghiem_parts/runtime/khung.js`:

```js
(function (root) {
  'use strict';

  // ---------- Logic thuần: chạy được cả trong trình duyệt lẫn Node ----------

  function taoNgauNhien(hatGiong) {
    var a = hatGiong | 0;
    return function () {
      a |= 0; a = a + 0x6D2B79F5 | 0;
      var t = Math.imul(a ^ a >>> 15, 1 | a);
      t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t;
      return ((t ^ t >>> 14) >>> 0) / 4294967296;
    };
  }

  function nhieuChuan(rng) {
    var u = 1 - rng();
    var v = rng();
    return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v);
  }

  function thamSoMacDinh(khaiBao) {
    var p = {};
    khaiBao.thamSo.forEach(function (ts) { p[ts.ma] = ts.macDinh; });
    return p;
  }

  function gopThamSo(khaiBao, vao) {
    var p = thamSoMacDinh(khaiBao);
    Object.keys(vao || {}).forEach(function (ma) { p[ma] = vao[ma]; });
    return p;
  }

  function tuKiem(khaiBao, moHinh) {
    var truot = [];
    var dongTruot = {};
    khaiBao.bangKiem.forEach(function (dong, chiSo) {
      var ketQua;
      try {
        ketQua = moHinh.tinh(gopThamSo(khaiBao, dong.vao));
      } catch (loi) {
        truot.push({ dong: chiSo + 1, ma: '*', mong: null, duoc: String(loi) });
        dongTruot[chiSo] = true;
        return;
      }
      Object.keys(dong.ra).forEach(function (ma) {
        var mong = dong.ra[ma];
        var duoc = ketQua[ma];
        var dat = mong === null
          ? duoc === null
          : typeof duoc === 'number' && isFinite(duoc) && Math.abs(duoc - mong) <= dong.saiSo;
        if (!dat) {
          truot.push({ dong: chiSo + 1, ma: ma, mong: mong, duoc: duoc === undefined ? null : duoc });
          dongTruot[chiSo] = true;
        }
      });
    });
    var tong = khaiBao.bangKiem.length;
    return { tong: tong, dat: tong - Object.keys(dongTruot).length, truot: truot };
  }

  function apDungPhep(phep, x) {
    if (x === null || x === undefined || !isFinite(x)) { return null; }
    var y;
    if (phep === 'binh-phuong') { y = x * x; }
    else if (phep === 'nghich-dao') { y = x === 0 ? NaN : 1 / x; }
    else if (phep === 'ln') { y = x > 0 ? Math.log(x) : NaN; }
    else if (phep === 'can') { y = x >= 0 ? Math.sqrt(x) : NaN; }
    else { y = x; }
    return isFinite(y) ? y : null;
  }

  function khopTuyenTinh(diem) {
    var n = diem.length;
    if (n < 2) { return null; }
    var tx = 0, ty = 0;
    diem.forEach(function (d) { tx += d[0]; ty += d[1]; });
    tx /= n; ty /= n;
    var sxx = 0, sxy = 0, syy = 0;
    diem.forEach(function (d) {
      sxx += (d[0] - tx) * (d[0] - tx);
      sxy += (d[0] - tx) * (d[1] - ty);
      syy += (d[1] - ty) * (d[1] - ty);
    });
    if (sxx === 0) { return null; }
    var heSoGoc = sxy / sxx;
    return {
      heSoGoc: heSoGoc,
      tungDoGoc: ty - heSoGoc * tx,
      tuongQuan: syy === 0 ? 1 : sxy / Math.sqrt(sxx * syy)
    };
  }

  function dinhDang(x, chuSo) {
    if (x === null || x === undefined || typeof x !== 'number' || !isFinite(x)) { return '—'; }
    var tron = Number(x.toFixed(chuSo));
    return (tron === 0 ? 0 : tron).toFixed(chuSo).replace('.', ',');
  }

  function danhDau(chu) {
    var sach = String(chu)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    return sach
      .replace(/\*\*(.+?)\*\*/g, '<b>$1</b>')
      .replace(/~([^~]+)~/g, '<sub>$1</sub>')
      .replace(/\^([^\^]+)\^/g, '<sup>$1</sup>');
  }

  function mauDo(khaiBao, ketQua, rng, batSaiSo) {
    var mau = {};
    khaiBao.daiLuongDo.forEach(function (d) {
      var thuc = ketQua[d.ma];
      if (thuc === null || thuc === undefined || !batSaiSo || !d.saiSo) { mau[d.ma] = thuc === undefined ? null : thuc; }
      else { mau[d.ma] = thuc + d.saiSo * nhieuChuan(rng); }
    });
    return mau;
  }

  function taoNhiemVu(cauHinh) {
    var nv = {
      buoc: 'du-doan',
      duDoan: null,
      lanDo: [],
      giaoVien: cauHinh.nguoiThaoTac === 'giao-vien'
    };
    nv.duocThaoTac = function () { return nv.giaoVien || nv.buoc !== 'du-doan'; };
    nv.chonDuDoan = function (giaTri) {
      if (nv.buoc !== 'du-doan' || giaTri === null || giaTri === undefined || String(giaTri).trim() === '') { return false; }
      nv.duDoan = String(giaTri).trim();
      nv.buoc = 'quan-sat';
      return true;
    };
    nv.ghiLanDo = function (dong) {
      if (!nv.duocThaoTac()) { return false; }
      nv.lanDo.push(dong);
      if (nv.buoc === 'quan-sat' && nv.lanDo.length >= cauHinh.quanSat.soLanDo) { nv.buoc = 'giai-thich'; }
      return true;
    };
    nv.xoaLanDo = function (chiSo) {
      if (chiSo < 0 || chiSo >= nv.lanDo.length) { return false; }
      nv.lanDo.splice(chiSo, 1);
      return true;
    };
    nv.duDoanDung = function () {
      if (nv.buoc !== 'giai-thich' || !cauHinh.duDoan.dapAn) { return null; }
      return nv.duDoan === cauHinh.duDoan.dapAn;
    };
    nv.batGiaoVien = function () { nv.giaoVien = true; };
    return nv;
  }

  function diemDoThi(doThi, lanDo) {
    var diem = [];
    lanDo.forEach(function (dong) {
      var x = apDungPhep(doThi.hoanh.phep, dong[doThi.hoanh.ma]);
      var y = apDungPhep(doThi.tung.phep, dong[doThi.tung.ma]);
      if (x !== null && y !== null) { diem.push([x, y]); }
    });
    return diem;
  }

  root.THI_NGHIEM_KHUNG = {
    taoNgauNhien: taoNgauNhien, nhieuChuan: nhieuChuan,
    thamSoMacDinh: thamSoMacDinh, gopThamSo: gopThamSo, tuKiem: tuKiem,
    apDungPhep: apDungPhep, khopTuyenTinh: khopTuyenTinh, dinhDang: dinhDang, danhDau: danhDau,
    mauDo: mauDo, taoNhiemVu: taoNhiemVu, diemDoThi: diemDoThi
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
```

`tools/vi/thi_nghiem_parts/runtime/chay_node.js`:

```js
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
```

- [ ] **Step 4: Chạy lại**

Run: `venv\Scripts\python.exe -m unittest tools.vi.tests.test_thi_nghiem_khung -v`
Expected: PASS (1 test; bên trong là 15 test Node).

- [ ] **Step 5: Commit**

```bash
git add tools/vi/thi_nghiem_parts/__init__.py tools/vi/thi_nghiem_parts/runtime tools/vi/tests/js tools/vi/tests/test_thi_nghiem_khung.py
git commit -m "feat(vi): add the pure logic of the virtual experiment runtime"
```

---

### Task 2: Bản tính lại bằng Python và test định luật

Tám hàm Python, mỗi hàm tính lại các đại lượng đo của một mô hình theo cách độc lập với JavaScript (ví dụ chuẩn độ acid yếu giải đa thức bậc ba theo [H⁺], cân bằng NO₂ giải theo độ phân li). `ngau_nhien` là mulberry32 khớp từng bit với `taoNgauNhien` của khung. Test ở task này không cần Node.

**Files:**
- Create: `tools/vi/thi_nghiem_parts/tham_chieu.py`
- Create: `tools/vi/tests/test_thi_nghiem_tham_chieu.py`

**Interfaces:**
- Produces: `tham_chieu.THAM_CHIEU: dict[str, Callable[[dict], dict]]` với 8 khoá `li-nem-xien`, `li-con-lac-don`, `li-mach-ohm`, `hoa-chuan-do`, `hoa-can-bang-no2`, `hoa-toc-do`, `toan-ham-so`, `toan-xac-suat`; và các hàm `li_nem_xien`, `li_con_lac_don`, `li_mach_ohm`, `hoa_chuan_do`, `hoa_can_bang_no2`, `hoa_toc_do`, `toan_ham_so`, `toan_xac_suat`, `ngau_nhien(hat_giong)`. Mỗi hàm nhận dict tham số đủ mọi mã, trả dict đại lượng đo (giá trị `None` khi không xác định).

- [ ] **Step 1: Viết test định luật**

`tools/vi/tests/test_thi_nghiem_tham_chieu.py`:

```python
"""Test định luật cho bản tính lại bằng Python của các mô hình thí nghiệm ảo."""

import math
import sys
import unittest
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from thi_nghiem_parts import tham_chieu  # noqa: E402

PROJECTILE = {"van-toc-dau": 20, "goc": 45, "do-cao-dau": 0, "g": 9.8}
PENDULUM_DEFAULTS = {"chieu-dai": 1.0, "g": 9.8, "goc-lech": 8, "khoi-luong": 0.2}
CIRCUIT = {"suat-dien-dong": 12, "dien-tro-1": 10, "dien-tro-2": 20, "kieu-mac": "noi-tiep"}
TITRATION = {"loai-acid": "hcl", "nong-do-acid": 0.1, "the-tich-acid": 20, "nong-do-base": 0.1,
             "the-tich-base": 0, "chi-thi": "phenolphtalein"}
RATE = {"nong-do": 0.1, "nhiet-do": 25, "xuc-tac": "khong"}


class ScienceLawTest(unittest.TestCase):
    """Định luật trên bản Python tham chiếu; test đối chiếu JavaScript bảo đảm JS khớp bản này."""

    def test_projectile_conserves_energy_and_peaks_at_45_degrees(self):
        p = {**PROJECTILE, "van-toc-dau": 20, "goc": 60, "do-cao-dau": 5}
        d = tham_chieu.li_nem_xien(p)
        vy = p["van-toc-dau"] * math.sin(math.radians(60))
        self.assertAlmostEqual(p["g"] * d["do-cao-cuc-dai"], p["g"] * 5 + vy * vy / 2, places=9)
        ranges = {goc: tham_chieu.li_nem_xien({**PROJECTILE, "goc": goc})["tam-xa"] for goc in (30, 45, 60)}
        self.assertAlmostEqual(ranges[30], ranges[60], places=9)
        self.assertGreater(ranges[45], ranges[30])

    def test_pendulum_period_squared_is_proportional_to_length_and_ignores_mass(self):
        base = PENDULUM_DEFAULTS
        short = tham_chieu.li_con_lac_don({**base, "chieu-dai": 0.5})["chu-ki"]
        long = tham_chieu.li_con_lac_don({**base, "chieu-dai": 2.0})["chu-ki"]
        self.assertAlmostEqual(long / short, 2.0, places=12)
        heavy = tham_chieu.li_con_lac_don({**base, "khoi-luong": 1.0})["chu-ki"]
        self.assertEqual(heavy, tham_chieu.li_con_lac_don(base)["chu-ki"])

    def test_circuit_obeys_kirchhoff(self):
        base = CIRCUIT
        series = tham_chieu.li_mach_ohm(base)
        self.assertAlmostEqual(series["hieu-dien-the-1"] + series["hieu-dien-the-2"], base["suat-dien-dong"], places=12)
        parallel = tham_chieu.li_mach_ohm({**base, "kieu-mac": "song-song"})
        self.assertAlmostEqual(parallel["cuong-do-1"] + parallel["cuong-do-2"], parallel["cuong-do-mach-chinh"], places=12)
        self.assertLess(parallel["dien-tro-tuong-duong"], min(base["dien-tro-1"], base["dien-tro-2"]))

    def test_titration_has_textbook_landmarks(self):
        base = TITRATION
        ph = lambda **change: tham_chieu.hoa_chuan_do({**base, **change})["ph"]  # noqa: E731
        self.assertAlmostEqual(ph(**{"the-tich-base": 0}), 1.0, places=6)
        self.assertAlmostEqual(ph(**{"the-tich-base": 20}), 7.0, places=6)
        weak = {"loai-acid": "ch3cooh"}
        self.assertAlmostEqual(ph(**weak, **{"the-tich-base": 10}), -math.log10(1.75e-5), places=2)
        self.assertGreater(ph(**weak, **{"the-tich-base": 20}), 8.0)
        curve = [ph(**weak, **{"the-tich-base": volume / 2}) for volume in range(0, 101)]
        self.assertEqual(curve, sorted(curve))

    def test_no2_equilibrium_follows_le_chatelier(self):
        value = lambda t, p: tham_chieu.hoa_can_bang_no2({"nhiet-do": t, "ap-suat": p})  # noqa: E731
        self.assertGreater(value(60, 1)["phan-mol-no2"], value(20, 1)["phan-mol-no2"])
        self.assertLess(value(25, 4)["phan-mol-no2"], value(25, 1)["phan-mol-no2"])
        self.assertGreater(value(25, 4)["nong-do-no2"], value(25, 1)["nong-do-no2"])
        here = value(25, 1)
        x = here["phan-mol-no2"]
        self.assertAlmostEqual(x * x * 1 / (1 - x), here["kp"], places=12)
        self.assertAlmostEqual(here["kp"], 0.146, places=2)

    def test_rate_follows_concentration_arrhenius_and_catalyst(self):
        base = RATE
        time = lambda **change: tham_chieu.hoa_toc_do({**base, **change})["thoi-gian"]  # noqa: E731
        self.assertAlmostEqual(time(), 40.0, places=9)
        self.assertAlmostEqual(time(**{"nong-do": 0.2}), 20.0, places=9)
        ratio = time() / time(**{"nhiet-do": 35})
        self.assertAlmostEqual(ratio, math.exp(50000 / 8.314462618 * (1 / 298.15 - 1 / 308.15)), places=9)
        self.assertLess(time(**{"xuc-tac": "co"}), time())

    def test_function_extrema_and_tangent(self):
        d = tham_chieu.toan_ham_so({"a": 0, "b": 2, "c": -8, "d": 1, "x0": 2})
        self.assertEqual((d["hoanh-do-cuc-tieu"], d["hoanh-do-cuc-dai"], d["he-so-goc"]), (2.0, None, 0.0))
        none = tham_chieu.toan_ham_so({"a": 1, "b": 0, "c": 3, "d": 0, "x0": 0})
        self.assertEqual((none["hoanh-do-cuc-dai"], none["hoanh-do-cuc-tieu"]), (None, None))

    def test_frequency_approaches_probability(self):
        for phep_thu, xac_suat in (("dong-xu-ngua", 1 / 2), ("hai-xuc-xac-tong-7", 1 / 6)):
            d = tham_chieu.toan_xac_suat({"phep-thu": phep_thu, "so-lan": 100000, "hat-giong": 3})
            self.assertLess(abs(d["tan-suat"] - xac_suat), 0.01)
            self.assertEqual(d["tan-so"], round(d["tan-suat"] * 100000))

if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Chạy để thấy trượt**

Run: `venv\Scripts\python.exe -m unittest tools.vi.tests.test_thi_nghiem_tham_chieu -v`
Expected: ERROR `ModuleNotFoundError` hoặc `ImportError: cannot import name 'tham_chieu'`.

- [ ] **Step 3: Viết bản tham chiếu**

`tools/vi/thi_nghiem_parts/tham_chieu.py`:

```python
"""Bản tính lại độc lập bằng Python của 8 mô hình thí nghiệm ảo.

Dùng cho hai việc: test so với JavaScript, và bảng số liệu lí tưởng ở trang giáo viên của phiếu học tập.
Mỗi hàm nhận dict tham số (đã đủ mặc định) và trả về dict đại lượng đo.
"""

from __future__ import annotations

import math

MASK = 0xFFFFFFFF
KW = 1.0e-14
KA = {"ch3cooh": 1.75e-5}
DELTA_H = 57200.0
DELTA_S = 175.83
R = 8.314462618
R_BAR = 0.08314462618
K25 = 0.00125
T25 = 298.15
EA = {"khong": 50000.0, "co": 45000.0}
LUONG_PHAN_UNG = 0.005
XAC_SUAT = {
    "dong-xu-ngua": 1 / 2, "xuc-xac-mat-6": 1 / 6, "xuc-xac-chan": 1 / 2,
    "hai-xuc-xac-tong-7": 1 / 6, "hai-xuc-xac-tong-12": 1 / 36,
}


def _imul(x: int, y: int) -> int:
    return (x * y) & MASK


def ngau_nhien(hat_giong: int):
    """mulberry32, khớp từng bit với taoNgauNhien trong khung.js."""
    state = hat_giong & MASK

    def rng() -> float:
        nonlocal state
        state = (state + 0x6D2B79F5) & MASK
        t = _imul(state ^ (state >> 15), 1 | state)
        t = ((t + _imul(t ^ (t >> 7), 61 | t)) & MASK) ^ t
        return ((t ^ (t >> 14)) & MASK) / 4294967296

    return rng


def li_nem_xien(p: dict) -> dict:
    goc = math.radians(p["goc"])
    vy = p["van-toc-dau"] * math.sin(goc)
    vx = p["van-toc-dau"] * math.cos(goc)
    thoi_gian = (vy + math.sqrt(vy * vy + 2 * p["g"] * p["do-cao-dau"])) / p["g"]
    return {
        "thoi-gian-bay": thoi_gian,
        "tam-xa": vx * thoi_gian,
        "do-cao-cuc-dai": p["do-cao-dau"] + vy * vy / (2 * p["g"]),
    }


def li_con_lac_don(p: dict) -> dict:
    chu_ki = 2 * math.pi * math.sqrt(p["chieu-dai"] / p["g"])
    return {"chu-ki": chu_ki, "thoi-gian-10-dao-dong": 10 * chu_ki}


def li_mach_ohm(p: dict) -> dict:
    e, r1, r2 = p["suat-dien-dong"], p["dien-tro-1"], p["dien-tro-2"]
    if p["kieu-mac"] == "song-song":
        return {
            "dien-tro-tuong-duong": 1 / (1 / r1 + 1 / r2),
            "cuong-do-mach-chinh": e / r1 + e / r2,
            "cuong-do-1": e / r1, "cuong-do-2": e / r2,
            "hieu-dien-the-1": e, "hieu-dien-the-2": e,
        }
    dong = e / (r1 + r2)
    return {
        "dien-tro-tuong-duong": r1 + r2,
        "cuong-do-mach-chinh": dong,
        "cuong-do-1": dong, "cuong-do-2": dong,
        "hieu-dien-the-1": dong * r1, "hieu-dien-the-2": dong * r2,
    }


def hoa_chuan_do(p: dict) -> dict:
    tong = p["the-tich-acid"] + p["the-tich-base"]
    acid = p["nong-do-acid"] * p["the-tich-acid"] / tong
    natri = p["nong-do-base"] * p["the-tich-base"] / tong
    if p["loai-acid"] == "hcl":
        lech = acid - natri
        h = (lech + math.sqrt(lech * lech + 4 * KW)) / 2
        return {"ph": -math.log10(h)}
    ka = KA[p["loai-acid"]]

    # Khác bản JS: khử mẫu thành đa thức bậc ba theo [H+], rồi chia đôi trên log10[H+].
    def da_thuc(h: float) -> float:
        return h ** 3 + (ka + natri) * h ** 2 + (ka * natri - KW - acid * ka) * h - ka * KW

    thap, cao = -14.0, 0.0
    for _ in range(200):
        giua = (thap + cao) / 2
        if da_thuc(10 ** giua) < 0:
            thap = giua
        else:
            cao = giua
    return {"ph": -(thap + cao) / 2}


def hoa_can_bang_no2(p: dict) -> dict:
    nhiet_do = p["nhiet-do"] + 273.15
    ap_suat = p["ap-suat"]
    kp = math.exp(-(DELTA_H - nhiet_do * DELTA_S) / (R * nhiet_do))
    # Giải theo độ phân li alpha rồi suy ra phần mol: cách khác với bản JS (JS giải thẳng phần mol).
    alpha = math.sqrt(kp / (kp + 4 * ap_suat))
    phan_mol = 2 * alpha / (1 + alpha)
    return {
        "kp": kp,
        "phan-mol-no2": phan_mol,
        "do-phan-li": alpha,
        "nong-do-no2": phan_mol * ap_suat / (R_BAR * nhiet_do),
    }


def hoa_toc_do(p: dict) -> dict:
    nhiet_do = p["nhiet-do"] + 273.15
    k = K25 * math.exp(EA["khong"] / (R * T25) - EA[p["xuc-tac"]] / (R * nhiet_do))
    return {"thoi-gian": LUONG_PHAN_UNG / (k * p["nong-do"])}


def toan_ham_so(p: dict) -> dict:
    a, b, c, d, x = p["a"], p["b"], p["c"], p["d"], p["x0"]
    cuc_dai = cuc_tieu = None
    nghiem = []
    if a != 0:
        biet = (2 * b) ** 2 - 4 * (3 * a) * c
        if biet > 0:
            nghiem = sorted(((-2 * b - math.sqrt(biet)) / (6 * a), (-2 * b + math.sqrt(biet)) / (6 * a)))
    elif b != 0:
        nghiem = [-c / (2 * b)]
    for diem in nghiem:
        dao_ham_cap_hai = 6 * a * diem + 2 * b
        if dao_ham_cap_hai < 0:
            cuc_dai = diem
        elif dao_ham_cap_hai > 0:
            cuc_tieu = diem
    return {
        "gia-tri": a * x ** 3 + b * x ** 2 + c * x + d,
        "he-so-goc": 3 * a * x ** 2 + 2 * b * x + c,
        "hoanh-do-cuc-dai": cuc_dai,
        "hoanh-do-cuc-tieu": cuc_tieu,
    }


def toan_xac_suat(p: dict) -> dict:
    rng = ngau_nhien(int(p["hat-giong"]))
    phep_thu = p["phep-thu"]
    tan_so = 0
    for _ in range(int(p["so-lan"])):
        if phep_thu == "dong-xu-ngua":
            xay_ra = rng() < 0.5
        else:
            mat = math.floor(rng() * 6) + 1
            if phep_thu == "xuc-xac-mat-6":
                xay_ra = mat == 6
            elif phep_thu == "xuc-xac-chan":
                xay_ra = mat % 2 == 0
            else:
                tong = mat + math.floor(rng() * 6) + 1
                xay_ra = tong == (7 if phep_thu == "hai-xuc-xac-tong-7" else 12)
        tan_so += xay_ra
    return {"tan-so": tan_so, "tan-suat": tan_so / p["so-lan"], "xac-suat-li-thuyet": XAC_SUAT[phep_thu]}


THAM_CHIEU = {
    "li-nem-xien": li_nem_xien,
    "li-con-lac-don": li_con_lac_don,
    "li-mach-ohm": li_mach_ohm,
    "hoa-chuan-do": hoa_chuan_do,
    "hoa-can-bang-no2": hoa_can_bang_no2,
    "hoa-toc-do": hoa_toc_do,
    "toan-ham-so": toan_ham_so,
    "toan-xac-suat": toan_xac_suat,
}
```

- [ ] **Step 4: Chạy lại**

Run: `venv\Scripts\python.exe -m unittest tools.vi.tests.test_thi_nghiem_tham_chieu -v`
Expected: PASS (8 test).

- [ ] **Step 5: Commit**

```bash
git add tools/vi/thi_nghiem_parts/tham_chieu.py tools/vi/tests/test_thi_nghiem_tham_chieu.py
git commit -m "feat(vi): add independent Python reference models with law tests"
```

---

### Task 3: Khuôn mô hình, bộ nạp, kiểm số qua Node, và mô hình con lắc đơn

**Files:**
- Create: `tools/vi/thi_nghiem_parts/thu_vien.py`
- Create: `tools/vi/thi_nghiem_parts/kiem_so.py`
- Create: `tools/vi/thi_nghiem_parts/mo_hinh/li-con-lac-don.json`
- Create: `tools/vi/thi_nghiem_parts/mo_hinh/li-con-lac-don.js`
- Create: `tools/vi/tests/thi_nghiem_mau.py`
- Create: `tools/vi/tests/test_thi_nghiem_mo_hinh.py`

**Interfaces:**
- Consumes: `THI_NGHIEM_KHUNG.tuKiem`, `chay_node.js` (Task 1); `tham_chieu.THAM_CHIEU` (Task 2).
- Produces: `thu_vien.Model(ma, khai_bao, js, json_path, js_path, moi)` với `tham_so(ma)`, `dai_luong(ma)`, `mac_dinh()`; `thu_vien.load(ma, folder) -> Model` (ném `ModelError`); `thu_vien.check_declaration(kb) -> list[str]`; `thu_vien.check_js(js, animation) -> list[str]`; `thu_vien.list_models() -> list[str]`; hằng `NEW_MODEL = "moi"`, `MIN_CHECK_ROWS = 5`.
- Produces: `kiem_so.find_node() -> str|None`; `kiem_so.bang_kiem(model) -> dict|None`; `kiem_so.luoi(model, diem) -> list[dict]|None`; `kiem_so.CheckError`. Trả `None` nghĩa là máy không có Node.
- Produces (test): `thi_nghiem_mau.{PENDULUM, HOOKE_JSON, HOOKE_JS, HOOKE_MD, grid(model, count, rng)}`.

- [ ] **Step 1: Viết dữ liệu mẫu và test**

`tools/vi/tests/thi_nghiem_mau.py`:

```python
"""Dữ liệu mẫu dùng chung cho các test thí nghiệm ảo."""

import random

PENDULUM = """---
tieu-de: Chu kì con lắc đơn
mon: Vật lí
lop: 11
mau: li-con-lac-don
nguoi-thao-tac: nhom
sai-so: bat
---

## Tham số
chieu-dai: 0.4..1.6 buoc 0.2 mac-dinh 1.0
g: co-dinh 9.8

## Dự đoán
cau: Khi tăng chiều dài dây gấp 4 lần, chu kì thay đổi thế nào?
A: Tăng 4 lần
B: Tăng 2 lần
C: Không đổi
dap-an: B

## Quan sát
so-lan-do: 5
do: chieu-dai, chu-ki
do-thi: chu-ki^2 theo chieu-dai

## Giải thích
cau: Từ đồ thị, T^2^ và l liên hệ thế nào?
goi-y-dap-an: T^2^ tỉ lệ thuận với l; T = 2π√(l/g).

## Kết luận
Chu kì con lắc đơn chỉ phụ thuộc chiều dài dây và g.
"""

HOOKE_JSON = {
    "ma": "moi", "ten": "Định luật Hooke", "mon": "Vật lí", "hoatHinh": "khong",
    "thamSo": [
        {"ma": "do-cung", "ten": "Độ cứng k", "kieu": "so", "donVi": "N/m", "min": 10, "max": 100, "buoc": 10, "macDinh": 50},
        {"ma": "do-gian", "ten": "Độ giãn x", "kieu": "so", "donVi": "m", "min": 0, "max": 0.2, "buoc": 0.01, "macDinh": 0.1},
    ],
    "daiLuongDo": [{"ma": "luc", "ten": "Lực đàn hồi F", "donVi": "N", "saiSo": 0.05, "chuSo": 2}],
    "congThuc": {"bieuThuc": "F = k·x", "dieuKien": "Lò xo còn trong giới hạn đàn hồi.", "nguon": ""},
    "bangKiem": [
        {"vao": {}, "ra": {"luc": 5.0}, "saiSo": 0.001},
        {"vao": {"do-gian": 0}, "ra": {"luc": 0.0}, "saiSo": 0.001},
        {"vao": {"do-gian": 0.2}, "ra": {"luc": 10.0}, "saiSo": 0.001},
        {"vao": {"do-cung": 100}, "ra": {"luc": 10.0}, "saiSo": 0.001},
        {"vao": {"do-cung": 10, "do-gian": 0.05}, "ra": {"luc": 0.5}, "saiSo": 0.001},
    ],
}
HOOKE_JS = """(function (root) {
  function tinh(p) { return { 'luc': p['do-cung'] * p['do-gian'] }; }
  function ve(ctx, p, t, kt, d) { ctx.fillRect(20, 20, 40 + 600 * p['do-gian'], 20); }
  root.THI_NGHIEM_MO_HINH = { tinh: tinh, ve: ve };
})(typeof globalThis !== 'undefined' ? globalThis : this);
"""
HOOKE_MD = """---
tieu-de: Định luật Hooke
mon: Vật lí
lop: 10
mau: moi
---

## Tham số
do-gian: 0..0.2 buoc 0.05 mac-dinh 0.1

## Dự đoán
cau: Độ giãn tăng gấp đôi thì lực đàn hồi thay đổi thế nào?

## Quan sát
so-lan-do: 4
do: do-gian, luc
do-thi: luc theo do-gian

## Giải thích
cau: Lực đàn hồi và độ giãn liên hệ thế nào?
goi-y-dap-an: F tỉ lệ thuận với x.

## Kết luận
Trong giới hạn đàn hồi, lực đàn hồi tỉ lệ thuận với độ giãn.
"""


def grid(model, count: int, rng: random.Random) -> list[dict]:
    points = []
    for _ in range(count):
        point = {}
        for ts in model.khai_bao["thamSo"]:
            if ts["kieu"] == "chon":
                point[ts["ma"]] = rng.choice(ts["luaChon"])["ma"]
            else:
                steps = round((ts["max"] - ts["min"]) / ts["buoc"])
                point[ts["ma"]] = round(ts["min"] + rng.randint(0, steps) * ts["buoc"], 6)
        points.append(point)
    return points
```

`tools/vi/tests/test_thi_nghiem_mo_hinh.py` (ở task này `MODELS` chỉ có một mẫu; Task 4–6 sẽ mở rộng):

```python
"""Test cho thư viện mô hình thí nghiệm ảo: khuôn, mã, và đối chiếu JavaScript với bản Python."""

import json
import random
import sys
import tempfile
import unittest
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from thi_nghiem_mau import HOOKE_JS, HOOKE_JSON, grid  # noqa: E402
from thi_nghiem_parts import kiem_so, tham_chieu, thu_vien  # noqa: E402

HAS_NODE = kiem_so.find_node() is not None
NEED_NODE = "máy không có Node nên không chạy được JavaScript"
MODELS = ("li-con-lac-don",)


class ModelLibraryTest(unittest.TestCase):
    def test_library_has_the_models(self):
        self.assertEqual(tuple(thu_vien.list_models()), MODELS)

    def test_every_library_model_loads_and_follows_the_contract(self):
        for ma in MODELS:
            with self.subTest(model=ma):
                model = thu_vien.load(ma, Path("."))
                self.assertFalse(model.moi)
                self.assertEqual(model.khai_bao["ma"], ma)
                self.assertGreaterEqual(len(model.khai_bao["bangKiem"]), thu_vien.MIN_CHECK_ROWS)
                self.assertIn(ma, tham_chieu.THAM_CHIEU)

    def test_unknown_model_lists_what_exists(self):
        with self.assertRaises(thu_vien.ModelError) as caught:
            thu_vien.load("li-khong-co", Path("."))
        self.assertIn("li-con-lac-don", str(caught.exception))
        self.assertIn("moi", str(caught.exception))

    def test_declaration_errors_are_named(self):
        cases = {
            "thiếu `congThuc.dieuKien`": lambda kb: kb["congThuc"].update(dieuKien=" "),
            "`bangKiem` phải có ít nhất 5 dòng": lambda kb: kb.update(bangKiem=kb["bangKiem"][:4]),
            "`hoatHinh` phải là một trong": lambda kb: kb.update(hoatHinh="nhanh"),
            "cần min < max": lambda kb: kb["thamSo"][0].update(min=200),
            "`vao` có tham số lạ `la`": lambda kb: kb["bangKiem"][0]["vao"].update(la=1),
            "`ra` có đại lượng lạ `la`": lambda kb: kb["bangKiem"][0]["ra"].update(la=1),
            "`saiSo` phải là số không âm": lambda kb: kb["daiLuongDo"][0].update(saiSo=-1),
            "`daiLuongDo` phải có ít nhất một": lambda kb: kb.update(daiLuongDo=[]),
        }
        for message, damage in cases.items():
            with self.subTest(message=message):
                kb = json.loads(json.dumps(HOOKE_JSON))
                damage(kb)
                self.assertTrue(any(message in error for error in thu_vien.check_declaration(kb)),
                                thu_vien.check_declaration(kb))
        self.assertEqual(thu_vien.check_declaration(HOOKE_JSON), [])

    def test_model_code_may_not_reach_the_network_or_the_page(self):
        for banned in ("fetch('x')", "https://cdn", "document.title", "eval('1')", "require('fs')", "</script>"):
            with self.subTest(banned=banned):
                self.assertTrue(thu_vien.check_js(HOOKE_JS + "// " + banned, "khong"))
        self.assertEqual(thu_vien.check_js(HOOKE_JS, "khong"), [])
        self.assertTrue(thu_vien.check_js(HOOKE_JS, "mot-lan"))

    def test_library_code_passes_its_own_rules(self):
        for ma in MODELS:
            model = thu_vien.load(ma, Path("."))
            self.assertEqual(thu_vien.check_js(model.js, model.khai_bao["hoatHinh"]), [], ma)

    def test_new_model_needs_both_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            (Path(tmp) / "mo-hinh.json").write_text(json.dumps(HOOKE_JSON), encoding="utf-8")
            with self.assertRaises(thu_vien.ModelError) as caught:
                thu_vien.load("moi", Path(tmp))
            self.assertIn("mo-hinh.js", str(caught.exception))


@unittest.skipUnless(HAS_NODE, NEED_NODE)
class JavaScriptAgreementTest(unittest.TestCase):
    def test_check_tables_pass(self):
        for ma in MODELS:
            with self.subTest(model=ma):
                check = kiem_so.bang_kiem(thu_vien.load(ma, Path(".")))
                self.assertEqual((check["dat"], check["truot"]), (check["tong"], []))

    def test_javascript_matches_the_python_reference_on_a_grid(self):
        rng = random.Random(20260920)
        for ma in MODELS:
            with self.subTest(model=ma):
                model = thu_vien.load(ma, Path("."))
                points = grid(model, 40 if ma == "toan-xac-suat" else 300, rng)
                for point, got in zip(points, kiem_so.luoi(model, points)):
                    expected = tham_chieu.THAM_CHIEU[ma](point)
                    for measure in model.khai_bao["daiLuongDo"]:
                        a, b = expected[measure["ma"]], got[measure["ma"]]
                        if a is None or b is None:
                            self.assertIsNone(a, point)
                            self.assertIsNone(b, point)
                        else:
                            self.assertLessEqual(abs(a - b), 1e-9 * max(1.0, abs(a)), (point, measure["ma"]))

    def test_a_wrong_model_fails_its_check_table(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            (folder / "mo-hinh.json").write_text(json.dumps(HOOKE_JSON), encoding="utf-8")
            (folder / "mo-hinh.js").write_text(HOOKE_JS.replace("p['do-cung'] * p['do-gian']", "p['do-cung'] + p['do-gian']"),
                                                encoding="utf-8")
            check = kiem_so.bang_kiem(thu_vien.load("moi", folder))
            self.assertLess(check["dat"], check["tong"])

    def test_a_crashing_model_is_a_check_error(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            (folder / "mo-hinh.json").write_text(json.dumps(HOOKE_JSON), encoding="utf-8")
            (folder / "mo-hinh.js").write_text(HOOKE_JS + "\nthis is not javascript(", encoding="utf-8")
            with self.assertRaises(kiem_so.CheckError):
                kiem_so.bang_kiem(thu_vien.load("moi", folder))

if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Chạy để thấy trượt**

Run: `venv\Scripts\python.exe -m unittest tools.vi.tests.test_thi_nghiem_mo_hinh -v`
Expected: ERROR `ImportError: cannot import name 'kiem_so'`.

- [ ] **Step 3: Viết bộ nạp và bộ kiểm số**

`tools/vi/thi_nghiem_parts/thu_vien.py`:

```python
"""Nạp và kiểm mô hình thí nghiệm ảo: mẫu trong thư viện hoặc mô hình mới trong thư mục thí nghiệm.

Mỗi mô hình gồm hai file: <mã>.json (khai báo) và <mã>.js (hàm tinh, ve). Khuôn ở docs/vi/tro-ly/mo-hinh-thi-nghiem.md.
Chỉ dùng thư viện chuẩn Python.
"""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path

LIBRARY_DIR = Path(__file__).resolve().parent / "mo_hinh"
NEW_MODEL = "moi"
NEW_JSON = "mo-hinh.json"
NEW_JS = "mo-hinh.js"
MIN_CHECK_ROWS = 5
ANIMATIONS = ("khong", "mot-lan", "lap")
BANNED_JS = (
    "http://", "https://", "fetch(", "XMLHttpRequest", "WebSocket", "import(", "import ", "require(",
    "eval(", "Function(", "process.", "document.", "window.", "localStorage", "</script",
)


class ModelError(Exception):
    """Mô hình không có, hoặc không đúng khuôn."""


@dataclass
class Model:
    ma: str
    khai_bao: dict
    js: str
    json_path: Path
    js_path: Path
    moi: bool

    def tham_so(self, ma: str) -> dict | None:
        return next((ts for ts in self.khai_bao["thamSo"] if ts["ma"] == ma), None)

    def dai_luong(self, ma: str) -> dict | None:
        return next((dl for dl in self.khai_bao["daiLuongDo"] if dl["ma"] == ma), None)

    def mac_dinh(self) -> dict:
        return {ts["ma"]: ts["macDinh"] for ts in self.khai_bao["thamSo"]}


def list_models() -> list[str]:
    return sorted(path.stem for path in LIBRARY_DIR.glob("*.json"))


def _number(value) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _text(value) -> bool:
    return isinstance(value, str) and value.strip() != ""


def check_declaration(kb) -> list[str]:
    """Trả về danh sách lỗi của file khai báo; rỗng là đạt."""
    if not isinstance(kb, dict):
        return ["file khai báo phải là một đối tượng JSON"]
    errors: list[str] = []
    for key in ("ten", "mon"):
        if not _text(kb.get(key)):
            errors.append(f"thiếu `{key}`")
    if kb.get("hoatHinh") not in ANIMATIONS:
        errors.append("`hoatHinh` phải là một trong: " + ", ".join(ANIMATIONS))

    params = kb.get("thamSo")
    param_codes: dict[str, dict] = {}
    if not isinstance(params, list) or not params:
        errors.append("`thamSo` phải có ít nhất một tham số")
        params = []
    for index, ts in enumerate(params, 1):
        where = f"thamSo[{index}]"
        if not isinstance(ts, dict) or not _text(ts.get("ma")) or not _text(ts.get("ten")):
            errors.append(f"{where}: thiếu `ma` hoặc `ten`")
            continue
        param_codes[ts["ma"]] = ts
        if ts.get("kieu") == "so":
            if not all(_number(ts.get(key)) for key in ("min", "max", "buoc", "macDinh")):
                errors.append(f"{where}: tham số số cần `min`, `max`, `buoc`, `macDinh` là số")
            elif not (ts["min"] < ts["max"] and ts["buoc"] > 0 and ts["min"] <= ts["macDinh"] <= ts["max"]):
                errors.append(f"{where}: cần min < max, buoc > 0 và macDinh nằm trong khoảng")
            if not isinstance(ts.get("donVi"), str):
                errors.append(f"{where}: thiếu `donVi` (để chuỗi rỗng nếu không có đơn vị)")
        elif ts.get("kieu") == "chon":
            options = ts.get("luaChon")
            codes = [o.get("ma") for o in options if isinstance(o, dict)] if isinstance(options, list) else []
            if len(codes) < 2 or not all(_text(o.get("ma")) and _text(o.get("ten")) for o in options):
                errors.append(f"{where}: tham số lựa chọn cần ít nhất hai mục có `ma` và `ten`")
            elif ts.get("macDinh") not in codes:
                errors.append(f"{where}: `macDinh` phải là một trong các lựa chọn")
        else:
            errors.append(f"{where}: `kieu` phải là `so` hoặc `chon`")

    measures = kb.get("daiLuongDo")
    measure_codes: set[str] = set()
    if not isinstance(measures, list) or not measures:
        errors.append("`daiLuongDo` phải có ít nhất một đại lượng đo")
        measures = []
    for index, dl in enumerate(measures, 1):
        where = f"daiLuongDo[{index}]"
        if not isinstance(dl, dict) or not _text(dl.get("ma")) or not _text(dl.get("ten")):
            errors.append(f"{where}: thiếu `ma` hoặc `ten`")
            continue
        measure_codes.add(dl["ma"])
        if not isinstance(dl.get("donVi"), str):
            errors.append(f"{where}: thiếu `donVi`")
        if not _number(dl.get("saiSo")) or dl["saiSo"] < 0:
            errors.append(f"{where}: `saiSo` phải là số không âm")
        if not isinstance(dl.get("chuSo"), int) or isinstance(dl.get("chuSo"), bool) or not 0 <= dl["chuSo"] <= 8:
            errors.append(f"{where}: `chuSo` phải là số nguyên từ 0 đến 8")
    if set(param_codes) & measure_codes:
        errors.append("mã tham số và mã đại lượng đo không được trùng nhau")

    formula = kb.get("congThuc")
    if not isinstance(formula, dict) or not _text(formula.get("bieuThuc")):
        errors.append("thiếu `congThuc.bieuThuc`")
    if not isinstance(formula, dict) or not _text(formula.get("dieuKien")):
        errors.append("thiếu `congThuc.dieuKien` (điều kiện áp dụng của công thức)")

    rows = kb.get("bangKiem")
    if not isinstance(rows, list) or len(rows) < MIN_CHECK_ROWS:
        errors.append(f"`bangKiem` phải có ít nhất {MIN_CHECK_ROWS} dòng")
        rows = rows if isinstance(rows, list) else []
    for index, row in enumerate(rows, 1):
        where = f"bangKiem[{index}]"
        if not isinstance(row, dict) or not isinstance(row.get("vao"), dict) or not isinstance(row.get("ra"), dict):
            errors.append(f"{where}: cần `vao` và `ra` là đối tượng")
            continue
        if not row["ra"]:
            errors.append(f"{where}: `ra` không được rỗng")
        if not _number(row.get("saiSo")) or row["saiSo"] < 0:
            errors.append(f"{where}: `saiSo` phải là số không âm")
        for code in row["vao"]:
            if code not in param_codes:
                errors.append(f"{where}: `vao` có tham số lạ `{code}`")
        for code, value in row["ra"].items():
            if code not in measure_codes:
                errors.append(f"{where}: `ra` có đại lượng lạ `{code}`")
            elif value is not None and not _number(value):
                errors.append(f"{where}: `ra.{code}` phải là số hoặc null")
    return errors


def check_js(js: str, animation: str) -> list[str]:
    errors = [f"file mô hình không được chứa `{token.strip()}`" for token in BANNED_JS if token in js]
    for name in ("THI_NGHIEM_MO_HINH", "tinh", "ve"):
        if name not in js:
            errors.append(f"file mô hình thiếu `{name}`")
    if animation == "mot-lan" and "thoiLuong" not in js:
        errors.append("mô hình `hoatHinh: mot-lan` phải có hàm `thoiLuong`")
    return errors


def load(ma: str, folder: Path) -> Model:
    """ma là mã mẫu trong thư viện, hoặc `moi` để lấy mo-hinh.json và mo-hinh.js trong folder."""
    moi = ma == NEW_MODEL
    if moi:
        json_path, js_path = folder / NEW_JSON, folder / NEW_JS
        missing = [path.name for path in (json_path, js_path) if not path.is_file()]
        if missing:
            raise ModelError("`mau: moi` cần có trong thư mục thí nghiệm: " + ", ".join(missing))
    else:
        json_path, js_path = LIBRARY_DIR / f"{ma}.json", LIBRARY_DIR / f"{ma}.js"
        if not json_path.is_file() or not js_path.is_file():
            raise ModelError(f"Không có mẫu `{ma}`. Các mẫu có sẵn: " + ", ".join(list_models()) + "; hoặc `moi`.")
    try:
        khai_bao = json.loads(json_path.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ModelError(f"Không đọc được {json_path.name}: {exc}") from exc
    try:
        js = js_path.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeDecodeError) as exc:
        raise ModelError(f"Không đọc được {js_path.name}: {exc}") from exc
    errors = check_declaration(khai_bao)
    if errors:
        raise ModelError(f"{json_path.name}: " + "; ".join(errors))
    errors = check_js(js, khai_bao["hoatHinh"])
    if errors:
        raise ModelError(f"{js_path.name}: " + "; ".join(errors))
    return Model(ma=ma, khai_bao=khai_bao, js=js, json_path=json_path, js_path=js_path, moi=moi)
```

`tools/vi/thi_nghiem_parts/kiem_so.py`:

```python
"""Chạy mô hình JavaScript qua Node để kiểm số. Máy không có Node thì trả về None, không báo lỗi."""

from __future__ import annotations

import json
import shutil
import subprocess
from pathlib import Path

RUNTIME_DIR = Path(__file__).resolve().parent / "runtime"
RUNNER = RUNTIME_DIR / "chay_node.js"
KHUNG_JS = RUNTIME_DIR / "khung.js"
TIMEOUT_S = 60


class CheckError(Exception):
    """Node có chạy nhưng mô hình hỏng: lỗi cú pháp, ném lỗi, hoặc quá thời gian."""


def find_node() -> str | None:
    return shutil.which("node")


def _run(model, mode: str, stdin: str = ""):
    node = find_node()
    if node is None:
        return None
    try:
        proc = subprocess.run(
            [node, str(RUNNER), str(KHUNG_JS), str(model.js_path), str(model.json_path), mode],
            input=stdin, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=TIMEOUT_S,
        )
    except subprocess.TimeoutExpired as exc:
        raise CheckError(f"Mô hình chạy quá {TIMEOUT_S} giây, có thể bị lặp vô hạn") from exc
    except OSError as exc:
        raise CheckError(f"Không chạy được Node: {exc}") from exc
    try:
        data = json.loads(proc.stdout.strip().splitlines()[-1])
    except (IndexError, json.JSONDecodeError) as exc:
        raise CheckError(f"Node không trả về JSON: {(proc.stderr or proc.stdout).strip()[:300]}") from exc
    if isinstance(data, dict) and "loi" in data:
        raise CheckError(f"Mô hình lỗi khi chạy: {data['loi']}")
    return data


def bang_kiem(model) -> dict | None:
    """{'tong', 'dat', 'truot': [{'dong', 'ma', 'mong', 'duoc'}]} hoặc None khi máy không có Node."""
    return _run(model, "bang-kiem")


def luoi(model, diem: list[dict]) -> list[dict] | None:
    """Giá trị các đại lượng đo tại từng bộ tham số; None khi máy không có Node."""
    return _run(model, "luoi", json.dumps(diem, ensure_ascii=False))
```

- [ ] **Step 4: Viết mô hình con lắc đơn**

`tools/vi/thi_nghiem_parts/mo_hinh/li-con-lac-don.json`:

```json
{
  "ma": "li-con-lac-don",
  "ten": "Con lắc đơn",
  "mon": "Vật lí",
  "hoatHinh": "lap",
  "thamSo": [
    {
      "ma": "chieu-dai",
      "ten": "Chiều dài dây l",
      "kieu": "so",
      "donVi": "m",
      "min": 0.2,
      "max": 2.0,
      "buoc": 0.05,
      "macDinh": 1.0
    },
    {
      "ma": "g",
      "ten": "Gia tốc trọng trường g",
      "kieu": "so",
      "donVi": "m/s^2^",
      "min": 1.6,
      "max": 24.8,
      "buoc": 0.1,
      "macDinh": 9.8
    },
    {
      "ma": "goc-lech",
      "ten": "Góc lệch ban đầu",
      "kieu": "so",
      "donVi": "°",
      "min": 2,
      "max": 15,
      "buoc": 1,
      "macDinh": 8
    },
    {
      "ma": "khoi-luong",
      "ten": "Khối lượng quả nặng m",
      "kieu": "so",
      "donVi": "kg",
      "min": 0.05,
      "max": 1.0,
      "buoc": 0.05,
      "macDinh": 0.2
    }
  ],
  "daiLuongDo": [
    {
      "ma": "chu-ki",
      "ten": "Chu kì T",
      "donVi": "s",
      "saiSo": 0.02,
      "chuSo": 3
    },
    {
      "ma": "thoi-gian-10-dao-dong",
      "ten": "Thời gian 10 dao động",
      "donVi": "s",
      "saiSo": 0.1,
      "chuSo": 2
    }
  ],
  "congThuc": {
    "bieuThuc": "T = 2π√(l/g)",
    "dieuKien": "Góc lệch nhỏ (không quá 15°); dây không giãn, khối lượng dây không đáng kể; bỏ qua sức cản.",
    "nguon": ""
  },
  "bangKiem": [
    {
      "vao": {
        "chieu-dai": 1.0
      },
      "ra": {
        "chu-ki": 2.007089923,
        "thoi-gian-10-dao-dong": 20.07089923
      },
      "saiSo": 0.0005
    },
    {
      "vao": {
        "chieu-dai": 0.25
      },
      "ra": {
        "chu-ki": 1.003544962,
        "thoi-gian-10-dao-dong": 10.03544962
      },
      "saiSo": 0.0005
    },
    {
      "vao": {
        "chieu-dai": 2.0
      },
      "ra": {
        "chu-ki": 2.83845379,
        "thoi-gian-10-dao-dong": 28.3845379
      },
      "saiSo": 0.0005
    },
    {
      "vao": {
        "chieu-dai": 1.0,
        "g": 1.6
      },
      "ra": {
        "chu-ki": 4.967294133,
        "thoi-gian-10-dao-dong": 49.67294133
      },
      "saiSo": 0.0005
    },
    {
      "vao": {
        "chieu-dai": 0.5,
        "khoi-luong": 1.0
      },
      "ra": {
        "chu-ki": 1.419226895,
        "thoi-gian-10-dao-dong": 14.19226895
      },
      "saiSo": 0.0005
    },
    {
      "vao": {
        "chieu-dai": 0.5,
        "khoi-luong": 0.05
      },
      "ra": {
        "chu-ki": 1.419226895,
        "thoi-gian-10-dao-dong": 14.19226895
      },
      "saiSo": 0.0005
    }
  ]
}
```

`tools/vi/thi_nghiem_parts/mo_hinh/li-con-lac-don.js`:

```js
(function (root) {
  'use strict';

  function tinh(p) {
    var chuKi = 2 * Math.PI * Math.sqrt(p['chieu-dai'] / p['g']);
    return { 'chu-ki': chuKi, 'thoi-gian-10-dao-dong': 10 * chuKi };
  }

  function ve(ctx, p, t, kt, d) {
    var K = root.THI_NGHIEM_KHUNG;
    var goc = p['goc-lech'] * Math.PI / 180 * Math.cos(2 * Math.PI * t / d['chu-ki']);
    var treo = [kt.rong / 2, 40];
    var day = p['chieu-dai'] / 2.0 * (kt.cao - 120);
    var vat = [treo[0] + day * Math.sin(goc), treo[1] + day * Math.cos(goc)];
    var banKinh = 10 + 14 * Math.sqrt(p['khoi-luong']);
    ctx.strokeStyle = K.MAU.net; ctx.lineWidth = 4;
    ctx.beginPath(); ctx.moveTo(treo[0] - 70, treo[1]); ctx.lineTo(treo[0] + 70, treo[1]); ctx.stroke();
    ctx.setLineDash([5, 5]); ctx.lineWidth = 1; ctx.strokeStyle = K.MAU.nhat;
    ctx.beginPath(); ctx.moveTo(treo[0], treo[1]); ctx.lineTo(treo[0], treo[1] + day + 30); ctx.stroke();
    ctx.setLineDash([]); ctx.lineWidth = 2; ctx.strokeStyle = K.MAU.net;
    ctx.beginPath(); ctx.moveTo(treo[0], treo[1]); ctx.lineTo(vat[0], vat[1]); ctx.stroke();
    ctx.fillStyle = K.MAU.chinh; ctx.beginPath(); ctx.arc(vat[0], vat[1], banKinh, 0, 2 * Math.PI); ctx.fill();
    ctx.fillStyle = K.MAU.net; ctx.font = '15px ' + K.PHONG;
    ctx.fillText('t = ' + K.dinhDang(t, 2) + ' s', 16, 24);
    ctx.fillText('Số dao động: ' + Math.floor(t / d['chu-ki']), 16, 46);
  }

  root.THI_NGHIEM_MO_HINH = { tinh: tinh, ve: ve };
})(typeof globalThis !== 'undefined' ? globalThis : this);
```

- [ ] **Step 5: Chạy lại**

Run: `venv\Scripts\python.exe -m unittest tools.vi.tests.test_thi_nghiem_mo_hinh -v`
Expected: PASS (11 test, không test nào bị bỏ qua).

- [ ] **Step 6: Commit**

```bash
git add tools/vi/thi_nghiem_parts/thu_vien.py tools/vi/thi_nghiem_parts/kiem_so.py tools/vi/thi_nghiem_parts/mo_hinh tools/vi/tests/thi_nghiem_mau.py tools/vi/tests/test_thi_nghiem_mo_hinh.py
git commit -m "feat(vi): add the experiment model contract, loader, and the pendulum model"
```

---

### Task 4: Hai mô hình Vật lí — ném xiên và đoạn mạch nối tiếp, song song

**Files:**
- Create: `tools/vi/thi_nghiem_parts/mo_hinh/li-nem-xien.json`, `li-nem-xien.js`
- Create: `tools/vi/thi_nghiem_parts/mo_hinh/li-mach-ohm.json`, `li-mach-ohm.js`
- Modify: `tools/vi/tests/test_thi_nghiem_mo_hinh.py` (hằng `MODELS`)

**Interfaces:**
- Consumes: khuôn mô hình (Task 3), `tham_chieu.li_nem_xien`, `tham_chieu.li_mach_ohm` (Task 2).
- Produces: mã mẫu `li-nem-xien` (`hoatHinh: mot-lan`, có `thoiLuong`), `li-mach-ohm` (tham số lựa chọn `kieu-mac`).

- [ ] **Step 1: Mở rộng `MODELS` trong test**

Trong `tools/vi/tests/test_thi_nghiem_mo_hinh.py` thay:

```python
MODELS = ("li-con-lac-don",)
```

bằng:

```python
MODELS = ("li-con-lac-don", "li-mach-ohm", "li-nem-xien")
```

- [ ] **Step 2: Chạy để thấy trượt**

Run: `venv\Scripts\python.exe -m unittest tools.vi.tests.test_thi_nghiem_mo_hinh -v`
Expected: FAIL ở `test_library_has_the_models` (thư viện mới có một mẫu).

- [ ] **Step 3: Viết hai mô hình**

`tools/vi/thi_nghiem_parts/mo_hinh/li-nem-xien.json`:

```json
{
  "ma": "li-nem-xien",
  "ten": "Chuyển động ném xiên",
  "mon": "Vật lí",
  "hoatHinh": "mot-lan",
  "thamSo": [
    {
      "ma": "van-toc-dau",
      "ten": "Vận tốc đầu v~0~",
      "kieu": "so",
      "donVi": "m/s",
      "min": 5,
      "max": 50,
      "buoc": 1,
      "macDinh": 20
    },
    {
      "ma": "goc",
      "ten": "Góc ném α",
      "kieu": "so",
      "donVi": "°",
      "min": 0,
      "max": 85,
      "buoc": 1,
      "macDinh": 45
    },
    {
      "ma": "do-cao-dau",
      "ten": "Độ cao ban đầu h",
      "kieu": "so",
      "donVi": "m",
      "min": 0,
      "max": 50,
      "buoc": 1,
      "macDinh": 0
    },
    {
      "ma": "g",
      "ten": "Gia tốc trọng trường g",
      "kieu": "so",
      "donVi": "m/s^2^",
      "min": 1.6,
      "max": 24.8,
      "buoc": 0.1,
      "macDinh": 9.8
    }
  ],
  "daiLuongDo": [
    {
      "ma": "tam-xa",
      "ten": "Tầm xa L",
      "donVi": "m",
      "saiSo": 0.2,
      "chuSo": 2
    },
    {
      "ma": "do-cao-cuc-dai",
      "ten": "Độ cao cực đại H",
      "donVi": "m",
      "saiSo": 0.1,
      "chuSo": 2
    },
    {
      "ma": "thoi-gian-bay",
      "ten": "Thời gian bay t",
      "donVi": "s",
      "saiSo": 0.02,
      "chuSo": 2
    }
  ],
  "congThuc": {
    "bieuThuc": "t = (v~0~sinα + √(v~0~^2^sin^2^α + 2gh))/g; L = v~0~cosα·t; H = h + v~0~^2^sin^2^α/(2g)",
    "dieuKien": "Bỏ qua sức cản không khí; g không đổi; vật coi là chất điểm.",
    "nguon": ""
  },
  "bangKiem": [
    {
      "vao": {
        "van-toc-dau": 20,
        "goc": 45
      },
      "ra": {
        "tam-xa": 40.81632653,
        "do-cao-cuc-dai": 10.20408163,
        "thoi-gian-bay": 2.886150127
      },
      "saiSo": 0.001
    },
    {
      "vao": {
        "van-toc-dau": 20,
        "goc": 30
      },
      "ra": {
        "tam-xa": 35.34797566,
        "do-cao-cuc-dai": 5.102040816,
        "thoi-gian-bay": 2.040816327
      },
      "saiSo": 0.001
    },
    {
      "vao": {
        "van-toc-dau": 20,
        "goc": 60
      },
      "ra": {
        "tam-xa": 35.34797566,
        "do-cao-cuc-dai": 15.30612245,
        "thoi-gian-bay": 3.534797566
      },
      "saiSo": 0.001
    },
    {
      "vao": {
        "van-toc-dau": 10,
        "goc": 0,
        "do-cao-dau": 20
      },
      "ra": {
        "tam-xa": 20.20305089,
        "do-cao-cuc-dai": 20.0,
        "thoi-gian-bay": 2.020305089
      },
      "saiSo": 0.001
    },
    {
      "vao": {
        "van-toc-dau": 30,
        "goc": 45,
        "do-cao-dau": 10,
        "g": 1.6
      },
      "ra": {
        "tam-xa": 572.3282756,
        "do-cao-cuc-dai": 150.625,
        "thoi-gian-bay": 26.97981365
      },
      "saiSo": 0.001
    },
    {
      "vao": {
        "van-toc-dau": 15,
        "goc": 85
      },
      "ra": {
        "tam-xa": 3.986820406,
        "do-cao-cuc-dai": 11.39239144,
        "thoi-gian-bay": 3.049575606
      },
      "saiSo": 0.001
    }
  ]
}
```

`tools/vi/thi_nghiem_parts/mo_hinh/li-nem-xien.js`:

```js
(function (root) {
  'use strict';

  function thanhPhan(p) {
    var goc = p['goc'] * Math.PI / 180;
    return { vx: p['van-toc-dau'] * Math.cos(goc), vy: p['van-toc-dau'] * Math.sin(goc) };
  }

  function tinh(p) {
    var v = thanhPhan(p);
    var g = p['g'];
    var thoiGian = (v.vy + Math.sqrt(v.vy * v.vy + 2 * g * p['do-cao-dau'])) / g;
    return {
      'thoi-gian-bay': thoiGian,
      'tam-xa': v.vx * thoiGian,
      'do-cao-cuc-dai': p['do-cao-dau'] + v.vy * v.vy / (2 * g)
    };
  }

  function ve(ctx, p, t, kt, d) {
    var K = root.THI_NGHIEM_KHUNG;
    var v = thanhPhan(p);
    var le = 48;
    var tiLe = Math.min((kt.rong - 2 * le) / Math.max(d['tam-xa'], 1), (kt.cao - 2 * le) / Math.max(d['do-cao-cuc-dai'], 1));
    function diem(thoiDiem) {
      var x = v.vx * thoiDiem;
      var y = p['do-cao-dau'] + v.vy * thoiDiem - p['g'] * thoiDiem * thoiDiem / 2;
      return [le + x * tiLe, kt.cao - le - Math.max(y, 0) * tiLe];
    }
    ctx.strokeStyle = K.MAU.net; ctx.lineWidth = 2;
    ctx.beginPath(); ctx.moveTo(le - 20, kt.cao - le); ctx.lineTo(kt.rong - le + 20, kt.cao - le); ctx.stroke();
    if (p['do-cao-dau'] > 0) {
      ctx.fillStyle = K.MAU.nhat;
      ctx.fillRect(le - 20, kt.cao - le - p['do-cao-dau'] * tiLe, 20, p['do-cao-dau'] * tiLe);
    }
    ctx.setLineDash([6, 6]); ctx.strokeStyle = K.MAU.nhat; ctx.beginPath();
    for (var i = 0; i <= 60; i += 1) {
      var q = diem(d['thoi-gian-bay'] * i / 60);
      if (i === 0) { ctx.moveTo(q[0], q[1]); } else { ctx.lineTo(q[0], q[1]); }
    }
    ctx.stroke(); ctx.setLineDash([]);
    var vat = diem(Math.min(t, d['thoi-gian-bay']));
    ctx.fillStyle = K.MAU.phu; ctx.beginPath(); ctx.arc(vat[0], vat[1], 9, 0, 2 * Math.PI); ctx.fill();
    ctx.fillStyle = K.MAU.net; ctx.font = '15px ' + K.PHONG;
    ctx.fillText('t = ' + K.dinhDang(Math.min(t, d['thoi-gian-bay']), 2) + ' s', le, 24);
  }

  root.THI_NGHIEM_MO_HINH = { tinh: tinh, ve: ve, thoiLuong: function (p, d) { return d['thoi-gian-bay']; } };
})(typeof globalThis !== 'undefined' ? globalThis : this);
```

`tools/vi/thi_nghiem_parts/mo_hinh/li-mach-ohm.json`:

```json
{
  "ma": "li-mach-ohm",
  "ten": "Đoạn mạch nối tiếp và song song",
  "mon": "Vật lí",
  "hoatHinh": "khong",
  "thamSo": [
    {
      "ma": "suat-dien-dong",
      "ten": "Hiệu điện thế nguồn U",
      "kieu": "so",
      "donVi": "V",
      "min": 1.5,
      "max": 24,
      "buoc": 0.5,
      "macDinh": 12
    },
    {
      "ma": "dien-tro-1",
      "ten": "Điện trở R~1~",
      "kieu": "so",
      "donVi": "Ω",
      "min": 1,
      "max": 100,
      "buoc": 1,
      "macDinh": 10
    },
    {
      "ma": "dien-tro-2",
      "ten": "Điện trở R~2~",
      "kieu": "so",
      "donVi": "Ω",
      "min": 1,
      "max": 100,
      "buoc": 1,
      "macDinh": 20
    },
    {
      "ma": "kieu-mac",
      "ten": "Kiểu mắc",
      "kieu": "chon",
      "luaChon": [
        {
          "ma": "noi-tiep",
          "ten": "Nối tiếp"
        },
        {
          "ma": "song-song",
          "ten": "Song song"
        }
      ],
      "macDinh": "noi-tiep"
    }
  ],
  "daiLuongDo": [
    {
      "ma": "cuong-do-mach-chinh",
      "ten": "Cường độ mạch chính I",
      "donVi": "A",
      "saiSo": 0.01,
      "chuSo": 3
    },
    {
      "ma": "cuong-do-1",
      "ten": "Cường độ qua R~1~",
      "donVi": "A",
      "saiSo": 0.01,
      "chuSo": 3
    },
    {
      "ma": "cuong-do-2",
      "ten": "Cường độ qua R~2~",
      "donVi": "A",
      "saiSo": 0.01,
      "chuSo": 3
    },
    {
      "ma": "hieu-dien-the-1",
      "ten": "Hiệu điện thế hai đầu R~1~",
      "donVi": "V",
      "saiSo": 0.05,
      "chuSo": 2
    },
    {
      "ma": "hieu-dien-the-2",
      "ten": "Hiệu điện thế hai đầu R~2~",
      "donVi": "V",
      "saiSo": 0.05,
      "chuSo": 2
    },
    {
      "ma": "dien-tro-tuong-duong",
      "ten": "Điện trở tương đương",
      "donVi": "Ω",
      "saiSo": 0,
      "chuSo": 2
    }
  ],
  "congThuc": {
    "bieuThuc": "Nối tiếp: R = R~1~ + R~2~, I = U/R. Song song: 1/R = 1/R~1~ + 1/R~2~, I = I~1~ + I~2~.",
    "dieuKien": "Nguồn có điện trở trong bằng 0; dây nối và ampe kế có điện trở không đáng kể; vôn kế có điện trở rất lớn.",
    "nguon": ""
  },
  "bangKiem": [
    {
      "vao": {},
      "ra": {
        "cuong-do-mach-chinh": 0.4,
        "cuong-do-1": 0.4,
        "cuong-do-2": 0.4,
        "hieu-dien-the-1": 4.0,
        "hieu-dien-the-2": 8.0,
        "dien-tro-tuong-duong": 30.0
      },
      "saiSo": 0.0005
    },
    {
      "vao": {
        "kieu-mac": "song-song"
      },
      "ra": {
        "cuong-do-mach-chinh": 1.8,
        "cuong-do-1": 1.2,
        "cuong-do-2": 0.6,
        "hieu-dien-the-1": 12.0,
        "hieu-dien-the-2": 12.0,
        "dien-tro-tuong-duong": 6.666666667
      },
      "saiSo": 0.0005
    },
    {
      "vao": {
        "suat-dien-dong": 6,
        "dien-tro-1": 5,
        "dien-tro-2": 5
      },
      "ra": {
        "cuong-do-mach-chinh": 0.6,
        "cuong-do-1": 0.6,
        "cuong-do-2": 0.6,
        "hieu-dien-the-1": 3.0,
        "hieu-dien-the-2": 3.0,
        "dien-tro-tuong-duong": 10.0
      },
      "saiSo": 0.0005
    },
    {
      "vao": {
        "suat-dien-dong": 6,
        "dien-tro-1": 5,
        "dien-tro-2": 5,
        "kieu-mac": "song-song"
      },
      "ra": {
        "cuong-do-mach-chinh": 2.4,
        "cuong-do-1": 1.2,
        "cuong-do-2": 1.2,
        "hieu-dien-the-1": 6.0,
        "hieu-dien-the-2": 6.0,
        "dien-tro-tuong-duong": 2.5
      },
      "saiSo": 0.0005
    },
    {
      "vao": {
        "suat-dien-dong": 24,
        "dien-tro-1": 100,
        "dien-tro-2": 1,
        "kieu-mac": "song-song"
      },
      "ra": {
        "cuong-do-mach-chinh": 24.24,
        "cuong-do-1": 0.24,
        "cuong-do-2": 24.0,
        "hieu-dien-the-1": 24.0,
        "hieu-dien-the-2": 24.0,
        "dien-tro-tuong-duong": 0.9900990099
      },
      "saiSo": 0.0005
    }
  ]
}
```

`tools/vi/thi_nghiem_parts/mo_hinh/li-mach-ohm.js`:

```js
(function (root) {
  'use strict';

  function tinh(p) {
    var u = p['suat-dien-dong'], r1 = p['dien-tro-1'], r2 = p['dien-tro-2'];
    if (p['kieu-mac'] === 'song-song') {
      return {
        'dien-tro-tuong-duong': r1 * r2 / (r1 + r2),
        'cuong-do-mach-chinh': u / r1 + u / r2,
        'cuong-do-1': u / r1, 'cuong-do-2': u / r2,
        'hieu-dien-the-1': u, 'hieu-dien-the-2': u
      };
    }
    var dong = u / (r1 + r2);
    return {
      'dien-tro-tuong-duong': r1 + r2,
      'cuong-do-mach-chinh': dong,
      'cuong-do-1': dong, 'cuong-do-2': dong,
      'hieu-dien-the-1': dong * r1, 'hieu-dien-the-2': dong * r2
    };
  }

  function dienTro(ctx, K, x, y, nhan) {
    ctx.fillStyle = K.MAU.nen; ctx.fillRect(x - 45, y - 14, 90, 28);
    ctx.strokeRect(x - 45, y - 14, 90, 28);
    ctx.fillStyle = K.MAU.net; ctx.textAlign = 'center'; ctx.fillText(nhan, x, y + 5); ctx.textAlign = 'left';
  }

  function ve(ctx, p, t, kt) {
    var K = root.THI_NGHIEM_KHUNG;
    var trai = 70, phai = kt.rong - 70, tren = 70, duoi = kt.cao - 70, giua = kt.rong / 2;
    ctx.strokeStyle = K.MAU.net; ctx.lineWidth = 2; ctx.font = '15px ' + K.PHONG;
    ctx.strokeRect(trai, tren, phai - trai, duoi - tren);
    ctx.fillStyle = K.MAU.nen; ctx.fillRect(giua - 14, duoi - 20, 28, 40);
    ctx.beginPath(); ctx.moveTo(giua - 8, duoi - 18); ctx.lineTo(giua - 8, duoi + 18);
    ctx.moveTo(giua + 8, duoi - 9); ctx.lineTo(giua + 8, duoi + 9); ctx.stroke();
    ctx.fillStyle = K.MAU.net;
    ctx.fillText('U = ' + K.dinhDang(p['suat-dien-dong'], 1) + ' V', giua + 22, duoi + 28);
    var nhan1 = 'R₁ = ' + p['dien-tro-1'] + ' Ω', nhan2 = 'R₂ = ' + p['dien-tro-2'] + ' Ω';
    if (p['kieu-mac'] === 'song-song') {
      ctx.beginPath(); ctx.moveTo(giua - 110, tren); ctx.lineTo(giua - 110, tren + 80); ctx.lineTo(giua + 110, tren + 80);
      ctx.lineTo(giua + 110, tren); ctx.stroke();
      dienTro(ctx, K, giua, tren, nhan1);
      dienTro(ctx, K, giua, tren + 80, nhan2);
    } else {
      dienTro(ctx, K, giua - 90, tren, nhan1);
      dienTro(ctx, K, giua + 90, tren, nhan2);
    }
    ctx.fillText(p['kieu-mac'] === 'song-song' ? 'Mắc song song' : 'Mắc nối tiếp', 16, 24);
  }

  root.THI_NGHIEM_MO_HINH = { tinh: tinh, ve: ve };
})(typeof globalThis !== 'undefined' ? globalThis : this);
```

- [ ] **Step 4: Chạy lại**

Run: `venv\Scripts\python.exe -m unittest tools.vi.tests.test_thi_nghiem_mo_hinh -v`
Expected: PASS (11 test). `test_javascript_matches_the_python_reference_on_a_grid` so 300 điểm mỗi mẫu.

- [ ] **Step 5: Commit**

```bash
git add tools/vi/thi_nghiem_parts/mo_hinh tools/vi/tests/test_thi_nghiem_mo_hinh.py
git commit -m "feat(vi): add projectile and resistor circuit experiment models"
```

---

### Task 5: Ba mô hình Hoá học — chuẩn độ, cân bằng N₂O₄ ⇌ 2NO₂, tốc độ phản ứng

Hằng số dùng trong ba mô hình này là dữ liệu khoa học, không được đổi: K~w~ = 1,0·10⁻¹⁴; K~a~(CH₃COOH) = 1,75·10⁻⁵; Δ~r~H° = 57 200 J/mol; Δ~r~S° = 175,83 J/(mol·K); R = 8,314462618 J/(mol·K) và 0,08314462618 L·bar/(mol·K); k(25 °C) = 1,25·10⁻³ s⁻¹; E~a~ = 50 000 và 45 000 J/mol; ΔC = 0,005 mol/L.

**Files:**
- Create: `tools/vi/thi_nghiem_parts/mo_hinh/hoa-chuan-do.json`, `hoa-chuan-do.js`
- Create: `tools/vi/thi_nghiem_parts/mo_hinh/hoa-can-bang-no2.json`, `hoa-can-bang-no2.js`
- Create: `tools/vi/thi_nghiem_parts/mo_hinh/hoa-toc-do.json`, `hoa-toc-do.js`
- Modify: `tools/vi/tests/test_thi_nghiem_mo_hinh.py` (hằng `MODELS`)

**Interfaces:**
- Consumes: khuôn mô hình (Task 3), `tham_chieu.hoa_chuan_do`, `hoa_can_bang_no2`, `hoa_toc_do` (Task 2).
- Produces: mã mẫu `hoa-chuan-do`, `hoa-can-bang-no2`, `hoa-toc-do` (`hoatHinh: mot-lan`).

- [ ] **Step 1: Mở rộng `MODELS` trong test**

Thay dòng `MODELS = ...` bằng:

```python
MODELS = ("hoa-can-bang-no2", "hoa-chuan-do", "hoa-toc-do", "li-con-lac-don", "li-mach-ohm", "li-nem-xien")
```

- [ ] **Step 2: Chạy để thấy trượt**

Run: `venv\Scripts\python.exe -m unittest tools.vi.tests.test_thi_nghiem_mo_hinh -v`
Expected: FAIL ở `test_library_has_the_models`.

- [ ] **Step 3: Viết ba mô hình**

`tools/vi/thi_nghiem_parts/mo_hinh/hoa-chuan-do.json`:

```json
{
  "ma": "hoa-chuan-do",
  "ten": "Chuẩn độ acid – base bằng dung dịch NaOH",
  "mon": "Hoá học",
  "hoatHinh": "khong",
  "thamSo": [
    {
      "ma": "loai-acid",
      "ten": "Acid cần chuẩn độ",
      "kieu": "chon",
      "luaChon": [
        {
          "ma": "hcl",
          "ten": "HCl (acid mạnh)"
        },
        {
          "ma": "ch3cooh",
          "ten": "CH~3~COOH (acid yếu)"
        }
      ],
      "macDinh": "hcl"
    },
    {
      "ma": "nong-do-acid",
      "ten": "Nồng độ acid",
      "kieu": "so",
      "donVi": "mol/L",
      "min": 0.01,
      "max": 0.5,
      "buoc": 0.01,
      "macDinh": 0.1
    },
    {
      "ma": "the-tich-acid",
      "ten": "Thể tích acid",
      "kieu": "so",
      "donVi": "mL",
      "min": 10,
      "max": 50,
      "buoc": 5,
      "macDinh": 20
    },
    {
      "ma": "nong-do-base",
      "ten": "Nồng độ NaOH",
      "kieu": "so",
      "donVi": "mol/L",
      "min": 0.01,
      "max": 0.5,
      "buoc": 0.01,
      "macDinh": 0.1
    },
    {
      "ma": "the-tich-base",
      "ten": "Thể tích NaOH đã nhỏ",
      "kieu": "so",
      "donVi": "mL",
      "min": 0,
      "max": 50,
      "buoc": 0.1,
      "macDinh": 0
    },
    {
      "ma": "chi-thi",
      "ten": "Chất chỉ thị",
      "kieu": "chon",
      "luaChon": [
        {
          "ma": "phenolphtalein",
          "ten": "Phenolphtalein"
        },
        {
          "ma": "metyl-da-cam",
          "ten": "Methyl da cam"
        },
        {
          "ma": "bromothymol",
          "ten": "Bromothymol xanh"
        }
      ],
      "macDinh": "phenolphtalein"
    }
  ],
  "daiLuongDo": [
    {
      "ma": "ph",
      "ten": "pH của dung dịch",
      "donVi": "",
      "saiSo": 0.05,
      "chuSo": 2
    }
  ],
  "congThuc": {
    "bieuThuc": "Bảo toàn điện tích: [H^+^] + [Na^+^] = [OH^−^] + [A^−^]; K~w~ = [H^+^][OH^−^]; với acid yếu [A^−^] = C·K~a~/(K~a~ + [H^+^]).",
    "dieuKien": "Dung dịch loãng ở 25 °C (K~w~ = 1,0·10^−14^); coi hoạt độ bằng nồng độ; thể tích cộng tính.",
    "nguon": "K~a~(CH~3~COOH) = 1,75·10^−5^ ở 25 °C (CRC Handbook of Chemistry and Physics)."
  },
  "bangKiem": [
    {
      "vao": {
        "the-tich-base": 0
      },
      "ra": {
        "ph": 1.0
      },
      "saiSo": 0.001
    },
    {
      "vao": {
        "the-tich-base": 10
      },
      "ra": {
        "ph": 1.477121255
      },
      "saiSo": 0.001
    },
    {
      "vao": {
        "the-tich-base": 20
      },
      "ra": {
        "ph": 7.0
      },
      "saiSo": 0.001
    },
    {
      "vao": {
        "the-tich-base": 30
      },
      "ra": {
        "ph": 12.30103056
      },
      "saiSo": 0.001
    },
    {
      "vao": {
        "loai-acid": "ch3cooh",
        "the-tich-base": 0
      },
      "ra": {
        "ph": 2.881353541
      },
      "saiSo": 0.001
    },
    {
      "vao": {
        "loai-acid": "ch3cooh",
        "the-tich-base": 10
      },
      "ra": {
        "ph": 4.757417468
      },
      "saiSo": 0.001
    },
    {
      "vao": {
        "loai-acid": "ch3cooh",
        "the-tich-base": 20
      },
      "ra": {
        "ph": 8.728018764
      },
      "saiSo": 0.001
    },
    {
      "vao": {
        "loai-acid": "ch3cooh",
        "the-tich-base": 25
      },
      "ra": {
        "ph": 12.04575758
      },
      "saiSo": 0.001
    }
  ]
}
```

`tools/vi/thi_nghiem_parts/mo_hinh/hoa-chuan-do.js`:

```js
(function (root) {
  'use strict';
  var KW = 1.0e-14;
  var KA = { 'ch3cooh': 1.75e-5 };
  // Khoảng đổi màu: [pH bắt đầu, pH kết thúc, màu dạng acid, màu dạng base] (màu là [r, g, b, độ đậm]).
  var CHI_THI = {
    'phenolphtalein': [8.2, 10.0, [255, 255, 255, 0.05], [236, 72, 153, 0.75]],
    'metyl-da-cam': [3.1, 4.4, [239, 68, 68, 0.7], [250, 204, 21, 0.7]],
    'bromothymol': [6.0, 7.6, [250, 204, 21, 0.7], [37, 99, 235, 0.7]]
  };

  function tinh(p) {
    var tong = p['the-tich-acid'] + p['the-tich-base'];
    var acid = p['nong-do-acid'] * p['the-tich-acid'] / tong;
    var natri = p['nong-do-base'] * p['the-tich-base'] / tong;
    var h;
    if (p['loai-acid'] === 'hcl') {
      var lech = acid - natri;
      h = (lech + Math.sqrt(lech * lech + 4 * KW)) / 2;
    } else {
      var ka = KA[p['loai-acid']];
      var thap = 0, cao = 14;
      for (var lan = 0; lan < 200; lan += 1) {
        var giua = (thap + cao) / 2;
        var thu = Math.pow(10, -giua);
        if (thu + natri - KW / thu - acid * ka / (ka + thu) > 0) { thap = giua; } else { cao = giua; }
      }
      h = Math.pow(10, -(thap + cao) / 2);
    }
    return { 'ph': -Math.log(h) / Math.LN10 };
  }

  function mauDungDich(chiThi, ph) {
    var ct = CHI_THI[chiThi];
    var k = Math.max(0, Math.min(1, (ph - ct[0]) / (ct[1] - ct[0])));
    var m = ct[2].map(function (a, i) { return a + (ct[3][i] - a) * k; });
    return 'rgba(' + Math.round(m[0]) + ',' + Math.round(m[1]) + ',' + Math.round(m[2]) + ',' + m[3].toFixed(2) + ')';
  }

  function ve(ctx, p, t, kt, d) {
    var K = root.THI_NGHIEM_KHUNG;
    var giua = kt.rong / 2, day = kt.cao - 50;
    ctx.strokeStyle = K.MAU.net; ctx.lineWidth = 2; ctx.font = '15px ' + K.PHONG;
    // Buret: vạch 0 ở trên, mực dung dịch hạ dần theo thể tích đã nhỏ (buret 50 mL).
    var buretTren = 30, buretCao = kt.cao * 0.45;
    ctx.strokeRect(giua - 9, buretTren, 18, buretCao);
    ctx.fillStyle = 'rgba(37,99,235,0.25)';
    var daNho = buretCao * p['the-tich-base'] / 50;
    ctx.fillRect(giua - 8, buretTren + daNho, 16, buretCao - daNho);
    ctx.beginPath(); ctx.moveTo(giua, buretTren + buretCao); ctx.lineTo(giua, buretTren + buretCao + 18); ctx.stroke();
    // Bình tam giác.
    var co = buretTren + buretCao + 24;
    ctx.beginPath(); ctx.moveTo(giua - 16, co); ctx.lineTo(giua - 16, co + 26); ctx.lineTo(giua - 90, day);
    ctx.lineTo(giua + 90, day); ctx.lineTo(giua + 16, co + 26); ctx.lineTo(giua + 16, co); ctx.stroke();
    var muc = Math.min(0.75, (p['the-tich-acid'] + p['the-tich-base']) / 130);
    var caoLong = (day - co - 26) * muc;
    // Tô hai lớp: nước nhạt để thấy mực dung dịch, rồi màu của chất chỉ thị.
    ['rgba(186,230,253,0.45)', mauDungDich(p['chi-thi'], d['ph'])].forEach(function (mau) {
      ctx.fillStyle = mau;
      ctx.beginPath(); ctx.moveTo(giua - 90, day); ctx.lineTo(giua + 90, day);
      ctx.lineTo(giua + 90 - 74 * caoLong / (day - co - 26), day - caoLong);
      ctx.lineTo(giua - 90 + 74 * caoLong / (day - co - 26), day - caoLong); ctx.closePath(); ctx.fill();
    });
    ctx.fillStyle = K.MAU.net;
    ctx.fillText('NaOH đã nhỏ: ' + K.dinhDang(p['the-tich-base'], 1) + ' mL', 16, 24);
  }

  root.THI_NGHIEM_MO_HINH = { tinh: tinh, ve: ve };
})(typeof globalThis !== 'undefined' ? globalThis : this);
```

`tools/vi/thi_nghiem_parts/mo_hinh/hoa-can-bang-no2.json`:

```json
{
  "ma": "hoa-can-bang-no2",
  "ten": "Cân bằng N₂O₄ ⇌ 2NO₂",
  "mon": "Hoá học",
  "hoatHinh": "khong",
  "thamSo": [
    {
      "ma": "nhiet-do",
      "ten": "Nhiệt độ",
      "kieu": "so",
      "donVi": "°C",
      "min": 0,
      "max": 100,
      "buoc": 5,
      "macDinh": 25
    },
    {
      "ma": "ap-suat",
      "ten": "Áp suất chung",
      "kieu": "so",
      "donVi": "bar",
      "min": 0.5,
      "max": 5,
      "buoc": 0.1,
      "macDinh": 1
    }
  ],
  "daiLuongDo": [
    {
      "ma": "phan-mol-no2",
      "ten": "Phần mol NO~2~",
      "donVi": "",
      "saiSo": 0.005,
      "chuSo": 3
    },
    {
      "ma": "do-phan-li",
      "ten": "Độ phân li của N~2~O~4~",
      "donVi": "",
      "saiSo": 0.005,
      "chuSo": 3
    },
    {
      "ma": "nong-do-no2",
      "ten": "Nồng độ NO~2~ (độ đậm màu nâu)",
      "donVi": "mol/L",
      "saiSo": 0.0005,
      "chuSo": 4
    },
    {
      "ma": "kp",
      "ten": "Hằng số cân bằng K~p~",
      "donVi": "",
      "saiSo": 0,
      "chuSo": 3
    }
  ],
  "congThuc": {
    "bieuThuc": "N~2~O~4~(g) ⇌ 2NO~2~(g), Δ~r~H° = +57,2 kJ; K~p~ = exp(−(Δ~r~H° − TΔ~r~S°)/RT); K~p~ = x^2^P/(1 − x) với x là phần mol NO~2~.",
    "dieuKien": "Hỗn hợp khí lí tưởng đã đạt cân bằng; coi Δ~r~H° và Δ~r~S° không đổi trong khoảng 0–100 °C; áp suất tính bằng bar.",
    "nguon": "Δ~r~H° = 57,20 kJ/mol và Δ~r~S° = 175,83 J/(mol·K), tính từ Δ~f~H° và S° ở 298 K (Atkins, Physical Chemistry, bảng dữ liệu nhiệt động)."
  },
  "bangKiem": [
    {
      "vao": {},
      "ra": {
        "phan-mol-no2": 0.3156788424,
        "do-phan-li": 0.1874220014,
        "nong-do-no2": 0.01273434104,
        "kp": 0.1456233384
      },
      "saiSo": 5e-05
    },
    {
      "vao": {
        "nhiet-do": 0
      },
      "ra": {
        "phan-mol-no2": 0.1242296573,
        "do-phan-li": 0.06622860724,
        "nong-do-no2": 0.005470032381,
        "kp": 0.0176222087
      },
      "saiSo": 5e-05
    },
    {
      "vao": {
        "nhiet-do": 100
      },
      "ra": {
        "phan-mol-no2": 0.9411256005,
        "do-phan-li": 0.8887981436,
        "nong-do-no2": 0.03033401913,
        "kp": 15.04418566
      },
      "saiSo": 5e-05
    },
    {
      "vao": {
        "ap-suat": 5
      },
      "ra": {
        "phan-mol-no2": 0.1567173475,
        "do-phan-li": 0.08502079012,
        "nong-do-no2": 0.03160953289,
        "kp": 0.1456233384
      },
      "saiSo": 5e-05
    },
    {
      "vao": {
        "nhiet-do": 50,
        "ap-suat": 0.5
      },
      "ra": {
        "phan-mol-no2": 0.7097801717,
        "do-phan-li": 0.5501234411,
        "nong-do-no2": 0.01320856148,
        "kp": 0.8679418892
      },
      "saiSo": 5e-05
    },
    {
      "vao": {
        "nhiet-do": 75,
        "ap-suat": 2
      },
      "ra": {
        "phan-mol-no2": 0.7321756231,
        "do-phan-li": 0.5775055571,
        "nong-do-no2": 0.05058766997,
        "kp": 4.003228901
      },
      "saiSo": 5e-05
    }
  ]
}
```

`tools/vi/thi_nghiem_parts/mo_hinh/hoa-can-bang-no2.js`:

```js
(function (root) {
  'use strict';
  var DELTA_H = 57200, DELTA_S = 175.83, R = 8.314462618, R_BAR = 0.08314462618;

  function tinh(p) {
    var nhietDo = p['nhiet-do'] + 273.15;
    var apSuat = p['ap-suat'];
    var kp = Math.exp(-(DELTA_H - nhietDo * DELTA_S) / (R * nhietDo));
    var phanMol = (-kp + Math.sqrt(kp * kp + 4 * kp * apSuat)) / (2 * apSuat);
    return {
      'kp': kp,
      'phan-mol-no2': phanMol,
      'do-phan-li': Math.sqrt(kp / (kp + 4 * apSuat)),
      'nong-do-no2': phanMol * apSuat / (R_BAR * nhietDo)
    };
  }

  function ve(ctx, p, t, kt, d) {
    var K = root.THI_NGHIEM_KHUNG;
    var trai = kt.rong * 0.2, rong = kt.rong * 0.5, day = kt.cao - 60, tren = 50;
    // Xi lanh: thể tích tỉ lệ nghịch với áp suất (so với 0,5 bar là đầy xi lanh).
    var caoKhi = (day - tren) * 0.5 / p['ap-suat'];
    ctx.font = '15px ' + K.PHONG; ctx.lineWidth = 2; ctx.strokeStyle = K.MAU.net;
    ctx.fillStyle = 'rgba(146,64,14,' + Math.min(0.9, d['nong-do-no2'] / 0.05).toFixed(3) + ')';
    ctx.fillRect(trai, day - caoKhi, rong, caoKhi);
    ctx.beginPath(); ctx.moveTo(trai, tren); ctx.lineTo(trai, day); ctx.lineTo(trai + rong, day); ctx.lineTo(trai + rong, tren); ctx.stroke();
    ctx.fillStyle = K.MAU.nhat; ctx.fillRect(trai + 2, day - caoKhi - 16, rong - 4, 16);
    ctx.fillRect(trai + rong / 2 - 6, tren - 20, 12, day - caoKhi - 16 - tren + 20);
    // Nhiệt kế.
    var nk = trai + rong + 60, cot = (day - tren) * p['nhiet-do'] / 100;
    ctx.strokeRect(nk, tren, 14, day - tren);
    ctx.fillStyle = K.MAU.xau; ctx.fillRect(nk + 2, day - cot, 10, cot);
    ctx.fillStyle = K.MAU.net;
    ctx.fillText(p['nhiet-do'] + ' °C', nk - 8, day + 22);
    ctx.fillText('P = ' + K.dinhDang(p['ap-suat'], 1) + ' bar', trai, day + 22);
    ctx.fillText('N₂O₄ (không màu) ⇌ 2NO₂ (nâu đỏ)', 16, 24);
  }

  root.THI_NGHIEM_MO_HINH = { tinh: tinh, ve: ve };
})(typeof globalThis !== 'undefined' ? globalThis : this);
```

`tools/vi/thi_nghiem_parts/mo_hinh/hoa-toc-do.json`:

```json
{
  "ma": "hoa-toc-do",
  "ten": "Các yếu tố ảnh hưởng đến tốc độ phản ứng",
  "mon": "Hoá học",
  "hoatHinh": "mot-lan",
  "thamSo": [
    {
      "ma": "nong-do",
      "ten": "Nồng độ chất tham gia",
      "kieu": "so",
      "donVi": "mol/L",
      "min": 0.02,
      "max": 0.5,
      "buoc": 0.02,
      "macDinh": 0.1
    },
    {
      "ma": "nhiet-do",
      "ten": "Nhiệt độ",
      "kieu": "so",
      "donVi": "°C",
      "min": 10,
      "max": 60,
      "buoc": 5,
      "macDinh": 25
    },
    {
      "ma": "xuc-tac",
      "ten": "Chất xúc tác",
      "kieu": "chon",
      "luaChon": [
        {
          "ma": "khong",
          "ten": "Không có"
        },
        {
          "ma": "co",
          "ten": "Có xúc tác"
        }
      ],
      "macDinh": "khong"
    }
  ],
  "daiLuongDo": [
    {
      "ma": "thoi-gian",
      "ten": "Thời gian đến khi vẩn đục che dấu X",
      "donVi": "s",
      "saiSo": 0.5,
      "chuSo": 1
    }
  ],
  "congThuc": {
    "bieuThuc": "v = k·C; k = A·exp(−E~a~/RT); thời gian t = ΔC/v với ΔC = 0,005 mol/L.",
    "dieuKien": "Phản ứng giả định bậc 1 theo chất tham gia, dùng tốc độ đầu; E~a~ = 50 kJ/mol khi không có xúc tác và 45 kJ/mol khi có; k = 1,25·10^−3^ s^−1^ ở 25 °C. Số liệu minh hoạ quy luật, không phải của một phản ứng cụ thể.",
    "nguon": ""
  },
  "bangKiem": [
    {
      "vao": {},
      "ra": {
        "thoi-gian": 40.0
      },
      "saiSo": 0.001
    },
    {
      "vao": {
        "nong-do": 0.2
      },
      "ra": {
        "thoi-gian": 20.0
      },
      "saiSo": 0.001
    },
    {
      "vao": {
        "nhiet-do": 35
      },
      "ra": {
        "thoi-gian": 20.78715968
      },
      "saiSo": 0.001
    },
    {
      "vao": {
        "nhiet-do": 45
      },
      "ra": {
        "thoi-gian": 11.25641706
      },
      "saiSo": 0.001
    },
    {
      "vao": {
        "xuc-tac": "co"
      },
      "ra": {
        "thoi-gian": 5.322282096
      },
      "saiSo": 0.001
    },
    {
      "vao": {
        "nong-do": 0.5,
        "nhiet-do": 60,
        "xuc-tac": "co"
      },
      "ra": {
        "thoi-gian": 0.1580840914
      },
      "saiSo": 0.001
    }
  ]
}
```

`tools/vi/thi_nghiem_parts/mo_hinh/hoa-toc-do.js`:

```js
(function (root) {
  'use strict';
  var R = 8.314462618, K25 = 0.00125, T25 = 298.15;
  var EA = { 'khong': 50000, 'co': 45000 };
  var LUONG_PHAN_UNG = 0.005;
  var A = K25 * Math.exp(EA['khong'] / (R * T25));

  function tinh(p) {
    var nhietDo = p['nhiet-do'] + 273.15;
    var k = A * Math.exp(-EA[p['xuc-tac']] / (R * nhietDo));
    return { 'thoi-gian': LUONG_PHAN_UNG / (k * p['nong-do']) };
  }

  function ve(ctx, p, t, kt, d) {
    var K = root.THI_NGHIEM_KHUNG;
    var giua = kt.rong / 2, tren = 70, day = kt.cao - 70, nuaRong = 110;
    var duc = Math.max(0, Math.min(1, t / d['thoi-gian']));
    ctx.font = '15px ' + K.PHONG; ctx.lineWidth = 2; ctx.strokeStyle = K.MAU.net;
    // Tờ giấy có dấu X đặt dưới cốc, nhìn từ trên xuống qua dung dịch.
    ctx.fillStyle = K.MAU.nen; ctx.fillRect(giua - nuaRong, tren, 2 * nuaRong, day - tren);
    ctx.lineWidth = 8; ctx.beginPath();
    ctx.moveTo(giua - 50, (tren + day) / 2 - 50); ctx.lineTo(giua + 50, (tren + day) / 2 + 50);
    ctx.moveTo(giua + 50, (tren + day) / 2 - 50); ctx.lineTo(giua - 50, (tren + day) / 2 + 50); ctx.stroke();
    ctx.fillStyle = 'rgba(253,230,138,' + (0.97 * duc).toFixed(3) + ')';
    ctx.fillRect(giua - nuaRong, tren, 2 * nuaRong, day - tren);
    ctx.lineWidth = 2; ctx.strokeRect(giua - nuaRong, tren, 2 * nuaRong, day - tren);
    ctx.fillStyle = K.MAU.net;
    ctx.fillText('Đồng hồ: ' + K.dinhDang(Math.min(t, d['thoi-gian']), 1) + ' s', 16, 24);
    ctx.fillText(p['xuc-tac'] === 'co' ? 'Có chất xúc tác' : 'Không có chất xúc tác', 16, 46);
    ctx.fillText(p['nhiet-do'] + ' °C', giua + nuaRong + 16, tren + 16);
  }

  root.THI_NGHIEM_MO_HINH = { tinh: tinh, ve: ve, thoiLuong: function (p, d) { return d['thoi-gian']; } };
})(typeof globalThis !== 'undefined' ? globalThis : this);
```

- [ ] **Step 4: Chạy lại**

Run: `venv\Scripts\python.exe -m unittest tools.vi.tests.test_thi_nghiem_mo_hinh tools.vi.tests.test_thi_nghiem_tham_chieu -v`
Expected: PASS (11 + 8 test).

- [ ] **Step 5: Commit**

```bash
git add tools/vi/thi_nghiem_parts/mo_hinh tools/vi/tests/test_thi_nghiem_mo_hinh.py
git commit -m "feat(vi): add titration, NO2 equilibrium and reaction rate experiment models"
```

---

### Task 6: Hai mô hình Toán — khảo sát hàm số và xác suất thực nghiệm

**Files:**
- Create: `tools/vi/thi_nghiem_parts/mo_hinh/toan-ham-so.json`, `toan-ham-so.js`
- Create: `tools/vi/thi_nghiem_parts/mo_hinh/toan-xac-suat.json`, `toan-xac-suat.js`
- Modify: `tools/vi/tests/test_thi_nghiem_mo_hinh.py` (hằng `MODELS`)

**Interfaces:**
- Consumes: khuôn mô hình (Task 3), `THI_NGHIEM_KHUNG.taoNgauNhien` (Task 1), `tham_chieu.toan_ham_so`, `toan_xac_suat` (Task 2).
- Produces: mã mẫu `toan-ham-so` (đại lượng có thể là `null`), `toan-xac-suat` (kết quả xác định theo `hat-giong`).

- [ ] **Step 1: Mở rộng `MODELS` trong test**

Thay dòng `MODELS = ...` bằng:

```python
MODELS = (
    "hoa-can-bang-no2", "hoa-chuan-do", "hoa-toc-do", "li-con-lac-don",
    "li-mach-ohm", "li-nem-xien", "toan-ham-so", "toan-xac-suat",
)
```

- [ ] **Step 2: Chạy để thấy trượt**

Run: `venv\Scripts\python.exe -m unittest tools.vi.tests.test_thi_nghiem_mo_hinh -v`
Expected: FAIL ở `test_library_has_the_models`.

- [ ] **Step 3: Viết hai mô hình**

`tools/vi/thi_nghiem_parts/mo_hinh/toan-ham-so.json`:

```json
{
  "ma": "toan-ham-so",
  "ten": "Khảo sát hàm số y = ax³ + bx² + cx + d",
  "mon": "Toán",
  "hoatHinh": "khong",
  "thamSo": [
    {
      "ma": "a",
      "ten": "Hệ số a",
      "kieu": "so",
      "donVi": "",
      "min": -5,
      "max": 5,
      "buoc": 0.5,
      "macDinh": 1
    },
    {
      "ma": "b",
      "ten": "Hệ số b",
      "kieu": "so",
      "donVi": "",
      "min": -5,
      "max": 5,
      "buoc": 0.5,
      "macDinh": 0
    },
    {
      "ma": "c",
      "ten": "Hệ số c",
      "kieu": "so",
      "donVi": "",
      "min": -5,
      "max": 5,
      "buoc": 0.5,
      "macDinh": -3
    },
    {
      "ma": "d",
      "ten": "Hệ số d",
      "kieu": "so",
      "donVi": "",
      "min": -5,
      "max": 5,
      "buoc": 0.5,
      "macDinh": 0
    },
    {
      "ma": "x0",
      "ten": "Hoành độ tiếp điểm x~0~",
      "kieu": "so",
      "donVi": "",
      "min": -5,
      "max": 5,
      "buoc": 0.1,
      "macDinh": 1
    }
  ],
  "daiLuongDo": [
    {
      "ma": "gia-tri",
      "ten": "Giá trị y(x~0~)",
      "donVi": "",
      "saiSo": 0,
      "chuSo": 3
    },
    {
      "ma": "he-so-goc",
      "ten": "Hệ số góc tiếp tuyến y'(x~0~)",
      "donVi": "",
      "saiSo": 0,
      "chuSo": 3
    },
    {
      "ma": "hoanh-do-cuc-dai",
      "ten": "Hoành độ điểm cực đại",
      "donVi": "",
      "saiSo": 0,
      "chuSo": 3
    },
    {
      "ma": "hoanh-do-cuc-tieu",
      "ten": "Hoành độ điểm cực tiểu",
      "donVi": "",
      "saiSo": 0,
      "chuSo": 3
    }
  ],
  "congThuc": {
    "bieuThuc": "y' = 3ax^2^ + 2bx + c; cực trị tại nghiệm của y' = 0 nơi y' đổi dấu; a = 0 thì là hàm bậc hai, đỉnh tại x = −c/(2b).",
    "dieuKien": "Hệ số thực; a = 0 và b = 0 thì hàm bậc nhất hoặc hằng, không có cực trị.",
    "nguon": ""
  },
  "bangKiem": [
    {
      "vao": {},
      "ra": {
        "gia-tri": -2.0,
        "he-so-goc": 0.0,
        "hoanh-do-cuc-dai": -1.0,
        "hoanh-do-cuc-tieu": 1.0
      },
      "saiSo": 1e-06
    },
    {
      "vao": {
        "x0": -1
      },
      "ra": {
        "gia-tri": 2.0,
        "he-so-goc": 0.0,
        "hoanh-do-cuc-dai": -1.0,
        "hoanh-do-cuc-tieu": 1.0
      },
      "saiSo": 1e-06
    },
    {
      "vao": {
        "a": -1,
        "b": 3,
        "c": 0,
        "d": 1,
        "x0": 2
      },
      "ra": {
        "gia-tri": 5.0,
        "he-so-goc": 0.0,
        "hoanh-do-cuc-dai": 2.0,
        "hoanh-do-cuc-tieu": 0.0
      },
      "saiSo": 1e-06
    },
    {
      "vao": {
        "a": 0,
        "b": 2,
        "c": -4,
        "d": 1,
        "x0": 0
      },
      "ra": {
        "gia-tri": 1.0,
        "he-so-goc": -4.0,
        "hoanh-do-cuc-dai": null,
        "hoanh-do-cuc-tieu": 1.0
      },
      "saiSo": 1e-06
    },
    {
      "vao": {
        "a": 1,
        "b": 0,
        "c": 3,
        "d": 0,
        "x0": 0.5
      },
      "ra": {
        "gia-tri": 1.625,
        "he-so-goc": 3.75,
        "hoanh-do-cuc-dai": null,
        "hoanh-do-cuc-tieu": null
      },
      "saiSo": 1e-06
    },
    {
      "vao": {
        "a": 0,
        "b": 0,
        "c": 2,
        "d": -1,
        "x0": 3
      },
      "ra": {
        "gia-tri": 5.0,
        "he-so-goc": 2.0,
        "hoanh-do-cuc-dai": null,
        "hoanh-do-cuc-tieu": null
      },
      "saiSo": 1e-06
    }
  ]
}
```

`tools/vi/thi_nghiem_parts/mo_hinh/toan-ham-so.js`:

```js
(function (root) {
  'use strict';

  function giaTri(p, x) { return ((p['a'] * x + p['b']) * x + p['c']) * x + p['d']; }
  function daoHam(p, x) { return (3 * p['a'] * x + 2 * p['b']) * x + p['c']; }

  function tinh(p) {
    var a = p['a'], b = p['b'], c = p['c'];
    var cucDai = null, cucTieu = null;
    if (a !== 0) {
      var biet = 4 * b * b - 12 * a * c;
      if (biet > 0) {
        var x1 = (-2 * b - Math.sqrt(biet)) / (6 * a);
        var x2 = (-2 * b + Math.sqrt(biet)) / (6 * a);
        var nho = Math.min(x1, x2), lon = Math.max(x1, x2);
        if (a > 0) { cucDai = nho; cucTieu = lon; } else { cucTieu = nho; cucDai = lon; }
      }
    } else if (b !== 0) {
      if (b < 0) { cucDai = -c / (2 * b); } else { cucTieu = -c / (2 * b); }
    }
    return {
      'gia-tri': giaTri(p, p['x0']),
      'he-so-goc': daoHam(p, p['x0']),
      'hoanh-do-cuc-dai': cucDai === null ? null : cucDai + 0,
      'hoanh-do-cuc-tieu': cucTieu === null ? null : cucTieu + 0
    };
  }

  function ve(ctx, p, t, kt, d) {
    var K = root.THI_NGHIEM_KHUNG;
    var BIEN = 6;
    var tiLe = Math.min(kt.rong, kt.cao) / (2 * BIEN);
    var goc = [kt.rong / 2, kt.cao / 2];
    function diem(x, y) { return [goc[0] + x * tiLe, goc[1] - y * tiLe]; }
    ctx.font = '13px ' + K.PHONG; ctx.lineWidth = 1; ctx.strokeStyle = '#e2e8f0';
    for (var i = -BIEN; i <= BIEN; i += 1) {
      ctx.beginPath(); ctx.moveTo(diem(i, 0)[0], 0); ctx.lineTo(diem(i, 0)[0], kt.cao);
      ctx.moveTo(0, diem(0, i)[1]); ctx.lineTo(kt.rong, diem(0, i)[1]); ctx.stroke();
    }
    ctx.strokeStyle = K.MAU.net; ctx.lineWidth = 1.5;
    ctx.beginPath(); ctx.moveTo(0, goc[1]); ctx.lineTo(kt.rong, goc[1]); ctx.moveTo(goc[0], 0); ctx.lineTo(goc[0], kt.cao); ctx.stroke();
    ctx.fillStyle = K.MAU.net; ctx.fillText('x', kt.rong - 14, goc[1] - 6); ctx.fillText('y', goc[0] + 6, 14);
    ctx.fillText('1', diem(1, 0)[0] - 3, goc[1] + 15); ctx.fillText('1', goc[0] - 14, diem(0, 1)[1] + 4);
    ctx.strokeStyle = K.MAU.chinh; ctx.lineWidth = 2.5; ctx.beginPath();
    var dangVe = false;
    for (var buoc = 0; buoc <= 600; buoc += 1) {
      var x = -BIEN * kt.rong / kt.cao + buoc * (2 * BIEN * kt.rong / kt.cao) / 600;
      var q = diem(x, giaTri(p, x));
      if (q[1] < -2000 || q[1] > kt.cao + 2000) { dangVe = false; continue; }
      if (dangVe) { ctx.lineTo(q[0], q[1]); } else { ctx.moveTo(q[0], q[1]); dangVe = true; }
    }
    ctx.stroke();
    var tiep = diem(p['x0'], d['gia-tri']);
    ctx.strokeStyle = K.MAU.phu; ctx.lineWidth = 2; ctx.beginPath();
    ctx.moveTo(tiep[0] - 3 * tiLe, tiep[1] + 3 * tiLe * d['he-so-goc']);
    ctx.lineTo(tiep[0] + 3 * tiLe, tiep[1] - 3 * tiLe * d['he-so-goc']); ctx.stroke();
    ctx.fillStyle = K.MAU.phu; ctx.beginPath(); ctx.arc(tiep[0], tiep[1], 5, 0, 2 * Math.PI); ctx.fill();
    ['hoanh-do-cuc-dai', 'hoanh-do-cuc-tieu'].forEach(function (ma) {
      if (d[ma] === null) { return; }
      var ct = diem(d[ma], giaTri(p, d[ma]));
      ctx.fillStyle = K.MAU.xau; ctx.beginPath(); ctx.arc(ct[0], ct[1], 4, 0, 2 * Math.PI); ctx.fill();
    });
  }

  root.THI_NGHIEM_MO_HINH = { tinh: tinh, ve: ve };
})(typeof globalThis !== 'undefined' ? globalThis : this);
```

`tools/vi/thi_nghiem_parts/mo_hinh/toan-xac-suat.json`:

```json
{
  "ma": "toan-xac-suat",
  "ten": "Xác suất thực nghiệm",
  "mon": "Toán",
  "hoatHinh": "khong",
  "thamSo": [
    {
      "ma": "phep-thu",
      "ten": "Phép thử và biến cố",
      "kieu": "chon",
      "luaChon": [
        {
          "ma": "dong-xu-ngua",
          "ten": "Tung đồng xu: mặt ngửa"
        },
        {
          "ma": "xuc-xac-mat-6",
          "ten": "Gieo xúc xắc: mặt 6 chấm"
        },
        {
          "ma": "xuc-xac-chan",
          "ten": "Gieo xúc xắc: số chấm chẵn"
        },
        {
          "ma": "hai-xuc-xac-tong-7",
          "ten": "Gieo hai xúc xắc: tổng bằng 7"
        },
        {
          "ma": "hai-xuc-xac-tong-12",
          "ten": "Gieo hai xúc xắc: tổng bằng 12"
        }
      ],
      "macDinh": "dong-xu-ngua"
    },
    {
      "ma": "so-lan",
      "ten": "Số lần thử n",
      "kieu": "so",
      "donVi": "lần",
      "min": 10,
      "max": 10000,
      "buoc": 10,
      "macDinh": 100
    },
    {
      "ma": "hat-giong",
      "ten": "Lượt gieo số",
      "kieu": "so",
      "donVi": "",
      "min": 1,
      "max": 999,
      "buoc": 1,
      "macDinh": 1
    }
  ],
  "daiLuongDo": [
    {
      "ma": "tan-so",
      "ten": "Số lần biến cố xảy ra",
      "donVi": "lần",
      "saiSo": 0,
      "chuSo": 0
    },
    {
      "ma": "tan-suat",
      "ten": "Tần suất",
      "donVi": "",
      "saiSo": 0,
      "chuSo": 4
    },
    {
      "ma": "xac-suat-li-thuyet",
      "ten": "Xác suất lí thuyết",
      "donVi": "",
      "saiSo": 0,
      "chuSo": 4
    }
  ],
  "congThuc": {
    "bieuThuc": "Tần suất f = (số lần biến cố xảy ra)/n; khi n lớn, f tiến gần xác suất P.",
    "dieuKien": "Đồng xu và xúc xắc cân đối, các lần thử độc lập; số ngẫu nhiên sinh bằng thuật toán có hạt giống nên cùng một lượt gieo cho cùng kết quả.",
    "nguon": ""
  },
  "bangKiem": [
    {
      "vao": {},
      "ra": {
        "tan-so": 53.0,
        "tan-suat": 0.53,
        "xac-suat-li-thuyet": 0.5
      },
      "saiSo": 1e-07
    },
    {
      "vao": {
        "so-lan": 1000
      },
      "ra": {
        "tan-so": 505.0,
        "tan-suat": 0.505,
        "xac-suat-li-thuyet": 0.5
      },
      "saiSo": 1e-07
    },
    {
      "vao": {
        "phep-thu": "xuc-xac-mat-6",
        "so-lan": 600,
        "hat-giong": 7
      },
      "ra": {
        "tan-so": 95.0,
        "tan-suat": 0.1583333333,
        "xac-suat-li-thuyet": 0.1666666667
      },
      "saiSo": 1e-07
    },
    {
      "vao": {
        "phep-thu": "xuc-xac-chan",
        "so-lan": 10,
        "hat-giong": 2
      },
      "ra": {
        "tan-so": 5.0,
        "tan-suat": 0.5,
        "xac-suat-li-thuyet": 0.5
      },
      "saiSo": 1e-07
    },
    {
      "vao": {
        "phep-thu": "hai-xuc-xac-tong-7",
        "so-lan": 3600,
        "hat-giong": 5
      },
      "ra": {
        "tan-so": 580.0,
        "tan-suat": 0.1611111111,
        "xac-suat-li-thuyet": 0.1666666667
      },
      "saiSo": 1e-07
    },
    {
      "vao": {
        "phep-thu": "hai-xuc-xac-tong-12",
        "so-lan": 10000,
        "hat-giong": 999
      },
      "ra": {
        "tan-so": 289.0,
        "tan-suat": 0.0289,
        "xac-suat-li-thuyet": 0.02777777778
      },
      "saiSo": 1e-07
    }
  ]
}
```

`tools/vi/thi_nghiem_parts/mo_hinh/toan-xac-suat.js`:

```js
(function (root) {
  'use strict';
  var XAC_SUAT = {
    'dong-xu-ngua': 1 / 2, 'xuc-xac-mat-6': 1 / 6, 'xuc-xac-chan': 1 / 2,
    'hai-xuc-xac-tong-7': 1 / 6, 'hai-xuc-xac-tong-12': 1 / 36
  };

  function xayRa(phepThu, rng) {
    if (phepThu === 'dong-xu-ngua') { return rng() < 0.5; }
    var mat = Math.floor(rng() * 6) + 1;
    if (phepThu === 'xuc-xac-mat-6') { return mat === 6; }
    if (phepThu === 'xuc-xac-chan') { return mat % 2 === 0; }
    var tong = mat + Math.floor(rng() * 6) + 1;
    return phepThu === 'hai-xuc-xac-tong-7' ? tong === 7 : tong === 12;
  }

  function tinh(p) {
    var rng = root.THI_NGHIEM_KHUNG.taoNgauNhien(p['hat-giong']);
    var tanSo = 0;
    for (var lan = 0; lan < p['so-lan']; lan += 1) {
      if (xayRa(p['phep-thu'], rng)) { tanSo += 1; }
    }
    return { 'tan-so': tanSo, 'tan-suat': tanSo / p['so-lan'], 'xac-suat-li-thuyet': XAC_SUAT[p['phep-thu']] };
  }

  function ve(ctx, p, t, kt, d) {
    var K = root.THI_NGHIEM_KHUNG;
    var trai = 56, phai = kt.rong - 24, tren = 40, day = kt.cao - 44;
    var tran = Math.min(1, Math.max(0.2, 2.5 * d['xac-suat-li-thuyet']));
    function y(f) { return day - (day - tren) * Math.min(f, tran) / tran; }
    ctx.font = '13px ' + K.PHONG; ctx.strokeStyle = K.MAU.net; ctx.lineWidth = 1.5;
    ctx.beginPath(); ctx.moveTo(trai, tren); ctx.lineTo(trai, day); ctx.lineTo(phai, day); ctx.stroke();
    ctx.fillStyle = K.MAU.net;
    ctx.fillText('Tần suất sau n lần thử', trai, tren - 14);
    ctx.fillText('n = ' + p['so-lan'], phai - 70, day + 20);
    ctx.fillText(K.dinhDang(tran, 2), 8, tren + 4); ctx.fillText('0', 38, day + 4);
    ctx.setLineDash([6, 5]); ctx.strokeStyle = K.MAU.tot; ctx.beginPath();
    ctx.moveTo(trai, y(d['xac-suat-li-thuyet'])); ctx.lineTo(phai, y(d['xac-suat-li-thuyet'])); ctx.stroke(); ctx.setLineDash([]);
    ctx.fillStyle = K.MAU.tot; ctx.fillText('P = ' + K.dinhDang(d['xac-suat-li-thuyet'], 4), phai - 90, y(d['xac-suat-li-thuyet']) - 6);
    var rng = K.taoNgauNhien(p['hat-giong']);
    var tanSo = 0, cach = Math.max(1, Math.floor(p['so-lan'] / 400));
    ctx.strokeStyle = K.MAU.chinh; ctx.lineWidth = 2; ctx.beginPath();
    for (var lan = 1; lan <= p['so-lan']; lan += 1) {
      if (xayRa(p['phep-thu'], rng)) { tanSo += 1; }
      if (lan % cach === 0 || lan === p['so-lan']) {
        var x = trai + (phai - trai) * lan / p['so-lan'];
        if (lan === cach) { ctx.moveTo(x, y(tanSo / lan)); } else { ctx.lineTo(x, y(tanSo / lan)); }
      }
    }
    ctx.stroke();
  }

  root.THI_NGHIEM_MO_HINH = { tinh: tinh, ve: ve };
})(typeof globalThis !== 'undefined' ? globalThis : this);
```

- [ ] **Step 4: Chạy lại**

Run: `venv\Scripts\python.exe -m unittest tools.vi.tests.test_thi_nghiem_mo_hinh -v`
Expected: PASS (11 test). Mẫu xác suất so 40 điểm, khớp tuyệt đối vì hai bộ sinh số khớp từng bit.

- [ ] **Step 5: Commit**

```bash
git add tools/vi/thi_nghiem_parts/mo_hinh tools/vi/tests/test_thi_nghiem_mo_hinh.py
git commit -m "feat(vi): add function study and empirical probability experiment models"
```

---

### Task 7: Bộ đọc `thi-nghiem.md`

Đọc file nguồn của thầy cô, kiểm từng dòng với khai báo của mô hình, và tạo ra cấu hình mà khung chạy đọc. Lỗi nào cũng kèm số dòng. Khoảng tham số không được vượt khoảng của mẫu: đó là cách điều kiện áp dụng của công thức được thực thi.

**Files:**
- Create: `tools/vi/thi_nghiem_parts/parse.py`
- Create: `tools/vi/tests/test_thi_nghiem_parse.py`

**Interfaces:**
- Consumes: `thu_vien.load(...).khai_bao` (Task 3); mẫu `li-con-lac-don`, `li-mach-ohm` (Task 3, 4); `thi_nghiem_mau.PENDULUM`.
- Produces: `parse.ParseError(line_no, message)` với `str()` dạng `Dòng N: ...`; `parse.read_meta(text) -> dict` (khoá `tieu-de`, `mon`, `lop`, `mau`, `nguoi-thao-tac`, `sai-so`); `parse.parse_experiment(text, khai_bao) -> Experiment`; `Experiment` có `meta`, `tham_so`, `du_doan_cau`, `lua_chon`, `dap_an`, `so_lan_do`, `cot`, `do_thi: tuple[Axis, Axis]|None` (trục tung trước), `giai_thich_cau`, `goi_y`, `ket_luan`, `warnings`, `sliders()`, `config()`; `Axis(ma, phep)`.

- [ ] **Step 1: Viết test**

`tools/vi/tests/test_thi_nghiem_parse.py`:

```python
"""Test cho bộ đọc thi-nghiem.md."""

import sys
import unittest
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from thi_nghiem_mau import PENDULUM  # noqa: E402
from thi_nghiem_parts import parse, thu_vien  # noqa: E402

class ParseTest(unittest.TestCase):
    def setUp(self):
        self.kb = thu_vien.load("li-con-lac-don", Path(".")).khai_bao

    def parse(self, text):
        return parse.parse_experiment(text, self.kb)

    def assertLine(self, text, line_no, fragment):
        with self.assertRaises(parse.ParseError) as caught:
            self.parse(text)
        self.assertEqual(caught.exception.line_no, line_no, str(caught.exception))
        self.assertIn(fragment, caught.exception.message)

    def test_valid_file_becomes_the_page_config(self):
        experiment = self.parse(PENDULUM)
        self.assertEqual(experiment.sliders(), ["chieu-dai"])
        config = experiment.config()
        self.assertEqual(config["thamSo"]["chieu-dai"], {"kieu": "truot", "min": 0.4, "max": 1.6, "buoc": 0.2, "macDinh": 1.0})
        self.assertEqual(config["thamSo"]["g"], {"kieu": "co-dinh", "giaTri": 9.8})
        self.assertEqual(config["thamSo"]["khoi-luong"], {"kieu": "co-dinh", "giaTri": 0.2})
        self.assertEqual(list(config["thamSo"]), ["chieu-dai", "g", "goc-lech", "khoi-luong"])
        self.assertEqual(config["duDoan"]["dapAn"], "B")
        self.assertEqual(config["quanSat"], {"soLanDo": 5, "cot": ["chieu-dai", "chu-ki"], "doThi": {
            "tung": {"ma": "chu-ki", "phep": "binh-phuong"}, "hoanh": {"ma": "chieu-dai", "phep": "khong"}}})
        self.assertTrue(config["saiSo"])
        self.assertEqual(config["nguoiThaoTac"], "nhom")

    def test_defaults_are_teacher_mode_and_no_noise(self):
        text = PENDULUM.replace("nguoi-thao-tac: nhom\n", "").replace("sai-so: bat\n", "")
        config = self.parse(text).config()
        self.assertEqual((config["nguoiThaoTac"], config["saiSo"]), ("giao-vien", False))

    def test_open_prediction_has_no_key(self):
        text = PENDULUM.replace("A: Tăng 4 lần\nB: Tăng 2 lần\nC: Không đổi\ndap-an: B\n", "")
        config = self.parse(text).config()
        self.assertEqual((config["duDoan"]["luaChon"], config["duDoan"]["dapAn"]), ([], None))

    def test_comma_decimals_are_accepted(self):
        text = PENDULUM.replace("0.4..1.6 buoc 0.2 mac-dinh 1.0", "0,4..1,6 buoc 0,2 mac-dinh 1,0")
        self.assertEqual(self.parse(text).tham_so["chieu-dai"]["min"], 0.4)

    def test_errors_name_the_line(self):
        self.assertLine(PENDULUM.replace("chieu-dai: 0.4..1.6", "chieu-dai: 0.1..1.6"), 11, "công thức của mẫu không còn đúng")
        self.assertLine(PENDULUM.replace("g: co-dinh 9.8", "g: co-dinh 50"), 12, "1.6..24.8")
        self.assertLine(PENDULUM.replace("g: co-dinh 9.8", "luc-can: co-dinh 1"), 12, "mẫu không có tham số `luc-can`")
        self.assertLine(PENDULUM.replace("buoc 0.2", "buoc 0"), 11, "`buoc`")
        self.assertLine(PENDULUM.replace("mac-dinh 1.0", "mac-dinh 3"), 11, "`mac-dinh`")
        self.assertLine(PENDULUM.replace("chieu-dai: 0.4..1.6 buoc 0.2 mac-dinh 1.0", "chieu-dai: tu 0.4 den 1.6"), 11, "cần dạng")
        self.assertLine(PENDULUM.replace("dap-an: B", "dap-an: D"), 19, "`dap-an`")
        self.assertLine(PENDULUM.replace("so-lan-do: 5", "so-lan-do: 2"), 22, "từ 3 đến 20")
        self.assertLine(PENDULUM.replace("do: chieu-dai, chu-ki", "do: chieu-dai, tan-so"), 23, "mẫu không có `tan-so`")
        self.assertLine(PENDULUM.replace("do: chieu-dai, chu-ki", "do: chieu-dai"), 23, "ít nhất một đại lượng đo")
        self.assertLine(PENDULUM.replace("chu-ki^2 theo chieu-dai", "chu-ki^3 theo chieu-dai"), 24, "không hiểu biểu thức")
        self.assertLine(PENDULUM.replace("chu-ki^2 theo chieu-dai", "chu-ki^2 theo g"), 24, "dòng `do:` không có")
        self.assertLine(PENDULUM.replace("chu-ki^2 theo chieu-dai", "chu-ki^2"), 24, "theo")
        self.assertLine(PENDULUM.replace("cau: Từ đồ thị", "xem https://example.com Từ đồ thị"), 27, "địa chỉ web")

    def test_structure_errors(self):
        for removed in ("## Giải thích\n", "## Kết luận\nChu kì con lắc đơn chỉ phụ thuộc chiều dài dây và g.\n"):
            text = PENDULUM.replace(removed, "")
            self.assertLine(text, len(text.splitlines()), "thiếu mục `" + removed.splitlines()[0] + "`")
        self.assertLine(PENDULUM.replace("## Quan sát", "## Đo đạc"), 21, "mục lạ")
        self.assertLine(PENDULUM.replace("goi-y-dap-an: T^2^ tỉ lệ thuận với l; T = 2π√(l/g).\n", ""), 26, "`goi-y-dap-an:`")
        self.assertLine(PENDULUM.replace("chieu-dai: 0.4..1.6 buoc 0.2 mac-dinh 1.0\n", ""), 10, "ít nhất một tham số thay đổi được")
        self.assertLine("## Tham số\n", 1, "`---`")

    def test_table_without_a_changing_parameter_warns(self):
        text = PENDULUM.replace("do: chieu-dai, chu-ki\ndo-thi: chu-ki^2 theo chieu-dai", "do: chu-ki")
        self.assertIn("không có tham số nào thay đổi được", self.parse(text).warnings[0])

    def test_meta_errors(self):
        for text, fragment in (
            (PENDULUM.replace("mau: li-con-lac-don\n", ""), "thiếu `mau`"),
            (PENDULUM.replace("sai-so: bat", "sai-so: co"), "`sai-so`"),
            (PENDULUM.replace("nguoi-thao-tac: nhom", "nguoi-thao-tac: hoc-sinh"), "`nguoi-thao-tac`"),
            (PENDULUM.replace("lop: 11", "khoi: 11"), "khoá lạ `khoi`"),
        ):
            with self.subTest(fragment=fragment):
                with self.assertRaises(parse.ParseError) as caught:
                    parse.read_meta(text)
                self.assertIn(fragment, caught.exception.message)

    def test_choice_parameters(self):
        kb = thu_vien.load("li-mach-ohm", Path(".")).khai_bao
        text = (PENDULUM.replace("mau: li-con-lac-don", "mau: li-mach-ohm")
                .replace("chieu-dai: 0.4..1.6 buoc 0.2 mac-dinh 1.0\ng: co-dinh 9.8", "kieu-mac: chon song-song, noi-tiep\ndien-tro-1: co-dinh 30")
                .replace("do: chieu-dai, chu-ki\ndo-thi: chu-ki^2 theo chieu-dai", "do: kieu-mac, cuong-do-mach-chinh"))
        experiment = parse.parse_experiment(text, kb)
        self.assertEqual(experiment.tham_so["kieu-mac"], {"kieu": "chon", "luaChon": ["song-song", "noi-tiep"], "macDinh": "song-song"})
        self.assertEqual(experiment.warnings, [])
        for bad, fragment in (("kieu-mac: chon song-song", "ít nhất hai lựa chọn"), ("kieu-mac: 1..2 buoc 1 mac-dinh 1", "tham số lựa chọn"),
                              ("kieu-mac: co-dinh cheo", "chỉ nhận"), ("dien-tro-2: chon 1, 2", "là tham số số")):
            with self.subTest(bad=bad):
                with self.assertRaises(parse.ParseError) as caught:
                    parse.parse_experiment(text.replace("kieu-mac: chon song-song, noi-tiep", bad), kb)
                self.assertIn(fragment, caught.exception.message)

if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Chạy để thấy trượt**

Run: `venv\Scripts\python.exe -m unittest tools.vi.tests.test_thi_nghiem_parse -v`
Expected: ERROR `ImportError: cannot import name 'parse'`.

- [ ] **Step 3: Viết bộ đọc**

`tools/vi/thi_nghiem_parts/parse.py`:

```python
"""Đọc file nguồn thi-nghiem.md thành cấu trúc dữ liệu và kiểm nó với khai báo của mô hình.

Ngữ pháp nằm trong docs/vi/tro-ly/thi-nghiem-ao.md. Chỉ dùng thư viện chuẩn Python.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

META_REQUIRED = ("tieu-de", "mon", "lop", "mau")
META_OPTIONAL = ("nguoi-thao-tac", "sai-so")
OPERATORS = ("giao-vien", "nhom")
SECTIONS = ("## Tham số", "## Dự đoán", "## Quan sát", "## Giải thích", "## Kết luận")
OPTION_KEYS = ("A", "B", "C", "D")
MIN_RUNS, MAX_RUNS = 3, 20
MAX_COLUMNS = 6
TRANSFORMS = (
    (re.compile(r"^([a-z0-9-]+)\^2$"), "binh-phuong"),
    (re.compile(r"^1/([a-z0-9-]+)$"), "nghich-dao"),
    (re.compile(r"^ln\(([a-z0-9-]+)\)$"), "ln"),
    (re.compile(r"^sqrt\(([a-z0-9-]+)\)$"), "can"),
    (re.compile(r"^([a-z0-9-]+)$"), "khong"),
)
_NUMBER = r"-?[0-9]+(?:[.,][0-9]+)?"
_SLIDER_RE = re.compile(rf"^({_NUMBER})\.\.({_NUMBER})\s+buoc\s+({_NUMBER})\s+mac-dinh\s+({_NUMBER})$")
_KEY_RE = re.compile(r"^([A-Za-z0-9-]+):\s*(.*)$")
_URL_RE = re.compile(r"https?://", re.IGNORECASE)


class ParseError(Exception):
    """Lỗi trong thi-nghiem.md; luôn kèm số dòng để AI sửa được đúng chỗ."""

    def __init__(self, line_no: int, message: str) -> None:
        super().__init__(f"Dòng {line_no}: {message}")
        self.line_no = line_no
        self.message = message


@dataclass
class Axis:
    ma: str
    phep: str


@dataclass
class Experiment:
    meta: dict[str, str]
    tham_so: dict[str, dict]
    du_doan_cau: str
    lua_chon: list[tuple[str, str]]
    dap_an: str | None
    so_lan_do: int
    cot: list[str]
    do_thi: tuple[Axis, Axis] | None  # (trục tung, trục hoành)
    giai_thich_cau: str
    goi_y: str
    ket_luan: str
    warnings: list[str] = field(default_factory=list)

    def sliders(self) -> list[str]:
        return [ma for ma, ch in self.tham_so.items() if ch["kieu"] != "co-dinh"]

    def config(self) -> dict:
        """Cấu hình nhúng vào file HTML; khung.js đọc đúng các khoá này."""
        do_thi = None
        if self.do_thi is not None:
            tung, hoanh = self.do_thi
            do_thi = {"tung": {"ma": tung.ma, "phep": tung.phep}, "hoanh": {"ma": hoanh.ma, "phep": hoanh.phep}}
        return {
            "tieuDe": self.meta["tieu-de"], "mon": self.meta["mon"], "lop": self.meta["lop"],
            "nguoiThaoTac": self.meta["nguoi-thao-tac"], "saiSo": self.meta["sai-so"] == "bat",
            "thamSo": self.tham_so,
            "duDoan": {"cau": self.du_doan_cau,
                       "luaChon": [{"ma": ma, "noiDung": noi_dung} for ma, noi_dung in self.lua_chon],
                       "dapAn": self.dap_an},
            "quanSat": {"soLanDo": self.so_lan_do, "cot": self.cot, "doThi": do_thi},
            "giaiThich": {"cau": self.giai_thich_cau, "goiY": self.goi_y},
            "ketLuan": self.ket_luan,
        }


def _number(text: str) -> float:
    return float(text.replace(",", "."))


def _lines(text: str) -> list[tuple[int, str]]:
    lines = [(index, line.rstrip()) for index, line in enumerate(text.lstrip("﻿").splitlines(), 1)]
    for line_no, line in lines:
        if _URL_RE.search(line):
            raise ParseError(line_no, "không chèn địa chỉ web vào thí nghiệm; file thí nghiệm chạy không cần mạng")
    return lines


def _split(text: str) -> tuple[dict[str, tuple[int, str]], dict[str, list[tuple[int, str]]], dict[str, int]]:
    lines = _lines(text)
    body = [(n, line) for n, line in lines if line.strip()]
    if not body or body[0][1].strip() != "---":
        raise ParseError(body[0][0] if body else 1, "file phải mở đầu bằng khối thông tin giữa hai dòng `---`")
    meta: dict[str, tuple[int, str]] = {}
    position = lines.index(body[0]) + 1
    closed = False
    while position < len(lines):
        line_no, line = lines[position]
        position += 1
        if line.strip() == "---":
            closed = True
            break
        if not line.strip():
            continue
        match = _KEY_RE.match(line.strip())
        if match is None:
            raise ParseError(line_no, "dòng trong khối thông tin phải có dạng `khoá: giá trị`")
        key, value = match.group(1), match.group(2).strip()
        if key not in META_REQUIRED + META_OPTIONAL:
            raise ParseError(line_no, f"khoá lạ `{key}`; chỉ dùng: " + ", ".join(META_REQUIRED + META_OPTIONAL))
        if key in meta:
            raise ParseError(line_no, f"khoá `{key}` bị lặp")
        meta[key] = (line_no, value)
    if not closed:
        raise ParseError(lines[-1][0], "khối thông tin chưa đóng bằng dòng `---`")

    sections: dict[str, list[tuple[int, str]]] = {}
    heading_lines: dict[str, int] = {}
    current: str | None = None
    for line_no, line in lines[position:]:
        if line.startswith("## "):
            heading = line.strip()
            if heading not in SECTIONS:
                raise ParseError(line_no, f"mục lạ `{heading}`; chỉ dùng: " + ", ".join(SECTIONS))
            if heading in sections:
                raise ParseError(line_no, f"mục `{heading}` bị lặp")
            current = heading
            sections[current] = []
            heading_lines[current] = line_no
        elif line.strip():
            if current is None:
                raise ParseError(line_no, "nội dung phải nằm dưới một mục `## ...`")
            sections[current].append((line_no, line.strip()))
    last = lines[-1][0]
    for heading in SECTIONS:
        if heading not in sections:
            raise ParseError(last, f"thiếu mục `{heading}`")
    return meta, sections, heading_lines


def read_meta(text: str) -> dict[str, str]:
    """Đọc riêng khối thông tin, để biết `mau` trước khi nạp mô hình."""
    meta, _, _ = _split(text)
    result: dict[str, str] = {}
    for key in META_REQUIRED:
        if key not in meta or not meta[key][1]:
            raise ParseError(1, f"khối thông tin thiếu `{key}`")
        result[key] = meta[key][1]
    result["nguoi-thao-tac"] = meta.get("nguoi-thao-tac", (0, "giao-vien"))[1] or "giao-vien"
    result["sai-so"] = meta.get("sai-so", (0, "tat"))[1] or "tat"
    if result["nguoi-thao-tac"] not in OPERATORS:
        raise ParseError(meta["nguoi-thao-tac"][0], "`nguoi-thao-tac` chỉ nhận giao-vien hoặc nhom")
    if result["sai-so"] not in ("bat", "tat"):
        raise ParseError(meta["sai-so"][0], "`sai-so` chỉ nhận bat hoặc tat")
    return result


def _pairs(lines: list[tuple[int, str]], allowed: tuple[str, ...], heading: str) -> dict[str, tuple[int, str]]:
    result: dict[str, tuple[int, str]] = {}
    for line_no, line in lines:
        match = _KEY_RE.match(line)
        if match is None or match.group(1) not in allowed:
            raise ParseError(line_no, f"mục `{heading}` chỉ nhận các dòng: " + ", ".join(f"`{key}:`" for key in allowed))
        if match.group(1) in result:
            raise ParseError(line_no, f"dòng `{match.group(1)}:` bị lặp")
        result[match.group(1)] = (line_no, match.group(2).strip())
    return result


def _required(pairs: dict[str, tuple[int, str]], key: str, heading_line: int, heading: str) -> tuple[int, str]:
    if key not in pairs or not pairs[key][1]:
        raise ParseError(pairs.get(key, (heading_line, ""))[0], f"mục `{heading}` thiếu dòng `{key}:`")
    return pairs[key]


def _parameters(lines: list[tuple[int, str]], heading_line: int, khai_bao: dict) -> dict[str, dict]:
    declared = {ts["ma"]: ts for ts in khai_bao["thamSo"]}
    chosen: dict[str, dict] = {}
    for line_no, line in lines:
        match = _KEY_RE.match(line)
        if match is None:
            raise ParseError(line_no, "dòng tham số phải có dạng `<mã>: ...`")
        ma, value = match.group(1), match.group(2).strip()
        if ma not in declared:
            raise ParseError(line_no, f"mẫu không có tham số `{ma}`; các tham số của mẫu: " + ", ".join(declared))
        if ma in chosen:
            raise ParseError(line_no, f"tham số `{ma}` bị lặp")
        ts = declared[ma]
        codes = [o["ma"] for o in ts.get("luaChon", [])]
        if value.startswith("co-dinh"):
            raw = value[len("co-dinh"):].strip()
            if ts["kieu"] == "chon":
                if raw not in codes:
                    raise ParseError(line_no, f"`{ma}` chỉ nhận: " + ", ".join(codes))
                chosen[ma] = {"kieu": "co-dinh", "giaTri": raw}
            else:
                if not re.fullmatch(_NUMBER, raw):
                    raise ParseError(line_no, f"`{ma}: co-dinh` cần một số, ví dụ `co-dinh {ts['macDinh']}`")
                number = _number(raw)
                if not ts["min"] <= number <= ts["max"]:
                    raise ParseError(line_no, f"`{ma}` phải nằm trong khoảng {ts['min']}..{ts['max']} {ts['donVi']} "
                                              "(ngoài khoảng này công thức của mẫu không còn đúng)")
                chosen[ma] = {"kieu": "co-dinh", "giaTri": number}
        elif value.startswith("chon"):
            if ts["kieu"] != "chon":
                raise ParseError(line_no, f"`{ma}` là tham số số; dùng `<nhỏ nhất>..<lớn nhất> buoc <bước> mac-dinh <giá trị>` "
                                          "hoặc `co-dinh <giá trị>`")
            picked = [item.strip() for item in value[len("chon"):].split(",") if item.strip()]
            unknown = [item for item in picked if item not in codes]
            if len(picked) < 2 or unknown or len(set(picked)) != len(picked):
                raise ParseError(line_no, f"`{ma}: chon` cần ít nhất hai lựa chọn khác nhau trong: " + ", ".join(codes))
            chosen[ma] = {"kieu": "chon", "luaChon": picked, "macDinh": picked[0]}
        else:
            if ts["kieu"] == "chon":
                raise ParseError(line_no, f"`{ma}` là tham số lựa chọn; dùng `chon {', '.join(codes)}` hoặc `co-dinh <mã>`")
            slider = _SLIDER_RE.match(value)
            if slider is None:
                raise ParseError(line_no, f"`{ma}` cần dạng `<nhỏ nhất>..<lớn nhất> buoc <bước> mac-dinh <giá trị>` "
                                          "hoặc `co-dinh <giá trị>`")
            low, high, step, default = (_number(part) for part in slider.groups())
            if not ts["min"] <= low < high <= ts["max"]:
                raise ParseError(line_no, f"khoảng của `{ma}` phải nằm trong {ts['min']}..{ts['max']} {ts['donVi']} "
                                          "(ngoài khoảng này công thức của mẫu không còn đúng) và nhỏ nhất < lớn nhất")
            if step <= 0 or step > high - low:
                raise ParseError(line_no, f"`buoc` của `{ma}` phải dương và không lớn hơn độ rộng khoảng")
            if not low <= default <= high:
                raise ParseError(line_no, f"`mac-dinh` của `{ma}` phải nằm trong khoảng {low:g}..{high:g}")
            chosen[ma] = {"kieu": "truot", "min": low, "max": high, "buoc": step, "macDinh": default}
    for ma, ts in declared.items():
        chosen.setdefault(ma, {"kieu": "co-dinh", "giaTri": ts["macDinh"]})
    if all(ch["kieu"] == "co-dinh" for ch in chosen.values()):
        raise ParseError(heading_line, "mục `## Tham số` cần ít nhất một tham số thay đổi được (thanh trượt hoặc `chon`)")
    return {ts["ma"]: chosen[ts["ma"]] for ts in khai_bao["thamSo"]}


def _axis(text: str, line_no: int, columns: list[str]) -> Axis:
    for pattern, name in TRANSFORMS:
        match = pattern.match(text.strip())
        if match is not None:
            if match.group(1) not in columns:
                raise ParseError(line_no, f"`do-thi` dùng `{match.group(1)}` nhưng dòng `do:` không có đại lượng này")
            return Axis(match.group(1), name)
    raise ParseError(line_no, f"không hiểu biểu thức `{text.strip()}`; dùng `<mã>`, `<mã>^2`, `1/<mã>`, `ln(<mã>)` "
                              "hoặc `sqrt(<mã>)`")


def parse_experiment(text: str, khai_bao: dict) -> Experiment:
    meta = read_meta(text)
    _, sections, heading_lines = _split(text)
    tham_so = _parameters(sections["## Tham số"], heading_lines["## Tham số"], khai_bao)

    heading = "## Dự đoán"
    pairs = _pairs(sections[heading], ("cau",) + OPTION_KEYS + ("dap-an",), heading)
    du_doan_cau = _required(pairs, "cau", heading_lines[heading], heading)[1]
    lua_chon = [(key, pairs[key][1]) for key in OPTION_KEYS if key in pairs]
    dap_an = None
    if lua_chon:
        if len(lua_chon) < 2 or any(not noi_dung for _, noi_dung in lua_chon):
            raise ParseError(pairs[lua_chon[0][0]][0], "dự đoán trắc nghiệm cần ít nhất hai lựa chọn có nội dung")
        line_no, dap_an = _required(pairs, "dap-an", heading_lines[heading], heading)
        if dap_an not in [key for key, _ in lua_chon]:
            raise ParseError(line_no, "`dap-an` phải là một trong các lựa chọn đã ghi")
    elif "dap-an" in pairs:
        raise ParseError(pairs["dap-an"][0], "`dap-an` chỉ dùng khi dự đoán có các lựa chọn A, B, C, D")

    heading = "## Quan sát"
    pairs = _pairs(sections[heading], ("so-lan-do", "do", "do-thi"), heading)
    line_no, raw = _required(pairs, "so-lan-do", heading_lines[heading], heading)
    if not raw.isdigit() or not MIN_RUNS <= int(raw) <= MAX_RUNS:
        raise ParseError(line_no, f"`so-lan-do` phải là số nguyên từ {MIN_RUNS} đến {MAX_RUNS}")
    so_lan_do = int(raw)
    line_no, raw = _required(pairs, "do", heading_lines[heading], heading)
    cot = [item.strip() for item in raw.split(",") if item.strip()]
    measures = [dl["ma"] for dl in khai_bao["daiLuongDo"]]
    known = list(tham_so) + measures
    unknown = [ma for ma in cot if ma not in known]
    if unknown:
        raise ParseError(line_no, f"mẫu không có `{unknown[0]}`; dùng được: " + ", ".join(known))
    if len(set(cot)) != len(cot) or len(cot) > MAX_COLUMNS:
        raise ParseError(line_no, f"`do` không được lặp và tối đa {MAX_COLUMNS} cột")
    if not any(ma in measures for ma in cot):
        raise ParseError(line_no, "`do` cần ít nhất một đại lượng đo: " + ", ".join(measures))
    do_thi = None
    if "do-thi" in pairs and pairs["do-thi"][1]:
        line_no, raw = pairs["do-thi"]
        parts = raw.split(" theo ")
        if len(parts) != 2:
            raise ParseError(line_no, "`do-thi` cần dạng `<biểu thức trục tung> theo <biểu thức trục hoành>`")
        do_thi = (_axis(parts[0], line_no, cot), _axis(parts[1], line_no, cot))
        for axis in do_thi:
            ts = next((item for item in khai_bao["thamSo"] if item["ma"] == axis.ma), None)
            if ts is not None and ts["kieu"] == "chon":
                raise ParseError(line_no, f"`{axis.ma}` là lựa chọn, không vẽ lên đồ thị được")

    heading = "## Giải thích"
    pairs = _pairs(sections[heading], ("cau", "goi-y-dap-an"), heading)
    giai_thich_cau = _required(pairs, "cau", heading_lines[heading], heading)[1]
    goi_y = _required(pairs, "goi-y-dap-an", heading_lines[heading], heading)[1]

    heading = "## Kết luận"
    if not sections[heading]:
        raise ParseError(heading_lines[heading], "mục `## Kết luận` chưa có nội dung")
    ket_luan = " ".join(line for _, line in sections[heading])

    warnings = []
    slider_columns = [ma for ma in cot if ma in tham_so and tham_so[ma]["kieu"] != "co-dinh"]
    if not slider_columns:
        warnings.append("Bảng số liệu không có tham số nào thay đổi được; học sinh sẽ không thấy quan hệ giữa các đại lượng.")
    return Experiment(meta=meta, tham_so=tham_so, du_doan_cau=du_doan_cau, lua_chon=lua_chon, dap_an=dap_an,
                      so_lan_do=so_lan_do, cot=cot, do_thi=do_thi, giai_thich_cau=giai_thich_cau, goi_y=goi_y,
                      ket_luan=ket_luan, warnings=warnings)
```

- [ ] **Step 4: Chạy lại**

Run: `venv\Scripts\python.exe -m unittest tools.vi.tests.test_thi_nghiem_parse -v`
Expected: PASS (9 test).

- [ ] **Step 5: Commit**

```bash
git add tools/vi/thi_nghiem_parts/parse.py tools/vi/tests/test_thi_nghiem_parse.py
git commit -m "feat(vi): parse thi-nghiem.md against the model declaration"
```

---

### Task 8: Giao diện khung chạy và bộ ghép HTML

Thêm phần giao diện vào `khung.js` (giữ nguyên phần logic của Task 1, chỉ mở rộng khối xuất ở cuối), viết `khung.css`, và `build_html.py` nhúng tất cả vào một file. Học sinh chưa chốt dự đoán thì thanh trượt, nút Chạy, nút Ghi lần đo bị khoá và ô số đo hiện `…`.

**Files:**
- Modify: `tools/vi/thi_nghiem_parts/runtime/khung.js` (thay toàn bộ file)
- Create: `tools/vi/thi_nghiem_parts/runtime/khung.css`
- Create: `tools/vi/thi_nghiem_parts/build_html.py`
- Create: `tools/vi/tests/test_thi_nghiem_html.py`

**Interfaces:**
- Consumes: `Experiment.config()` (Task 7), `Model.js`, `Model.khai_bao` (Task 3).
- Produces: `THI_NGHIEM_KHUNG.{MAU, PHONG, khoiDong()}` thêm vào các hàm của Task 1; `build_html.build(experiment, model) -> str`; `build_html.write(experiment, model, folder) -> Path`; `build_html.BuildError`; `build_html.FILENAME = "thi-nghiem.html"`; `build_html.NETWORK_MARKS`.
- Trang HTML có các id: `tieu-de`, `phu-de`, `khung-ve`, `dieu-khien-chay`, `so-do`, `tham-so`, `nhiem-vu`, `bang`, `do-thi`, `khop`, `cong-thuc`, `dieu-kien`, `tu-kiem`, `du-lieu`.

- [ ] **Step 1: Viết test**

`tools/vi/tests/test_thi_nghiem_html.py`:

```python
"""Test cho bộ ghép file HTML thí nghiệm ảo."""

import json
import sys
import unittest
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from thi_nghiem_mau import PENDULUM  # noqa: E402
from thi_nghiem_parts import build_html, parse, thu_vien  # noqa: E402


class BuildHtmlTest(unittest.TestCase):
    def setUp(self):
        self.model = thu_vien.load("li-con-lac-don", Path("."))
        self.experiment = parse.parse_experiment(PENDULUM, self.model.khai_bao)

    def test_page_is_self_contained_and_offline(self):
        page = build_html.build(self.experiment, self.model)
        for mark in build_html.NETWORK_MARKS + ("<link", "src="):
            self.assertNotIn(mark, page)
        for needle in ("THI_NGHIEM_KHUNG.khoiDong();", "THI_NGHIEM_MO_HINH", 'id="du-lieu"', "<title>Chu kì con lắc đơn</title>"):
            self.assertIn(needle, page)

    def test_embedded_data_round_trips_and_cannot_close_the_script(self):
        self.experiment.meta["tieu-de"] = "H~2~O </script><b>x"
        page = build_html.build(self.experiment, self.model)
        raw = page.split('<script id="du-lieu" type="application/json">', 1)[1].split("</script>", 1)[0]
        self.assertNotIn("<", raw)
        data = json.loads(raw)
        self.assertEqual(data["cauHinh"]["tieuDe"], "H~2~O </script><b>x")
        self.assertEqual(data["khaiBao"]["ma"], "li-con-lac-don")
        self.assertIn("<title>H2O &lt;/script&gt;&lt;b&gt;x</title>", page)

    def test_network_marks_in_model_code_stop_the_build(self):
        self.model.js += "\n// url(x)"
        with self.assertRaises(build_html.BuildError):
            build_html.build(self.experiment, self.model)

if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Chạy để thấy trượt**

Run: `venv\Scripts\python.exe -m unittest tools.vi.tests.test_thi_nghiem_html -v`
Expected: ERROR `ImportError: cannot import name 'build_html'`.

- [ ] **Step 3: Thay `khung.js` bằng bản đủ giao diện**

`tools/vi/thi_nghiem_parts/runtime/khung.js` (phần logic giống hệt Task 1; phần mới bắt đầu từ dòng chú thích `Giao diện`):

```js
(function (root) {
  'use strict';

  // ---------- Logic thuần: chạy được cả trong trình duyệt lẫn Node ----------

  function taoNgauNhien(hatGiong) {
    var a = hatGiong | 0;
    return function () {
      a |= 0; a = a + 0x6D2B79F5 | 0;
      var t = Math.imul(a ^ a >>> 15, 1 | a);
      t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t;
      return ((t ^ t >>> 14) >>> 0) / 4294967296;
    };
  }

  function nhieuChuan(rng) {
    var u = 1 - rng();
    var v = rng();
    return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v);
  }

  function thamSoMacDinh(khaiBao) {
    var p = {};
    khaiBao.thamSo.forEach(function (ts) { p[ts.ma] = ts.macDinh; });
    return p;
  }

  function gopThamSo(khaiBao, vao) {
    var p = thamSoMacDinh(khaiBao);
    Object.keys(vao || {}).forEach(function (ma) { p[ma] = vao[ma]; });
    return p;
  }

  function tuKiem(khaiBao, moHinh) {
    var truot = [];
    var dongTruot = {};
    khaiBao.bangKiem.forEach(function (dong, chiSo) {
      var ketQua;
      try {
        ketQua = moHinh.tinh(gopThamSo(khaiBao, dong.vao));
      } catch (loi) {
        truot.push({ dong: chiSo + 1, ma: '*', mong: null, duoc: String(loi) });
        dongTruot[chiSo] = true;
        return;
      }
      Object.keys(dong.ra).forEach(function (ma) {
        var mong = dong.ra[ma];
        var duoc = ketQua[ma];
        var dat = mong === null
          ? duoc === null
          : typeof duoc === 'number' && isFinite(duoc) && Math.abs(duoc - mong) <= dong.saiSo;
        if (!dat) {
          truot.push({ dong: chiSo + 1, ma: ma, mong: mong, duoc: duoc === undefined ? null : duoc });
          dongTruot[chiSo] = true;
        }
      });
    });
    var tong = khaiBao.bangKiem.length;
    return { tong: tong, dat: tong - Object.keys(dongTruot).length, truot: truot };
  }

  function apDungPhep(phep, x) {
    if (x === null || x === undefined || !isFinite(x)) { return null; }
    var y;
    if (phep === 'binh-phuong') { y = x * x; }
    else if (phep === 'nghich-dao') { y = x === 0 ? NaN : 1 / x; }
    else if (phep === 'ln') { y = x > 0 ? Math.log(x) : NaN; }
    else if (phep === 'can') { y = x >= 0 ? Math.sqrt(x) : NaN; }
    else { y = x; }
    return isFinite(y) ? y : null;
  }

  function khopTuyenTinh(diem) {
    var n = diem.length;
    if (n < 2) { return null; }
    var tx = 0, ty = 0;
    diem.forEach(function (d) { tx += d[0]; ty += d[1]; });
    tx /= n; ty /= n;
    var sxx = 0, sxy = 0, syy = 0;
    diem.forEach(function (d) {
      sxx += (d[0] - tx) * (d[0] - tx);
      sxy += (d[0] - tx) * (d[1] - ty);
      syy += (d[1] - ty) * (d[1] - ty);
    });
    if (sxx === 0) { return null; }
    var heSoGoc = sxy / sxx;
    return {
      heSoGoc: heSoGoc,
      tungDoGoc: ty - heSoGoc * tx,
      tuongQuan: syy === 0 ? 1 : sxy / Math.sqrt(sxx * syy)
    };
  }

  function dinhDang(x, chuSo) {
    if (x === null || x === undefined || typeof x !== 'number' || !isFinite(x)) { return '—'; }
    var tron = Number(x.toFixed(chuSo));
    return (tron === 0 ? 0 : tron).toFixed(chuSo).replace('.', ',');
  }

  function danhDau(chu) {
    var sach = String(chu)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    return sach
      .replace(/\*\*(.+?)\*\*/g, '<b>$1</b>')
      .replace(/~([^~]+)~/g, '<sub>$1</sub>')
      .replace(/\^([^\^]+)\^/g, '<sup>$1</sup>');
  }

  function mauDo(khaiBao, ketQua, rng, batSaiSo) {
    var mau = {};
    khaiBao.daiLuongDo.forEach(function (d) {
      var thuc = ketQua[d.ma];
      if (thuc === null || thuc === undefined || !batSaiSo || !d.saiSo) { mau[d.ma] = thuc === undefined ? null : thuc; }
      else { mau[d.ma] = thuc + d.saiSo * nhieuChuan(rng); }
    });
    return mau;
  }

  function taoNhiemVu(cauHinh) {
    var nv = {
      buoc: 'du-doan',
      duDoan: null,
      lanDo: [],
      giaoVien: cauHinh.nguoiThaoTac === 'giao-vien'
    };
    nv.duocThaoTac = function () { return nv.giaoVien || nv.buoc !== 'du-doan'; };
    nv.chonDuDoan = function (giaTri) {
      if (nv.buoc !== 'du-doan' || giaTri === null || giaTri === undefined || String(giaTri).trim() === '') { return false; }
      nv.duDoan = String(giaTri).trim();
      nv.buoc = 'quan-sat';
      return true;
    };
    nv.ghiLanDo = function (dong) {
      if (!nv.duocThaoTac()) { return false; }
      nv.lanDo.push(dong);
      if (nv.buoc === 'quan-sat' && nv.lanDo.length >= cauHinh.quanSat.soLanDo) { nv.buoc = 'giai-thich'; }
      return true;
    };
    nv.xoaLanDo = function (chiSo) {
      if (chiSo < 0 || chiSo >= nv.lanDo.length) { return false; }
      nv.lanDo.splice(chiSo, 1);
      return true;
    };
    nv.duDoanDung = function () {
      if (nv.buoc !== 'giai-thich' || !cauHinh.duDoan.dapAn) { return null; }
      return nv.duDoan === cauHinh.duDoan.dapAn;
    };
    nv.batGiaoVien = function () { nv.giaoVien = true; };
    return nv;
  }

  function diemDoThi(doThi, lanDo) {
    var diem = [];
    lanDo.forEach(function (dong) {
      var x = apDungPhep(doThi.hoanh.phep, dong[doThi.hoanh.ma]);
      var y = apDungPhep(doThi.tung.phep, dong[doThi.tung.ma]);
      if (x !== null && y !== null) { diem.push([x, y]); }
    });
    return diem;
  }

  // ---------- Giao diện: chỉ chạy khi có document ----------

  var MAU = { nen: '#ffffff', net: '#1e293b', nhat: '#94a3b8', chinh: '#2563eb', phu: '#f59e0b', tot: '#16a34a', xau: '#dc2626' };
  var PHONG = '"Segoe UI", Arial, sans-serif';

  function tao(the, lop, html) {
    var nut = document.createElement(the);
    if (lop) { nut.className = lop; }
    if (html !== undefined) { nut.innerHTML = html; }
    return nut;
  }

  function soChuSo(buoc) {
    var chu = String(buoc);
    return chu.indexOf('.') < 0 ? 0 : chu.length - chu.indexOf('.') - 1;
  }

  function chuanBiKhung(canvas, tiLeCao) {
    canvas.style.height = 'auto';
    var rong = canvas.clientWidth || 640;
    var cao = Math.round(Math.min(rong * tiLeCao, 460));
    var mat = window.devicePixelRatio || 1;
    canvas.style.height = cao + 'px';
    canvas.width = Math.round(rong * mat); canvas.height = Math.round(cao * mat);
    var ctx = canvas.getContext('2d');
    ctx.setTransform(mat, 0, 0, mat, 0, 0);
    ctx.fillStyle = MAU.nen; ctx.fillRect(0, 0, rong, cao);
    return { ctx: ctx, kt: { rong: rong, cao: cao } };
  }

  function vachChia(nho, lon) {
    if (nho === lon) { nho -= 1; lon += 1; }
    var tho = (lon - nho) / 5;
    var bac = Math.pow(10, Math.floor(Math.log(tho) / Math.LN10));
    var buoc = [1, 2, 5, 10].map(function (m) { return m * bac; }).filter(function (b) { return b >= tho; })[0];
    var dau = Math.floor(nho / buoc) * buoc, vach = [];
    for (var v = dau; v <= lon + buoc * 0.5; v += buoc) { vach.push(Number(v.toPrecision(12))); }
    return vach;
  }

  function veDoThi(canvas, diem, khop, nhanHoanh, nhanTung) {
    var k = chuanBiKhung(canvas, 0.6), ctx = k.ctx, kt = k.kt;
    var trai = 64, phai = kt.rong - 20, tren = 34, day = kt.cao - 46;
    ctx.font = '13px ' + PHONG; ctx.fillStyle = MAU.net; ctx.strokeStyle = MAU.net; ctx.lineWidth = 1.5;
    ctx.beginPath(); ctx.moveTo(trai, tren); ctx.lineTo(trai, day); ctx.lineTo(phai, day); ctx.stroke();
    ctx.fillText(nhanTung, 8, 14); ctx.textAlign = 'right'; ctx.fillText(nhanHoanh, phai, kt.cao - 8); ctx.textAlign = 'left';
    if (!diem.length) { ctx.fillStyle = MAU.nhat; ctx.fillText('Chưa có số liệu', trai + 16, tren + 30); return; }
    var xs = diem.map(function (d) { return d[0]; }), ys = diem.map(function (d) { return d[1]; });
    var vx = vachChia(Math.min.apply(null, xs), Math.max.apply(null, xs));
    var vy = vachChia(Math.min.apply(null, ys), Math.max.apply(null, ys));
    function px(x) { return trai + (phai - trai) * (x - vx[0]) / (vx[vx.length - 1] - vx[0]); }
    function py(y) { return day - (day - tren) * (y - vy[0]) / (vy[vy.length - 1] - vy[0]); }
    ctx.strokeStyle = '#e2e8f0'; ctx.lineWidth = 1;
    vx.forEach(function (v) {
      ctx.beginPath(); ctx.moveTo(px(v), tren); ctx.lineTo(px(v), day); ctx.stroke();
      ctx.textAlign = 'center'; ctx.fillText(String(v).replace('.', ','), px(v), day + 16);
    });
    vy.forEach(function (v) {
      ctx.beginPath(); ctx.moveTo(trai, py(v)); ctx.lineTo(phai, py(v)); ctx.stroke();
      ctx.textAlign = 'right'; ctx.fillText(String(v).replace('.', ','), trai - 6, py(v) + 4);
    });
    ctx.textAlign = 'left';
    if (khop) {
      ctx.strokeStyle = MAU.phu; ctx.lineWidth = 2; ctx.beginPath();
      ctx.moveTo(px(vx[0]), py(khop.tungDoGoc + khop.heSoGoc * vx[0]));
      ctx.lineTo(px(vx[vx.length - 1]), py(khop.tungDoGoc + khop.heSoGoc * vx[vx.length - 1])); ctx.stroke();
    }
    ctx.fillStyle = MAU.chinh;
    diem.forEach(function (d) { ctx.beginPath(); ctx.arc(px(d[0]), py(d[1]), 5, 0, 2 * Math.PI); ctx.fill(); });
  }

  function khoiDong() {
    var duLieu = JSON.parse(document.getElementById('du-lieu').textContent);
    var cauHinh = duLieu.cauHinh, khaiBao = duLieu.khaiBao, moHinh = root.THI_NGHIEM_MO_HINH;
    var nv = taoNhiemVu(cauHinh);
    var rng = taoNgauNhien(Date.now() % 2147483647);
    var tt = { p: {}, t: 0, dangChay: false, daXong: khaiBao.hoatHinh !== 'mot-lan', d: null, mau: null, mocThoiGian: 0 };
    var theoMa = {};
    khaiBao.thamSo.forEach(function (ts) { theoMa[ts.ma] = ts; });
    khaiBao.daiLuongDo.forEach(function (dl) { theoMa[dl.ma] = dl; });
    Object.keys(cauHinh.thamSo).forEach(function (ma) {
      var ch = cauHinh.thamSo[ma];
      tt.p[ma] = ch.kieu === 'co-dinh' ? ch.giaTri : ch.macDinh;
    });

    document.getElementById('tieu-de').innerHTML = danhDau(cauHinh.tieuDe);
    document.getElementById('phu-de').innerHTML = danhDau(cauHinh.mon + ' ' + cauHinh.lop + ' · ' + khaiBao.ten);
    document.getElementById('cong-thuc').innerHTML = '<b>Mô hình:</b> ' + danhDau(khaiBao.congThuc.bieuThuc);
    document.getElementById('dieu-kien').innerHTML = '<b>Điều kiện lí tưởng hoá:</b> ' + danhDau(khaiBao.congThuc.dieuKien) +
      (khaiBao.congThuc.nguon ? ' <b>Nguồn số liệu:</b> ' + danhDau(khaiBao.congThuc.nguon) : '');

    var kiem = tuKiem(khaiBao, moHinh);
    var dongKiem = document.getElementById('tu-kiem');
    if (kiem.dat === kiem.tong) {
      dongKiem.textContent = 'Tự kiểm: ' + kiem.dat + '/' + kiem.tong + ' đạt.';
    } else {
      dongKiem.textContent = 'Tự kiểm: ' + kiem.dat + '/' + kiem.tong + ' đạt; trượt dòng ' +
        kiem.truot.map(function (m) { return m.dong; }).join(', ') + '.';
      var canhBao = tao('div', 'dai-do', 'Mô hình không qua tự kiểm — không dùng để dạy.');
      document.body.insertBefore(canhBao, document.body.firstChild);
    }

    var khungVe = document.getElementById('khung-ve');
    var oSoDo = document.getElementById('so-do');
    var oChay = document.getElementById('dieu-khien-chay');
    var oThamSo = document.getElementById('tham-so');
    var oNhiemVu = document.getElementById('nhiem-vu');
    var oBang = document.getElementById('bang');
    var oKhop = document.getElementById('khop');
    var khungDoThi = document.getElementById('do-thi');
    var cacNutKhoa = [];

    function nhan(ma) { var m = theoMa[ma]; return danhDau(m.ten) + (m.donVi ? ' (' + danhDau(m.donVi) + ')' : ''); }
    function nhanTho(ma) { var m = theoMa[ma]; return (m.ten + (m.donVi ? ' (' + m.donVi + ')' : '')).replace(/[~^*]/g, ''); }
    function chuSoCua(ma) { var m = theoMa[ma]; return m.chuSo !== undefined ? m.chuSo : soChuSo(m.buoc || 1); }
    function hienGiaTri(ma, giaTri) {
      var m = theoMa[ma];
      if (m.kieu === 'chon') { return danhDau(m.luaChon.filter(function (lc) { return lc.ma === giaTri; })[0].ten); }
      return dinhDang(giaTri, chuSoCua(ma));
    }
    function nhanBieuThuc(bt) {
      var goc = nhanTho(bt.ma);
      return { 'binh-phuong': '(' + goc + ')²', 'nghich-dao': '1/(' + goc + ')', 'ln': 'ln(' + goc + ')', 'can': '√(' + goc + ')' }[bt.phep] || goc;
    }

    function layMau() { tt.mau = mauDo(khaiBao, tt.d, rng, cauHinh.saiSo); }

    function veLai() {
      var k = chuanBiKhung(khungVe, 0.62);
      moHinh.ve(k.ctx, tt.p, tt.t, k.kt, tt.d);
    }

    function hienSoDo() {
      oSoDo.innerHTML = '';
      khaiBao.daiLuongDo.forEach(function (dl) {
        var o = tao('div', 'o-do');
        o.appendChild(tao('span', 'ten-do', nhan(dl.ma)));
        o.appendChild(tao('span', 'gia-tri-do', tt.daXong && nv.duocThaoTac() ? dinhDang(tt.mau[dl.ma], dl.chuSo) : '…'));
        oSoDo.appendChild(o);
      });
    }

    function tinhLai(giuThoiGian) {
      tt.d = moHinh.tinh(tt.p);
      if (!giuThoiGian) { tt.t = 0; tt.dangChay = false; tt.daXong = khaiBao.hoatHinh !== 'mot-lan'; }
      layMau(); veLai(); hienSoDo(); capNhatKhoa();
    }

    function nhip(moc) {
      if (!tt.dangChay) { return; }
      tt.t += (moc - tt.mocThoiGian) / 1000; tt.mocThoiGian = moc;
      if (khaiBao.hoatHinh === 'mot-lan' && tt.t >= moHinh.thoiLuong(tt.p, tt.d)) {
        tt.t = moHinh.thoiLuong(tt.p, tt.d); tt.dangChay = false; tt.daXong = true;
        layMau(); hienSoDo(); capNhatKhoa();
      }
      veLai();
      if (tt.dangChay) { window.requestAnimationFrame(nhip); }
    }

    function capNhatKhoa() {
      var mo = nv.duocThaoTac();
      cacNutKhoa.forEach(function (nut) { nut.disabled = !mo; });
      if (nutGhi) { nutGhi.disabled = !mo || !tt.daXong; }
      if (nutChay) { nutChay.textContent = tt.dangChay ? 'Dừng' : 'Chạy'; }
      document.body.className = mo ? '' : 'dang-khoa';
    }

    // Tham số
    var oCoDinh = tao('ul', 'co-dinh');
    khaiBao.thamSo.forEach(function (ts) {
      var ch = cauHinh.thamSo[ts.ma];
      if (ch.kieu === 'co-dinh') {
        oCoDinh.appendChild(tao('li', '', nhan(ts.ma) + ': <b>' + hienGiaTri(ts.ma, ch.giaTri) + '</b>'));
        return;
      }
      var hang = tao('label', 'hang-tham-so');
      var giaTri = tao('b', 'gia-tri', hienGiaTri(ts.ma, tt.p[ts.ma]));
      hang.appendChild(tao('span', '', nhan(ts.ma) + ': '));
      hang.appendChild(giaTri);
      var nhap;
      if (ch.kieu === 'chon') {
        nhap = tao('select');
        ch.luaChon.forEach(function (ma) {
          var muc = tao('option', '', hienGiaTri(ts.ma, ma)); muc.value = ma; nhap.appendChild(muc);
        });
        nhap.value = ch.macDinh;
      } else {
        nhap = tao('input'); nhap.type = 'range';
        nhap.min = ch.min; nhap.max = ch.max; nhap.step = ch.buoc; nhap.value = ch.macDinh;
      }
      nhap.addEventListener('input', function () {
        tt.p[ts.ma] = ch.kieu === 'chon' ? nhap.value : Number(nhap.value);
        giaTri.innerHTML = hienGiaTri(ts.ma, tt.p[ts.ma]);
        tinhLai(false);
      });
      cacNutKhoa.push(nhap);
      hang.appendChild(nhap);
      oThamSo.appendChild(hang);
    });
    if (oCoDinh.childNodes.length) {
      oThamSo.appendChild(tao('p', 'nho', 'Giữ cố định:'));
      oThamSo.appendChild(oCoDinh);
    }

    // Chạy, ghi lần đo
    var nutChay = null, nutGhi = tao('button', 'nut chinh', 'Ghi lần đo');
    if (khaiBao.hoatHinh !== 'khong') {
      nutChay = tao('button', 'nut', 'Chạy');
      nutChay.addEventListener('click', function () {
        if (tt.dangChay) { tt.dangChay = false; capNhatKhoa(); return; }
        if (khaiBao.hoatHinh === 'mot-lan' && tt.daXong) { tt.t = 0; tt.daXong = false; hienSoDo(); }
        tt.dangChay = true; tt.mocThoiGian = window.performance.now(); capNhatKhoa();
        window.requestAnimationFrame(nhip);
      });
      var nutDatLai = tao('button', 'nut', 'Đặt lại');
      nutDatLai.addEventListener('click', function () { tinhLai(false); });
      cacNutKhoa.push(nutChay, nutDatLai);
      oChay.appendChild(nutChay); oChay.appendChild(nutDatLai);
    }
    nutGhi.addEventListener('click', function () {
      var dong = {};
      cauHinh.quanSat.cot.forEach(function (ma) { dong[ma] = theoMa[ma].saiSo !== undefined ? tt.mau[ma] : tt.p[ma]; });
      if (nv.ghiLanDo(dong)) { layMau(); hienSoDo(); veBang(); veNhiemVu(); }
    });
    oChay.appendChild(nutGhi);
    if (cauHinh.saiSo) { oChay.appendChild(tao('span', 'nho', 'Đang bật sai số đo: mỗi lần đo lệch ngẫu nhiên một chút, như đo thật.')); }

    // Bảng số liệu và đồ thị
    function veBang() {
      var cot = cauHinh.quanSat.cot, doThi = cauHinh.quanSat.doThi;
      var html = '<table><thead><tr><th>Lần</th>' + cot.map(function (ma) { return '<th>' + nhan(ma) + '</th>'; }).join('') + '<th></th></tr></thead><tbody>';
      nv.lanDo.forEach(function (dong, chiSo) {
        html += '<tr><td>' + (chiSo + 1) + '</td>' + cot.map(function (ma) { return '<td>' + hienGiaTri(ma, dong[ma]) + '</td>'; }).join('') +
          '<td><button class="xoa" data-chi-so="' + chiSo + '" title="Xoá lần đo này">×</button></td></tr>';
      });
      oBang.innerHTML = html + '</tbody></table>';
      Array.prototype.forEach.call(oBang.querySelectorAll('.xoa'), function (nut) {
        nut.addEventListener('click', function () { nv.xoaLanDo(Number(nut.getAttribute('data-chi-so'))); veBang(); veNhiemVu(); });
      });
      var nutChep = tao('button', 'nut', 'Chép số liệu');
      nutChep.addEventListener('click', function () {
        var dongChu = [['Lần'].concat(cot.map(nhanTho)).join('\t')];
        nv.lanDo.forEach(function (dong, chiSo) {
          dongChu.push([chiSo + 1].concat(cot.map(function (ma) { return hienGiaTri(ma, dong[ma]).replace(/<[^>]+>/g, ''); })).join('\t'));
        });
        var vung = tao('textarea'); vung.value = dongChu.join('\n'); document.body.appendChild(vung);
        vung.select(); document.execCommand('copy'); document.body.removeChild(vung);
        nutChep.textContent = 'Đã chép, dán vào Excel';
      });
      oBang.appendChild(nutChep);
      if (!doThi) { khungDoThi.style.display = 'none'; oKhop.textContent = ''; return; }
      var diem = diemDoThi(doThi, nv.lanDo), khop = khopTuyenTinh(diem);
      veDoThi(khungDoThi, diem, khop, nhanBieuThuc(doThi.hoanh), nhanBieuThuc(doThi.tung));
      oKhop.textContent = khop
        ? 'Đường thẳng khớp: hệ số góc = ' + dinhDang(khop.heSoGoc, 4) + '; tung độ gốc = ' + dinhDang(khop.tungDoGoc, 4) +
          '; hệ số tương quan r = ' + dinhDang(khop.tuongQuan, 4)
        : 'Cần ít nhất hai lần đo khác nhau để vẽ đường thẳng khớp.';
    }

    // Ba bước nhiệm vụ
    function veNhiemVu() {
      oNhiemVu.innerHTML = '';
      var b1 = tao('div', 'buoc' + (nv.buoc === 'du-doan' ? ' dang-lam' : ''));
      b1.appendChild(tao('h2', '', '1. Dự đoán'));
      b1.appendChild(tao('p', '', danhDau(cauHinh.duDoan.cau)));
      if (nv.buoc === 'du-doan') {
        var oNhap;
        if (cauHinh.duDoan.luaChon.length) {
          oNhap = tao('div');
          cauHinh.duDoan.luaChon.forEach(function (lc) {
            var dong = tao('label', 'lua-chon');
            var o = tao('input'); o.type = 'radio'; o.name = 'du-doan'; o.value = lc.ma;
            dong.appendChild(o); dong.appendChild(tao('span', '', '<b>' + lc.ma + '.</b> ' + danhDau(lc.noiDung)));
            oNhap.appendChild(dong);
          });
        } else { oNhap = tao('textarea'); oNhap.rows = 3; oNhap.placeholder = 'Viết dự đoán của em'; }
        b1.appendChild(oNhap);
        var nutChot = tao('button', 'nut chinh', 'Chốt dự đoán');
        nutChot.addEventListener('click', function () {
          var chon = cauHinh.duDoan.luaChon.length ? (oNhap.querySelector('input:checked') || {}).value : oNhap.value;
          if (nv.chonDuDoan(chon)) { veNhiemVu(); hienSoDo(); capNhatKhoa(); }
        });
        b1.appendChild(nutChot);
        b1.appendChild(tao('p', 'nho', 'Chốt dự đoán xong mới làm được thí nghiệm. Dự đoán không sửa được.'));
      } else { b1.appendChild(tao('p', 'da-chon', 'Dự đoán của em: <b>' + danhDau(nv.duDoan || '(giáo viên trình diễn)') + '</b>')); }
      oNhiemVu.appendChild(b1);

      var b2 = tao('div', 'buoc' + (nv.buoc === 'quan-sat' ? ' dang-lam' : ''));
      b2.appendChild(tao('h2', '', '2. Quan sát'));
      b2.appendChild(tao('p', '', 'Thay đổi tham số, làm thí nghiệm và bấm <b>Ghi lần đo</b>. Đã ghi <b>' + nv.lanDo.length +
        '/' + cauHinh.quanSat.soLanDo + '</b> lần đo.'));
      oNhiemVu.appendChild(b2);

      var b3 = tao('div', 'buoc' + (nv.buoc === 'giai-thich' ? ' dang-lam' : ''));
      b3.appendChild(tao('h2', '', '3. Giải thích'));
      if (nv.buoc === 'giai-thich' || nv.giaoVien) {
        b3.appendChild(tao('p', '', danhDau(cauHinh.giaiThich.cau)));
        var dung = nv.duDoanDung();
        if (dung !== null) {
          b3.appendChild(tao('p', dung ? 'dung' : 'sai', dung ? 'Dự đoán của em khớp với kết quả thí nghiệm.'
            : 'Dự đoán của em chưa khớp với kết quả. Hãy dùng số liệu để giải thích vì sao.'));
        }
        var oGoiY = tao('div', 'goi-y'); oGoiY.style.display = 'none';
        oGoiY.innerHTML = '<p><b>Gợi ý đáp án:</b> ' + danhDau(cauHinh.giaiThich.goiY) + '</p><p><b>Kết luận:</b> ' + danhDau(cauHinh.ketLuan) + '</p>';
        var nutGoiY = tao('button', 'nut', 'Xem gợi ý đáp án và kết luận');
        nutGoiY.addEventListener('click', function () { oGoiY.style.display = 'block'; });
        b3.appendChild(nutGoiY); b3.appendChild(oGoiY);
      } else { b3.appendChild(tao('p', 'nho', 'Mở ra khi đã ghi đủ số lần đo.')); }
      oNhiemVu.appendChild(b3);

      if (!nv.giaoVien) {
        var nutGiaoVien = tao('button', 'nut nho-nut', 'Chế độ giáo viên (bỏ khoá)');
        nutGiaoVien.addEventListener('click', function () { nv.batGiaoVien(); veNhiemVu(); hienSoDo(); capNhatKhoa(); });
        oNhiemVu.appendChild(nutGiaoVien);
      }
    }

    window.addEventListener('resize', function () { veLai(); veBang(); });
    tinhLai(false); veBang(); veNhiemVu();
  }

  root.THI_NGHIEM_KHUNG = {
    taoNgauNhien: taoNgauNhien, nhieuChuan: nhieuChuan,
    thamSoMacDinh: thamSoMacDinh, gopThamSo: gopThamSo, tuKiem: tuKiem,
    apDungPhep: apDungPhep, khopTuyenTinh: khopTuyenTinh, dinhDang: dinhDang, danhDau: danhDau,
    mauDo: mauDo, taoNhiemVu: taoNhiemVu, diemDoThi: diemDoThi,
    MAU: MAU, PHONG: PHONG, khoiDong: khoiDong
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
```

- [ ] **Step 4: Viết `khung.css`**

`tools/vi/thi_nghiem_parts/runtime/khung.css`:

```css
* { box-sizing: border-box; }
body { margin: 0; font-family: "Segoe UI", Arial, sans-serif; font-size: 17px; line-height: 1.5; color: #1e293b; background: #f1f5f9; }
.dau { background: #1b365d; color: #ffffff; padding: 14px 20px; }
.dau h1 { margin: 0; font-size: 26px; }
.dau p { margin: 2px 0 0; opacity: 0.85; }
.dai-do { background: #dc2626; color: #ffffff; font-weight: 700; padding: 12px 20px; font-size: 19px; }
.bo-cuc { display: grid; grid-template-columns: minmax(0, 3fr) minmax(0, 2fr); gap: 16px; padding: 16px; max-width: 1500px; margin: 0 auto; }
.bo-cuc > section { background: #ffffff; border: 1px solid #cbd5e1; border-radius: 12px; padding: 14px; min-width: 0; }
.so-lieu { grid-column: 1 / -1; }
canvas { display: block; width: 100%; border-radius: 8px; }
#khung-ve { border: 1px solid #e2e8f0; }
#dieu-khien-chay { display: flex; flex-wrap: wrap; align-items: center; gap: 10px; margin: 12px 0; }
#so-do { display: flex; flex-wrap: wrap; gap: 10px; }
.o-do { flex: 1 1 180px; background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 10px; padding: 8px 12px; }
.ten-do { display: block; font-size: 14px; color: #475569; }
.gia-tri-do { font-size: 26px; font-weight: 700; color: #1b365d; font-variant-numeric: tabular-nums; }
.hang-tham-so { display: block; margin-bottom: 14px; }
.hang-tham-so input[type="range"], .hang-tham-so select { display: block; width: 100%; margin-top: 6px; font: inherit; }
.hang-tham-so input[type="range"] { height: 28px; }
.gia-tri { color: #2563eb; }
.co-dinh { margin: 4px 0 0; padding-left: 20px; }
.nho { font-size: 14px; color: #64748b; margin: 6px 0; }
.nut { font: inherit; font-weight: 600; padding: 9px 16px; border-radius: 10px; border: 1px solid #94a3b8; background: #ffffff; color: #1e293b; cursor: pointer; }
.nut.chinh { background: #2563eb; border-color: #2563eb; color: #ffffff; }
.nut:disabled, input:disabled, select:disabled { opacity: 0.45; cursor: not-allowed; }
.nho-nut { font-size: 14px; font-weight: 400; margin-top: 8px; }
.buoc { border-left: 5px solid #cbd5e1; padding: 2px 0 2px 12px; margin: 14px 0; }
.buoc.dang-lam { border-left-color: #f59e0b; }
.buoc h2 { font-size: 19px; margin: 0 0 4px; }
.buoc p { margin: 4px 0; }
.buoc textarea { width: 100%; font: inherit; padding: 8px; border-radius: 8px; border: 1px solid #94a3b8; }
.lua-chon { display: flex; gap: 8px; align-items: baseline; padding: 4px 0; cursor: pointer; }
.lua-chon input { transform: scale(1.3); }
.da-chon { background: #eff6ff; border-radius: 8px; padding: 6px 10px; }
.dung { color: #15803d; font-weight: 600; }
.sai { color: #b45309; font-weight: 600; }
.goi-y { background: #f0fdf4; border: 1px solid #86efac; border-radius: 10px; padding: 4px 12px; margin-top: 8px; }
table { border-collapse: collapse; width: 100%; margin-bottom: 10px; font-variant-numeric: tabular-nums; }
th, td { border: 1px solid #cbd5e1; padding: 6px 10px; text-align: center; }
th { background: #f1f5f9; font-size: 15px; }
.xoa { border: 0; background: none; color: #dc2626; font-size: 20px; cursor: pointer; line-height: 1; }
#do-thi { border: 1px solid #e2e8f0; margin-top: 12px; max-width: 760px; }
#khop { font-size: 15px; color: #334155; margin-top: 6px; }
.dang-khoa #khung-ve { opacity: 0.55; }
footer { max-width: 1500px; margin: 0 auto; padding: 0 20px 24px; font-size: 14px; color: #475569; }
footer p { margin: 4px 0; }
@media (max-width: 820px) {
  body { font-size: 16px; }
  .bo-cuc { grid-template-columns: minmax(0, 1fr); padding: 10px; }
  .dau h1 { font-size: 21px; }
}
```

- [ ] **Step 5: Viết bộ ghép HTML**

`tools/vi/thi_nghiem_parts/build_html.py`:

```python
"""Ghép khung chạy, mô hình và cấu hình thành một file HTML tự chứa, không gọi Internet."""

from __future__ import annotations

import html
import json
import re
from pathlib import Path

RUNTIME_DIR = Path(__file__).resolve().parent / "runtime"
FILENAME = "thi-nghiem.html"
NETWORK_MARKS = ("http://", "https://", "//cdn", "@import", "url(")

TEMPLATE = """<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<style>
{css}
</style>
</head>
<body>
<header class="dau"><h1 id="tieu-de"></h1><p id="phu-de"></p></header>
<main class="bo-cuc">
<section class="mo-phong"><canvas id="khung-ve"></canvas><div id="dieu-khien-chay"></div><div id="so-do"></div></section>
<section class="ben-phai"><div id="tham-so"></div><div id="nhiem-vu"></div></section>
<section class="so-lieu"><div id="bang"></div><canvas id="do-thi"></canvas><div id="khop"></div></section>
</main>
<footer><p id="cong-thuc"></p><p id="dieu-kien"></p><p id="tu-kiem"></p></footer>
<script>
{khung_js}
</script>
<script>
{model_js}
</script>
<script id="du-lieu" type="application/json">{data}</script>
<script>THI_NGHIEM_KHUNG.khoiDong();</script>
</body>
</html>
"""


class BuildError(Exception):
    """File HTML ghép ra không đạt điều kiện chạy không cần mạng."""


def _embed_json(data: dict) -> str:
    """JSON an toàn bên trong thẻ <script>: không để lọt `</script` hay `<!--`."""
    return json.dumps(data, ensure_ascii=False).replace("<", "\\u003c")


def build(experiment, model) -> str:
    khung_js = (RUNTIME_DIR / "khung.js").read_text(encoding="utf-8")
    css = (RUNTIME_DIR / "khung.css").read_text(encoding="utf-8")
    page = TEMPLATE.format(
        title=html.escape(re.sub(r"[~^*]", "", experiment.meta["tieu-de"])),
        css=css,
        khung_js=khung_js,
        model_js=model.js,
        data=_embed_json({"cauHinh": experiment.config(), "khaiBao": model.khai_bao}),
    )
    found = [mark for mark in NETWORK_MARKS if mark in page]
    if found:
        raise BuildError("File HTML chứa dấu hiệu tải từ Internet: " + ", ".join(found))
    return page


def write(experiment, model, folder: Path) -> Path:
    path = folder / FILENAME
    path.write_text(build(experiment, model), encoding="utf-8", newline="\n")
    return path
```

- [ ] **Step 6: Chạy lại, kể cả test Node của Task 1**

Run: `venv\Scripts\python.exe -m unittest tools.vi.tests.test_thi_nghiem_html tools.vi.tests.test_thi_nghiem_khung tools.vi.tests.test_thi_nghiem_mo_hinh -v`
Expected: PASS (3 + 1 + 11 test). Test Node vẫn qua chứng tỏ phần logic không bị đổi.

- [ ] **Step 7: Commit**

```bash
git add tools/vi/thi_nghiem_parts/runtime tools/vi/thi_nghiem_parts/build_html.py tools/vi/tests/test_thi_nghiem_html.py
git commit -m "feat(vi): add the experiment page UI and the offline HTML builder"
```

---

### Task 9: Phiếu học tập Word

**Files:**
- Create: `tools/vi/thi_nghiem_parts/phieu.py`
- Create: `tools/vi/tests/test_thi_nghiem_phieu.py`

**Interfaces:**
- Consumes: `word_parts.base.{new_document, write, grid_borders, set_widths, fill_cell}` (sẵn có); `Experiment`, `Axis` (Task 7); `Model` (Task 3); `tham_chieu.THAM_CHIEU` (Task 2); `kiem_so.luoi`, `kiem_so.CheckError` (Task 3).
- Produces: `phieu.build(experiment, model, folder) -> tuple[Path, list[str]]` (đường dẫn file, các cảnh báo); `phieu.ideal_table(experiment, model) -> tuple[list[list[str]]|None, str]`; `phieu.FILENAME = "phieu-hoc-tap.docx"`; `phieu.GRID_ROWS`, `phieu.GRID_COLUMNS`.

- [ ] **Step 1: Viết test**

`tools/vi/tests/test_thi_nghiem_phieu.py`:

```python
"""Test cho phiếu học tập Word của thí nghiệm ảo."""

import sys
import tempfile
import unittest
from pathlib import Path

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from thi_nghiem_mau import PENDULUM  # noqa: E402
from thi_nghiem_parts import parse, thu_vien  # noqa: E402


class WorksheetTest(unittest.TestCase):
    def setUp(self):
        from docx import Document
        from thi_nghiem_parts import phieu
        self.Document, self.phieu = Document, phieu
        self.model = thu_vien.load("li-con-lac-don", Path("."))
        self.experiment = parse.parse_experiment(PENDULUM, self.model.khai_bao)

    def test_ideal_table_uses_the_reference_model(self):
        rows, warning = self.phieu.ideal_table(self.experiment, self.model)
        self.assertEqual(warning, "")
        self.assertEqual([row[0] for row in rows], ["0,40", "0,80", "1,00", "1,40", "1,60"])
        self.assertEqual(rows[2][1], "2,007")

    def test_document_has_student_table_grid_and_teacher_page(self):
        with tempfile.TemporaryDirectory() as tmp:
            path, warnings = self.phieu.build(self.experiment, self.model, Path(tmp))
            document = self.Document(str(path))
        self.assertEqual(warnings, [])
        student, graph, ideal = document.tables
        self.assertEqual((len(student.rows), len(student.columns)), (6, 3))
        self.assertEqual([cell.text for cell in student.rows[0].cells], ["Lần", "Chiều dài dây l (m)", "Chu kì T (s)"])
        self.assertEqual((len(graph.rows), len(graph.columns)), (self.phieu.GRID_ROWS, self.phieu.GRID_COLUMNS))
        self.assertEqual(len(ideal.rows), 6)
        text = "\n".join(paragraph.text for paragraph in document.paragraphs)
        for needle in ("PHIẾU HỌC TẬP", "DÀNH CHO GIÁO VIÊN", "Đáp án phần dự đoán: B", "Trục tung: (Chu kì T (s))2", "T = 2π√(l/g)"):
            self.assertIn(needle, text)
        self.assertLess(text.index("Dự đoán của em"), text.index("DÀNH CHO GIÁO VIÊN"))

    def test_no_slider_column_means_no_ideal_table(self):
        self.experiment.cot = ["chu-ki"]
        rows, warning = self.phieu.ideal_table(self.experiment, self.model)
        self.assertIsNone(rows)
        self.assertIn("thanh trượt", warning)

if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Chạy để thấy trượt**

Run: `venv\Scripts\python.exe -m unittest tools.vi.tests.test_thi_nghiem_phieu -v`
Expected: ERROR `ImportError: cannot import name 'phieu'`.

- [ ] **Step 3: Viết bộ dựng phiếu**

`tools/vi/thi_nghiem_parts/phieu.py`:

```python
"""Dựng phiếu học tập Word cho một thí nghiệm ảo: trang học sinh và trang giáo viên."""

from __future__ import annotations

import math
from pathlib import Path

from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.shared import Cm, Pt

from word_parts import base

from . import kiem_so, tham_chieu

FILENAME = "phieu-hoc-tap.docx"
DOTS = "." * 118
GRID_ROWS, GRID_COLUMNS = 14, 20
GRID_CELL_CM = 0.75
TRANSFORM_LABELS = {"binh-phuong": "({})^2^", "nghich-dao": "1/({})", "ln": "ln({})", "can": "√({})", "khong": "{}"}


def label(model, ma: str) -> str:
    item = model.tham_so(ma) or model.dai_luong(ma)
    return item["ten"] + (f" ({item['donVi']})" if item.get("donVi") else "")


def axis_label(model, axis) -> str:
    return TRANSFORM_LABELS[axis.phep].format(label(model, axis.ma))


def _format(value, digits: int) -> str:
    if value is None:
        return "—"
    if isinstance(value, str):
        return value
    return f"{value:.{digits}f}".replace(".", ",")


def _digits(model, ma: str) -> int:
    measure = model.dai_luong(ma)
    if measure is not None:
        return measure["chuSo"]
    step = str(model.tham_so(ma).get("buoc", 1))
    return len(step.split(".")[1]) if "." in step else 0


def _cell(model, ma: str, point: dict, result: dict) -> str:
    parameter = model.tham_so(ma)
    if parameter is None:
        return _format(result.get(ma), _digits(model, ma))
    if parameter["kieu"] == "chon":
        return next(option["ten"] for option in parameter["luaChon"] if option["ma"] == point[ma])
    return _format(point[ma], _digits(model, ma))


def ideal_table(experiment, model) -> tuple[list[list[str]] | None, str]:
    """Bảng số liệu lí tưởng cho giáo viên. Trả về (bảng hoặc None, cảnh báo hoặc chuỗi rỗng)."""
    varied = next((ma for ma in experiment.cot
                   if ma in experiment.tham_so and experiment.tham_so[ma]["kieu"] == "truot"), None)
    if varied is None:
        return None, "Phiếu không có bảng số liệu lí tưởng vì bảng đo không có tham số dạng thanh trượt."
    chosen = experiment.tham_so[varied]
    count = experiment.so_lan_do
    steps = round((chosen["max"] - chosen["min"]) / chosen["buoc"])
    # floor(x + 0.5) thay cho round(): round() của Python làm tròn về số chẵn nên các mốc bị lệch về một phía.
    values = sorted({round(chosen["min"] + math.floor(index * steps / (count - 1) + 0.5) * chosen["buoc"], 10)
                     for index in range(count)})
    points = []
    for value in values:
        point = {ma: (ch["giaTri"] if ch["kieu"] == "co-dinh" else ch["macDinh"]) for ma, ch in experiment.tham_so.items()}
        point[varied] = value
        points.append(point)
    if model.ma in tham_chieu.THAM_CHIEU:
        results = [tham_chieu.THAM_CHIEU[model.ma](point) for point in points]
    else:
        try:
            results = kiem_so.luoi(model, points)
        except kiem_so.CheckError as exc:
            return None, f"Phiếu không có bảng số liệu lí tưởng: {exc}"
        if results is None:
            return None, "Phiếu không có bảng số liệu lí tưởng vì máy này không có Node để chạy mô hình mới."
    return [[_cell(model, ma, point, result) for ma in experiment.cot] for point, result in zip(points, results)], ""


def _heading(document, text: str) -> None:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(8)
    base.write(paragraph, text, bold=True)


def _line(document, text: str, **kwargs) -> None:
    base.write(document.add_paragraph(), text, **kwargs)


def _dots(document, count: int) -> None:
    for _ in range(count):
        document.add_paragraph(DOTS)


def _data_table(document, headers: list[str], rows: list[list[str]]) -> None:
    table = document.add_table(rows=1 + len(rows), cols=len(headers))
    base.grid_borders(table)
    for index, header in enumerate(headers):
        base.fill_cell(table.rows[0].cells[index], [header])
    for row_index, row in enumerate(rows, 1):
        for column, value in enumerate(row):
            base.fill_cell(table.rows[row_index].cells[column], [value], bold_first=False)
        table.rows[row_index].height = Cm(0.9)


def _grid(document) -> None:
    table = document.add_table(rows=GRID_ROWS, cols=GRID_COLUMNS)
    base.grid_borders(table)
    base.set_widths(table, [Cm(GRID_CELL_CM)] * GRID_COLUMNS)
    for row in table.rows:
        row.height = Cm(GRID_CELL_CM)
        row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY


def build(experiment, model, folder: Path) -> tuple[Path, list[str]]:
    warnings: list[str] = []
    document = base.new_document(margins_cm=(1.8, 1.8, 2.0, 1.5), size_pt=13, line_spacing=1.2)
    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    base.write(title, "PHIẾU HỌC TẬP", bold=True)
    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    base.write(subtitle, f"{experiment.meta['tieu-de']} — {experiment.meta['mon']} {experiment.meta['lop']}", bold=True)
    _line(document, "Họ và tên / Nhóm: ................................................................  Lớp: ..................")

    _heading(document, "1. Dự đoán (làm trước khi chạy thí nghiệm)")
    _line(document, experiment.du_doan_cau)
    for ma, noi_dung in experiment.lua_chon:
        _line(document, f"**{ma}.** {noi_dung}")
    _line(document, "Dự đoán của em: " + "." * 90)

    _heading(document, "2. Quan sát")
    _line(document, f"Thay đổi tham số, làm thí nghiệm và ghi {experiment.so_lan_do} lần đo vào bảng.")
    headers = ["Lần"] + [label(model, ma) for ma in experiment.cot]
    _data_table(document, headers, [[str(index)] + [""] * len(experiment.cot) for index in range(1, experiment.so_lan_do + 1)])
    if experiment.do_thi is not None:
        tung, hoanh = experiment.do_thi
        _line(document, f"Vẽ đồ thị. Trục tung: {axis_label(model, tung)}. Trục hoành: {axis_label(model, hoanh)}.")
        _grid(document)

    _heading(document, "3. Giải thích")
    _line(document, experiment.giai_thich_cau)
    _dots(document, 4)
    _heading(document, "Kết luận")
    _dots(document, 3)

    document.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    teacher = document.add_paragraph()
    teacher.alignment = WD_ALIGN_PARAGRAPH.CENTER
    base.write(teacher, "DÀNH CHO GIÁO VIÊN — KHÔNG PHÁT CHO HỌC SINH", bold=True)
    if experiment.dap_an:
        _line(document, f"**Đáp án phần dự đoán:** {experiment.dap_an}")
    _line(document, f"**Gợi ý phần giải thích:** {experiment.goi_y}")
    _line(document, f"**Kết luận:** {experiment.ket_luan}")
    formula = model.khai_bao["congThuc"]
    _line(document, f"**Mô hình của thí nghiệm ảo:** {formula['bieuThuc']}")
    _line(document, f"**Điều kiện lí tưởng hoá:** {formula['dieuKien']}")
    if formula.get("nguon"):
        _line(document, f"**Nguồn số liệu:** {formula['nguon']}")
    rows, warning = ideal_table(experiment, model)
    if rows is None:
        warnings.append(warning)
    else:
        _heading(document, "Số liệu lí tưởng (không có sai số đo)")
        _data_table(document, [label(model, ma) for ma in experiment.cot], rows)

    path = folder / FILENAME
    document.save(str(path))
    return path, warnings
```

- [ ] **Step 4: Chạy lại**

Run: `venv\Scripts\python.exe -m unittest tools.vi.tests.test_thi_nghiem_phieu -v`
Expected: PASS (3 test).

- [ ] **Step 5: Commit**

```bash
git add tools/vi/thi_nghiem_parts/phieu.py tools/vi/tests/test_thi_nghiem_phieu.py
git commit -m "feat(vi): add the Word worksheet for virtual experiments"
```

---

### Task 10: Công cụ dòng lệnh `thi_nghiem.py`

**Files:**
- Create: `tools/vi/thi_nghiem.py`
- Create: `tools/vi/tests/test_thi_nghiem_cong_cu.py`

**Interfaces:**
- Consumes: `parse.read_meta`, `parse.parse_experiment`, `parse.ParseError` (Task 7); `thu_vien.load`, `thu_vien.ModelError` (Task 3); `kiem_so.bang_kiem`, `kiem_so.CheckError` (Task 3); `build_html.write`, `build_html.BuildError` (Task 8); `phieu.build` (Task 9).
- Produces: `python tools/vi/thi_nghiem.py <thư_mục> [--phan tat-ca|html|phieu] [--plan-only]`. stdout một dòng JSON với khoá `ready`, `files`, `mau`, `tham_so`, `so_lan_do`, `kiem_so{chay, dat, tong}`, `warnings`, `error{step, message, fix}|null`; `error.step` ∈ `input`, `parse`, `model`, `check`, `docx`, `write`, `internal`. Ghi `thi-nghiem.html`, `phieu-hoc-tap.docx`, `can-soat.md` vào chính thư mục đó. Hàm `thi_nghiem.main(argv) -> int`, `thi_nghiem.load_phieu()`.

- [ ] **Step 1: Viết test**

`tools/vi/tests/test_thi_nghiem_cong_cu.py`:

```python
"""Test cho công cụ dòng lệnh thi_nghiem.py."""

import contextlib
import io
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

TOOLS_VI = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(TOOLS_VI))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from thi_nghiem_mau import HOOKE_JS, HOOKE_JSON, HOOKE_MD, PENDULUM  # noqa: E402
from thi_nghiem_parts import build_html, kiem_so, thu_vien  # noqa: E402

import thi_nghiem  # noqa: E402

HAS_NODE = kiem_so.find_node() is not None
NEED_NODE = "máy không có Node nên không chạy được JavaScript"


class CliCase(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.folder = Path(self._tmp.name) / "con-lac"
        self.folder.mkdir()
        (self.folder / "thi-nghiem.md").write_text(PENDULUM, encoding="utf-8")

    def tearDown(self):
        self._tmp.cleanup()

    def run_tool(self, *argv):
        out = io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(io.StringIO()):
            code = thi_nghiem.main([str(arg) for arg in argv])
        lines = out.getvalue().splitlines()
        self.assertEqual(len(lines), 1, out.getvalue())
        return code, json.loads(lines[0])

    def new_model(self, js=HOOKE_JS, kb=None):
        (self.folder / "thi-nghiem.md").write_text(HOOKE_MD, encoding="utf-8")
        (self.folder / "mo-hinh.json").write_text(json.dumps(kb or HOOKE_JSON, ensure_ascii=False), encoding="utf-8")
        (self.folder / "mo-hinh.js").write_text(js, encoding="utf-8")


class CliTest(CliCase):
    def test_builds_page_worksheet_and_review_file(self):
        code, data = self.run_tool(self.folder)
        self.assertEqual(code, 0, data)
        self.assertTrue(data["ready"])
        self.assertEqual([Path(path).name for path in data["files"]], ["thi-nghiem.html", "phieu-hoc-tap.docx", "can-soat.md"])
        self.assertEqual((data["mau"], data["tham_so"], data["so_lan_do"]), ("li-con-lac-don", ["chieu-dai"], 5))
        review = (self.folder / "can-soat.md").read_text(encoding="utf-8")
        self.assertIn("T = 2π√(l/g)", review)
        self.assertNotIn("do AI viết", review)

    def test_plan_only_and_part_selection(self):
        code, data = self.run_tool(self.folder, "--plan-only")
        self.assertEqual((code, data["files"]), (0, []))
        self.assertFalse((self.folder / "thi-nghiem.html").exists())
        code, data = self.run_tool(self.folder, "--phan", "html")
        self.assertEqual([Path(path).name for path in data["files"]], ["thi-nghiem.html", "can-soat.md"])
        code, data = self.run_tool(self.folder, "--phan", "word")
        self.assertEqual((code, data["error"]["step"]), (1, "input"))

    def test_input_errors(self):
        code, data = self.run_tool(self.folder / "khong-co")
        self.assertEqual((code, data["error"]["step"]), (1, "input"))
        (self.folder / "thi-nghiem.md").unlink()
        code, data = self.run_tool(self.folder)
        self.assertEqual((code, data["error"]["step"]), (1, "input"))
        code, data = self.run_tool()
        self.assertEqual((code, data["error"]["step"]), (1, "input"))

    def test_parse_error_names_the_line(self):
        (self.folder / "thi-nghiem.md").write_text(PENDULUM.replace("so-lan-do: 5", "so-lan-do: 99"), encoding="utf-8")
        code, data = self.run_tool(self.folder)
        self.assertEqual((code, data["error"]["step"]), (1, "parse"))
        self.assertIn("Dòng 22", data["error"]["message"])

    def test_unknown_model_is_a_model_error(self):
        (self.folder / "thi-nghiem.md").write_text(PENDULUM.replace("mau: li-con-lac-don", "mau: li-con-lac"), encoding="utf-8")
        code, data = self.run_tool(self.folder)
        self.assertEqual((code, data["error"]["step"]), (1, "model"))

    def test_new_model_is_flagged_for_the_teacher(self):
        self.new_model()
        code, data = self.run_tool(self.folder)
        self.assertEqual(code, 0, data)
        self.assertTrue(any("do AI viết" in warning for warning in data["warnings"]))
        review = (self.folder / "can-soat.md").read_text(encoding="utf-8")
        for needle in ("chưa có người duyệt", "F = k·x", "Bảng số kiểm do AI viết", "máy tính cầm tay"):
            self.assertIn(needle, review)

    def test_new_model_without_conditions_or_with_network_code_is_blocked(self):
        kb = json.loads(json.dumps(HOOKE_JSON))
        kb["congThuc"]["dieuKien"] = ""
        self.new_model(kb=kb)
        code, data = self.run_tool(self.folder)
        self.assertEqual((code, data["error"]["step"]), (1, "model"))
        self.assertIn("dieuKien", data["error"]["message"])
        self.new_model(js=HOOKE_JS + "\nfetch('x');")
        code, data = self.run_tool(self.folder)
        self.assertEqual((code, data["error"]["step"]), (1, "model"))

    @unittest.skipUnless(HAS_NODE, NEED_NODE)
    def test_wrong_physics_is_a_check_error(self):
        self.new_model(js=HOOKE_JS.replace("p['do-cung'] * p['do-gian']", "p['do-cung'] * p['do-gian'] * 2"))
        code, data = self.run_tool(self.folder)
        self.assertEqual((code, data["error"]["step"]), (1, "check"))
        self.assertIn("dòng 1 `luc`", data["error"]["message"])
        self.assertFalse((self.folder / "thi-nghiem.html").exists())

    def test_machine_without_node_still_builds_and_says_so(self):
        with mock.patch.object(kiem_so, "find_node", return_value=None):
            code, data = self.run_tool(self.folder, "--phan", "html")
        self.assertEqual(code, 0, data)
        self.assertEqual(data["kiem_so"], {"chay": False, "dat": 0, "tong": 6})
        self.assertTrue(any("không có Node" in warning for warning in data["warnings"]))
        self.assertIn("chưa chạy trên máy này", (self.folder / "can-soat.md").read_text(encoding="utf-8"))

    def test_missing_python_docx_is_a_docx_error(self):
        error = ImportError("No module named 'docx'", name="docx")
        with mock.patch.object(thi_nghiem, "load_phieu", side_effect=error):
            code, data = self.run_tool(self.folder)
        self.assertEqual((code, data["error"]["step"]), (1, "docx"))
        self.assertIn("requirements-vi.txt", data["error"]["fix"])

    def test_write_and_internal_errors(self):
        with mock.patch.object(build_html, "write", side_effect=PermissionError("đang mở")):
            code, data = self.run_tool(self.folder)
        self.assertEqual((code, data["error"]["step"]), (1, "write"))
        with mock.patch.object(thu_vien, "load", side_effect=RuntimeError("boom")):
            code, data = self.run_tool(self.folder)
        self.assertEqual((code, data["error"]["step"]), (1, "internal"))

if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Chạy để thấy trượt**

Run: `venv\Scripts\python.exe -m unittest tools.vi.tests.test_thi_nghiem_cong_cu -v`
Expected: ERROR `ModuleNotFoundError: No module named 'thi_nghiem'`.

- [ ] **Step 3: Viết công cụ**

`tools/vi/thi_nghiem.py`:

```python
#!/usr/bin/env python3
"""Tạo thí nghiệm ảo (file HTML chạy không cần mạng) và phiếu học tập Word từ thi-nghiem.md.

Cách dùng:
    python tools/vi/thi_nghiem.py <thư_mục_thí_nghiệm> [--phan tat-ca|html|phieu] [--plan-only]

stdout: đúng một dòng JSON. Tiến trình đi ra stderr.
Mã thoát: 0 khi xong, 1 khi lỗi.
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from thi_nghiem_parts import build_html, kiem_so, parse, thu_vien  # noqa: E402

SOURCE_NAME = "thi-nghiem.md"
REVIEW_NAME = "can-soat.md"
PART_CHOICES = ("html", "phieu")
MAX_PATH = 200
GUIDE = "docs/vi/tro-ly/thi-nghiem-ao.md"
MODEL_GUIDE = "docs/vi/tro-ly/mo-hinh-thi-nghiem.md"
FIX_SOURCE = f"Viết file {SOURCE_NAME} trong thư mục thí nghiệm theo {GUIDE} rồi chạy lại."
FIX_PARSE = f"Sửa đúng dòng đó trong {SOURCE_NAME} theo {GUIDE} rồi chạy lại."
FIX_MODEL = f"Sửa mô hình theo khuôn trong {MODEL_GUIDE}; không bỏ công thức, điều kiện áp dụng hay bảng số kiểm."
FIX_CHECK = ("Sửa hàm tinh trong mo-hinh.js cho khớp bảng số kiểm. Chỉ sửa bảng số kiểm khi chính bảng sai, "
             f"và khi đó ghi rõ dòng đã sửa vào mục cần soát để thầy cô kiểm lại ({MODEL_GUIDE}).")
FIX_WRITE = "Đóng file Word hoặc trình duyệt đang mở file cũ rồi chạy lại; kiểm tra ổ đĩa còn trống."
FIX_INTERNAL = "Gửi nguyên dòng error.message cho người bảo trì."
FIX_ARGS = "Chạy: python tools/vi/thi_nghiem.py <thư_mục_thí_nghiệm> [--phan tat-ca|html|phieu] [--plan-only]"
NEW_MODEL_NOTE = (
    "Mô hình này do AI viết, chưa có người duyệt. Bảng số kiểm cũng do AI tự tính, nên nó chỉ bắt được lỗi lập trình, "
    "không bắt được lỗi hiểu sai kiến thức. Thầy cô kiểm lại công thức dưới đây và tính lại ít nhất hai dòng của bảng "
    "số kiểm bằng máy tính cầm tay trước khi dùng trên lớp."
)


def fix_docx() -> str:
    return (
        f'Cài thư viện bằng: "{sys.executable}" -m pip install -r tools/vi/requirements-vi.txt '
        "(hoặc chạy lại CAI-DAT.bat)"
    )


class ArgumentError(Exception):
    """Tham số dòng lệnh sai; báo bằng JSON thay vì để argparse tự thoát."""


class JsonArgumentParser(argparse.ArgumentParser):
    def error(self, message):
        raise ArgumentError(message)


def log(text: str) -> None:
    print(text, file=sys.stderr, flush=True)


def emit(payload: dict) -> None:
    """In đúng một dòng JSON và không bao giờ ném lỗi."""
    text = json.dumps(payload, ensure_ascii=False) + "\n"
    try:
        try:
            sys.stdout.write(text)
        except UnicodeEncodeError:
            buffer = getattr(sys.stdout, "buffer", None)
            if buffer is not None:
                buffer.write(text.encode("utf-8", errors="replace"))
            else:
                sys.stdout.write(json.dumps(payload, ensure_ascii=True) + "\n")
        sys.stdout.flush()
    except OSError:
        pass


def configure_streams() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


def result(*, ready: bool, files=(), mau: str = "", tham_so=(), so_lan_do: int = 0,
           kiem_so_info: dict | None = None, warnings=(), error: dict | None = None) -> dict:
    return {
        "ready": ready,
        "files": [str(path) for path in files],
        "mau": mau,
        "tham_so": list(tham_so),
        "so_lan_do": so_lan_do,
        "kiem_so": kiem_so_info or {"chay": False, "dat": 0, "tong": 0},
        "warnings": list(warnings),
        "error": error,
    }


def failure(step: str, message: str, fix: str, **rest) -> dict:
    return result(ready=False, error={"step": step, "message": message, "fix": fix}, **rest)


def select_parts(value: str) -> list[str]:
    if value.strip() == "tat-ca":
        return list(PART_CHOICES)
    chosen = [item.strip() for item in value.split(",") if item.strip()]
    if not chosen or any(item not in PART_CHOICES for item in chosen):
        raise ValueError("--phan chỉ nhận tat-ca hoặc " + ", ".join(PART_CHOICES) + f"; gặp {value!r}")
    return [part for part in PART_CHOICES if part in chosen]


def load_phieu():
    """Import muộn để thiếu python-docx vẫn báo được lỗi dạng JSON."""
    from thi_nghiem_parts import phieu

    return phieu


def review_text(experiment, model, check: dict | None, warnings: list[str]) -> str:
    formula = model.khai_bao["congThuc"]
    lines = [f"# Cần thầy cô soát — {experiment.meta['tieu-de']}", ""]
    if model.moi:
        lines += [NEW_MODEL_NOTE, ""]
    lines += [
        f"- Mẫu: `{model.ma}` — {model.khai_bao['ten']}",
        f"- Công thức của mô hình: {formula['bieuThuc']}",
        f"- Điều kiện lí tưởng hoá: {formula['dieuKien']}",
    ]
    if formula.get("nguon"):
        lines.append(f"- Nguồn số liệu: {formula['nguon']}")
    if check is None:
        lines.append("- Kiểm số: chưa chạy trên máy này vì không có Node; file HTML tự kiểm mỗi lần mở.")
    else:
        lines.append(f"- Kiểm số: {check['dat']}/{check['tong']} dòng của bảng số kiểm đạt.")
    if model.moi:
        lines += ["", "## Bảng số kiểm do AI viết", ""]
        for index, row in enumerate(model.khai_bao["bangKiem"], 1):
            lines.append(f"{index}. vào {json.dumps(row['vao'], ensure_ascii=False)} → ra "
                         f"{json.dumps(row['ra'], ensure_ascii=False)} (sai số cho phép {row['saiSo']})")
    lines += ["", "## Cảnh báo", ""]
    lines += [f"- {warning}" for warning in warnings] or ["Không có cảnh báo."]
    return "\n".join(lines) + "\n"


def run(args) -> int:
    try:
        parts = select_parts(args.phan)
    except ValueError as exc:
        emit(failure("input", str(exc), "Chạy lại với --phan tat-ca"))
        return 1
    folder = args.folder.expanduser().resolve()
    if not folder.is_dir():
        emit(failure("input", f"Không có thư mục thí nghiệm: {folder}", FIX_SOURCE))
        return 1
    source = folder / SOURCE_NAME
    if not source.is_file():
        emit(failure("input", f"Không có file {SOURCE_NAME} trong {folder}", FIX_SOURCE))
        return 1
    try:
        text = source.read_text(encoding="utf-8-sig")
    except (OSError, UnicodeDecodeError) as exc:
        emit(failure("input", f"Không đọc được {source}: {exc}", f"Lưu lại {SOURCE_NAME} bằng bảng mã UTF-8 rồi chạy lại"))
        return 1

    try:
        meta = parse.read_meta(text)
    except parse.ParseError as exc:
        emit(failure("parse", str(exc), FIX_PARSE))
        return 1
    try:
        model = thu_vien.load(meta["mau"], folder)
    except thu_vien.ModelError as exc:
        emit(failure("model", str(exc), FIX_MODEL, mau=meta["mau"]))
        return 1
    try:
        experiment = parse.parse_experiment(text, model.khai_bao)
    except parse.ParseError as exc:
        emit(failure("parse", str(exc), FIX_PARSE, mau=model.ma))
        return 1

    summary = {"mau": model.ma, "tham_so": experiment.sliders(), "so_lan_do": experiment.so_lan_do}
    warnings = list(experiment.warnings)
    try:
        check = kiem_so.bang_kiem(model)
    except kiem_so.CheckError as exc:
        emit(failure("check", str(exc), FIX_CHECK, **summary, warnings=warnings))
        return 1
    if check is None:
        kiem_so_info = {"chay": False, "dat": 0, "tong": len(model.khai_bao["bangKiem"])}
        warnings.append("Chưa chạy kiểm số trên máy này vì không có Node; file HTML sẽ tự kiểm khi mở.")
    else:
        kiem_so_info = {"chay": True, "dat": check["dat"], "tong": check["tong"]}
        if check["dat"] != check["tong"]:
            detail = "; ".join(f"dòng {m['dong']} `{m['ma']}`: mong {m['mong']}, được {m['duoc']}" for m in check["truot"])
            emit(failure("check", f"Bảng số kiểm trượt {check['tong'] - check['dat']}/{check['tong']} dòng: {detail}",
                         FIX_CHECK, **summary, kiem_so_info=kiem_so_info, warnings=warnings))
            return 1
    if model.moi:
        warnings.append(f"Mô hình do AI viết, chưa có người duyệt: thầy cô soát công thức và bảng số kiểm trong {REVIEW_NAME}.")
    if len(str(folder)) > MAX_PATH:
        warnings.append(f"Đường dẫn thư mục dài {len(str(folder))} ký tự; Windows có thể không ghi được file. "
                        "Chuyển bộ công cụ ra ổ đĩa gần gốc, ví dụ D:\\PPTmaster.")

    if args.plan_only:
        log(f"Chỉ kiểm {SOURCE_NAME}, không ghi file.")
        emit(result(ready=True, **summary, kiem_so_info=kiem_so_info, warnings=warnings))
        return 0

    phieu = None
    if "phieu" in parts:
        try:
            phieu = load_phieu()
        except ImportError as exc:
            if getattr(exc, "name", None) in ("docx", "lxml"):
                emit(failure("docx", f"Chưa cài thư viện python-docx ({exc})", fix_docx(), **summary,
                             kiem_so_info=kiem_so_info, warnings=warnings))
            else:
                emit(failure("internal", f"Lỗi ngoài dự kiến khi nạp bộ dựng Word: {exc}", FIX_INTERNAL, **summary,
                             kiem_so_info=kiem_so_info, warnings=warnings))
            return 1

    files: list[Path] = []
    try:
        if "html" in parts:
            files.append(build_html.write(experiment, model, folder))
        if phieu is not None:
            path, phieu_warnings = phieu.build(experiment, model, folder)
            files.append(path)
            warnings += phieu_warnings
        review = folder / REVIEW_NAME
        review.write_text(review_text(experiment, model, check, warnings), encoding="utf-8", newline="\n")
        files.append(review)
    except build_html.BuildError as exc:
        emit(failure("model", str(exc), FIX_MODEL, **summary, kiem_so_info=kiem_so_info, warnings=warnings))
        return 1
    except OSError as exc:
        emit(failure("write", f"Không ghi được file: {exc}", FIX_WRITE, **summary, kiem_so_info=kiem_so_info,
                     warnings=warnings))
        return 1

    log(f"Đã tạo {len(files)} file.")
    emit(result(ready=True, files=files, **summary, kiem_so_info=kiem_so_info, warnings=warnings))
    return 0


def main(argv: list[str] | None = None) -> int:
    configure_streams()
    parser = JsonArgumentParser(description="Tạo thí nghiệm ảo HTML và phiếu học tập Word")
    parser.add_argument("folder", type=Path, help="Thư mục thí nghiệm, chứa file thi-nghiem.md")
    parser.add_argument("--phan", default="tat-ca", help="tat-ca, hoặc html,phieu")
    parser.add_argument("--plan-only", action="store_true", help="Chỉ kiểm, không ghi file")
    try:
        args = parser.parse_args(argv)
    except ArgumentError as exc:
        emit(failure("input", f"Tham số không hợp lệ: {exc}", FIX_ARGS))
        return 1
    try:
        return run(args)
    except Exception as exc:  # noqa: BLE001 - stdout không bao giờ được để trống
        emit(failure("internal", f"Lỗi ngoài dự kiến: {exc}", FIX_INTERNAL))
        return 1


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: Chạy lại toàn bộ test thí nghiệm ảo**

Run:
```
venv\Scripts\python.exe -m unittest tools.vi.tests.test_thi_nghiem_cong_cu tools.vi.tests.test_thi_nghiem_khung tools.vi.tests.test_thi_nghiem_tham_chieu tools.vi.tests.test_thi_nghiem_mo_hinh tools.vi.tests.test_thi_nghiem_parse tools.vi.tests.test_thi_nghiem_html tools.vi.tests.test_thi_nghiem_phieu
```
Expected: PASS (11 + 1 + 8 + 11 + 9 + 3 + 3 = 46 test), không test nào bị bỏ qua.

- [ ] **Step 5: Chạy thử bằng tay một lần**

```
venv\Scripts\python.exe -c "import pathlib,sys; sys.path.insert(0,'tools/vi/tests'); import thi_nghiem_mau as m; d=pathlib.Path('projects/_thi-nghiem/_thu-con-lac'); d.mkdir(parents=True, exist_ok=True); (d/'thi-nghiem.md').write_text(m.PENDULUM, encoding='utf-8')"
venv\Scripts\python.exe tools\vi\thi_nghiem.py projects\_thi-nghiem\_thu-con-lac
```
Expected: một dòng JSON có `"ready": true`, ba file, `"kiem_so": {"chay": true, "dat": 6, "tong": 6}`. `git status` không được hiện gì trong `projects/`.

- [ ] **Step 6: Commit**

```bash
git add tools/vi/thi_nghiem.py tools/vi/tests/test_thi_nghiem_cong_cu.py
git commit -m "feat(vi): add the thi_nghiem.py command line tool"
```

---

### Task 11: Hướng dẫn cho AI, luật Antigravity và test lớp hướng dẫn

Bài học vi.7: Antigravity chép nguyên `AGENTS.md` vào luật nhưng **không** chép nội dung file nhắc bằng `@`. Vì vậy dòng loại việc thứ 9 và cách làm phải nằm thẳng trong `.agents/rules/ppt-master-vi.md`, và file đó phải dưới 12.000 ký tự. Test sẵn có khoá bảng loại việc trong file luật phải khớp từng chữ với bảng trong `quy-trinh-hoi.md` (chỉ khác cột đường dẫn).

**Files:**
- Create: `docs/vi/tro-ly/thi-nghiem-ao.md`
- Create: `docs/vi/tro-ly/mo-hinh-thi-nghiem.md`
- Modify: `docs/vi/tro-ly/quy-trinh-hoi.md`, `docs/vi/tro-ly/mau-brief.md`, `AGENTS.vi.md`, `.agents/rules/ppt-master-vi.md`
- Modify: `tools/vi/tests/test_vi_layer.py`

**Interfaces:**
- Consumes: `parse.read_meta`, `parse.parse_experiment`, `thu_vien.load`, `thu_vien.check_declaration`, `thu_vien.check_js`, `thu_vien.list_models` (Task 3, 7); lệnh và bảng `error.step` của `thi_nghiem.py` (Task 10).
- Produces: tiêu đề `## 14. Làm thí nghiệm ảo` là mục cuối của `AGENTS.vi.md`; dòng bảng `| Thí nghiệm ảo | "thí nghiệm ảo", "mô phỏng thí nghiệm", "mô phỏng tương tác" | ... |` giống nhau ở `quy-trinh-hoi.md` và file luật.

- [ ] **Step 1: Viết test lớp hướng dẫn**

Trong `tools/vi/tests/test_vi_layer.py`:

(a) Thay `self.assertEqual(len(common), 8)` bằng `self.assertEqual(len(common), 9)`.

(b) Thay cả hàm `test_agents_vi_keeps_the_four_task_sections_last_in_order` bằng:

```python
    def test_agents_vi_keeps_the_five_task_sections_last_in_order(self):
        headings = h2_headings(read("AGENTS.vi.md"))
        self.assertEqual(headings[-5], AGENTS_VI_ASSISTANT_HEADING)
        self.assertEqual(headings[-4], AGENTS_VI_VIDEO_HEADING)
        self.assertEqual(headings[-3], AGENTS_VI_EXAM_HEADING)
        self.assertEqual(headings[-2], AGENTS_VI_LESSON_HEADING)
        self.assertEqual(headings[-1], AGENTS_VI_EXPERIMENT_HEADING)
```

(c) Trong `test_agents_vi_environment_section_points_to_guide`, thay `self.assertEqual(h2_headings(text)[-4], AGENTS_VI_ASSISTANT_HEADING)` bằng `self.assertEqual(h2_headings(text)[-5], AGENTS_VI_ASSISTANT_HEADING)`.

(d) Trong `test_task_type_count_matches_the_table`, thay hai chỗ `"8 loại"` bằng `"9 loại"` và hai chỗ `("5 loại", "6 loại", "7 loại")` bằng `("5 loại", "6 loại", "7 loại", "8 loại")`.

(e) Thêm ngay trước dòng `if __name__ == "__main__":` ở cuối file:

```python
AGENTS_VI_EXPERIMENT_HEADING = "## 14. Làm thí nghiệm ảo"
EXPERIMENT_GUIDE = "docs/vi/tro-ly/thi-nghiem-ao.md"
MODEL_GUIDE = "docs/vi/tro-ly/mo-hinh-thi-nghiem.md"
EXPERIMENT_COMMAND = r"python tools\vi\thi_nghiem.py"
EXPERIMENT_GUIDE_HEADINGS = (
    "## Khi nào dùng",
    "## Câu hỏi bắt buộc",
    "## Câu hỏi tuỳ chọn",
    "## Tạo nhanh",
    "## Cấu trúc thi-nghiem.md",
    "## Đầu ra",
    "## Nối vào bài giảng",
    "## Ghi vào brief",
)


class ExperimentGuideTest(unittest.TestCase):
    def test_guide_has_its_own_sections_in_order(self):
        self.assertEqual(h2_headings(read(EXPERIMENT_GUIDE)), list(EXPERIMENT_GUIDE_HEADINGS))

    def test_guide_questions_are_limited_and_have_suggestions(self):
        items = numbered_items(section(read(EXPERIMENT_GUIDE), "## Câu hỏi bắt buộc"))
        self.assertTrue(1 <= len(items) <= 7, f"{len(items)} câu")
        for item in items:
            self.assertIn("Gợi ý:", item)
        quick = numbered_items(section(read(EXPERIMENT_GUIDE), "## Tạo nhanh"))
        self.assertTrue(2 <= len(quick) <= 3, f"{len(quick)} câu")

    def test_guide_examples_parse_against_the_real_models(self):
        """Bài học từ gói đề thi: ngữ pháp trong hướng dẫn phải khớp parser thật, chứng minh bằng file mẫu."""
        from thi_nghiem_parts import parse, thu_vien

        body = section(read(EXPERIMENT_GUIDE), "## Cấu trúc thi-nghiem.md")
        blocks = re.findall(r"```[a-z]*\n(---\n.*?)```", body, re.S)
        self.assertEqual(len(blocks), 3, "cần một ví dụ cho mỗi môn")
        subjects = set()
        for block in blocks:
            meta = parse.read_meta(block)
            model = thu_vien.load(meta["mau"], REPO_ROOT)
            experiment = parse.parse_experiment(block, model.khai_bao)
            self.assertEqual(experiment.warnings, [], meta["mau"])
            subjects.add(model.khai_bao["mon"])
        self.assertEqual(subjects, {"Toán", "Vật lí", "Hoá học"})

    def test_guide_states_the_grammar_and_the_limits(self):
        body = section(read(EXPERIMENT_GUIDE), "## Cấu trúc thi-nghiem.md")
        for phrase in ("tieu-de", "nguoi-thao-tac", "sai-so", "co-dinh", "mac-dinh", "buoc", "chon", "so-lan-do",
                       "do-thi", "theo", "ln(", "sqrt(", "goi-y-dap-an", "công thức của mẫu không còn đúng",
                       "Không chèn địa chỉ web"):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, body)

    def test_guide_names_outputs_and_forbids_hand_written_pages(self):
        body = section(read(EXPERIMENT_GUIDE), "## Đầu ra")
        for phrase in (EXPERIMENT_COMMAND, "thi-nghiem.html", "phieu-hoc-tap.docx", "can-soat.md", "mo-hinh.json",
                       "mo-hinh.js", "Không viết file HTML bằng tay", "projects\\_thi-nghiem\\"):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, body)

    def test_guide_skips_the_pptx_only_steps(self):
        body = section(read(EXPERIMENT_GUIDE), "## Ghi vào brief")
        for phrase in ("projects/_thi-nghiem/", "brief.md", "import-sources", "dòng chốt"):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, body)

    def test_guides_have_no_markdown_links(self):
        for name in (EXPERIMENT_GUIDE, MODEL_GUIDE):
            self.assertEqual(LINK_RE.findall(read(name)), [], name)

    def test_model_guide_lists_every_code_of_every_library_model(self):
        from thi_nghiem_parts import thu_vien

        text = read(MODEL_GUIDE)
        self.assertEqual(len(thu_vien.list_models()), 8)
        for ma in thu_vien.list_models():
            model = thu_vien.load(ma, REPO_ROOT)
            self.assertIn(f"### `{ma}`", text)
            for item in model.khai_bao["thamSo"] + model.khai_bao["daiLuongDo"]:
                self.assertIn(f"`{item['ma']}`", text, f"{ma}: {item['ma']}")
            self.assertIn(model.khai_bao["congThuc"]["dieuKien"], text, ma)

    def test_model_guide_example_follows_the_contract(self):
        from thi_nghiem_parts import thu_vien

        text = section(read(MODEL_GUIDE), "## Khuôn mô hình mới")
        declaration = json.loads(re.search(r"```json\n(.*?)```", text, re.S).group(1))
        code = re.search(r"```js\n(.*?)```", text, re.S).group(1)
        self.assertEqual(thu_vien.check_declaration(declaration), [])
        self.assertEqual(thu_vien.check_js(code, declaration["hoatHinh"]), [])

    def test_model_guide_keeps_the_teacher_review_rule(self):
        body = section(read(MODEL_GUIDE), "## Khi mô hình do AI viết")
        for phrase in ("can-soat.md", "máy tính cầm tay", "không lược bỏ", "`check`", "`model`",
                       "không bắt được lỗi hiểu sai kiến thức"):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, body)


class ExperimentWiringTest(unittest.TestCase):
    def test_common_rules_table_lists_the_experiment_task(self):
        body = section(read("docs/vi/tro-ly/quy-trinh-hoi.md"), "## Khi nào áp dụng")
        self.assertIn("thi-nghiem-ao.md", body)
        for keyword in ("thí nghiệm ảo", "mô phỏng thí nghiệm"):
            self.assertIn(keyword, body)
        self.assertIn('Loại việc "Thí nghiệm ảo" không tạo PPTX', body)

    def test_image_rule_does_not_apply_to_experiments(self):
        self.assertIn("Thí nghiệm ảo", section(read("docs/vi/tro-ly/quy-trinh-hoi.md"), "## Ảnh minh hoạ"))

    def test_agents_vi_section_explains_the_order_and_the_command(self):
        text = read("AGENTS.vi.md")
        self.assertIn(AGENTS_VI_EXPERIMENT_HEADING, h2_headings(text))
        body = section(text, AGENTS_VI_EXPERIMENT_HEADING)
        self.assertIn(EXPERIMENT_COMMAND, body)
        for phrase in ("(docs/vi/tro-ly/thi-nghiem-ao.md)", "(docs/vi/tro-ly/mo-hinh-thi-nghiem.md)", "projects/_thi-nghiem/",
                       "thi-nghiem.md", "--plan-only", "can-soat.md", "kiem_so", r"venv\Scripts\python.exe"):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, body)

    def test_agents_vi_section_maps_every_error_step(self):
        body = section(read("AGENTS.vi.md"), AGENTS_VI_EXPERIMENT_HEADING)
        for step in ("input", "parse", "model", "check", "docx", "write", "internal"):
            self.assertIn(f"`{step}`", body)
        self.assertIn("requirements-vi.txt", body)

    def test_agents_vi_section_bans_shortcuts(self):
        body = section(read("AGENTS.vi.md"), AGENTS_VI_EXPERIMENT_HEADING)
        for phrase in ("không viết file HTML bằng tay", "không bỏ bảng số kiểm", "không chèn thư viện", "project_manager.py init",
                       "không chạm `skills/`", "không commit gì trong `projects/`"):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, body)

    def test_agents_vi_triggers_and_table_include_the_experiment_task(self):
        text = read("AGENTS.vi.md")
        self.assertIn('"thí nghiệm ảo"', section(text, "## 3. Câu lệnh tiếng Việt kích hoạt skill `ppt-master`"))
        assistant = section(text, AGENTS_VI_ASSISTANT_HEADING)
        self.assertIn("(docs/vi/tro-ly/thi-nghiem-ao.md)", assistant)
        self.assertIn("mục 14", assistant)

    def test_antigravity_rule_carries_the_experiment_rules(self):
        rule = read(".agents/rules/ppt-master-vi.md")
        for phrase in ("## Thí nghiệm ảo", "docs/vi/tro-ly/thi-nghiem-ao.md", r"tools\vi\thi_nghiem.py", "9 loại việc",
                       "Không viết file HTML bằng tay", "can-soat.md"):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, rule)
        self.assertLessEqual(len(rule), ANTIGRAVITY_RULE_LIMIT)

    def test_brief_template_lists_the_experiment_task(self):
        self.assertIn("Thí nghiệm ảo", read("docs/vi/tro-ly/mau-brief.md"))
```

(f) Thêm `import json` vào khối import đầu file (giữa `import fnmatch` và `import re`).

- [ ] **Step 2: Chạy để thấy trượt**

Run: `venv\Scripts\python.exe -m unittest tools.vi.tests.test_vi_layer`
Expected: FAIL/ERROR ở các test vừa thêm và vừa sửa (thiếu file hướng dẫn, thiếu mục 14, còn chữ "8 loại").

- [ ] **Step 3: Viết hai file hướng dẫn cho AI**

`docs/vi/tro-ly/thi-nghiem-ao.md`:

````markdown
# Loại việc: Thí nghiệm ảo

File dành cho AI. Luôn đọc `docs/vi/tro-ly/quy-trinh-hoi.md` trước file này. Đầu ra của loại việc này là một file HTML thí nghiệm ảo chạy không cần mạng và một phiếu học tập Word, không phải PPTX.

## Khi nào dùng

Thầy cô cần một thí nghiệm ảo, mô phỏng tương tác cho Toán, Vật lí hoặc Hoá học để học sinh thay đổi tham số, đo, ghi số liệu và rút ra kết luận.

Ví dụ câu lệnh:
- "Làm thí nghiệm ảo con lắc đơn cho Vật lí 11"
- "Tạo mô phỏng chuẩn độ acid – base cho Hoá 11"
- "Làm thí nghiệm ảo gieo xúc xắc cho bài xác suất lớp 10"

Thầy cô chỉ cần hình minh hoạ một thí nghiệm trên slide thì không dùng loại việc này; đó là việc của bài giảng.

## Câu hỏi bắt buộc

1. Môn, lớp, tên bài, và thí nghiệm nào?
   Gợi ý: mẫu gần nhất trong `docs/vi/tro-ly/mo-hinh-thi-nghiem.md`. Thí nghiệm ngoài danh mục thì nói rõ: em sẽ viết mô hình mới, thầy cô cần soát công thức trước khi dùng.
2. Thí nghiệm dùng ở hoạt động nào: mở đầu, hình thành kiến thức, hay luyện tập?
   Gợi ý: hình thành kiến thức.
3. Ai thao tác: giáo viên trên máy chiếu, hay học sinh theo nhóm trên máy tính hoặc điện thoại?
   Gợi ý: giáo viên trình diễn, các nhóm ghi số liệu vào phiếu.
4. Học sinh cần rút ra kết luận gì từ thí nghiệm?
   Gợi ý: AI đề xuất một kết luận theo yêu cầu cần đạt của bài để thầy cô sửa.
5. Tham số nào được thay đổi, trong khoảng nào, và đo bao nhiêu lần?
   Gợi ý: một tham số chính theo khoảng mặc định của mẫu, 5 lần đo.
6. Có bật sai số đo để học sinh tập xử lí số liệu không?
   Gợi ý: tắt với THCS, bật với THPT.
7. Có cần phiếu học tập Word để in không?
   Gợi ý: có.

## Câu hỏi tuỳ chọn

- Thầy cô có địa chỉ web đã đưa file lên (ví dụ Netlify) để gắn vào slide không? Chỉ hỏi khi thầy cô nhắc tới việc cho học sinh dùng điện thoại.

## Tạo nhanh

1. Môn, lớp, tên bài và thí nghiệm (câu hỏi bắt buộc 1).
2. Ai thao tác (câu hỏi bắt buộc 3).

## Cấu trúc thi-nghiem.md

Viết một file `thi-nghiem.md` trong `projects/_thi-nghiem/<tên_thí_nghiệm>/`. File mở đầu bằng khối thông tin giữa hai dòng `---`, rồi đúng năm mục theo thứ tự: `## Tham số`, `## Dự đoán`, `## Quan sát`, `## Giải thích`, `## Kết luận`.

Khối thông tin:

| Khoá | Bắt buộc | Giá trị |
|---|---|---|
| `tieu-de` | có | tên thí nghiệm hiện trên trang |
| `mon`, `lop` | có | môn và lớp |
| `mau` | có | mã mẫu trong `mo-hinh-thi-nghiem.md`, hoặc `moi` khi AI viết mô hình mới |
| `nguoi-thao-tac` | không | `giao-vien` (mặc định, không khoá) hoặc `nhom` (khoá cho tới khi học sinh chốt dự đoán) |
| `sai-so` | không | `tat` (mặc định) hoặc `bat` |

Mục `## Tham số`, mỗi dòng một tham số của mẫu; tham số không ghi thì giữ cố định ở mặc định của mẫu:

- Thanh trượt: `<mã>: <nhỏ nhất>..<lớn nhất> buoc <bước> mac-dinh <giá trị>`
- Giữ cố định: `<mã>: co-dinh <giá trị>`
- Tham số lựa chọn: `<mã>: chon <mã_1>, <mã_2>` (mục đầu là mặc định) hoặc `<mã>: co-dinh <mã_lựa_chọn>`
- Khoảng không được vượt khoảng của mẫu, vì ngoài khoảng đó công thức của mẫu không còn đúng. Phải có ít nhất một tham số thay đổi được.

Mục `## Dự đoán`: dòng `cau:`; nếu là trắc nghiệm thì thêm `A:` đến `D:` (ít nhất hai) và `dap-an:`. Không có lựa chọn thì là câu mở, học sinh tự viết.

Mục `## Quan sát`:

- `so-lan-do:` số nguyên từ 3 đến 20.
- `do:` các cột của bảng số liệu, là mã tham số hoặc mã đại lượng đo, cách nhau bằng dấu phẩy, tối đa 6 cột, có ít nhất một đại lượng đo.
- `do-thi:` (tuỳ chọn) `<trục tung> theo <trục hoành>`. Mỗi trục là một mã có trong `do:`, tuỳ chọn kèm một phép biến đổi: `<mã>^2`, `1/<mã>`, `ln(<mã>)`, `sqrt(<mã>)`. Chọn phép biến đổi để đồ thị thành đường thẳng thì học sinh đọc được quy luật từ hệ số góc. Đây là cú pháp tính toán, khác quy ước hiển thị `T^2^` dùng trong các dòng chữ.

Mục `## Giải thích`: dòng `cau:` và dòng `goi-y-dap-an:`. Mục `## Kết luận`: một đoạn văn.

Quy ước chữ: `H~2~SO~4~` cho chỉ số dưới, `m/s^2^` cho chỉ số trên, `**in đậm**`. Không chèn địa chỉ web vào file: thí nghiệm chạy không cần mạng.

Ví dụ Vật lí (học sinh thao tác theo nhóm, có sai số đo):

```
---
tieu-de: Chu kì con lắc đơn
mon: Vật lí
lop: 11
mau: li-con-lac-don
nguoi-thao-tac: nhom
sai-so: bat
---

## Tham số
chieu-dai: 0.4..1.6 buoc 0.2 mac-dinh 1.0
g: co-dinh 9.8

## Dự đoán
cau: Khi tăng chiều dài dây gấp 4 lần, chu kì thay đổi thế nào?
A: Tăng 4 lần
B: Tăng 2 lần
C: Không đổi
dap-an: B

## Quan sát
so-lan-do: 5
do: chieu-dai, chu-ki
do-thi: chu-ki^2 theo chieu-dai

## Giải thích
cau: Từ đồ thị, T^2^ và l liên hệ thế nào? Suy ra công thức tính chu kì.
goi-y-dap-an: T^2^ tỉ lệ thuận với l, hệ số góc bằng 4π^2^/g; suy ra T = 2π√(l/g).

## Kết luận
Chu kì con lắc đơn chỉ phụ thuộc chiều dài dây và gia tốc trọng trường, không phụ thuộc khối lượng quả nặng.
```

Ví dụ Hoá học (giáo viên trình diễn, dự đoán câu mở):

```
---
tieu-de: Chuyển dịch cân bằng N~2~O~4~ ⇌ 2NO~2~
mon: Hoá học
lop: 11
mau: hoa-can-bang-no2
---

## Tham số
nhiet-do: 0..100 buoc 10 mac-dinh 20
ap-suat: co-dinh 1

## Dự đoán
cau: Ngâm bình khí vào nước nóng thì màu nâu đỏ của bình thay đổi thế nào? Vì sao?

## Quan sát
so-lan-do: 6
do: nhiet-do, phan-mol-no2, nong-do-no2

## Giải thích
cau: Phản ứng thuận thu nhiệt hay toả nhiệt? Dùng nguyên lí Le Chatelier để giải thích số liệu.
goi-y-dap-an: Tăng nhiệt độ thì phần mol NO~2~ tăng, màu đậm hơn; cân bằng chuyển dịch theo chiều thu nhiệt, vậy phản ứng thuận thu nhiệt.

## Kết luận
Khi tăng nhiệt độ, cân bằng chuyển dịch theo chiều phản ứng thu nhiệt; khi giảm nhiệt độ, cân bằng chuyển dịch theo chiều toả nhiệt.
```

Ví dụ Toán (tham số lựa chọn giữ cố định):

```
---
tieu-de: Tần suất và xác suất khi gieo hai xúc xắc
mon: Toán
lop: 10
mau: toan-xac-suat
nguoi-thao-tac: nhom
---

## Tham số
phep-thu: co-dinh hai-xuc-xac-tong-7
so-lan: 100..10000 buoc 100 mac-dinh 100
hat-giong: 1..999 buoc 1 mac-dinh 1

## Dự đoán
cau: Gieo hai xúc xắc càng nhiều lần thì tần suất xuất hiện tổng bằng 7 thay đổi thế nào?
A: Tăng dần tới 1
B: Dao động rồi ổn định quanh một số
C: Giảm dần về 0
dap-an: B

## Quan sát
so-lan-do: 6
do: so-lan, tan-so, tan-suat

## Giải thích
cau: Tần suất ổn định quanh giá trị nào? So sánh với xác suất tính bằng cách đếm số kết quả thuận lợi.
goi-y-dap-an: Có 6 trên 36 kết quả cho tổng bằng 7, nên xác suất là 1/6 ≈ 0,1667; tần suất tiến gần giá trị này khi số lần gieo lớn.

## Kết luận
Khi số lần thử đủ lớn, tần suất của biến cố xấp xỉ xác suất của biến cố đó.
```

## Đầu ra

Chạy `python tools\vi\thi_nghiem.py projects\_thi-nghiem\<tên_thí_nghiệm>`. Công cụ ghi vào chính thư mục đó:

- `thi-nghiem.html`: mở bằng trình duyệt là chạy, không cần mạng. Có ba bước Dự đoán – Quan sát – Giải thích, bảng số liệu, đồ thị, nút chép số liệu sang Excel, và tự kiểm mô hình mỗi lần mở.
- `phieu-hoc-tap.docx`: trang học sinh và trang đáp án dành cho giáo viên.
- `can-soat.md`: công thức, điều kiện lí tưởng hoá, kết quả kiểm số, và các điểm thầy cô cần soát. Đọc nguyên văn file này cho thầy cô.

Thí nghiệm ngoài danh mục: đặt `mau: moi`, rồi viết `mo-hinh.json` và `mo-hinh.js` trong cùng thư mục theo `docs/vi/tro-ly/mo-hinh-thi-nghiem.md`. Không viết file HTML bằng tay, không chèn thư viện từ Internet, không bỏ công thức, điều kiện áp dụng hay bảng số kiểm.

## Nối vào bài giảng

Thầy cô làm slide cho cùng bài thì thêm một trang "Thí nghiệm ảo": sơ đồ thí nghiệm vẽ trên slide, một câu nhiệm vụ, và liên kết mở `thi-nghiem.html`. Chép file HTML vào cạnh file PPTX và dùng đường dẫn tương đối; có địa chỉ web thầy cô cung cấp thì dùng địa chỉ đó và có thể thêm mã QR. Không làm mã QR tới file trong máy.

## Ghi vào brief

- Loại việc: Thí nghiệm ảo.
- Thầy cô yêu cầu: môn, lớp, bài, thí nghiệm, hoạt động, người thao tác, kết luận cần rút ra, tham số và số lần đo, sai số đo, phiếu học tập — ghi đúng lời thầy cô.
- AI đề xuất (thầy cô đã đồng ý): mẫu chọn dùng, khoảng tham số, câu dự đoán và các gợi ý thầy cô chấp nhận.
- Loại việc này không tạo PPTX: ghi brief vào `projects/_thi-nghiem/<tên_thí_nghiệm>/brief.md`; không có dòng chốt cách xác nhận, không chạy `import-sources`, không có bước xác nhận của upstream.
- Viết theo mẫu `docs/vi/tro-ly/mau-brief.md`.
````

`docs/vi/tro-ly/mo-hinh-thi-nghiem.md`:

````markdown
# Mô hình thí nghiệm ảo: danh mục mẫu và khuôn viết mô hình mới

File dành cho AI. Đọc cùng `docs/vi/tro-ly/thi-nghiem-ao.md`. Phần đầu là danh mục 8 mẫu đã kiểm trong thư viện; phần sau là khuôn bắt buộc khi thầy cô cần một thí nghiệm ngoài danh mục.

## Danh mục mẫu

Dùng đúng các mã dưới đây trong `thi-nghiem.md`. Khoảng ở cột cuối là giới hạn áp dụng của công thức; `thi-nghiem.md` chỉ được chọn khoảng nằm trong đó.

### `hoa-can-bang-no2` — Cân bằng N₂O₄ ⇌ 2NO₂ (Hoá học)

| Tham số | Mã | Khoảng, bước, mặc định |
|---|---|---|
| Nhiệt độ | `nhiet-do` | 0 đến 100 °C, bước 5, mặc định 25 |
| Áp suất chung | `ap-suat` | 0,5 đến 5 bar, bước 0,1, mặc định 1 |

| Đại lượng đo | Mã | Sai số đo |
|---|---|---|
| Phần mol NO~2~ | `phan-mol-no2` | 0,005 |
| Độ phân li của N~2~O~4~ | `do-phan-li` | 0,005 |
| Nồng độ NO~2~ (độ đậm màu nâu) (mol/L) | `nong-do-no2` | 0,0005 |
| Hằng số cân bằng K~p~ | `kp` | 0 |

- Mô hình: N~2~O~4~(g) ⇌ 2NO~2~(g), Δ~r~H° = +57,2 kJ; K~p~ = exp(−(Δ~r~H° − TΔ~r~S°)/RT); K~p~ = x^2^P/(1 − x) với x là phần mol NO~2~.
- Điều kiện lí tưởng hoá: Hỗn hợp khí lí tưởng đã đạt cân bằng; coi Δ~r~H° và Δ~r~S° không đổi trong khoảng 0–100 °C; áp suất tính bằng bar.
- Nguồn số liệu: Δ~r~H° = 57,20 kJ/mol và Δ~r~S° = 175,83 J/(mol·K), tính từ Δ~f~H° và S° ở 298 K (Atkins, Physical Chemistry, bảng dữ liệu nhiệt động).

### `hoa-chuan-do` — Chuẩn độ acid – base bằng dung dịch NaOH (Hoá học)

| Tham số | Mã | Khoảng, bước, mặc định |
|---|---|---|
| Acid cần chuẩn độ | `loai-acid` | lựa chọn: `hcl` (HCl (acid mạnh)), `ch3cooh` (CH~3~COOH (acid yếu)); mặc định `hcl` |
| Nồng độ acid | `nong-do-acid` | 0,01 đến 0,5 mol/L, bước 0,01, mặc định 0,1 |
| Thể tích acid | `the-tich-acid` | 10 đến 50 mL, bước 5, mặc định 20 |
| Nồng độ NaOH | `nong-do-base` | 0,01 đến 0,5 mol/L, bước 0,01, mặc định 0,1 |
| Thể tích NaOH đã nhỏ | `the-tich-base` | 0 đến 50 mL, bước 0,1, mặc định 0 |
| Chất chỉ thị | `chi-thi` | lựa chọn: `phenolphtalein` (Phenolphtalein), `metyl-da-cam` (Methyl da cam), `bromothymol` (Bromothymol xanh); mặc định `phenolphtalein` |

| Đại lượng đo | Mã | Sai số đo |
|---|---|---|
| pH của dung dịch | `ph` | 0,05 |

- Mô hình: Bảo toàn điện tích: [H^+^] + [Na^+^] = [OH^−^] + [A^−^]; K~w~ = [H^+^][OH^−^]; với acid yếu [A^−^] = C·K~a~/(K~a~ + [H^+^]).
- Điều kiện lí tưởng hoá: Dung dịch loãng ở 25 °C (K~w~ = 1,0·10^−14^); coi hoạt độ bằng nồng độ; thể tích cộng tính.
- Nguồn số liệu: K~a~(CH~3~COOH) = 1,75·10^−5^ ở 25 °C (CRC Handbook of Chemistry and Physics).

### `hoa-toc-do` — Các yếu tố ảnh hưởng đến tốc độ phản ứng (Hoá học)

| Tham số | Mã | Khoảng, bước, mặc định |
|---|---|---|
| Nồng độ chất tham gia | `nong-do` | 0,02 đến 0,5 mol/L, bước 0,02, mặc định 0,1 |
| Nhiệt độ | `nhiet-do` | 10 đến 60 °C, bước 5, mặc định 25 |
| Chất xúc tác | `xuc-tac` | lựa chọn: `khong` (Không có), `co` (Có xúc tác); mặc định `khong` |

| Đại lượng đo | Mã | Sai số đo |
|---|---|---|
| Thời gian đến khi vẩn đục che dấu X (s) | `thoi-gian` | 0,5 |

- Mô hình: v = k·C; k = A·exp(−E~a~/RT); thời gian t = ΔC/v với ΔC = 0,005 mol/L.
- Điều kiện lí tưởng hoá: Phản ứng giả định bậc 1 theo chất tham gia, dùng tốc độ đầu; E~a~ = 50 kJ/mol khi không có xúc tác và 45 kJ/mol khi có; k = 1,25·10^−3^ s^−1^ ở 25 °C. Số liệu minh hoạ quy luật, không phải của một phản ứng cụ thể.

### `li-con-lac-don` — Con lắc đơn (Vật lí)

| Tham số | Mã | Khoảng, bước, mặc định |
|---|---|---|
| Chiều dài dây l | `chieu-dai` | 0,2 đến 2 m, bước 0,05, mặc định 1 |
| Gia tốc trọng trường g | `g` | 1,6 đến 24,8 m/s^2^, bước 0,1, mặc định 9,8 |
| Góc lệch ban đầu | `goc-lech` | 2 đến 15 °, bước 1, mặc định 8 |
| Khối lượng quả nặng m | `khoi-luong` | 0,05 đến 1 kg, bước 0,05, mặc định 0,2 |

| Đại lượng đo | Mã | Sai số đo |
|---|---|---|
| Chu kì T (s) | `chu-ki` | 0,02 |
| Thời gian 10 dao động (s) | `thoi-gian-10-dao-dong` | 0,1 |

- Mô hình: T = 2π√(l/g)
- Điều kiện lí tưởng hoá: Góc lệch nhỏ (không quá 15°); dây không giãn, khối lượng dây không đáng kể; bỏ qua sức cản.

### `li-mach-ohm` — Đoạn mạch nối tiếp và song song (Vật lí)

| Tham số | Mã | Khoảng, bước, mặc định |
|---|---|---|
| Hiệu điện thế nguồn U | `suat-dien-dong` | 1,5 đến 24 V, bước 0,5, mặc định 12 |
| Điện trở R~1~ | `dien-tro-1` | 1 đến 100 Ω, bước 1, mặc định 10 |
| Điện trở R~2~ | `dien-tro-2` | 1 đến 100 Ω, bước 1, mặc định 20 |
| Kiểu mắc | `kieu-mac` | lựa chọn: `noi-tiep` (Nối tiếp), `song-song` (Song song); mặc định `noi-tiep` |

| Đại lượng đo | Mã | Sai số đo |
|---|---|---|
| Cường độ mạch chính I (A) | `cuong-do-mach-chinh` | 0,01 |
| Cường độ qua R~1~ (A) | `cuong-do-1` | 0,01 |
| Cường độ qua R~2~ (A) | `cuong-do-2` | 0,01 |
| Hiệu điện thế hai đầu R~1~ (V) | `hieu-dien-the-1` | 0,05 |
| Hiệu điện thế hai đầu R~2~ (V) | `hieu-dien-the-2` | 0,05 |
| Điện trở tương đương (Ω) | `dien-tro-tuong-duong` | 0 |

- Mô hình: Nối tiếp: R = R~1~ + R~2~, I = U/R. Song song: 1/R = 1/R~1~ + 1/R~2~, I = I~1~ + I~2~.
- Điều kiện lí tưởng hoá: Nguồn có điện trở trong bằng 0; dây nối và ampe kế có điện trở không đáng kể; vôn kế có điện trở rất lớn.

### `li-nem-xien` — Chuyển động ném xiên (Vật lí)

| Tham số | Mã | Khoảng, bước, mặc định |
|---|---|---|
| Vận tốc đầu v~0~ | `van-toc-dau` | 5 đến 50 m/s, bước 1, mặc định 20 |
| Góc ném α | `goc` | 0 đến 85 °, bước 1, mặc định 45 |
| Độ cao ban đầu h | `do-cao-dau` | 0 đến 50 m, bước 1, mặc định 0 |
| Gia tốc trọng trường g | `g` | 1,6 đến 24,8 m/s^2^, bước 0,1, mặc định 9,8 |

| Đại lượng đo | Mã | Sai số đo |
|---|---|---|
| Tầm xa L (m) | `tam-xa` | 0,2 |
| Độ cao cực đại H (m) | `do-cao-cuc-dai` | 0,1 |
| Thời gian bay t (s) | `thoi-gian-bay` | 0,02 |

- Mô hình: t = (v~0~sinα + √(v~0~^2^sin^2^α + 2gh))/g; L = v~0~cosα·t; H = h + v~0~^2^sin^2^α/(2g)
- Điều kiện lí tưởng hoá: Bỏ qua sức cản không khí; g không đổi; vật coi là chất điểm.

### `toan-ham-so` — Khảo sát hàm số y = ax³ + bx² + cx + d (Toán)

| Tham số | Mã | Khoảng, bước, mặc định |
|---|---|---|
| Hệ số a | `a` | -5 đến 5, bước 0,5, mặc định 1 |
| Hệ số b | `b` | -5 đến 5, bước 0,5, mặc định 0 |
| Hệ số c | `c` | -5 đến 5, bước 0,5, mặc định -3 |
| Hệ số d | `d` | -5 đến 5, bước 0,5, mặc định 0 |
| Hoành độ tiếp điểm x~0~ | `x0` | -5 đến 5, bước 0,1, mặc định 1 |

| Đại lượng đo | Mã | Sai số đo |
|---|---|---|
| Giá trị y(x~0~) | `gia-tri` | 0 |
| Hệ số góc tiếp tuyến y'(x~0~) | `he-so-goc` | 0 |
| Hoành độ điểm cực đại | `hoanh-do-cuc-dai` | 0 |
| Hoành độ điểm cực tiểu | `hoanh-do-cuc-tieu` | 0 |

- Mô hình: y' = 3ax^2^ + 2bx + c; cực trị tại nghiệm của y' = 0 nơi y' đổi dấu; a = 0 thì là hàm bậc hai, đỉnh tại x = −c/(2b).
- Điều kiện lí tưởng hoá: Hệ số thực; a = 0 và b = 0 thì hàm bậc nhất hoặc hằng, không có cực trị.

### `toan-xac-suat` — Xác suất thực nghiệm (Toán)

| Tham số | Mã | Khoảng, bước, mặc định |
|---|---|---|
| Phép thử và biến cố | `phep-thu` | lựa chọn: `dong-xu-ngua` (Tung đồng xu: mặt ngửa), `xuc-xac-mat-6` (Gieo xúc xắc: mặt 6 chấm), `xuc-xac-chan` (Gieo xúc xắc: số chấm chẵn), `hai-xuc-xac-tong-7` (Gieo hai xúc xắc: tổng bằng 7), `hai-xuc-xac-tong-12` (Gieo hai xúc xắc: tổng bằng 12); mặc định `dong-xu-ngua` |
| Số lần thử n | `so-lan` | 10 đến 10000 lần, bước 10, mặc định 100 |
| Lượt gieo số | `hat-giong` | 1 đến 999, bước 1, mặc định 1 |

| Đại lượng đo | Mã | Sai số đo |
|---|---|---|
| Số lần biến cố xảy ra (lần) | `tan-so` | 0 |
| Tần suất | `tan-suat` | 0 |
| Xác suất lí thuyết | `xac-suat-li-thuyet` | 0 |

- Mô hình: Tần suất f = (số lần biến cố xảy ra)/n; khi n lớn, f tiến gần xác suất P.
- Điều kiện lí tưởng hoá: Đồng xu và xúc xắc cân đối, các lần thử độc lập; số ngẫu nhiên sinh bằng thuật toán có hạt giống nên cùng một lượt gieo cho cùng kết quả.

## Khuôn mô hình mới

Chỉ viết mô hình mới khi không mẫu nào trong danh mục dùng được. Đặt `mau: moi` trong `thi-nghiem.md` và viết hai file trong cùng thư mục thí nghiệm.

`mo-hinh.json` khai bốn thứ, cùng `ten`, `mon` và `hoatHinh`:

| Khoá | Nội dung |
|---|---|
| `hoatHinh` | `khong` (hình tĩnh, đổi theo tham số), `mot-lan` (chạy một lần rồi mới có số đo) hoặc `lap` (chạy lặp lại) |
| `thamSo` | mỗi tham số có `ma`, `ten`, `kieu`. `kieu: so` cần `donVi`, `min`, `max`, `buoc`, `macDinh`. `kieu: chon` cần `luaChon` (ít nhất hai mục có `ma`, `ten`) và `macDinh`. `min` và `max` là giới hạn áp dụng của công thức. |
| `daiLuongDo` | mỗi đại lượng có `ma`, `ten`, `donVi`, `saiSo` (độ lệch chuẩn của nhiễu khi bật sai số đo; 0 nếu không có nhiễu), `chuSo` (số chữ số thập phân hiển thị) |
| `congThuc` | `bieuThuc` (công thức bằng chữ), `dieuKien` (điều kiện áp dụng, bắt buộc), `nguon` (nguồn của hằng số; để chuỗi rỗng nếu không dùng hằng số tra cứu) |
| `bangKiem` | ít nhất 5 dòng `{"vao": ..., "ra": ..., "saiSo": ...}`. `vao` chỉ ghi tham số khác mặc định; `ra` là giá trị đúng của đại lượng đo, tính tay từ công thức; `saiSo` là sai số tuyệt đối cho phép. Chọn các dòng ở biên và ở giữa khoảng, không chỉ ở mặc định. |

`mo-hinh.js` gán `THI_NGHIEM_MO_HINH` với hai hàm, thêm `thoiLuong` khi `hoatHinh` là `mot-lan`:

- `tinh(p)`: nhận đối tượng tham số theo mã, trả về đối tượng có mọi mã trong `daiLuongDo`. Hàm thuần: không phụ thuộc thời gian, không dùng số ngẫu nhiên ngoài `THI_NGHIEM_KHUNG.taoNgauNhien(hạt_giống)`, trả `null` cho đại lượng không xác định.
- `ve(ctx, p, t, kt, d)`: vẽ lên `canvas` 2D; `t` là thời gian tính bằng giây, `kt.rong` và `kt.cao` là kích thước khung, `d` là kết quả của `tinh(p)`. Dùng `THI_NGHIEM_KHUNG.MAU` cho màu, `THI_NGHIEM_KHUNG.PHONG` cho font, `THI_NGHIEM_KHUNG.dinhDang(số, chữ_số)` để in số.
- `thoiLuong(p, d)`: số giây của một lần chạy.
- Không dùng `document`, `window`, `fetch`, `import`, `require`, `eval`, địa chỉ web hay thư viện ngoài: công cụ chặn các file như vậy.

Ví dụ đủ khuôn, định luật Hooke. File `mo-hinh.json`:

```json
{
  "ma": "moi",
  "ten": "Định luật Hooke",
  "mon": "Vật lí",
  "hoatHinh": "khong",
  "thamSo": [
    {
      "ma": "do-cung",
      "ten": "Độ cứng k",
      "kieu": "so",
      "donVi": "N/m",
      "min": 10,
      "max": 100,
      "buoc": 10,
      "macDinh": 50
    },
    {
      "ma": "do-gian",
      "ten": "Độ giãn x",
      "kieu": "so",
      "donVi": "m",
      "min": 0,
      "max": 0.2,
      "buoc": 0.01,
      "macDinh": 0.1
    }
  ],
  "daiLuongDo": [
    {
      "ma": "luc",
      "ten": "Lực đàn hồi F",
      "donVi": "N",
      "saiSo": 0.05,
      "chuSo": 2
    }
  ],
  "congThuc": {
    "bieuThuc": "F = k·x",
    "dieuKien": "Lò xo còn trong giới hạn đàn hồi.",
    "nguon": ""
  },
  "bangKiem": [
    {
      "vao": {},
      "ra": {
        "luc": 5.0
      },
      "saiSo": 0.001
    },
    {
      "vao": {
        "do-gian": 0
      },
      "ra": {
        "luc": 0.0
      },
      "saiSo": 0.001
    },
    {
      "vao": {
        "do-gian": 0.2
      },
      "ra": {
        "luc": 10.0
      },
      "saiSo": 0.001
    },
    {
      "vao": {
        "do-cung": 100
      },
      "ra": {
        "luc": 10.0
      },
      "saiSo": 0.001
    },
    {
      "vao": {
        "do-cung": 10,
        "do-gian": 0.05
      },
      "ra": {
        "luc": 0.5
      },
      "saiSo": 0.001
    }
  ]
}
```

File `mo-hinh.js`:

```js
(function (root) {
  function tinh(p) { return { 'luc': p['do-cung'] * p['do-gian'] }; }
  function ve(ctx, p, t, kt, d) { ctx.fillRect(20, 20, 40 + 600 * p['do-gian'], 20); }
  root.THI_NGHIEM_MO_HINH = { tinh: tinh, ve: ve };
})(typeof globalThis !== 'undefined' ? globalThis : this);
```

## Khi mô hình do AI viết

- Công cụ chặn mô hình thiếu `congThuc.dieuKien`, thiếu `bangKiem` đủ 5 dòng, hoặc có lệnh mạng (`error.step` là `model`).
- Máy có Node thì công cụ chạy `bangKiem`; trượt là `error.step` `check`. Khi đó sửa hàm `tinh` cho khớp bảng. Chỉ sửa bảng số kiểm khi chính bảng sai, và ghi rõ dòng đã sửa cùng lý do vào tin nhắn báo thầy cô.
- Bảng số kiểm do AI tự tính chỉ bắt được lỗi lập trình, không bắt được lỗi hiểu sai kiến thức. Vì vậy `can-soat.md` luôn nhắc thầy cô soát công thức và tính lại ít nhất hai dòng bằng máy tính cầm tay; AI đọc nguyên văn lời nhắc đó cho thầy cô, không lược bỏ.
- Không viết mô hình cho hiện tượng mà AI không nêu được công thức và điều kiện áp dụng; khi đó nói rõ với thầy cô và đề xuất mẫu gần nhất trong danh mục.
````

- [ ] **Step 4: Sửa `quy-trinh-hoi.md` và `mau-brief.md`**

Trong `docs/vi/tro-ly/quy-trinh-hoi.md`:

1. Thay `thuộc một trong 8 loại việc dưới đây` bằng `thuộc một trong 9 loại việc dưới đây`.
2. Thêm ngay sau dòng bảng của "Soạn giáo án tích hợp năng lực số và AI":

```
| Thí nghiệm ảo | "thí nghiệm ảo", "mô phỏng thí nghiệm", "mô phỏng tương tác" | [thi-nghiem-ao.md](thi-nghiem-ao.md) |
```

3. Thêm ngay sau gạch đầu dòng bắt đầu bằng `- Loại việc "Soạn giáo án tích hợp năng lực số và AI" không tạo PPTX`:

```
- Loại việc "Thí nghiệm ảo" không tạo PPTX; nó ghi brief như các loại khác nhưng không đi vào quy trình của upstream. Các bước chỉ dành cho PPTX (dòng chốt cách xác nhận cuối tin nhắn hỏi, `import-sources`, bước xác nhận của upstream, `quick-generate.md`) và mục "Ảnh minh hoạ" không áp dụng; làm theo mục "Ghi vào brief" của docs/vi/tro-ly/thi-nghiem-ao.md.
```

4. Trong mục `## Ảnh minh hoạ`, thay `Soạn đề KHTN tiếng Anh, Soạn giáo án và Video bài giảng.` bằng `Soạn đề KHTN tiếng Anh, Soạn giáo án, Thí nghiệm ảo và Video bài giảng.`

Trong `docs/vi/tro-ly/mau-brief.md`, thay `| Soạn giáo án tích hợp năng lực số và AI>` bằng `| Soạn giáo án tích hợp năng lực số và AI | Thí nghiệm ảo>`.

- [ ] **Step 5: Sửa `AGENTS.vi.md`**

1. Mục 3: thay `"kế hoạch bài dạy", "KHBD".` bằng `"kế hoạch bài dạy", "KHBD", "thí nghiệm ảo", "mô phỏng thí nghiệm".`
2. Mục 10: thay `thuộc một trong 8 loại việc dưới đây` bằng `thuộc một trong 9 loại việc dưới đây`; thêm dòng bảng sau dòng "Soạn giáo án tích hợp năng lực số và AI":

```
| Thí nghiệm ảo | [docs/vi/tro-ly/thi-nghiem-ao.md](docs/vi/tro-ly/thi-nghiem-ao.md) |
```

3. Mục 10: thay đoạn `hoặc thuộc loại việc "Soạn đề KHTN tiếng Anh" (mục 12) hoặc "Soạn giáo án tích hợp năng lực số và năng lực AI" (mục 13) — hai loại việc đó không có bước xác nhận của upstream, xem mục 12 và mục 13.` bằng `hoặc thuộc loại việc "Soạn đề KHTN tiếng Anh" (mục 12), "Soạn giáo án tích hợp năng lực số và năng lực AI" (mục 13) hoặc "Thí nghiệm ảo" (mục 14) — ba loại việc đó không có bước xác nhận của upstream, xem mục 12, mục 13 và mục 14.`
4. Mục 10: thay `Yêu cầu không thuộc 8 loại` bằng `Yêu cầu không thuộc 9 loại`.
5. Thêm vào cuối file (sau dòng "Điều cấm" của mục 13, cách một dòng trống):

```markdown
## 14. Làm thí nghiệm ảo

Khi người dùng cần một thí nghiệm ảo hoặc mô phỏng tương tác cho Toán, Vật lí, Hoá học, đọc [docs/vi/tro-ly/thi-nghiem-ao.md](docs/vi/tro-ly/thi-nghiem-ao.md) và [docs/vi/tro-ly/mo-hinh-thi-nghiem.md](docs/vi/tro-ly/mo-hinh-thi-nghiem.md) rồi làm đúng thứ tự dưới. Như mục 4: có `venv\Scripts\python.exe` ở thư mục gốc repo thì dùng nó cho mọi lệnh Python dưới đây, không có thì dùng `python`.

1. Hỏi một lượt theo file hướng dẫn, chờ trả lời. Chọn mẫu gần nhất trong danh mục; chỉ viết mô hình mới khi không mẫu nào dùng được, và nói trước với thầy cô rằng mô hình mới cần thầy cô soát công thức.
2. Tạo `projects/_thi-nghiem/<tên_thí_nghiệm>/` và viết `thi-nghiem.md` theo đúng ngữ pháp trong file hướng dẫn. Mô hình mới thì đặt `mau: moi` và viết thêm `mo-hinh.json`, `mo-hinh.js` theo khuôn.
3. Chạy `python tools\vi\thi_nghiem.py projects\_thi-nghiem\<tên_thí_nghiệm>`; thêm `--plan-only` khi chỉ muốn kiểm; thêm `--phan html` khi thầy cô không cần phiếu học tập.
4. Đọc dòng JSON ở stdout. `ready` là `true` thì báo thầy cô đường dẫn các file, mẫu đã dùng, tham số thay đổi được, số lần đo, kết quả `kiem_so` (đạt bao nhiêu trên bao nhiêu dòng, hoặc chưa chạy vì máy không có Node), đọc nguyên văn `warnings` và nội dung `can-soat.md`. Dặn thầy cô: mở `thi-nghiem.html` bằng trình duyệt là chạy, không cần mạng; muốn học sinh dùng điện thoại thì đưa file lên web.
5. Thầy cô làm slide cho cùng bài: thêm trang nối sang thí nghiệm theo mục "Nối vào bài giảng" của file hướng dẫn.

| `error.step` | Xử lý |
|---|---|
| `input` | Chưa có thư mục hoặc `thi-nghiem.md`, viết file rồi chạy lại. |
| `parse` | Sửa đúng dòng `error.message` nêu rồi chạy lại. Khoảng tham số vượt khoảng của mẫu thì thu hẹp khoảng, không đổi mẫu. |
| `model` | Mã mẫu không có, hoặc mô hình mới sai khuôn: sửa theo `error.message` và khuôn trong file hướng dẫn mô hình. |
| `check` | Bảng số kiểm trượt: sửa hàm `tinh` của mô hình mới cho khớp bảng. Chỉ sửa bảng khi chính bảng sai, và khi đó nói rõ với thầy cô dòng nào đã sửa. |
| `docx` | Có `venv\Scripts\python.exe` ở thư mục gốc repo thì chạy `venv\Scripts\python.exe -m pip install -r tools/vi/requirements-vi.txt`; không thì chạy `python -m pip install -r tools/vi/requirements-vi.txt`, rồi chạy lại, tối đa một lần. |
| `write` | Xin thầy cô đóng file Word hoặc tab trình duyệt đang mở file cũ rồi chạy lại. |
| `internal` | Lỗi ngoài dự kiến; dán nguyên `error.message` để báo cho người bảo trì, không tự đoán cách sửa. |

Điều cấm: không viết file HTML bằng tay; không bỏ bảng số kiểm, công thức hay điều kiện áp dụng để qua được công cụ; không chèn thư viện, font hay địa chỉ web từ Internet; không sửa file trong `tools/vi/thi_nghiem_parts/`; không chạy `project_manager.py init`; không tạo SVG; không chạm `skills/`; không commit gì trong `projects/`.
```

- [ ] **Step 6: Sửa file luật Antigravity**

Trong `.agents/rules/ppt-master-vi.md`:

1. Thay `yêu cầu thuộc một trong 8 loại việc ở bảng dưới` bằng `yêu cầu thuộc một trong 9 loại việc ở bảng dưới`.
2. Thêm dòng bảng sau dòng "Soạn giáo án tích hợp năng lực số và AI":

```
| Thí nghiệm ảo | "thí nghiệm ảo", "mô phỏng thí nghiệm", "mô phỏng tương tác" | `docs/vi/tro-ly/thi-nghiem-ao.md` |
```

3. Thêm ngay trước tiêu đề `## Các việc khác`:

```markdown
## Thí nghiệm ảo

Đầu ra là file HTML chạy không cần mạng và phiếu học tập Word, không phải PPTX.

- Đọc `docs/vi/tro-ly/thi-nghiem-ao.md` và `docs/vi/tro-ly/mo-hinh-thi-nghiem.md`, hỏi một lượt, viết `thi-nghiem.md` trong `projects/_thi-nghiem/<tên>/`, rồi chạy `tools\vi\thi_nghiem.py <thư_mục>`.
- Không viết file HTML bằng tay, không chèn thư viện hay địa chỉ web. Thí nghiệm ngoài danh mục thì viết mô hình mới theo khuôn, có công thức, điều kiện áp dụng và bảng số kiểm.
- Đọc nguyên văn `can-soat.md` cho thầy cô; mô hình do AI viết thì thầy cô phải soát công thức trước khi dùng.

```

- [ ] **Step 7: Chạy lại**

Run: `venv\Scripts\python.exe -m unittest tools.vi.tests.test_vi_layer`
Expected: PASS (157 test cũ + 18 test mới = 175 test). Nếu `test_antigravity_rule_task_table_matches_common_rules` trượt, so từng chữ hai dòng bảng "Thí nghiệm ảo".

- [ ] **Step 8: Commit**

```bash
git add docs/vi/tro-ly AGENTS.vi.md .agents/rules/ppt-master-vi.md tools/vi/tests/test_vi_layer.py
git commit -m "docs(vi): add the virtual experiment task type for agents"
```

---

### Task 12: Tài liệu cho thầy cô

**Files:**
- Create: `docs/vi/thi-nghiem-ao.md`
- Modify: `docs/vi/xu-ly-loi.md`, `docs/vi/cau-lenh-mau.md`, `docs/vi/bat-dau-nhanh.md`, `README.md`
- Modify: `tools/vi/tests/test_vi_layer.py`

**Interfaces:**
- Consumes: tên file đầu ra và `error.step` của `thi_nghiem.py` (Task 10); 8 mã mẫu (Task 3–6).

- [ ] **Step 1: Viết test**

Trong `tools/vi/tests/test_vi_layer.py`, thêm `"thi-nghiem-ao.md",` vào cuối tuple `REQUIRED_DOCS`, và thêm ngay trước dòng `if __name__ == "__main__":`:

```python
class ExperimentUserDocsTest(unittest.TestCase):
    def test_doc_explains_use_limits_and_review(self):
        text = read("docs/vi/thi-nghiem-ao.md")
        for phrase in ("thi-nghiem.html", "phieu-hoc-tap.docx", "can-soat.md", "không cần mạng", "USB", "Netlify",
                       "điện thoại", "Chế độ giáo viên", "Ghi lần đo", "Chép số liệu", "sai số đo", "Tự kiểm",
                       "mô hình lí tưởng", "soát công thức", "Node"):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, text)

    def test_doc_lists_the_eight_models(self):
        text = read("docs/vi/thi-nghiem-ao.md")
        for name in ("Ném xiên", "Con lắc đơn", "nối tiếp và song song", "Chuẩn độ", "N₂O₄ ⇌ 2NO₂", "Tốc độ phản ứng",
                     "Khảo sát hàm số", "Xác suất thực nghiệm"):
            self.assertIn(name, text)

    def test_quick_start_mentions_the_experiment_task(self):
        text = read("docs/vi/bat-dau-nhanh.md")
        self.assertNotIn("8 loại", text)
        headings = h2_headings(text)
        self.assertIn("## Làm thí nghiệm ảo", headings)
        self.assertLess(headings.index("## Làm thí nghiệm ảo"), headings.index("## Lấy file kết quả"))
        self.assertIn("(thi-nghiem-ao.md)", section(text, "## Làm thí nghiệm ảo"))

    def test_sample_commands_have_an_experiment_section(self):
        body = section(read("docs/vi/cau-lenh-mau.md"), "## Thí nghiệm ảo")
        self.assertIn("thí nghiệm ảo", body)
        self.assertIn("phiếu học tập", body)

    def test_troubleshooting_has_the_experiment_section(self):
        text = read("docs/vi/xu-ly-loi.md")
        headings = h2_headings(text)
        self.assertEqual(headings.index("## Tạo thí nghiệm ảo thất bại"), headings.index("## Kiểm hiệu ứng không đạt") - 1)
        body = section(text, "## Tạo thí nghiệm ảo thất bại")
        for phrase in ("`input`", "`parse`", "`model`", "`check`", "`docx`", "`write`", "`internal`", "Dòng",
                       "requirements-vi.txt", "dải đỏ", "Node"):
            with self.subTest(phrase=phrase):
                self.assertIn(phrase, body)

    def test_readme_mentions_the_experiment_feature(self):
        readme = read("README.md")
        self.assertIn("thí nghiệm ảo", section(readme, "## Làm được gì"))
        self.assertIn("(docs/vi/thi-nghiem-ao.md)", section(readme, "## Tài liệu"))
```

- [ ] **Step 2: Chạy để thấy trượt**

Run: `venv\Scripts\python.exe -m unittest tools.vi.tests.test_vi_layer`
Expected: FAIL ở `ExperimentUserDocsTest` và `test_required_vietnamese_docs_exist`.

- [ ] **Step 3: Viết tài liệu cho thầy cô**

`docs/vi/thi-nghiem-ao.md`:

```markdown
# Làm thí nghiệm ảo

Thí nghiệm ảo là một trang web nhỏ chạy ngay trên máy của thầy cô: học sinh đổi tham số, đo, ghi số liệu, vẽ đồ thị rồi rút ra kết luận. Bộ công cụ tạo cho mỗi thí nghiệm ba file:

| File | Dùng để |
|---|---|
| `thi-nghiem.html` | Mở bằng trình duyệt (Chrome, Edge) là chạy, **không cần mạng** |
| `phieu-hoc-tap.docx` | Phiếu in cho học sinh; trang cuối là đáp án dành cho giáo viên |
| `can-soat.md` | Công thức của mô hình, điều kiện lí tưởng hoá và những điểm thầy cô cần soát |

## Cách yêu cầu

Nhắn cho AI, ví dụ: `Làm thí nghiệm ảo con lắc đơn cho Vật lí 11, học sinh làm theo nhóm, có phiếu học tập`. AI hỏi một lượt ngắn (thí nghiệm nào, ai thao tác, học sinh cần rút ra kết luận gì, đo mấy lần, có bật sai số đo không), rồi tạo file trong `projects\_thi-nghiem\<tên_thí_nghiệm>\`.

## Các thí nghiệm có sẵn

| Môn | Thí nghiệm | Học sinh thay đổi | Học sinh đo |
|---|---|---|---|
| Vật lí | Ném xiên | vận tốc đầu, góc ném, độ cao, g | tầm xa, độ cao cực đại, thời gian bay |
| Vật lí | Con lắc đơn | chiều dài dây, g, khối lượng | chu kì, thời gian 10 dao động |
| Vật lí | Đoạn mạch nối tiếp và song song | hiệu điện thế, hai điện trở, kiểu mắc | cường độ dòng điện, hiệu điện thế |
| Hoá học | Chuẩn độ acid – base | acid mạnh hay yếu, nồng độ, thể tích NaOH nhỏ vào, chất chỉ thị | pH |
| Hoá học | Cân bằng N₂O₄ ⇌ 2NO₂ | nhiệt độ, áp suất | phần mol NO₂, độ đậm màu nâu |
| Hoá học | Tốc độ phản ứng | nồng độ, nhiệt độ, chất xúc tác | thời gian phản ứng |
| Toán | Khảo sát hàm số | các hệ số, tiếp điểm | giá trị, hệ số góc tiếp tuyến, cực trị |
| Toán | Xác suất thực nghiệm | phép thử, số lần gieo | tần số, tần suất |

Tám thí nghiệm này đã được kiểm bằng số: mỗi mô hình được tính lại độc lập bằng một chương trình khác và so trên hàng trăm bộ số, cộng với kiểm định luật (bảo toàn cơ năng, pH = 7 tại điểm tương đương, tần suất tiến về xác suất…).

Thầy cô cần thí nghiệm khác thì AI vẫn viết được mô hình mới. Khi đó `can-soat.md` ghi rõ **mô hình do AI viết, chưa có người duyệt**: thầy cô soát công thức và tính lại ít nhất hai dòng của bảng số kiểm bằng máy tính cầm tay trước khi dùng trên lớp.

## Dùng trên lớp

- **Máy chiếu:** chép thư mục thí nghiệm sang USB, bấm đúp `thi-nghiem.html`. Không cần mạng, không cần cài gì.
- **Học sinh dùng điện thoại:** điện thoại không mở được file từ USB. Thầy cô đưa `thi-nghiem.html` lên một trang web tĩnh (ví dụ kéo thả vào Netlify) rồi gửi địa chỉ cho học sinh.
- **Ba bước trên trang:** *Dự đoán* → *Quan sát* → *Giải thích*. Học sinh phải chốt dự đoán thì mới làm được thí nghiệm; ghi đủ số lần đo thì mục Giải thích mới mở và lúc đó trang mới đối chiếu dự đoán với kết quả. Nút **Chế độ giáo viên** bỏ mọi khoá để thầy cô trình diễn tự do.
- **Ghi lần đo** đưa số liệu đang hiện vào bảng; đồ thị vẽ từ chính bảng đó, kèm đường thẳng khớp và hệ số góc. **Chép số liệu** rồi dán thẳng vào Excel.
- **Sai số đo:** bật thì mỗi lần đo lệch ngẫu nhiên một chút như đo thật, để học sinh tập lấy trung bình và xử lí số liệu.
- Trang không lưu điểm và không gửi dữ liệu đi đâu.

## Những điều cần biết

- Thí nghiệm ảo là **mô hình lí tưởng**, không thay thí nghiệm thật. Cuối trang luôn ghi điều kiện lí tưởng hoá (bỏ qua sức cản, góc lệch nhỏ, dung dịch loãng…); ngoài các điều kiện đó kết quả thật sẽ khác.
- Mỗi lần mở, trang **Tự kiểm** mô hình của mình và ghi kết quả ở cuối trang. Nếu hiện dải đỏ "Mô hình không qua tự kiểm" thì không dùng file đó để dạy; nhờ AI tạo lại.
- Máy có cài Node thì công cụ kiểm số ngay lúc tạo file; máy không có Node thì bước này bỏ qua và trang tự kiểm khi mở. Không cần cài Node để dùng.
- Muốn gắn thí nghiệm vào bài giảng PowerPoint, nhờ AI thêm một trang có liên kết mở `thi-nghiem.html`; nhớ chép file HTML đi cùng file PPTX.
```

- [ ] **Step 4: Sửa các tài liệu liên quan**

`docs/vi/xu-ly-loi.md` — thêm ngay trước tiêu đề `## Kiểm hiệu ứng không đạt`:

```markdown
## Tạo thí nghiệm ảo thất bại

Xem dòng kết quả AI đọc được, phần `error`:

- `input`: chưa có thư mục thí nghiệm hoặc file `thi-nghiem.md`. Nhờ AI viết file theo docs/vi/tro-ly/thi-nghiem-ao.md rồi chạy lại.
- `parse`: `error.message` nêu đúng số **Dòng** trong `thi-nghiem.md` cần sửa. Hay gặp nhất: khoảng tham số vượt khoảng của mẫu (ví dụ góc lệch con lắc trên 15°), vì ngoài khoảng đó công thức không còn đúng.
- `model`: mã mẫu không có trong danh mục, hoặc mô hình AI tự viết thiếu công thức, điều kiện áp dụng hay bảng số kiểm. AI sửa theo khuôn; không bỏ các phần đó.
- `check`: mô hình AI tự viết cho kết quả lệch bảng số kiểm. AI sửa mô hình; nếu AI sửa bảng số kiểm thì phải nói rõ với thầy cô dòng nào đã sửa.
- `docx`: máy chưa có thư viện `python-docx`. Có `venv\Scripts\python.exe` ở thư mục gốc repo thì chạy `venv\Scripts\python.exe -m pip install -r tools/vi/requirements-vi.txt`; không thì chạy `python -m pip install -r tools/vi/requirements-vi.txt`, hoặc bấm đúp `CAI-DAT.bat`.
- `write`: phiếu học tập đang mở trong Word, hoặc ổ đĩa hết dung lượng, hoặc đường dẫn quá 200 ký tự (xem mục **Đường dẫn quá dài**).
- `internal`: lỗi ngoài dự kiến. Dán nguyên dòng `error.message` gửi người bảo trì.

Cảnh báo "chưa chạy kiểm số vì không có Node" không phải lỗi: trang HTML tự kiểm mỗi lần mở. Trang hiện **dải đỏ** "Mô hình không qua tự kiểm" thì không dùng để dạy; nhờ AI tạo lại.

```

`docs/vi/cau-lenh-mau.md` — thêm ngay trước tiêu đề `## Tạo nhanh`:

````markdown
## Thí nghiệm ảo

Thí nghiệm có sẵn trong danh mục:

```
Làm thí nghiệm ảo con lắc đơn cho Vật lí 11, học sinh làm theo nhóm, có phiếu học tập
```

Thí nghiệm cho giáo viên trình diễn trên máy chiếu:

```
Làm thí nghiệm ảo chuyển dịch cân bằng N2O4 ⇌ 2NO2 cho Hoá 11, giáo viên trình diễn, không cần phiếu học tập
```

````

`docs/vi/bat-dau-nhanh.md`:

1. Thay `Với 8 loại việc trên` bằng `Với 9 loại việc trên`.
2. Thêm ngay trước tiêu đề `## Lấy file kết quả`:

```markdown
## Làm thí nghiệm ảo

Cần một mô phỏng để học sinh đổi tham số, đo và vẽ đồ thị, nhắn `Làm thí nghiệm ảo con lắc đơn cho Vật lí 11, có phiếu học tập`. AI tạo một file HTML mở bằng trình duyệt là chạy, không cần mạng, kèm phiếu học tập Word. Chi tiết trong [Làm thí nghiệm ảo](thi-nghiem-ao.md).

```

`README.md`:

1. Trong mục `## Làm được gì`, thêm sau dòng "Soạn **giáo án**…":

```
- Tạo **thí nghiệm ảo** Toán, Vật lí, Hoá học: một file HTML chạy không cần mạng, có bảng số liệu, đồ thị và phiếu học tập Word; 8 mô hình đã kiểm bằng số.
```

2. Trong bảng mục `## Tài liệu`, thêm sau dòng "Soạn giáo án":

```
| [Làm thí nghiệm ảo](docs/vi/thi-nghiem-ao.md) | File HTML tương tác chạy không cần mạng, kèm phiếu học tập |
```

- [ ] **Step 5: Chạy lại toàn bộ test lớp Việt**

Run:
```
venv\Scripts\python.exe -m unittest tools.vi.tests.test_vi_layer
venv\Scripts\python.exe -m unittest tools.vi.tests.test_de_thi tools.vi.tests.test_doctor tools.vi.tests.test_giao_an tools.vi.tests.test_installer tools.vi.tests.test_kiem_hieu_ung tools.vi.tests.test_video
```
Expected: PASS. `test_vi_layer` có 181 test (175 + 6). `test_changes_limited_to_vietnamese_layer` phải qua: mọi file mới nằm trong `tools/vi/`, `docs/vi/`, `AGENTS.vi.md`, `README.md`, `.agents/rules/`.

- [ ] **Step 6: Commit**

```bash
git add docs/vi README.md tools/vi/tests/test_vi_layer.py
git commit -m "docs(vi): document virtual experiments for teachers"
```

---

### Task 13: Chạy thật 8 mẫu, kiểm bằng mắt, ghi biên bản (controller làm, không giao subagent)

Bước này **mở cửa sổ trình duyệt** khi chủ repo tự xem; phần chụp ảnh chạy ngầm, không mở cửa sổ. Báo trước cho chủ repo.

**Files:**
- Create: `docs/vi/phat-trien/2026-09-20-thi-nghiem-ao-kiem-thu.md`
- Đầu ra thử ở `projects/_thi-nghiem/_thu-*/` (không commit)

- [ ] **Step 1: Tạo 8 thí nghiệm thử, mỗi mẫu một cái**

Viết `projects/_thi-nghiem/_thu-<mã>/thi-nghiem.md` cho từng mẫu trong 8 mẫu. Ba mẫu `li-con-lac-don`, `hoa-can-bang-no2`, `toan-xac-suat` dùng nguyên ba ví dụ trong `docs/vi/tro-ly/thi-nghiem-ao.md`. Năm mẫu còn lại dùng các file dưới đây.

`projects/_thi-nghiem/_thu-li-nem-xien/thi-nghiem.md`:

```text
---
tieu-de: Ném xiên
mon: Vật lí
lop: 11
mau: li-nem-xien
---

## Tham số
goc: 15..75 buoc 15 mac-dinh 45

## Dự đoán
cau: Em dự đoán kết quả thay đổi thế nào?

## Quan sát
so-lan-do: 4
do: goc, tam-xa
do-thi: tam-xa theo goc

## Giải thích
cau: Giải thích kết quả.
goi-y-dap-an: Theo công thức của mô hình.

## Kết luận
Kết luận của thí nghiệm.
```

`projects/_thi-nghiem/_thu-li-mach-ohm/thi-nghiem.md`:

```text
---
tieu-de: Mạch điện
mon: Vật lí
lop: 11
mau: li-mach-ohm
---

## Tham số
kieu-mac: chon song-song, noi-tiep
dien-tro-2: 10..50 buoc 10 mac-dinh 20

## Dự đoán
cau: Em dự đoán kết quả thay đổi thế nào?

## Quan sát
so-lan-do: 4
do: dien-tro-2, cuong-do-mach-chinh
do-thi: cuong-do-mach-chinh theo 1/dien-tro-2

## Giải thích
cau: Giải thích kết quả.
goi-y-dap-an: Theo công thức của mô hình.

## Kết luận
Kết luận của thí nghiệm.
```

`projects/_thi-nghiem/_thu-hoa-chuan-do/thi-nghiem.md`:

```text
---
tieu-de: Chuẩn độ
mon: Hoá học
lop: 11
mau: hoa-chuan-do
---

## Tham số
loai-acid: co-dinh ch3cooh
the-tich-base: 0..40 buoc 0.5 mac-dinh 12

## Dự đoán
cau: Em dự đoán kết quả thay đổi thế nào?

## Quan sát
so-lan-do: 4
do: the-tich-base, ph
do-thi: ph theo the-tich-base

## Giải thích
cau: Giải thích kết quả.
goi-y-dap-an: Theo công thức của mô hình.

## Kết luận
Kết luận của thí nghiệm.
```

`projects/_thi-nghiem/_thu-hoa-toc-do/thi-nghiem.md`:

```text
---
tieu-de: Tốc độ phản ứng
mon: Hoá học
lop: 11
mau: hoa-toc-do
---

## Tham số
nhiet-do: 10..60 buoc 10 mac-dinh 20
xuc-tac: chon khong, co

## Dự đoán
cau: Em dự đoán kết quả thay đổi thế nào?

## Quan sát
so-lan-do: 4
do: nhiet-do, thoi-gian
do-thi: ln(thoi-gian) theo nhiet-do

## Giải thích
cau: Giải thích kết quả.
goi-y-dap-an: Theo công thức của mô hình.

## Kết luận
Kết luận của thí nghiệm.
```

`projects/_thi-nghiem/_thu-toan-ham-so/thi-nghiem.md`:

```text
---
tieu-de: Khảo sát hàm số
mon: Toán
lop: 11
mau: toan-ham-so
---

## Tham số
a: -2..2 buoc 0.5 mac-dinh 1
x0: -3..3 buoc 0.1 mac-dinh 1.5

## Dự đoán
cau: Em dự đoán kết quả thay đổi thế nào?

## Quan sát
so-lan-do: 4
do: x0, gia-tri, he-so-goc

## Giải thích
cau: Giải thích kết quả.
goi-y-dap-an: Theo công thức của mô hình.

## Kết luận
Kết luận của thí nghiệm.
```

Chạy `venv\Scripts\python.exe tools\vi\thi_nghiem.py projects\_thi-nghiem\_thu-<mã>` cho từng thư mục.
Expected: 8 dòng JSON đều `"ready": true`, `kiem_so.chay` là `true`, `dat` bằng `tong`.

- [ ] **Step 2: Chụp ảnh ngầm và kiểm DOM**

Với mỗi file, dùng `chrome-headless-shell.exe` trong `%LOCALAPPDATA%\ms-playwright\chromium_headless_shell-*\` (nếu có; không có thì mở tay bằng Edge):

```
chrome-headless-shell.exe --no-sandbox --disable-gpu --window-size=1366,1500 --virtual-time-budget=3000 --screenshot=<thư_mục>\rong.png file:///<đường_dẫn>/thi-nghiem.html
chrome-headless-shell.exe --no-sandbox --disable-gpu --window-size=400,1900 --virtual-time-budget=3000 --screenshot=<thư_mục>\hep.png file:///<đường_dẫn>/thi-nghiem.html
chrome-headless-shell.exe --no-sandbox --disable-gpu --virtual-time-budget=3000 --dump-dom file:///<đường_dẫn>/thi-nghiem.html
```

Expected trong DOM: `id="tu-kiem">Tự kiểm: N/N đạt.`; ba khối `class="buoc`; không có `class="dai-do"` trong phần thân. Xem từng ảnh: khung vẽ không tràn, chữ không méo, bố cục hẹp xếp dọc.

- [ ] **Step 3: Chạy thử luồng bấm cho một mẫu có khoá**

Chép `thi-nghiem.html` của mẫu con lắc (bản `nguoi-thao-tac: nhom`) thành `auto.html`, chèn trước `</body>` một đoạn script chọn đáp án, bấm "Chốt dự đoán", đổi thanh trượt 5 lần và bấm "Ghi lần đo" mỗi lần, bấm "Xem gợi ý", rồi ghi `document.title = 'XONG ' + số_dòng + ' ' + nội_dung_ô_khop`. Dump DOM.
Expected: tiêu đề `XONG 5 Đường thẳng khớp: hệ số góc = 4,0…` (khoảng 3,9–4,1 khi bật sai số; lí thuyết 4π²/g = 4,03).

- [ ] **Step 4: Mở phiếu Word**

Mở `phieu-hoc-tap.docx` của hai mẫu bằng Word: bảng đúng số cột, lưới đồ thị vuông vức trong một trang, trang giáo viên ở trang riêng, chỉ số trên dưới hiển thị đúng.

- [ ] **Step 5: Ghi biên bản kiểm thử**

Viết `docs/vi/phat-trien/2026-09-20-thi-nghiem-ao-kiem-thu.md`: bảng 8 mẫu × (JSON ready, kiểm số, tự kiểm trong trình duyệt, ảnh rộng, ảnh hẹp), kết quả luồng bấm, kết quả mở Word, và mọi lỗi tìm thấy cùng commit sửa. Không chép đường dẫn riêng của máy chủ repo vào file.

- [ ] **Step 6: Gửi chủ repo soát ba mẫu Hoá**

Gửi ảnh và `can-soat.md` của `hoa-chuan-do`, `hoa-can-bang-no2`, `hoa-toc-do`; nêu rõ các hằng số (K~a~, Δ~r~H°, Δ~r~S°, E~a~) và nguồn. Chờ chủ repo xác nhận hoặc yêu cầu sửa trước khi sang Task 14.

- [ ] **Step 7: Commit biên bản**

```bash
git add docs/vi/phat-trien/2026-09-20-thi-nghiem-ao-kiem-thu.md
git commit -m "docs(vi): record the virtual experiment acceptance run"
```

---

### Task 14: Phát hành v6.3.2-vi.8 (controller làm, chỉ sau khi chủ repo đồng ý)

**Files:**
- Modify: `CHANGELOG-VI.md`, `README.md`

- [ ] **Step 1: Viết mục nhật ký thay đổi**

Thêm vào đầu `CHANGELOG-VI.md`, ngay dưới dòng tiêu đề, mục `## 6.3.2-vi.8 — <ngày phát hành>` gồm: một câu tóm tắt; `### Thêm` (công cụ `thi_nghiem.py` và ba file đầu ra; 8 mô hình và cách kiểm ba lớp; khung Dự đoán – Quan sát – Giải thích, bảng số liệu, đồ thị, sai số đo; phiếu học tập; mô hình mới do AI viết và quy tắc soát; loại việc thứ 9, mục 14 của `AGENTS.vi.md`, luật Antigravity; tài liệu thầy cô và mục xử lý lỗi); `### Không thay đổi` (lõi PPT Master v6.3.2 giữ nguyên; không thêm thư viện; Node không bắt buộc); `### Rủi ro` (mô hình lí tưởng không thay thí nghiệm thật; mô hình do AI viết chỉ được bảo vệ bằng bảng kiểm do AI viết và lời nhắc soát; điện thoại cần đưa file lên web; máy không có Node thì chỉ còn lớp tự kiểm trong trang; nêu trung thực phần nào chủ repo chưa tự nghiệm thu).

- [ ] **Step 2: Đổi số phiên bản**

Trong `README.md` thay `Phiên bản: **6.3.2-vi.7**` bằng `Phiên bản: **6.3.2-vi.8**`.

- [ ] **Step 3: Chạy toàn bộ test và kiểm toàn vẹn**

Run: mọi module `tools.vi.tests.test_*` và `venv\Scripts\python.exe skills/ppt-master/scripts/attribution_guard.py`.
Expected: tất cả PASS; guard thoát mã 0.

- [ ] **Step 4: Commit, gộp, gắn tag, phát hành**

```bash
git add CHANGELOG-VI.md README.md
git commit -m "docs(vi): release 6.3.2-vi.8"
git switch main && git merge --ff-only feat/vi-thi-nghiem-ao
git push origin main
git tag -a v6.3.2-vi.8 -m "PPT Master ban Viet 6.3.2-vi.8"
git push origin v6.3.2-vi.8
gh release create v6.3.2-vi.8 --repo luonghaianh1208/PPTmaster --title "PPT Master bản Việt 6.3.2-vi.8" --notes-file <file chứa đúng mục 6.3.2-vi.8 của CHANGELOG-VI.md> --verify-tag
```

Nếu nhánh có commit chứa đường dẫn riêng của chủ repo thì gộp bằng squash thay cho fast-forward, như đã làm ở vi.6. Sau khi phát hành: `gh release list --repo luonghaianh1208/PPTmaster --limit 2` phải hiện `v6.3.2-vi.8` là Latest, và kiểm repo gốc không bị tạo release.
