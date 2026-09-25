(function (root) {
  'use strict';

  var TOC_DO_VIET = 20;
  var LAU_BANG = 0.5;
  var HINH_TOI_THIEU = 1.2;
  var CO_TAY = 0.85;
  var NS = 'http://www.w3.org/2000/svg';
  var DANH_DAU = /\*\*(.+?)\*\*|~([^~]+)~|\^([^\^]+)\^/g;

  function kep(x, a, b) { return x < a ? a : (x > b ? b : x); }
  function tienDo(t, batDau, thoiLuong) {
    if (thoiLuong <= 0) { return t >= batDau ? 1 : 0; }
    return kep((t - batDau) / thoiLuong, 0, 1);
  }
  function thoat(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function tach(chu) {
    var kq = [];
    var vt = 0;
    var m;
    chu = String(chu);
    DANH_DAU.lastIndex = 0;
    while ((m = DANH_DAU.exec(chu))) {
      if (m.index > vt) { kq.push({ the: '', chu: chu.slice(vt, m.index) }); }
      kq.push(m[1] !== undefined ? { the: 'b', chu: m[1] } : (m[2] !== undefined ? { the: 'sub', chu: m[2] } : { the: 'sup', chu: m[3] }));
      vt = m.index + m[0].length;
    }
    if (vt < chu.length) { kq.push({ the: '', chu: chu.slice(vt) }); }
    return kq;
  }
  function demKyTu(chu) { return tach(chu).reduce(function (n, o) { return n + o.chu.length; }, 0); }
  // Hiện n ký tự đầu; khi 0 < n < tổng, chèn span.ngoi rỗng ngay sau ký tự thứ n (điểm ngòi bút).
  function catDanhDau(chu, n) {
    var con = n;
    var html = '';
    var giua = n > 0 && n < demKyTu(chu);
    tach(chu).forEach(function (o) {
      var hien = con > 0 ? o.chu.slice(0, con) : '';
      var an = o.chu.slice(hien.length);
      con -= hien.length;
      var ngoi = giua && hien && con === 0 ? '<span class="ngoi"></span>' : '';
      if (ngoi) { giua = false; }
      var trong = thoat(hien) + ngoi + (an ? '<span class="an">' + thoat(an) + '</span>' : '');
      html += o.the ? '<' + o.the + '>' + trong + '</' + o.the + '>' : trong;
    });
    return html;
  }
  function thoiGianViet(chu) { return Math.max(0.5, demKyTu(chu) / TOC_DO_VIET); }

  function rng(hat) {
    var a = hat | 0;
    return function () {
      a = a + 0x6D2B79F5 | 0;
      var t = Math.imul(a ^ a >>> 15, 1 | a);
      t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t;
      return ((t ^ t >>> 14) >>> 0) / 4294967296;
    };
  }
  function lam(x) { return Math.round(x * 10) / 10; }
  function duongQua(diem, hat) {
    var r = rng(hat);
    var d = '';
    for (var i = 0; i < diem.length; i++) {
      var p = diem[i];
      if (i === 0) { d += 'M' + lam(p[0]) + ' ' + lam(p[1]); continue; }
      var q = diem[i - 1];
      var dx = p[0] - q[0];
      var dy = p[1] - q[1];
      var dai = Math.sqrt(dx * dx + dy * dy) || 1;
      var nx = -dy / dai;
      var ny = dx / dai;
      var n = Math.max(2, Math.round(dai / 70));
      for (var j = 1; j <= n; j++) {
        var u = j / n;
        var lech = j === n ? 0 : (r() - 0.5) * 4;
        d += ' L' + lam(q[0] + dx * u + nx * lech) + ' ' + lam(q[1] + dy * u + ny * lech);
      }
    }
    return d;
  }
  function hopQua(x, y, w, h, hat) { return duongQua([[x, y], [x + w, y], [x + w, y + h], [x, y + h], [x, y - 2]], hat); }
  function vongTron(cx, cy, r) {
    return 'M' + lam(cx - r) + ' ' + lam(cy) + ' a' + r + ' ' + r + ' 0 1 0 ' + (2 * r) + ' 0 a' + r + ' ' + r + ' 0 1 0 ' + (-2 * r) + ' 0';
  }
  function muiTen(x1, y1, x2, y2, hat) {
    var goc = Math.atan2(y2 - y1, x2 - x1);
    var c = function (lech) { return [x2 - 16 * Math.cos(goc + lech), y2 - 16 * Math.sin(goc + lech)]; };
    var a = c(0.5);
    var b = c(-0.5);
    return duongQua([[x1, y1], [x2, y2]], hat) + ' M' + lam(a[0]) + ' ' + lam(a[1]) + ' L' + x2 + ' ' + y2 + ' L' + lam(b[0]) + ' ' + lam(b[1]);
  }

  function gan(dich, tuy) {
    Object.keys(tuy || {}).forEach(function (k) { dich[k] = tuy[k]; });
    return dich;
  }
  // Đặt khung w×h giữ tỉ lệ rong:cao vừa trong ô (x, y, o, c), căn giữa (object-fit: contain).
  function vuaKhung(rong, cao, x, y, o, c) {
    var s = Math.min(o / rong, c / cao);
    var w = rong * s, h = cao * s;
    return { x: x + (o - w) / 2, y: y + (c - h) / 2, rong: w, cao: h };
  }
  function tao(du) {
    var gh = du.thoiLuong;
    var sau = du.co && du.co.lauBang ? LAU_BANG + 0.05 : 0;
    function dau(batDau, lui) { return Math.min(Math.max(batDau, sau), gh - lui); }
    function chu(id, noiDung, x, y, rong, cao, co, batDau, tuy) {
      batDau = dau(batDau, 0.6);
      var dai = Math.max(0.3, Math.min(thoiGianViet(noiDung), gh - 0.2 - batDau));
      return gan({ id: id, kieu: 'chu', chu: noiDung, x: x, y: y, rong: rong, cao: cao, co: co, batDau: batDau, thoiLuong: dai, can: 'trai', mau: '' }, tuy);
    }
    function net(id, d, batDau, dai, tuy) {
      batDau = dau(batDau, 0.6);
      dai = Math.max(0.2, Math.min(dai, gh - 0.2 - batDau));
      return gan({ id: id, kieu: 'net', d: d, batDau: batDau, thoiLuong: dai, mau: '', day: 4 }, tuy);
    }
    // Biểu tượng vẽ lần lượt từng phần tử; cả hình ít nhất 1,2 s.
    function hinh(id, h, x, y, kich, batDau, tuy) {
      batDau = dau(batDau, 0.2 + HINH_TOI_THIEU);
      var dai = Math.min(3, Math.max(HINH_TOI_THIEU, 0.45 * h.phanTu.length));
      dai = Math.max(0.3, Math.min(dai, gh - 0.2 - batDau));
      return gan({ id: id, kieu: 'hinh', phanTu: h.phanTu, viewBox: h.viewBox, x: x, y: y, kich: kich, batDau: batDau, thoiLuong: dai, mau: '' }, tuy);
    }
    // Ảnh hiện dần 0,4 s rồi phóng/lướt chậm tới cuối cảnh; khung giữ tỉ lệ ảnh trong ô.
    function anh(id, a, x, y, rong, cao, batDau) {
      batDau = dau(batDau, 0.6);
      var k = vuaKhung(a.rong, a.cao, x, y, rong, cao);
      return { id: id, kieu: 'anh', dataUrl: a.dataUrl, nguon: a.nguon, x: k.x, y: k.y, rong: k.rong, cao: k.cao,
        batDau: batDau, thoiLuong: Math.min(0.4, gh - 0.2 - batDau) };
    }
    function tieuDe(noiDung, batDau, rong) {
      rong = rong || 1160;
      var co = rong < 1160 && demKyTu(noiDung) > 40 ? 32 : 40;
      var c = chu('tieu-de', noiDung, 60, 30, rong, 116, co, batDau, { mau: 'nhan', day: true });
      var w = Math.min(rong, Math.max(240, demKyTu(noiDung) * co * 0.6));
      return [c, net('gach', duongQua([[60, 160], [60 + w, 160]], 7), c.batDau + c.thoiLuong, 0.4, { mau: 'nhan', quay: false })];
    }
    // Cột hình bên phải (ô x 900, y 200, rộng 320, cao 380) cho cảnh có `hinh` hoặc `anh`.
    var coCot = !!(du.hinh || du.anh);
    function cot() {
      var batDau = du.moc && du.moc.length ? du.moc[0] : 1.0;
      if (du.hinh) { return [hinh('hinh', du.hinh, 910, 240, 300, batDau)]; }
      if (du.anh) { return [anh('anh', du.anh, 900, 200, 320, 380, batDau)]; }
      return [];
    }
    return { chu: chu, net: net, hinh: hinh, anh: anh, tieuDe: tieuDe, cot: cot, coCot: coCot, gh: gh };
  }

  function matTran(s) {
    if (Math.abs(s.z - 1) < 1e-9 && Math.abs(s.tx) < 1e-6 && Math.abs(s.ty) < 1e-6) { return 'none'; }
    return 'matrix(' + s.z + ',0,0,' + s.z + ',' + s.tx + ',' + s.ty + ')';
  }

  function khoiDong(du) {
    var H = root.THI_HINH;
    var T = root.THI_BAN_TAY;
    var Q = root.THI_MAY_QUAY;
    var khung = document.getElementById('khung');
    var loai = root.THI_CANH[du.loai];
    var muc = loai.muc(du);
    var co = du.co || {};
    var gh = du.thoiLuong;
    var thiNghiem = du.loai === 'thi-nghiem';
    var nen = null;
    if (co.lauBang && du.nenTruoc) {
      nen = document.createElement('img');
      nen.id = 'nen-truoc';
      nen.src = du.nenTruoc;
      khung.appendChild(nen);
    }
    var goc = document.createElement('div');
    goc.id = 'bang';
    khung.appendChild(goc);
    var svg = document.createElementNS(NS, 'svg');
    svg.setAttribute('class', 've');
    svg.setAttribute('viewBox', '0 0 1280 720');
    goc.appendChild(svg);
    var ds = muc.map(function (m) {
      var el;
      if (m.kieu === 'net') {
        el = document.createElementNS(NS, 'path');
        el.setAttribute('d', m.d);
        el.setAttribute('pathLength', '1');
        el.setAttribute('class', 'net ' + m.mau);
        el.style.strokeWidth = String(m.day);
        el.style.strokeDasharray = '1';
        svg.appendChild(el);
      } else if (m.kieu === 'hinh') {
        el = H.taoHinh(svg, m);
      } else if (m.kieu === 'anh') {
        el = H.taoAnh(goc, m);
      } else {
        el = document.createElement('div');
        el.className = 'chu ' + m.can + ' ' + m.mau + (m.day ? ' day' : '');
        el.setAttribute('data-id', m.id);
        el.style.left = m.x + 'px';
        el.style.top = m.y + 'px';
        el.style.width = m.rong + 'px';
        el.style.height = m.cao + 'px';
        el.style.fontSize = m.co + 'px';
        goc.appendChild(el);
      }
      return { m: m, el: el };
    });
    var theoId = {};
    ds.forEach(function (o) { theoId[o.m.id] = o; });
    if (loai.dung) { loai.dung(goc, du); }
    var tay = null;
    if (co.banTay && !thiNghiem) {
      goc.insertAdjacentHTML('beforeend', T.SVG);
      tay = goc.lastElementChild;
    }

    function datMuc(t) {
      ds.forEach(function (o) {
        var p = tienDo(t, o.m.batDau, o.m.thoiLuong);
        if (o.m.kieu === 'net') {
          o.el.style.strokeDashoffset = String(1 - p);
          o.el.style.opacity = p > 0 ? '1' : '0';
        } else if (o.m.kieu === 'hinh') {
          H.datHinh(o.el, p);
        } else if (o.m.kieu === 'anh') {
          H.datAnh(o.el, o.m, p, t, gh, du.so);
        } else if (!o.m.dong) {
          var html = catDanhDau(o.m.chu, Math.round(p * demKyTu(o.m.chu)));
          o.el.innerHTML = o.m.day ? '<span class="trong">' + html + '</span>' : html;
        }
      });
      if (loai.capNhat) { loai.capNhat(goc, du, t); }
    }

    // Hộp bao và điểm đầu/cuối của từng mục ở trạng thái cuối, đo khi Z = 1. Đo một lần, sau khi font đã nạp.
    var hop = null;
    var dauCuoi = {};
    function doHop() {
      goc.style.transform = 'none';
      datMuc(1e6);
      var k = khung.getBoundingClientRect();
      var kq = {};
      ds.forEach(function (o) {
        var m = o.m;
        if (m.dong) { return; }
        if (m.kieu === 'net') {
          var b = o.el.getBBox();
          kq[m.id] = { x: b.x, y: b.y, w: b.width, h: b.height };
        } else if (m.kieu === 'hinh') {
          kq[m.id] = { x: m.x, y: m.y, w: m.kich, h: m.kich };
        } else if (m.kieu === 'anh') {
          kq[m.id] = { x: m.x, y: m.y, w: m.rong, h: m.cao };
        } else {
          var r = document.createRange();
          r.selectNodeContents(o.el);
          var cac = r.getClientRects();
          var bao = r.getBoundingClientRect();
          if (!cac.length || bao.width <= 0) { return; }
          kq[m.id] = { x: bao.left - k.left, y: bao.top - k.top, w: bao.width, h: bao.height };
          var a = cac[0], z = cac[cac.length - 1];
          dauCuoi[m.id] = {
            dau: { x: a.left - k.left, y: a.top + 0.8 * a.height - k.top },
            cuoi: { x: z.right - k.left, y: z.top + 0.8 * z.height - k.top }
          };
        }
      });
      return kq;
    }

    function ngoiCua(cam) {
      var k = khung.getBoundingClientRect();
      return function (m, p) {
        var o = theoId[m.id];
        if (m.kieu === 'net') {
          var L = o.el.getTotalLength();
          var d = o.el.getPointAtLength(kep(p, 0, 1) * L);
          return { x: d.x, y: d.y };
        }
        if (m.kieu === 'hinh') { return H.ngoiHinh(o.el, m, p); }
        var dc = dauCuoi[m.id] || { dau: { x: m.x, y: m.y + m.co }, cuoi: { x: m.x, y: m.y + m.co } };
        var s = p > 0 && p < 1 ? o.el.querySelector('.ngoi') : null;
        if (!s) { return p >= 0.5 ? dc.cuoi : dc.dau; }
        var r = s.getBoundingClientRect();
        return { x: (r.left - k.left - cam.tx) / cam.z, y: (r.top + 0.8 * r.height - k.top - cam.ty) / cam.z };
      };
    }

    function datNen(t) {
      if (!nen) { return; }
      if (t > LAU_BANG) { nen.style.display = 'none'; return; }
      nen.style.display = 'block';
      nen.style.clipPath = 'inset(0 0 0 ' + kep(-120 + 1400 * t / LAU_BANG, 0, 1280) + 'px)';
    }

    function dat(t, noiBo) {
      if (!noiBo && !hop) { hop = doHop(); }
      datMuc(t);
      datNen(t);
      var cam = noiBo ? { z: 1, tx: 0, ty: 0 } : Q.tinh(muc, hop, t, gh, { mayQuay: co.mayQuay === true, day: thiNghiem });
      goc.style.transform = matTran(cam);
      if (!tay) { return; }
      var v = noiBo ? { hien: false } : T.viTri(muc, t, ngoiCua(cam), T.NGHI, { lauBang: !!co.lauBang, gh: gh });
      tay.style.display = v.hien ? 'block' : 'none';
      if (!v.hien) { return; }
      tay.setAttribute('data-kieu', v.kieu);
      // Tay nằm trong lớp bảng để theo máy quay, nhưng giữ cỡ trên màn hình không đổi khi phóng.
      tay.style.transform = 'translate(' + v.x + 'px,' + v.y + 'px) scale(' + (CO_TAY / cam.z) + ')';
    }
    root.datThoiDiem = function (t) { dat(t, false); };
    root.THI_VIDEO.thoiDiemCuoi = function () {
      return du.thoiLuong - 0.2;
    };
    // Kiểm tràn ở trạng thái cuối, máy quay Z = 1, không bàn tay, không nền cảnh trước.
    root.THI_VIDEO.kiemTran = function () {
      dat(1e6, false);
      goc.style.transform = 'none';
      if (tay) { tay.style.display = 'none'; }
      if (nen) { nen.style.display = 'none'; }
      var loi = [];
      var cacO = goc.querySelectorAll('.chu');
      for (var i = 0; i < cacO.length; i++) {
        var el = cacO[i];
        var r = el.getBoundingClientRect();
        if (el.scrollHeight > el.clientHeight + 1 || el.scrollWidth > el.clientWidth + 1 || r.right > 1281 || r.bottom > 721) {
          loi.push(el.getAttribute('data-id'));
        }
      }
      return loi;
    };
    dat(0, true);
    root.THI_VIDEO.san = true;
  }

  root.THI_CANH = root.THI_CANH || {};
  root.THI_VIDEO = {
    LAU_BANG: LAU_BANG,
    kep: kep, tienDo: tienDo, thoat: thoat, demKyTu: demKyTu, catDanhDau: catDanhDau,
    thoiGianViet: thoiGianViet, duongQua: duongQua, hopQua: hopQua, vongTron: vongTron, muiTen: muiTen,
    rng: rng, vuaKhung: vuaKhung, tao: tao, khoiDong: khoiDong, san: false
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
