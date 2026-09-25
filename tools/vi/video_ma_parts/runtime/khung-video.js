(function (root) {
  'use strict';

  var DAN_DAU = 0.7;
  var TOC_DO_VIET = 20;
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
  function catDanhDau(chu, n) {
    var con = n;
    var html = '';
    tach(chu).forEach(function (o) {
      var hien = con > 0 ? o.chu.slice(0, con) : '';
      var an = o.chu.slice(hien.length);
      con -= hien.length;
      var trong = thoat(hien) + (an ? '<span class="an">' + thoat(an) + '</span>' : '');
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
  function tao(du) {
    var gh = du.thoiLuong;
    function chu(id, noiDung, x, y, rong, cao, co, batDau, tuy) {
      batDau = Math.min(batDau, gh - 0.6);
      var dai = Math.max(0.3, Math.min(thoiGianViet(noiDung), gh - 0.2 - batDau));
      return gan({ id: id, kieu: 'chu', chu: noiDung, x: x, y: y, rong: rong, cao: cao, co: co, batDau: batDau, thoiLuong: dai, can: 'trai', mau: '' }, tuy);
    }
    function net(id, d, batDau, dai, tuy) {
      batDau = Math.min(batDau, gh - 0.6);
      dai = Math.max(0.2, Math.min(dai, gh - 0.2 - batDau));
      return gan({ id: id, kieu: 'net', d: d, batDau: batDau, thoiLuong: dai, mau: '', day: 4 }, tuy);
    }
    function tieuDe(noiDung, batDau) {
      var c = chu('tieu-de', noiDung, 60, 30, 1160, 116, 40, batDau, { mau: 'nhan' });
      var w = Math.min(1160, Math.max(240, demKyTu(noiDung) * 40 * 0.6));
      return [c, net('gach', duongQua([[60, 160], [60 + w, 160]], 7), c.batDau + c.thoiLuong, 0.4, { mau: 'nhan' })];
    }
    return { chu: chu, net: net, tieuDe: tieuDe, gh: gh };
  }

  function khoiDong(du) {
    var goc = document.getElementById('khung');
    var loai = root.THI_CANH[du.loai];
    var muc = loai.muc(du);
    var svg = document.createElementNS(NS, 'svg');
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
      } else {
        el = document.createElement('div');
        el.className = 'chu ' + m.can + ' ' + m.mau;
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
    if (loai.dung) { loai.dung(goc, du); }
    function dat(t) {
      ds.forEach(function (o) {
        var p = tienDo(t, o.m.batDau, o.m.thoiLuong);
        if (o.m.kieu === 'net') {
          o.el.style.strokeDashoffset = String(1 - p);
          o.el.style.opacity = p > 0 ? '1' : '0';
        } else if (!o.m.dong) {
          o.el.innerHTML = catDanhDau(o.m.chu, Math.round(p * demKyTu(o.m.chu)));
        }
      });
      if (loai.capNhat) { loai.capNhat(goc, du, t); }
    }
    root.datThoiDiem = dat;
    root.THI_VIDEO.thoiDiemCuoi = function () {
      return muc.reduce(function (cao, m) { return Math.max(cao, m.batDau + m.thoiLuong); }, 0) + 0.3;
    };
    root.THI_VIDEO.kiemTran = function () {
      dat(1e6);
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
    dat(0);
    root.THI_VIDEO.san = true;
  }

  root.THI_CANH = root.THI_CANH || {};
  root.THI_VIDEO = {
    DAN_DAU: DAN_DAU, kep: kep, tienDo: tienDo, thoat: thoat, demKyTu: demKyTu, catDanhDau: catDanhDau,
    thoiGianViet: thoiGianViet, duongQua: duongQua, hopQua: hopQua, vongTron: vongTron, muiTen: muiTen,
    tao: tao, khoiDong: khoiDong, san: false
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
