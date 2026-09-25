(function (root) {
  'use strict';

  // Vẽ dần một biểu tượng nét (từng phần tử một) và ảnh thật có Ken Burns. Mọi thứ là hàm của t.
  var NS = 'http://www.w3.org/2000/svg';
  var BO_QUA = { style: 1, 'class': 1, stroke: 1, 'stroke-width': 1, fill: 1, 'stroke-dasharray': 1, 'stroke-dashoffset': 1 };
  var PHONG = 0.12;
  var LUOT = 0.03;
  var HIEN = 0.4;

  function kep(x, a, b) { return x < a ? a : (x > b ? b : x); }
  function khung(vb) {
    var so = String(vb || '0 0 24 24').trim().split(/[\s,]+/).map(Number);
    return { x: so[0] || 0, y: so[1] || 0, w: so[2] || 24, h: so[3] || 24 };
  }

  function taoHinh(svg, m) {
    var vb = khung(m.viewBox);
    var s = m.kich / Math.max(vb.w, vb.h);
    var g = document.createElementNS(NS, 'g');
    g.setAttribute('class', 'hinh ' + m.mau);
    g.setAttribute('data-hinh', m.id);
    g.setAttribute('transform', 'translate(' + m.x + ' ' + m.y + ') scale(' + s + ') translate(' + (-vb.x) + ' ' + (-vb.y) + ')');
    var cac = m.phanTu.map(function (p) {
      var el = document.createElementNS(NS, p.the);
      Object.keys(p.thuocTinh || {}).forEach(function (k) {
        // Bỏ thuộc tính có không gian tên (`{ns}tên` từ ElementTree) và mọi thuộc tính sự kiện `on…`.
        if (!BO_QUA[k] && k.charAt(0) !== '{' && !/^on/i.test(k)) { el.setAttribute(k, p.thuocTinh[k]); }
      });
      el.setAttribute('pathLength', '1');
      el.style.strokeWidth = String(Math.max(3, m.kich / 60) / s);
      el.style.strokeDasharray = '1';
      g.appendChild(el);
      return el;
    });
    svg.appendChild(g);
    return { g: g, cac: cac, vb: vb, s: s };
  }

  function datHinh(o, p) {
    var n = o.cac.length;
    o.cac.forEach(function (el, i) {
      var pi = kep(p * n - i, 0, 1);
      el.style.strokeDashoffset = String(1 - pi);
      el.style.opacity = pi > 0 ? '1' : '0';
    });
  }

  // Ngòi bút của hình tại tiến độ p, toạ độ lớp bảng.
  function ngoiHinh(o, m, p) {
    var n = o.cac.length;
    var i = Math.min(n - 1, Math.floor(kep(p, 0, 1) * n));
    var pi = kep(p * n - i, 0, 1);
    var el = o.cac[i];
    var d = el.getPointAtLength(pi * el.getTotalLength());
    return { x: m.x + (d.x - o.vb.x) * o.s, y: m.y + (d.y - o.vb.y) * o.s };
  }

  // Ken Burns: phóng 1,0 → 1,12 và lướt tối đa 3% theo hướng xác định bằng hạt giống.
  function kenBurns(hat, q) {
    var goc = root.THI_VIDEO.rng(hat * 7919 + 17)() * 2 * Math.PI;
    q = kep(q, 0, 1);
    return { s: 1 + PHONG * q, dx: Math.cos(goc) * LUOT * q, dy: Math.sin(goc) * LUOT * q };
  }

  function taoAnh(goc, m) {
    var V = root.THI_VIDEO;
    var el = document.createElement('div');
    el.className = 'anh';
    el.setAttribute('data-anh', m.id);
    el.style.left = m.x + 'px';
    el.style.top = m.y + 'px';
    el.style.width = m.rong + 'px';
    el.style.height = m.cao + 'px';
    var cua = document.createElement('div');
    cua.className = 'cua-anh';
    var img = document.createElement('img');
    img.src = m.dataUrl;
    img.alt = '';
    cua.appendChild(img);
    el.appendChild(cua);
    var vien = document.createElementNS(NS, 'svg');
    vien.setAttribute('class', 'vien-anh');
    vien.setAttribute('width', String(m.rong + 20));
    vien.setAttribute('height', String(m.cao + 20));
    vien.setAttribute('viewBox', '-10 -10 ' + (m.rong + 20) + ' ' + (m.cao + 20));
    var net = document.createElementNS(NS, 'path');
    net.setAttribute('d', V.hopQua(-5, -5, m.rong + 10, m.cao + 10, 13));
    net.setAttribute('pathLength', '1');
    net.setAttribute('class', 'net');
    net.style.strokeWidth = '4';
    net.style.strokeDasharray = '1';
    vien.appendChild(net);
    el.appendChild(vien);
    var nguon = document.createElement('div');
    // Chỗ dòng nguồn do bố cục gọi quyết định: `duoi` dưới khung, trải trong ô `oNguon` (cột phải);
    // `canh` bên phải khung (tiêu đề có ảnh); mặc định trong góc dưới phải của khung (cảnh `anh`).
    var vt = m.viTriNguon === 'duoi' || m.viTriNguon === 'canh' ? m.viTriNguon : 'trong';
    nguon.className = 'nguon ' + vt;
    if (vt === 'duoi' && m.oNguon) {
      nguon.style.left = (m.oNguon.x - m.x) + 'px';
      nguon.style.width = m.oNguon.rong + 'px';
    }
    nguon.textContent = m.nguon;
    el.appendChild(nguon);
    goc.appendChild(el);
    return { el: el, img: img, net: net };
  }

  function datAnh(o, m, p, t, gh, hat) {
    o.el.style.opacity = String(kep((t - m.batDau) / HIEN, 0, 1));
    o.net.style.strokeDashoffset = String(1 - p);
    var q = gh > m.batDau ? (t - m.batDau) / (gh - m.batDau) : 1;
    var k = kenBurns(hat || 1, q);
    o.img.style.transform = 'translate(' + (k.dx * m.rong) + 'px,' + (k.dy * m.cao) + 'px) scale(' + k.s + ')';
  }

  root.THI_HINH = { taoHinh: taoHinh, datHinh: datHinh, ngoiHinh: ngoiHinh, kenBurns: kenBurns, taoAnh: taoAnh, datAnh: datAnh };
})(typeof globalThis !== 'undefined' ? globalThis : this);
