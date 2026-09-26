(function (root) {
  'use strict';

  var TOC_DO_VIET = 20;
  var LAU_BANG = 0.5; // mặc định khi chạy ngoài trang (Node); trang lấy du.giayLauBang từ lich.LAU_BANG của Python
  var HINH_TOI_THIEU = 1.2;
  var CO_TAY = 0.85;
  var NS = 'http://www.w3.org/2000/svg';
  var DANH_DAU = /\*\*(.+?)\*\*|~([^~]+)~|\^([^\^]+)\^/g;
  // Cùng ngữ pháp với kiem.py (_CUM_RE, _SO_DUNG). `____` (ô trống) và ` == ` có khoảng trắng hai bên là chữ thường.
  var CUM = /(?<!=)==(?![=\s])(.+?)(?<![=\s])==(?!=)|\(\((.+?)\)\)|(?<!_)__(?![_\s])(.+?)(?<![_\s])__(?!_)/g;
  var SO = /\{\{(-?\d+(?:\.\d+)?)\}\}/g;

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
  // Số chạy `{{1500.5}}`: hiện kiểu Việt (dấu phẩy thập phân), giữ số chữ số thập phân của số gốc.
  function soCuoi(goc) {
    var chuSo = (goc.split('.')[1] || '').length;
    var tron = Number(Number(goc).toFixed(chuSo));
    return (tron === 0 ? 0 : tron).toFixed(chuSo).replace('.', ',');
  }
  // Đoạn chữ trong một phần: đoạn thường {the, chu} hoặc đoạn số {the, chu, so: k, gtri, chuSo}.
  function tachDoan(chu, dem) {
    var kq = [];
    tach(chu).forEach(function (o) {
      var vt = 0;
      var m;
      SO.lastIndex = 0;
      while ((m = SO.exec(o.chu))) {
        if (m.index > vt) { kq.push({ the: o.the, chu: o.chu.slice(vt, m.index) }); }
        kq.push({ the: o.the, chu: soCuoi(m[1]), so: dem.so++, gtri: Number(m[1]), chuSo: (m[1].split('.')[1] || '').length });
        vt = m.index + m[0].length;
      }
      if (vt < o.chu.length) { kq.push({ the: o.the, chu: o.chu.slice(vt) }); }
    });
    return kq;
  }
  // Chữ -> các phần {cum: null | {kieu, k}, doan: [...]}. Cụm nhấn `==x==` tô, `((x))` khoanh, `__x__` gạch.
  // `khongCum` (biểu thức của cong-thuc): không tách cụm, `((` và `__` giữ nguyên là chữ.
  function phanTich(chu, khongCum) {
    chu = String(chu);
    if (khongCum) { return chu ? [{ cum: null, doan: tachDoan(chu, { so: 0, cum: 0 }) }] : []; }
    var phan = [];
    var dem = { so: 0, cum: 0 };
    var vt = 0;
    var m;
    CUM.lastIndex = 0;
    while ((m = CUM.exec(chu))) {
      if (m.index > vt) { phan.push({ cum: null, doan: tachDoan(chu.slice(vt, m.index), dem) }); }
      var kieu = m[1] !== undefined ? 'to' : (m[2] !== undefined ? 'khoanh' : 'gach');
      phan.push({ cum: { kieu: kieu, k: dem.cum++ }, doan: tachDoan(m[1] !== undefined ? m[1] : (m[2] !== undefined ? m[2] : m[3]), dem) });
      vt = m.index + m[0].length;
    }
    if (vt < chu.length) { phan.push({ cum: null, doan: tachDoan(chu.slice(vt), dem) }); }
    return phan;
  }
  function demPhan(phan) {
    return phan.reduce(function (n, p) { return p.doan.reduce(function (s, o) { return s + o.chu.length; }, n); }, 0);
  }
  function demKyTu(chu, khongCum) { return demPhan(phanTich(chu, khongCum)); }
  // Bề rộng ước lượng theo số ký tự: đệm hai đầu cụm khoanh (~1,2 ký tự) và lề cụm khi chu-dong (~0,5 ký tự) cũng chiếm chỗ.
  function demRong(chu, chuDong) {
    var phan = phanTich(chu);
    return phan.reduce(function (n, p) {
      return n + (p.cum ? (p.cum.kieu === 'khoanh' ? 1.2 : 0) + (chuDong ? 0.5 : 0) : 0);
    }, demPhan(phan));
  }
  // Vị trí (theo chữ hiển thị) và giá trị của từng số chạy.
  function viTriSo(chu, khongCum) {
    var kq = [];
    var vt = 0;
    phanTich(chu, khongCum).forEach(function (p) {
      p.doan.forEach(function (o) {
        if (o.so !== undefined) { kq.push({ k: o.so, viTri: vt, dai: o.chu.length, gtri: o.gtri, chuSo: o.chuSo, chu: o.chu }); }
        vt += o.chu.length;
      });
    });
    return kq;
  }
  function htmlSo(o, hien, so) {
    var mo = '<span class="so" data-so="' + o.so + '">';
    if (!hien) { return mo + '<span class="an">' + thoat(o.chu) + '</span></span>'; }
    var chay = so && so[o.so] !== undefined && so[o.so] !== null ? String(so[o.so]) : o.chu;
    if (chay === o.chu) { return mo + thoat(o.chu) + '</span>'; }
    // Chữ số cuối giữ chỗ (ẩn) để bề rộng không đổi khi số đang chạy.
    return mo + '<span class="so-cuoi">' + thoat(o.chu) + '</span><span class="so-chay">' + thoat(chay) + '</span></span>';
  }
  function kieuNay(v) {
    if (v.a === 1 && v.s === 1 && v.y === 0) { return ''; }
    return ' style="opacity:' + lam3(v.a) + ';transform:translateY(' + lam3(v.y) + 'px) scale(' + lam3(v.s) + ')"';
  }
  function lam3(x) { return Math.round(x * 1000) / 1000; }
  // Chế độ nảy: mỗi ký tự một span.nay; ký tự liền nhau (không cách) gói trong span.tu để từ không bị ngắt dòng.
  function htmlNay(chu, batDauI, nay) {
    var kq = '';
    var tu = '';
    for (var j = 0; j < chu.length; j++) {
      var c = chu[j];
      if (/\s/.test(c)) {
        if (tu) { kq += '<span class="tu">' + tu + '</span>'; tu = ''; }
        kq += c;
      } else {
        tu += '<span class="nay"' + kieuNay(nay(batDauI + j)) + '>' + thoat(c) + '</span>';
      }
    }
    return tu ? kq + '<span class="tu">' + tu + '</span>' : kq;
  }
  // Hiện n ký tự đầu; khi 0 < n < tổng, chèn span.ngoi rỗng ngay sau ký tự thứ n (điểm ngòi bút).
  // Dấu đánh dấu không đếm. Cụm bọc span.cum, số bọc span.so; phần chưa viết nằm trong span.an để bố cục không nhảy.
  // `so`: chuỗi đang hiện của từng số chạy (bỏ trống là giá trị cuối). `nay(i)`: {s, y, a} của ký tự i (chế độ nảy).
  function catDanhDau(chu, n, so, nay, khongCum) {
    var phan = phanTich(chu, khongCum);
    var con = nay ? Infinity : n;
    var giua = !nay && n > 0 && n < demPhan(phan);
    var i = 0;
    var html = '';
    phan.forEach(function (p) {
      var trongCum = '';
      p.doan.forEach(function (o) {
        var trong;
        if (nay) {
          trong = o.so !== undefined
            ? '<span class="tu"><span class="nay"' + kieuNay(nay(i)) + '>' + htmlSo(o, true, so) + '</span></span>'
            : htmlNay(o.chu, i, nay);
          i += o.chu.length;
        } else if (o.so !== undefined) {
          var hienSo = con > 0;
          con -= Math.min(con, o.chu.length);
          trong = htmlSo(o, hienSo, so) + (giua && hienSo && con === 0 ? '<span class="ngoi"></span>' : '');
          if (hienSo && con === 0) { giua = false; }
        } else {
          var hien = con > 0 ? o.chu.slice(0, con) : '';
          var an = o.chu.slice(hien.length);
          con -= hien.length;
          var ngoi = giua && hien && con === 0 ? '<span class="ngoi"></span>' : '';
          if (ngoi) { giua = false; }
          trong = thoat(hien) + ngoi + (an ? '<span class="an">' + thoat(an) + '</span>' : '');
        }
        trongCum += o.the ? '<' + o.the + '>' + trong + '</' + o.the + '>' : trong;
      });
      html += p.cum ? '<span class="cum ' + p.cum.kieu + '" data-cum="' + p.cum.k + '">' + trongCum + '</span>' : trongCum;
    });
    return html;
  }
  function thoiGianViet(chu, khongCum) { return Math.max(0.5, demKyTu(chu, khongCum) / TOC_DO_VIET); }

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
  function giayLau(du) { return typeof du.giayLauBang === 'number' ? du.giayLauBang : LAU_BANG; }
  function tao(du) {
    var gh = du.thoiLuong;
    var sau = du.co && du.co.lauBang ? giayLau(du) + 0.05 : 0;
    function dau(batDau, lui) { return Math.min(Math.max(batDau, sau), gh - lui); }
    function chu(id, noiDung, x, y, rong, cao, co, batDau, tuy) {
      batDau = dau(batDau, 0.6);
      // Chữ nảy (tiêu đề khi chu-dong): nảy từng ký tự thay cho bút viết; không có bàn tay.
      var nay = !!(tuy && tuy.nay);
      var dai = Math.max(0.3, Math.min(nay ? root.THI_DONG.thoiGianNay(demKyTu(noiDung)) : thoiGianViet(noiDung, tuy && tuy.khongCum), gh - 0.2 - batDau));
      if (nay) { tuy = gan({ tay: false }, tuy); }
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
    // `tuy.viTriNguon`: `duoi` (dưới khung, trong ô `tuy.oNguon`), `canh` (bên phải khung), bỏ trống là trong khung.
    function anh(id, a, x, y, rong, cao, batDau, tuy) {
      batDau = dau(batDau, 0.6);
      var k = vuaKhung(a.rong, a.cao, x, y, rong, cao);
      return gan({ id: id, kieu: 'anh', dataUrl: a.dataUrl, nguon: a.nguon, x: k.x, y: k.y, rong: k.rong, cao: k.cao,
        batDau: batDau, thoiLuong: Math.min(0.4, gh - 0.2 - batDau) }, tuy);
    }
    function tieuDe(noiDung, batDau, rong) {
      rong = rong || 1160;
      var co = rong < 1160 && demKyTu(noiDung) > 40 ? 32 : 40;
      var c = chu('tieu-de', noiDung, 60, 30, rong, 116, co, batDau, { mau: 'nhan', day: true });
      var w = Math.min(rong, Math.max(240, demKyTu(noiDung) * co * 0.6));
      return [c, net('gach', duongQua([[60, 160], [60 + w, 160]], 7), c.batDau + c.thoiLuong, 0.4, { mau: 'nhan', quay: false })];
    }
    // Cột hình bên phải (ô x 900, y 200, rộng 320, cao 380) cho cảnh có `hinh` hoặc `anh`.
    // Ảnh cao 350 để hai dòng nguồn dưới khung (trải hết bề rộng cột) vẫn nằm trên y = 620.
    var coCot = !!(du.hinh || du.anh);
    function cot() {
      var batDau = du.moc && du.moc.length ? du.moc[0] : 1.0;
      if (du.hinh) { return [hinh('hinh', du.hinh, 910, 240, 300, batDau)]; }
      if (du.anh) { return [anh('anh', du.anh, 900, 200, 320, 350, batDau, { viTriNguon: 'duoi', oNguon: { x: 900, rong: 320 } })]; }
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
    var lau = giayLau(du);
    var nen = null;
    if (co.lauBang && du.nenTruoc) {
      nen = document.createElement('img');
      nen.id = 'nen-truoc';
      nen.src = du.nenTruoc;
      khung.appendChild(nen);
    }
    if (co.chuDong) { khung.className = 'chu-dong'; }
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
    // Chữ: số chạy và cụm nhấn. Thời điểm tính trước từ mục và du.tu (thuần); hộp cụm đo trong doHop.
    var D = root.THI_DONG;
    var N = root.THI_NHAN;
    var hatNhan = 0;
    ds.forEach(function (o) {
      var m = o.m;
      if (m.kieu !== 'chu' || m.dong) { return; }
      var tong = demKyTu(m.chu, m.khongCum);
      var keo = m.nay ? Math.max(1, D.thoiGianNay(tong) / m.thoiLuong) : 1;
      // Lúc ký tự thứ i hiện ra (viết tay: khi round(p * tổng) > i; nảy: lúc ký tự i bắt đầu nảy).
      var luc = function (i) {
        return m.nay ? m.batDau + D.LECH_NAY * i / keo : m.batDau + m.thoiLuong * (i + 0.5) / Math.max(1, tong);
      };
      o.tong = tong;
      o.keo = keo;
      o.so = viTriSo(m.chu, m.khongCum).map(function (s) {
        var bd = luc(s.viTri);
        return { gtri: s.gtri, chuSo: s.chuSo, batDau: bd, dai: Math.min(0.8, Math.max(0, gh - 0.2 - bd)) };
      });
      o.cum = (m.khongCum ? [] : N.tachCum(m.chu)).map(function (c) {
        var cuoi = c.viTri + Math.max(1, c.dai) - 1;
        var xong = m.nay ? m.batDau + (D.LECH_NAY * cuoi + D.NAY) / keo : luc(cuoi);
        var lich = N.lichNo(N.thoiDiemNhan(c, du.tu || [], m.batDau, xong), xong, gh);
        c.no = lich.no;
        c.daiNo = lich.dai;
        c.hat = ++hatNhan;
        if (c.kieu !== 'to') {
          c.net = document.createElementNS(NS, 'path');
          c.net.setAttribute('pathLength', '1');
          c.net.setAttribute('class', 'nhan-net ' + c.kieu);
          c.net.setAttribute('data-nhan', m.id + '-' + c.k);
          c.net.style.strokeDasharray = '1';
          svg.appendChild(c.net);
        }
        return c;
      });
    });
    function datChu(o, t) {
      var m = o.m;
      var so = o.so.map(function (s) { return D.soChay(t, s.batDau, s.dai, s.gtri, s.chuSo); });
      var html;
      if (m.nay) {
        var tt = m.batDau + (t - m.batDau) * o.keo;
        html = catDanhDau(m.chu, o.tong, so, function (i) { return D.nayChu(i, o.tong, tt, m.batDau); });
      } else {
        html = catDanhDau(m.chu, Math.round(tienDo(t, m.batDau, m.thoiLuong) * o.tong), so, null, m.khongCum);
      }
      o.el.innerHTML = m.day ? '<span class="trong">' + html + '</span>' : html;
      o.cum.forEach(function (c) {
        var p = tienDo(t, c.no, c.daiNo);
        var pe = D.easeInOut(p);
        var sp = o.el.querySelector('[data-cum="' + c.k + '"]');
        if (c.kieu === 'to') {
          sp.style.backgroundSize = lam3(100 * pe) + '% 100%';
        } else {
          c.net.style.strokeDashoffset = String(1 - pe);
          c.net.style.opacity = pe > 0 ? '1' : '0';
        }
        // Cụm một dòng nảy nhẹ 1 → 1,12 → 1 (cần inline-block; cụm nhiều dòng giữ nguyên để không đổi ngắt dòng).
        // Độ phình giới hạn theo chỗ trống hai bên (đo trong doHop) để không chạm chữ bên cạnh.
        if (co.chuDong && c.motDong) {
          sp.style.display = 'inline-block';
          sp.style.transform = p > 0 && p < 1 ? 'scale(' + lam3(1 + c.nay * Math.sin(Math.PI * p)) + ')' : 'none';
        }
      });
    }
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
          datChu(o, t);
        }
      });
      if (loai.capNhat) { loai.capNhat(goc, du, t); }
    }

    // Chỗ trống (px) giữa cụm `sp` và chữ gần nhất cùng dòng bên trái/phải (hoặc mép ô chữ).
    function choTrong(el, sp) {
      var r = sp.getBoundingClientRect();
      var eb = el.getBoundingClientRect();
      var trai = r.left - eb.left;
      var phai = eb.right - r.right;
      var di = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
      var rg = document.createRange();
      var n;
      while ((n = di.nextNode())) {
        if (sp.contains(n)) { continue; }
        for (var i = 0; i < n.data.length; i++) {
          if (/\s/.test(n.data[i])) { continue; }
          rg.setStart(n, i);
          rg.setEnd(n, i + 1);
          var c = rg.getBoundingClientRect();
          if (c.width <= 0 || c.bottom <= r.top + 2 || c.top >= r.bottom - 2) { continue; }
          if (c.right <= r.left + 0.5) { trai = Math.min(trai, r.left - c.right); }
          else if (c.left >= r.right - 0.5) { phai = Math.min(phai, c.left - r.right); }
        }
      }
      return Math.min(trai, phai);
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
      // Hộp cụm nhấn (từng dòng) ở trạng thái cuối: vòng khoanh, nét gạch, và cụm nào nằm gọn một dòng.
      ds.forEach(function (o) {
        (o.cum || []).forEach(function (c) {
          var sp = o.el.querySelector('[data-cum="' + c.k + '"]');
          var cac = Array.prototype.map.call(sp.getClientRects(), function (r) {
            return { x: r.left - k.left, y: r.top - k.top, w: r.width, h: r.height };
          }).filter(function (r) { return r.w > 0; });
          c.motDong = cac.length === 1;
          c.nay = c.motDong && co.chuDong ? Math.max(0, Math.min(0.12, 2 * (choTrong(o.el, sp) - 1) / Math.max(1, cac[0].w))) : 0;
          if (!c.net || !cac.length) { return; }
          if (c.kieu === 'gach') {
            c.net.setAttribute('d', N.duongGach(cac, c.hat));
          } else {
            // Một vòng mỗi dòng; cao hơn dòng chữ tối đa 3 px, không chạm dòng trên/dưới (khoảng giữa hai dòng − 1 px).
            var dong = parseFloat(getComputedStyle(o.el).lineHeight) || 0;
            var day = Math.min(3, Math.min.apply(null, cac.map(function (r) { return dong - r.h; })) - 1);
            c.net.setAttribute('d', N.duongKhoanh(cac, c.hat, day));
          }
        });
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
      if (t > lau) { nen.style.display = 'none'; return; }
      nen.style.display = 'block';
      nen.style.clipPath = 'inset(0 0 0 ' + kep(-120 + 1400 * t / lau, 0, 1280) + 'px)';
    }

    function dat(t, noiBo) {
      if (!noiBo && !hop) { hop = doHop(); }
      datMuc(t);
      datNen(t);
      var cam = noiBo ? { z: 1, tx: 0, ty: 0 } : Q.tinh(muc, hop, t, gh, { mayQuay: co.mayQuay === true, day: thiNghiem });
      goc.style.transform = matTran(cam);
      if (!tay) { return; }
      var v = noiBo ? { hien: false } : T.viTri(muc, t, ngoiCua(cam), T.NGHI, { lauBang: !!co.lauBang, giayLau: lau, gh: gh });
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
      // Dòng nguồn ảnh: nằm trong khung hình và trên vạch phụ đề (y = 620).
      var ng = goc.querySelector('.anh .nguon');
      if (ng) {
        var rn = ng.getBoundingClientRect();
        if (rn.left < -1 || rn.top < -1 || rn.right > 1281 || rn.bottom > 621) { loi.push('nguon'); }
      }
      // Vòng khoanh và nét gạch của cụm nhấn: trong khung hình và trên vạch phụ đề.
      var cacNet = svg.querySelectorAll('path.nhan-net');
      for (var j = 0; j < cacNet.length; j++) {
        var b = cacNet[j].getBBox();
        if (b.x < -1 || b.y < -1 || b.x + b.width > 1281 || b.y + b.height > 621) { loi.push('nhan-' + cacNet[j].getAttribute('data-nhan')); }
      }
      return loi;
    };
    dat(0, true);
    root.THI_VIDEO.san = true;
  }

  root.THI_CANH = root.THI_CANH || {};
  root.THI_VIDEO = {
    LAU_BANG: LAU_BANG,
    kep: kep, tienDo: tienDo, thoat: thoat, demKyTu: demKyTu, catDanhDau: catDanhDau, phanTich: phanTich, demRong: demRong, viTriSo: viTriSo,
    thoiGianViet: thoiGianViet, duongQua: duongQua, hopQua: hopQua, vongTron: vongTron, muiTen: muiTen,
    rng: rng, vuaKhung: vuaKhung, tao: tao, khoiDong: khoiDong, san: false
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
