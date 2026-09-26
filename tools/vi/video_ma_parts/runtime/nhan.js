(function (root) {
  'use strict';

  // Cụm nhấn: tách cụm, tìm lúc giọng đọc tới cụm, vẽ vòng khoanh / nét gạch. Hàm thuần, chạy được trong Node.
  var V = root.THI_VIDEO;
  var TREO = 0.3;           // không có trong lời: nổ 0,3 s sau khi cụm viết xong
  var NO = 0.45;            // hiệu ứng nổ trong 0,45 s
  var BIEN = { trai: 4, phai: 1276, tren: 4, duoi: 616 }; // vòng khoanh, nét gạch nằm trong khung và trên vạch phụ đề

  function kep(x, a, b) { return x < a ? a : (x > b ? b : x); }
  function lam(x) { return Math.round(x * 10) / 10; }

  // Khoá so khớp: chữ thường, bỏ dấu câu, giữ dấu thanh (như `khoa` của lich.py).
  function khoa(s) {
    return String(s).normalize('NFC').toLowerCase().replace(/[^\p{L}\p{M}\p{N}_\s]/gu, '').replace(/\s+/g, ' ').trim();
  }

  function tachCum(chu) {
    var kq = [];
    var vt = 0;
    V.phanTich(chu).forEach(function (p) {
      var noi = p.doan.map(function (o) { return o.chu; }).join('');
      if (p.cum) { kq.push({ kieu: p.cum.kieu, noiDung: noi, viTri: vt, dai: noi.length, k: p.cum.k }); }
      vt += noi.length;
    });
    return kq;
  }

  // Lần đầu dãy khoá của cụm xuất hiện trong lời từ `batDauMuc` trở đi; không có thì `xongMuc + 0,3`.
  function thoiDiemNhan(cum, tu, batDauMuc, xongMuc) {
    var can = khoa(typeof cum === 'string' ? cum : cum.noiDung).split(' ').filter(Boolean);
    var ds = [];
    (tu || []).forEach(function (w) {
      khoa(w.khoa !== undefined && w.khoa !== null ? w.khoa : w.chu).split(' ').forEach(function (k) {
        if (k) { ds.push({ t: w.t, k: k }); }
      });
    });
    if (can.length) {
      for (var i = 0; i + can.length <= ds.length; i++) {
        if (ds[i].t < batDauMuc - 1e-9) { continue; }
        var khop = true;
        for (var j = 0; j < can.length && khop; j++) { khop = ds[i + j].k === can[j]; }
        if (khop) { return ds[i].t; }
      }
    }
    return xongMuc + TREO;
  }

  function diem(x, y) { return lam(kep(x, BIEN.trai, BIEN.phai)) + ' ' + lam(kep(y, BIEN.tren, BIEN.duoi)); }

  // Vòng elip vẽ tay quanh hộp {x, y, w, h}: hơi xoắn, vẽ quá một vòng một chút; kẹp trong khung, trên y 616.
  function duongKhoanh(b, hat) {
    var r = V.rng(hat * 7919 + 11);
    var cx = b.x + b.w / 2;
    var cy = b.y + b.h / 2;
    var rx = b.w / 2 * 1.06 + 14;
    var ry = b.h / 2 * 1.2 + 8;
    var goc0 = -2.5 + r() * 0.4;
    var quet = 2 * Math.PI + 0.5;
    var n = 48;
    var d = '';
    for (var j = 0; j <= n; j++) {
      var u = j / n;
      var g = goc0 + quet * u;
      // Xoắn nhẹ: đầu nét ngoài hơn cuối nét ~11 px để chỗ vẽ chồng thấy rõ hai nét, không thành vệt đậm.
      var lech = 7 - 11 * u + (r() - 0.5) * 1.5;
      d += (j ? ' L' : 'M') + diem(cx + (rx + lech) * Math.cos(g), cy + (ry + lech) * Math.sin(g));
    }
    return d;
  }

  // Nét gạch dưới từng dòng của cụm (hộp từng dòng), nối thành một đường để vẽ dần.
  function duongGach(hop, hat) {
    return hop.map(function (b, i) {
      var y = Math.min(b.y + b.h + 3, BIEN.duoi - 2);
      var x1 = kep(b.x - 2, BIEN.trai, BIEN.phai);
      var x2 = kep(b.x + b.w + 2, BIEN.trai, BIEN.phai);
      return V.duongQua([[x1, y], [x2, y]], hat * 31 + i);
    }).join(' ');
  }

  root.THI_NHAN = { NO: NO, TREO: TREO, BIEN: BIEN, khoa: khoa, tachCum: tachCum, thoiDiemNhan: thoiDiemNhan,
    duongKhoanh: duongKhoanh, duongGach: duongGach };
})(typeof globalThis !== 'undefined' ? globalThis : this);
