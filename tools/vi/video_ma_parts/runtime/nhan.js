(function (root) {
  'use strict';

  // Cụm nhấn: tách cụm, tìm lúc giọng đọc tới cụm, vẽ vòng khoanh / nét gạch. Hàm thuần, chạy được trong Node.
  var V = root.THI_VIDEO;
  var TREO = 0.3;           // không có trong lời: nổ 0,3 s sau khi cụm viết xong
  var NO = 0.45;            // hiệu ứng nổ trong 0,45 s

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

  // Lịch nổ: không bao giờ trước khi cụm viết xong (`xong`); gần cuối cảnh thì dời sớm lại, và nếu vẫn không đủ chỗ
  // thì rút ngắn hiệu ứng để xong trước gh − 0,2.
  function lichNo(thoiDiem, xong, gh) {
    var het = gh - 0.2;
    var no = Math.max(xong, Math.min(thoiDiem, het - NO));
    return { no: no, dai: Math.max(0, Math.min(NO, het - no)) };
  }

  function diem(x, y) { return lam(x) + ' ' + lam(y); }

  // Vòng elip vẽ tay quanh từng hộp dòng {x, y, w, h} của cụm (hộp đã gồm phần đệm ngang): rộng hơn hộp 4 px mỗi bên,
  // cao hơn `day` px (mặc định 3) trên và dưới; nét trôi vào trong (không ra ngoài) để chỗ vẽ chồng thấy hai nét.
  // Không kẹp vào khung: vượt khung thì kiemTran báo.
  function duongKhoanh(hop, hat, day) {
    if (!Array.isArray(hop)) { hop = [hop]; }
    day = typeof day === 'number' ? Math.max(0, day) : 3;
    return hop.map(function (b, i) {
      var r = V.rng(hat * 7919 + 11 + i * 101);
      var cx = b.x + b.w / 2;
      var cy = b.y + b.h / 2;
      var rx = b.w / 2 + 4;
      var ry = b.h / 2 + day;
      var goc0 = -2.5 + r() * 0.4;
      var quet = 2 * Math.PI + 0.5;
      var n = 48;
      var d = '';
      for (var j = 0; j <= n; j++) {
        var u = j / n;
        var g = goc0 + quet * u;
        var vao = 4 * u + r() * 0.5;
        d += (j ? ' L' : 'M') + diem(cx + (rx - vao) * Math.cos(g), cy + (ry - vao / 4) * Math.sin(g));
      }
      return d;
    }).join(' ');
  }

  // Nét gạch dưới từng dòng của cụm (hộp từng dòng), nối thành một đường để vẽ dần. Không kẹp vào khung.
  function duongGach(hop, hat) {
    return hop.map(function (b, i) {
      var y = b.y + b.h + 3;
      return V.duongQua([[b.x - 2, y], [b.x + b.w + 2, y]], hat * 31 + i);
    }).join(' ');
  }

  root.THI_NHAN = { NO: NO, TREO: TREO, khoa: khoa, tachCum: tachCum, thoiDiemNhan: thoiDiemNhan,
    lichNo: lichNo, duongKhoanh: duongKhoanh, duongGach: duongGach };
})(typeof globalThis !== 'undefined' ? globalThis : this);
