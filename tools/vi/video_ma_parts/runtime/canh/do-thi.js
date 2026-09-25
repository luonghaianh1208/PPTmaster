(function (root) {
  'use strict';
  var V = root.THI_VIDEO;
  var X0 = 190, X1 = 1090, Y0 = 560, Y1 = 270;

  function soVN(v) { return String(Math.round(v * 1000) / 1000).replace('.', ','); }
  function tyLe(v, thap, cao, a, b) { return cao === thap ? (a + b) / 2 : a + (v - thap) * (b - a) / (cao - thap); }

  root.THI_CANH['do-thi'] = {
    muc: function (du) {
      var B = V.tao(du);
      var t = du.truong;
      var xs = du.diem.map(function (p) { return p[0]; });
      var ys = du.diem.map(function (p) { return p[1]; });
      var xMin = Math.min.apply(null, xs), xMax = Math.max.apply(null, xs);
      var yMin = Math.min.apply(null, ys), yMax = Math.max.apply(null, ys);
      var kq = B.tieuDe(t['tieu-de'][0], 0.2);
      kq.push(B.net('truc-x', V.duongQua([[140, 600], [1140, 600]], 1), 0.3, 0.6, {}));
      kq.push(B.net('truc-y', V.duongQua([[140, 600], [140, 230]], 2), 0.4, 0.6, {}));
      kq.push(B.chu('truc-ngang', t['truc-ngang'][0], 700, 650, 440, 40, 24, 1.0, { can: 'phai', mau: 'nhan' }));
      kq.push(B.chu('truc-doc', t['truc-doc'][0], 160, 186, 700, 40, 24, 1.0, { mau: 'nhan' }));
      kq.push(B.chu('x-min', soVN(xMin), X0 - 60, 606, 120, 30, 20, 1.0, { can: 'giua' }));
      kq.push(B.chu('x-max', soVN(xMax), X1 - 60, 606, 120, 30, 20, 1.0, { can: 'giua' }));
      kq.push(B.chu('y-min', soVN(yMin), 20, Y0 - 15, 110, 30, 20, 1.0, { can: 'phai' }));
      kq.push(B.chu('y-max', soVN(yMax), 20, Y1 - 15, 110, 30, 20, 1.0, { can: 'phai' }));
      var truoc = null;
      du.diem.forEach(function (p, k) {
        var px = tyLe(p[0], xMin, xMax, X0, X1);
        var py = tyLe(p[1], yMin, yMax, Y0, Y1);
        if (truoc) { kq.push(B.net('doan-' + k, V.duongQua([truoc, [px, py]], 40 + k), du.moc[k], 0.4, { mau: 'nhan' })); }
        kq.push(B.net('diem-' + k, V.vongTron(px, py, 9), du.moc[k], 0.3, { mau: 'do' }));
        truoc = [px, py];
      });
      return kq;
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
