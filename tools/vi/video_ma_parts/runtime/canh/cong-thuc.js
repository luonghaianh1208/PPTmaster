(function (root) {
  'use strict';
  var V = root.THI_VIDEO;
  // `bieu-thuc` có ` | ` thì tách phần: phần k viết ở mốc câu k, các phần nối nhau trên cùng dòng (cách một khoảng
  // trắng); dòng giải thích k khi đó ở mốc câu (số phần + k).
  var TACH = ' | ';
  root.THI_CANH['cong-thuc'] = {
    muc: function (du) {
      var B = V.tao(du);
      var t = du.truong;
      var cot = B.coCot;
      var phan = t['bieu-thuc'][0].split(TACH);
      var soPhan = phan.length > 1 ? phan.length : 0;
      var bieuThuc = phan.join(' ');
      var kq = [];
      kq.push(B.net('khung', V.hopQua(100, 190, cot ? 760 : 1080, 150, 21), 0.1, 0.7, {}));
      var tuy = { can: 'giua', mau: 'nhan', khongCum: true };
      if (soPhan) {
        var tong = V.demKyTu(bieuThuc, true);
        var ky = phan.map(function (p, k) { return V.demKyTu(p, true) + (k < soPhan - 1 ? 1 : 0); });
        ky[soPhan - 1] = tong - ky.slice(0, -1).reduce(function (a, b) { return a + b; }, 0);
        tuy.phan = phan.map(function (_, k) { return { batDau: du.moc[k], ky: ky[k] }; });
      }
      var bt = B.chu('bieu-thuc', bieuThuc, 120, 215, cot ? 720 : 1040, 120, cot && V.demKyTu(bieuThuc, true) > 30 ? 30 : 40, 0.9, tuy);
      kq.push(bt);
      (t['giai-thich'] || []).forEach(function (g, k) {
        var bd = Math.max(du.moc[soPhan + k], bt.batDau + bt.thoiLuong + 0.3);
        kq.push(B.chu('giai-thich-' + k, g, 110, 356 + k * 68, cot ? 750 : 1060, 66, cot ? 24 : 28, bd, {}));
      });
      return kq.concat(B.cot());
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
