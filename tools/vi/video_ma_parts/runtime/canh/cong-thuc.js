(function (root) {
  'use strict';
  var V = root.THI_VIDEO;
  root.THI_CANH['cong-thuc'] = {
    muc: function (du) {
      var B = V.tao(du);
      var t = du.truong;
      var cot = B.coCot;
      var bieuThuc = t['bieu-thuc'][0];
      var kq = [];
      kq.push(B.net('khung', V.hopQua(100, 190, cot ? 760 : 1080, 150, 21), 0.1, 0.7, {}));
      var bt = B.chu('bieu-thuc', bieuThuc, 120, 215, cot ? 720 : 1040, 120, cot && V.demKyTu(bieuThuc) > 30 ? 30 : 40, 0.9, { can: 'giua', mau: 'nhan' });
      kq.push(bt);
      (t['giai-thich'] || []).forEach(function (g, k) {
        var bd = Math.max(du.moc[k], bt.batDau + bt.thoiLuong + 0.3);
        kq.push(B.chu('giai-thich-' + k, g, 110, 356 + k * 68, cot ? 750 : 1060, 66, cot ? 24 : 28, bd, {}));
      });
      return kq.concat(B.cot());
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
