(function (root) {
  'use strict';
  var V = root.THI_VIDEO;
  root.THI_CANH['cong-thuc'] = {
    muc: function (du) {
      var B = V.tao(du);
      var t = du.truong;
      var kq = [];
      kq.push(B.net('khung', V.hopQua(100, 190, 1080, 150, 21), 0.1, 0.7, {}));
      var bt = B.chu('bieu-thuc', t['bieu-thuc'][0], 120, 215, 1040, 110, 40, 0.9, { can: 'giua', mau: 'nhan' });
      kq.push(bt);
      (t['giai-thich'] || []).forEach(function (g, k) {
        var bd = Math.max(du.moc[k], bt.batDau + bt.thoiLuong + 0.3);
        kq.push(B.chu('giai-thich-' + k, g, 110, 370 + k * 84, 1060, 76, 28, bd, {}));
      });
      return kq;
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
