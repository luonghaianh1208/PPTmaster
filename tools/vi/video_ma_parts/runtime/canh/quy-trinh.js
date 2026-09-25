(function (root) {
  'use strict';
  var V = root.THI_VIDEO;
  root.THI_CANH['quy-trinh'] = {
    muc: function (du) {
      var B = V.tao(du);
      var t = du.truong;
      var kq = B.tieuDe(t['tieu-de'][0], 0.2);
      var n = t.buoc.length;
      var w = (1160 - (n - 1) * 70) / n;
      t.buoc.forEach(function (b, k) {
        var x = 60 + k * (w + 70);
        kq.push(B.net('hop-' + k, V.hopQua(x, 230, w, 300, 20 + k), Math.max(0.1, du.moc[k] - 0.4), 0.5, {}));
        kq.push(B.chu('buoc-' + k, b, x + 14, 250, w - 28, 260, 24, du.moc[k], {}));
        if (k > 0) {
          kq.push(B.net('mui-' + k, V.muiTen(x - 64, 380, x - 6, 380, 30 + k), Math.max(0.1, du.moc[k] - 0.6), 0.3, { mau: 'nhan' }));
        }
      });
      return kq;
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
