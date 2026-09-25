(function (root) {
  'use strict';
  var V = root.THI_VIDEO;
  root.THI_CANH['so-sanh'] = {
    muc: function (du) {
      var B = V.tao(du);
      var t = du.truong;
      var nt = t['y-trai'].length;
      var kq = B.tieuDe(t['tieu-de'][0], 0.2);
      kq.push(B.net('giua', V.duongQua([[640, 190], [640, 660]], 9), 0.3, 0.6, {}));
      kq.push(B.chu('trai', t.trai[0], 60, 190, 540, 70, 34, 0.3, { mau: 'nhan' }));
      kq.push(B.chu('phai', t.phai[0], 680, 190, 540, 70, 34, Math.max(0.3, du.moc[nt] - 0.9), { mau: 'nhan' }));
      t['y-trai'].forEach(function (y, k) { kq.push(B.chu('y-trai-' + k, y, 80, 285 + k * 92, 520, 88, 28, du.moc[k], {})); });
      t['y-phai'].forEach(function (y, k) { kq.push(B.chu('y-phai-' + k, y, 700, 285 + k * 92, 520, 88, 28, du.moc[nt + k], {})); });
      return kq;
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
