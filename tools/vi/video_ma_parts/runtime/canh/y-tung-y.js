(function (root) {
  'use strict';
  var V = root.THI_VIDEO;
  root.THI_CANH['y-tung-y'] = {
    muc: function (du) {
      var B = V.tao(du);
      var t = du.truong;
      var kq = B.tieuDe(t['tieu-de'][0], 0.2);
      t.y.forEach(function (y, k) {
        var top = 186 + k * 74;
        kq.push(B.net('cham-' + k, V.vongTron(96, top + 22, 9), du.moc[k], 0.3, { mau: 'nhan' }));
        kq.push(B.chu('y-' + k, y, 124, top, 1100, 72, 30, du.moc[k], {}));
      });
      return kq;
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
