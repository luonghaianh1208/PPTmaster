(function (root) {
  'use strict';
  var V = root.THI_VIDEO;
  root.THI_CANH['y-tung-y'] = {
    muc: function (du) {
      var B = V.tao(du);
      var t = du.truong;
      var cot = B.coCot;
      var kq = B.tieuDe(t['tieu-de'][0], 0.2, cot ? 800 : 1160);
      var dai = Math.max.apply(null, t.y.map(function (y) { return V.demKyTu(y); }));
      var co = cot && dai > 45 ? 26 : 30;
      t.y.forEach(function (y, k) {
        var top = 186 + k * 74;
        kq.push(B.net('cham-' + k, V.vongTron(96, top + 22, 9), du.moc[k], 0.3, { mau: 'nhan' }));
        kq.push(B.chu('y-' + k, y, 124, top, cot ? 736 : 1100, 72, co, du.moc[k], {}));
      });
      return kq.concat(B.cot());
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
