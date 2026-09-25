(function (root) {
  'use strict';
  var V = root.THI_VIDEO;
  root.THI_CANH['tieu-de'] = {
    muc: function (du) {
      var B = V.tao(du);
      var t = du.truong;
      var kq = [];
      var c = B.chu('chu', t.chu[0], 80, 60, 1120, 320, 60, 0.3, { can: 'giua', mau: 'nhan', day: true });
      kq.push(c);
      kq.push(B.net('gach', V.duongQua([[340, 400], [940, 400]], 3), c.batDau + c.thoiLuong, 0.4, { mau: 'nhan' }));
      if (t.phu) { kq.push(B.chu('phu', t.phu[0], 80, 430, 1120, 100, 34, c.batDau + c.thoiLuong + 0.4, { can: 'giua' })); }
      return kq;
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
