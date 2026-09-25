(function (root) {
  'use strict';
  var V = root.THI_VIDEO;
  root.THI_CANH['khai-niem'] = {
    muc: function (du) {
      var B = V.tao(du);
      var t = du.truong;
      var kq = [B.net('khung', V.hopQua(60, 190, 1160, 450, 11), 0.1, 0.9, {})];
      var tn = B.chu('thuat-ngu', t['thuat-ngu'][0], 100, 200, 1080, 125, 40, 1.0, { mau: 'nhan', day: true });
      kq.push(tn);
      kq.push(B.net('gach', V.duongQua([[100, 335], [700, 335]], 4), tn.batDau + tn.thoiLuong, 0.4, { mau: 'nhan' }));
      kq.push(B.chu('dinh-nghia', t['dinh-nghia'][0], 100, 360, 1080, 260, 32, tn.batDau + tn.thoiLuong + 0.5, {}));
      return kq;
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
