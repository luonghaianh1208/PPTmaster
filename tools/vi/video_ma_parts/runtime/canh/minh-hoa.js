(function (root) {
  'use strict';
  var V = root.THI_VIDEO;
  root.THI_CANH['minh-hoa'] = {
    muc: function (du) {
      var B = V.tao(du);
      var t = du.truong;
      var kq = B.tieuDe(t['tieu-de'][0], 0.2);
      var n = du.hinhs.length;
      var o = 1160 / n;
      du.hinhs.forEach(function (h, k) {
        var giua = 60 + (k + 0.5) * o;
        var hk = B.hinh('hinh-' + k, h, giua - 120, 230, 240, du.moc.length > k ? du.moc[k] : 1.0 + k);
        kq.push(hk);
        kq.push(B.chu('nhan-' + k, h.nhan, giua - (o - 40) / 2, 480, o - 40, 80, 28, hk.batDau + hk.thoiLuong, { can: 'giua' }));
      });
      return kq;
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
