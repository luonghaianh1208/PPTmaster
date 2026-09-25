(function (root) {
  'use strict';
  var V = root.THI_VIDEO;
  root.THI_CANH['tieu-de'] = {
    muc: function (du) {
      var B = V.tao(du);
      var t = du.truong;
      var kq = [];
      if (!B.coCot) {
        var c = B.chu('chu', t.chu[0], 80, 60, 1120, 320, 60, 0.3, { can: 'giua', mau: 'nhan', day: true });
        kq.push(c);
        kq.push(B.net('gach', V.duongQua([[340, 400], [940, 400]], 3), c.batDau + c.thoiLuong, 0.4, { mau: 'nhan', quay: false }));
        if (t.phu) { kq.push(B.chu('phu', t.phu[0], 80, 430, 1120, 100, 34, c.batDau + c.thoiLuong + 0.4, { can: 'giua' })); }
        return kq;
      }
      // Có hình: hình 180×180 ở giữa phía trên (y 40–220), vẽ trước; tiêu đề dời xuống ngay dưới.
      // Dòng nguồn của ảnh nằm bên phải khung, không chen vào ô tiêu đề hai dòng.
      var h = du.hinh ? B.hinh('hinh', du.hinh, 550, 40, 180, 0.3) : B.anh('anh', du.anh, 550, 40, 180, 180, 0.3, { viTriNguon: 'canh' });
      kq.push(h);
      var co = V.demKyTu(t.chu[0]) <= 40 ? 56 : 44;
      var c2 = B.chu('chu', t.chu[0], 80, 225, 1120, 150, co, h.batDau + h.thoiLuong, { can: 'giua', mau: 'nhan', day: true });
      kq.push(c2);
      kq.push(B.net('gach', V.duongQua([[340, 395], [940, 395]], 3), c2.batDau + c2.thoiLuong, 0.4, { mau: 'nhan', quay: false }));
      if (t.phu) { kq.push(B.chu('phu', t.phu[0], 80, 420, 1120, 100, 34, c2.batDau + c2.thoiLuong + 0.4, { can: 'giua' })); }
      return kq;
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
