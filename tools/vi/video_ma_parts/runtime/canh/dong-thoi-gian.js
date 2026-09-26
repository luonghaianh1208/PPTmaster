(function (root) {
  'use strict';
  var V = root.THI_VIDEO;
  // Trục ngang ở y = AY từ AX0 tới AX1 (mũi tên). Đến 4 mốc: nhãn trên trục, mô tả dưới trục, mỗi mốc một ô rộng
  // bằng khoảng giữa hai mốc. Từ 5 mốc: mốc chẵn cả nhãn và mô tả ở trên, mốc lẻ ở dưới (xen kẽ cao thấp), nên
  // mỗi ô rộng gần gấp đôi mà không chạm ô cùng phía.
  var AX0 = 60, AX1 = 1220, AY = 400;

  function tach(v) {
    var i = v.indexOf('|');
    return [v.slice(0, i).trim(), v.slice(i + 1).trim()];
  }

  root.THI_CANH['dong-thoi-gian'] = {
    muc: function (du) {
      var B = V.tao(du);
      var t = du.truong;
      var kq = B.tieuDe(t['tieu-de'][0], 0.2);
      kq.push(B.net('truc', V.muiTen(AX0, AY, AX1, AY, 9), 0.3, 0.6, { quay: false }));
      var n = t.moc.length;
      var o = (AX1 - 40 - AX0) / n;
      var xen = n > 4;
      var w = xen ? 2 * o - 24 : o - 20;
      t.moc.forEach(function (v, k) {
        var hai = tach(v);
        var x = AX0 + (k + 0.5) * o;
        var trai = Math.max(24, x - w / 2), phai = Math.min(1256, x + w / 2);
        var duoi = xen && k % 2 === 1;
        kq.push(B.net('cham-' + k, V.vongTron(Math.round(x * 10) / 10, AY, 10), du.moc[k], 0.3, { mau: 'do', am: 'ting' }));
        var nhan = B.chu('nhan-' + k, hai[0], trai, duoi ? AY + 26 : AY - 64, phai - trai, 38, 26, du.moc[k] + 0.25,
          { can: 'giua', mau: 'nhan' });
        kq.push(nhan);
        var moTa = xen && !duoi
          ? [AY - 204, { can: 'giua', day: true }]
          : [duoi ? AY + 68 : AY + 26, { can: 'giua' }];
        kq.push(B.chu('mo-ta-' + k, hai[1], trai, moTa[0], phai - trai, xen ? 132 : 150, xen ? 20 : 22,
          nhan.batDau + nhan.thoiLuong, moTa[1]));
      });
      return kq;
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
