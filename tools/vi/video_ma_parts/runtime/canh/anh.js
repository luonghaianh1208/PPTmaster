(function (root) {
  'use strict';
  var V = root.THI_VIDEO;
  root.THI_CANH.anh = {
    muc: function (du) {
      var B = V.tao(du);
      var t = du.truong;
      var chuThich = t['chu-thich'][0];
      var n = V.demKyTu(chuThich);
      return [
        B.anh('anh', du.anh, 80, 70, 1120, 490, 0.2),
        B.chu('chu-thich', chuThich, 80, 575, 1120, 50, n <= 60 ? 30 : (n <= 75 ? 26 : 24), du.moc.length ? du.moc[0] : 1.0, { can: 'giua' })
      ];
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
