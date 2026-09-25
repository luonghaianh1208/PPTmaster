(function (root) {
  'use strict';
  var V = root.THI_VIDEO;
  var RONG = 792, CAO = 462;

  function thamSoTai(du, t) {
    var K = root.THI_NGHIEM_KHUNG;
    var p = K.thamSoMacDinh(du.khaiBao);
    Object.keys(du.thamSo).forEach(function (ma) {
      var ds = du.thamSo[ma];
      var cuoi = ds[ds.length - 1];
      if (t < ds[0][0]) { return; }
      if (t >= cuoi[0]) { p[ma] = cuoi[1]; return; }
      for (var i = 1; i < ds.length; i++) {
        if (t < ds[i][0]) {
          var a = ds[i - 1], b = ds[i];
          p[ma] = a[1] + (b[1] - a[1]) * (t - a[0]) / (b[0] - a[0]);
          return;
        }
      }
    });
    return p;
  }

  function giaTri(du, t) {
    var p = thamSoTai(du, t);
    return { tham: p, dai: root.THI_NGHIEM_MO_HINH.tinh(p) };
  }

  function soThapPhan(buoc) {
    var s = String(buoc);
    var i = s.indexOf('.');
    return i < 0 ? 0 : Math.min(3, s.length - i - 1);
  }
  function tim(ds, ma) { return ds.filter(function (x) { return x.ma === ma; })[0]; }

  root.THI_CANH['thi-nghiem'] = {
    thamSoTai: thamSoTai,
    giaTri: giaTri,
    muc: function (du) {
      var B = V.tao(du);
      var kb = du.khaiBao;
      var kq = B.tieuDe(kb.ten, 0.2);
      kq.push(B.net('khung', V.hopQua(40, 190, 800, 470, 5), 0.1, 0.7, {}));
      kq.push(B.chu('nhan-tham-so', 'Thông số', 880, 190, 360, 40, 26, 0.4, { mau: 'nhan' }));
      Object.keys(du.thamSo).slice(0, 3).forEach(function (ma, k) {
        kq.push(B.chu('ts-' + k, '', 880, 236 + k * 60, 360, 56, 20, 0.4, { dong: true }));
      });
      kq.push(B.chu('nhan-do', 'Số đo', 880, 430, 360, 40, 26, 0.4, { mau: 'nhan' }));
      du.do.slice(0, 3).forEach(function (ma, k) {
        kq.push(B.chu('do-' + k, '', 880, 476 + k * 70, 360, 66, 22, 0.4, { dong: true, mau: 'do' }));
      });
      return kq;
    },
    dung: function (goc) {
      var c = document.createElement('canvas');
      c.id = 'ban-ve';
      c.width = RONG;
      c.height = CAO;
      c.style.position = 'absolute';
      c.style.left = '44px';
      c.style.top = '194px';
      c.style.background = '#ffffff';
      goc.appendChild(c);
    },
    capNhat: function (goc, du, t) {
      var K = root.THI_NGHIEM_KHUNG;
      var M = root.THI_NGHIEM_MO_HINH;
      var g = giaTri(du, t);
      var ctx = goc.querySelector('#ban-ve').getContext('2d');
      ctx.clearRect(0, 0, RONG, CAO);
      var tm = Math.max(0, t - du.danDau);
      if (typeof M.thoiLuong === 'function') { tm = Math.min(tm, M.thoiLuong(g.tham, g.dai)); }
      M.ve(ctx, g.tham, tm, { rong: RONG, cao: CAO }, g.dai);
      Object.keys(du.thamSo).slice(0, 3).forEach(function (ma, k) {
        var ts = tim(du.khaiBao.thamSo, ma);
        goc.querySelector('[data-id="ts-' + k + '"]').innerHTML =
          K.danhDau(ts.ten) + ' = ' + K.dinhDang(g.tham[ma], soThapPhan(ts.buoc)) + ' ' + K.danhDau(ts.donVi || '');
      });
      du.do.slice(0, 3).forEach(function (ma, k) {
        var dl = tim(du.khaiBao.daiLuongDo, ma);
        goc.querySelector('[data-id="do-' + k + '"]').innerHTML =
          K.danhDau(dl.ten) + ' = ' + K.dinhDang(g.dai[ma], dl.chuSo) + ' ' + K.danhDau(dl.donVi || '');
      });
    }
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
