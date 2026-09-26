(function (root) {
  'use strict';

  // Hàm làm mượt, lò xo, nảy chữ và số chạy: thuần theo t, chạy được trong Node. Không thư viện.
  var LECH_NAY = 0.04; // chữ sau bắt đầu nảy muộn hơn chữ trước 0,04 s
  var NAY = 0.5;       // mỗi chữ nảy trong 0,5 s
  var CAO_NAY = 36;    // chữ rơi từ 36 px phía trên xuống

  function kep01(x) { return x < 0 ? 0 : (x > 1 ? 1 : x); }

  function easeOutBack(x) {
    x = kep01(x);
    if (x === 1) { return 1; }
    var c1 = 1.70158;
    var c3 = c1 + 1;
    return 1 + c3 * Math.pow(x - 1, 3) + c1 * Math.pow(x - 1, 2);
  }

  function easeInOut(x) {
    x = kep01(x);
    return x < 0.5 ? 2 * x * x : 1 - Math.pow(-2 * x + 2, 2) / 2;
  }

  // Lò xo tắt dần có vượt đích; nhân (1 − x) để đúng 1 tại x = 1.
  function lo_xo(x, cung, tat) {
    x = kep01(x);
    if (x === 1) { return 1; }
    cung = typeof cung === 'number' ? cung : 12;
    tat = typeof tat === 'number' ? tat : 0.35;
    return 1 - Math.exp(-tat * cung * x) * Math.cos(cung * x) * (1 - x);
  }

  // Số kiểu Việt như dinhDang của thí nghiệm ảo: dấu phẩy thập phân, không phân cách nghìn.
  function dinhDang(x, chuSo) {
    var tron = Number(x.toFixed(chuSo));
    return (tron === 0 ? 0 : tron).toFixed(chuSo).replace('.', ',');
  }

  function soChay(t, batDau, dai, gtri, chuSo) {
    var p = dai > 0 ? kep01((t - batDau) / dai) : (t >= batDau ? 1 : 0);
    var v = p >= 1 ? gtri : gtri * (1 - Math.pow(1 - p, 3));
    return dinhDang(v, chuSo);
  }

  // Nảy đàn hồi của một khối (tỉ lệ): 1 tại p = 0, phình tới khoảng 1 + 0,56·bien, dao động tắt dần, đúng 1 tại p = 1.
  function nayDanHoi(p, bien) {
    p = kep01(p);
    if (p === 0 || p === 1) { return 1; }
    return 1 + bien * Math.sin(3 * Math.PI * p) * Math.exp(-4 * p);
  }

  function thoiGianNay(n) { return LECH_NAY * Math.max(0, n - 1) + NAY; }

  // Chữ thứ i trong n chữ: s (tỉ lệ), y (px, âm là phía trên), a (độ đục).
  function nayChu(i, n, t, batDau) {
    var p = kep01((t - (batDau + LECH_NAY * i)) / NAY);
    if (p >= 1) { return { s: 1, y: 0, a: 1 }; }
    return { s: 0.4 + 0.6 * easeOutBack(p), y: -CAO_NAY * (1 - lo_xo(p)), a: kep01(p / 0.25) };
  }

  root.THI_DONG = {
    LECH_NAY: LECH_NAY, NAY: NAY,
    easeOutBack: easeOutBack, easeInOut: easeInOut, lo_xo: lo_xo,
    dinhDang: dinhDang, soChay: soChay, nayChu: nayChu, thoiGianNay: thoiGianNay, nayDanHoi: nayDanHoi
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
