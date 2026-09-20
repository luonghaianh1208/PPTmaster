(function (root) {
  'use strict';
  var XAC_SUAT = {
    'dong-xu-ngua': 1 / 2, 'xuc-xac-mat-6': 1 / 6, 'xuc-xac-chan': 1 / 2,
    'hai-xuc-xac-tong-7': 1 / 6, 'hai-xuc-xac-tong-12': 1 / 36
  };

  function xayRa(phepThu, rng) {
    if (phepThu === 'dong-xu-ngua') { return rng() < 0.5; }
    var mat = Math.floor(rng() * 6) + 1;
    if (phepThu === 'xuc-xac-mat-6') { return mat === 6; }
    if (phepThu === 'xuc-xac-chan') { return mat % 2 === 0; }
    var tong = mat + Math.floor(rng() * 6) + 1;
    return phepThu === 'hai-xuc-xac-tong-7' ? tong === 7 : tong === 12;
  }

  function tinh(p) {
    var rng = root.THI_NGHIEM_KHUNG.taoNgauNhien(p['hat-giong']);
    var tanSo = 0;
    for (var lan = 0; lan < p['so-lan']; lan += 1) {
      if (xayRa(p['phep-thu'], rng)) { tanSo += 1; }
    }
    return { 'tan-so': tanSo, 'tan-suat': tanSo / p['so-lan'], 'xac-suat-li-thuyet': XAC_SUAT[p['phep-thu']] };
  }

  function ve(ctx, p, t, kt, d) {
    var K = root.THI_NGHIEM_KHUNG;
    var trai = 56, phai = kt.rong - 24, tren = 40, day = kt.cao - 44;
    var tran = Math.min(1, Math.max(0.2, 2.5 * d['xac-suat-li-thuyet']));
    function y(f) { return day - (day - tren) * Math.min(f, tran) / tran; }
    ctx.font = '13px ' + K.PHONG; ctx.strokeStyle = K.MAU.net; ctx.lineWidth = 1.5;
    ctx.beginPath(); ctx.moveTo(trai, tren); ctx.lineTo(trai, day); ctx.lineTo(phai, day); ctx.stroke();
    ctx.fillStyle = K.MAU.net;
    ctx.fillText('Tần suất sau n lần thử', trai, tren - 14);
    ctx.fillText('n = ' + p['so-lan'], phai - 70, day + 20);
    ctx.fillText(K.dinhDang(tran, 2), 8, tren + 4); ctx.fillText('0', 38, day + 4);
    ctx.setLineDash([6, 5]); ctx.strokeStyle = K.MAU.tot; ctx.beginPath();
    ctx.moveTo(trai, y(d['xac-suat-li-thuyet'])); ctx.lineTo(phai, y(d['xac-suat-li-thuyet'])); ctx.stroke(); ctx.setLineDash([]);
    ctx.fillStyle = K.MAU.tot; ctx.fillText('P = ' + K.dinhDang(d['xac-suat-li-thuyet'], 4), phai - 90, y(d['xac-suat-li-thuyet']) - 6);
    var rng = K.taoNgauNhien(p['hat-giong']);
    var tanSo = 0, cach = Math.max(1, Math.floor(p['so-lan'] / 400));
    ctx.strokeStyle = K.MAU.chinh; ctx.lineWidth = 2; ctx.beginPath();
    for (var lan = 1; lan <= p['so-lan']; lan += 1) {
      if (xayRa(p['phep-thu'], rng)) { tanSo += 1; }
      if (lan % cach === 0 || lan === p['so-lan']) {
        var x = trai + (phai - trai) * lan / p['so-lan'];
        if (lan === cach) { ctx.moveTo(x, y(tanSo / lan)); } else { ctx.lineTo(x, y(tanSo / lan)); }
      }
    }
    ctx.stroke();
  }

  root.THI_NGHIEM_MO_HINH = { tinh: tinh, ve: ve };
})(typeof globalThis !== 'undefined' ? globalThis : this);
