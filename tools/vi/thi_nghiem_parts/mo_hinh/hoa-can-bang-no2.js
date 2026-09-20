(function (root) {
  'use strict';
  var DELTA_H = 57200, DELTA_S = 175.83, R = 8.314462618, R_BAR = 0.08314462618;

  function tinh(p) {
    var nhietDo = p['nhiet-do'] + 273.15;
    var apSuat = p['ap-suat'];
    var kp = Math.exp(-(DELTA_H - nhietDo * DELTA_S) / (R * nhietDo));
    var phanMol = (-kp + Math.sqrt(kp * kp + 4 * kp * apSuat)) / (2 * apSuat);
    return {
      'kp': kp,
      'phan-mol-no2': phanMol,
      'do-phan-li': Math.sqrt(kp / (kp + 4 * apSuat)),
      'nong-do-no2': phanMol * apSuat / (R_BAR * nhietDo)
    };
  }

  function ve(ctx, p, t, kt, d) {
    var K = root.THI_NGHIEM_KHUNG;
    var trai = kt.rong * 0.2, rong = kt.rong * 0.5, day = kt.cao - 60, tren = 50;
    // Xi lanh: thể tích tỉ lệ nghịch với áp suất (so với 0,5 bar là đầy xi lanh).
    var caoKhi = (day - tren) * 0.5 / p['ap-suat'];
    ctx.font = '15px ' + K.PHONG; ctx.lineWidth = 2; ctx.strokeStyle = K.MAU.net;
    ctx.fillStyle = 'rgba(146,64,14,' + Math.min(0.9, d['nong-do-no2'] / 0.05).toFixed(3) + ')';
    ctx.fillRect(trai, day - caoKhi, rong, caoKhi);
    ctx.beginPath(); ctx.moveTo(trai, tren); ctx.lineTo(trai, day); ctx.lineTo(trai + rong, day); ctx.lineTo(trai + rong, tren); ctx.stroke();
    ctx.fillStyle = K.MAU.nhat; ctx.fillRect(trai + 2, day - caoKhi - 16, rong - 4, 16);
    ctx.fillRect(trai + rong / 2 - 6, tren - 20, 12, day - caoKhi - 16 - tren + 20);
    // Nhiệt kế.
    var nk = trai + rong + 60, cot = (day - tren) * p['nhiet-do'] / 100;
    ctx.strokeRect(nk, tren, 14, day - tren);
    ctx.fillStyle = K.MAU.xau; ctx.fillRect(nk + 2, day - cot, 10, cot);
    ctx.fillStyle = K.MAU.net;
    ctx.fillText(p['nhiet-do'] + ' °C', nk - 8, day + 22);
    ctx.fillText('P = ' + K.dinhDang(p['ap-suat'], 1) + ' bar', trai, day + 22);
    ctx.fillText('N₂O₄ (không màu) ⇌ 2NO₂ (nâu đỏ)', 16, 24);
  }

  root.THI_NGHIEM_MO_HINH = { tinh: tinh, ve: ve };
})(typeof globalThis !== 'undefined' ? globalThis : this);
