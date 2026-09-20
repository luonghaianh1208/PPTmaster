(function (root) {
  'use strict';
  var R = 8.314462618, K25 = 0.00125, T25 = 298.15;
  var EA = { 'khong': 50000, 'co': 45000 };
  var LUONG_PHAN_UNG = 0.005;
  var A = K25 * Math.exp(EA['khong'] / (R * T25));

  function tinh(p) {
    var nhietDo = p['nhiet-do'] + 273.15;
    var k = A * Math.exp(-EA[p['xuc-tac']] / (R * nhietDo));
    return { 'thoi-gian': LUONG_PHAN_UNG / (k * p['nong-do']) };
  }

  function ve(ctx, p, t, kt, d) {
    var K = root.THI_NGHIEM_KHUNG;
    var giua = kt.rong / 2, tren = 70, day = kt.cao - 70, nuaRong = 110;
    var duc = Math.max(0, Math.min(1, t / d['thoi-gian']));
    ctx.font = '15px ' + K.PHONG; ctx.lineWidth = 2; ctx.strokeStyle = K.MAU.net;
    // Tờ giấy có dấu X đặt dưới cốc, nhìn từ trên xuống qua dung dịch.
    ctx.fillStyle = K.MAU.nen; ctx.fillRect(giua - nuaRong, tren, 2 * nuaRong, day - tren);
    ctx.lineWidth = 8; ctx.beginPath();
    ctx.moveTo(giua - 50, (tren + day) / 2 - 50); ctx.lineTo(giua + 50, (tren + day) / 2 + 50);
    ctx.moveTo(giua + 50, (tren + day) / 2 - 50); ctx.lineTo(giua - 50, (tren + day) / 2 + 50); ctx.stroke();
    ctx.fillStyle = 'rgba(253,230,138,' + (0.97 * duc).toFixed(3) + ')';
    ctx.fillRect(giua - nuaRong, tren, 2 * nuaRong, day - tren);
    ctx.lineWidth = 2; ctx.strokeRect(giua - nuaRong, tren, 2 * nuaRong, day - tren);
    ctx.fillStyle = K.MAU.net;
    ctx.fillText('Đồng hồ: ' + K.dinhDang(Math.min(t, d['thoi-gian']), 1) + ' s', 16, 24);
    ctx.fillText(p['xuc-tac'] === 'co' ? 'Có chất xúc tác' : 'Không có chất xúc tác', 16, 46);
    ctx.fillText(p['nhiet-do'] + ' °C', giua + nuaRong + 16, tren + 16);
  }

  root.THI_NGHIEM_MO_HINH = { tinh: tinh, ve: ve, thoiLuong: function (p, d) { return d['thoi-gian']; } };
})(typeof globalThis !== 'undefined' ? globalThis : this);
