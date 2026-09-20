(function (root) {
  'use strict';
  var KW = 1.0e-14;
  var KA = { 'ch3cooh': 1.75e-5 };
  // Khoảng đổi màu: [pH bắt đầu, pH kết thúc, màu dạng acid, màu dạng base] (màu là [r, g, b, độ đậm]).
  var CHI_THI = {
    'phenolphtalein': [8.2, 10.0, [255, 255, 255, 0.05], [236, 72, 153, 0.75]],
    'metyl-da-cam': [3.1, 4.4, [239, 68, 68, 0.7], [250, 204, 21, 0.7]],
    'bromothymol': [6.0, 7.6, [250, 204, 21, 0.7], [37, 99, 235, 0.7]]
  };

  function tinh(p) {
    var tong = p['the-tich-acid'] + p['the-tich-base'];
    var acid = p['nong-do-acid'] * p['the-tich-acid'] / tong;
    var natri = p['nong-do-base'] * p['the-tich-base'] / tong;
    var h;
    if (p['loai-acid'] === 'hcl') {
      var lech = acid - natri;
      h = (lech + Math.sqrt(lech * lech + 4 * KW)) / 2;
    } else {
      var ka = KA[p['loai-acid']];
      var thap = 0, cao = 14;
      for (var lan = 0; lan < 200; lan += 1) {
        var giua = (thap + cao) / 2;
        var thu = Math.pow(10, -giua);
        if (thu + natri - KW / thu - acid * ka / (ka + thu) > 0) { thap = giua; } else { cao = giua; }
      }
      h = Math.pow(10, -(thap + cao) / 2);
    }
    return { 'ph': -Math.log(h) / Math.LN10 };
  }

  function mauDungDich(chiThi, ph) {
    var ct = CHI_THI[chiThi];
    var k = Math.max(0, Math.min(1, (ph - ct[0]) / (ct[1] - ct[0])));
    var m = ct[2].map(function (a, i) { return a + (ct[3][i] - a) * k; });
    return 'rgba(' + Math.round(m[0]) + ',' + Math.round(m[1]) + ',' + Math.round(m[2]) + ',' + m[3].toFixed(2) + ')';
  }

  function ve(ctx, p, t, kt, d) {
    var K = root.THI_NGHIEM_KHUNG;
    var giua = kt.rong / 2, day = kt.cao - 50;
    ctx.strokeStyle = K.MAU.net; ctx.lineWidth = 2; ctx.font = '15px ' + K.PHONG;
    // Buret: vạch 0 ở trên, mực dung dịch hạ dần theo thể tích đã nhỏ (buret 50 mL).
    var buretTren = 30, buretCao = kt.cao * 0.45;
    ctx.strokeRect(giua - 9, buretTren, 18, buretCao);
    ctx.fillStyle = 'rgba(37,99,235,0.25)';
    var daNho = buretCao * p['the-tich-base'] / 50;
    ctx.fillRect(giua - 8, buretTren + daNho, 16, buretCao - daNho);
    ctx.beginPath(); ctx.moveTo(giua, buretTren + buretCao); ctx.lineTo(giua, buretTren + buretCao + 18); ctx.stroke();
    // Bình tam giác.
    var co = buretTren + buretCao + 24;
    ctx.beginPath(); ctx.moveTo(giua - 16, co); ctx.lineTo(giua - 16, co + 26); ctx.lineTo(giua - 90, day);
    ctx.lineTo(giua + 90, day); ctx.lineTo(giua + 16, co + 26); ctx.lineTo(giua + 16, co); ctx.stroke();
    var muc = Math.min(0.75, (p['the-tich-acid'] + p['the-tich-base']) / 130);
    var caoLong = (day - co - 26) * muc;
    // Tô hai lớp: nước nhạt để thấy mực dung dịch, rồi màu của chất chỉ thị.
    ['rgba(186,230,253,0.45)', mauDungDich(p['chi-thi'], d['ph'])].forEach(function (mau) {
      ctx.fillStyle = mau;
      ctx.beginPath(); ctx.moveTo(giua - 90, day); ctx.lineTo(giua + 90, day);
      ctx.lineTo(giua + 90 - 74 * caoLong / (day - co - 26), day - caoLong);
      ctx.lineTo(giua - 90 + 74 * caoLong / (day - co - 26), day - caoLong); ctx.closePath(); ctx.fill();
    });
    ctx.fillStyle = K.MAU.net;
    ctx.fillText('NaOH đã nhỏ: ' + K.dinhDang(p['the-tich-base'], 1) + ' mL', 16, 24);
  }

  root.THI_NGHIEM_MO_HINH = { tinh: tinh, ve: ve };
})(typeof globalThis !== 'undefined' ? globalThis : this);
