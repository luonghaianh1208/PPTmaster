(function (root) {
  'use strict';

  function tinh(p) {
    var chuKi = 2 * Math.PI * Math.sqrt(p['chieu-dai'] / p['g']);
    return { 'chu-ki': chuKi, 'thoi-gian-10-dao-dong': 10 * chuKi };
  }

  function ve(ctx, p, t, kt, d) {
    var K = root.THI_NGHIEM_KHUNG;
    var goc = p['goc-lech'] * Math.PI / 180 * Math.cos(2 * Math.PI * t / d['chu-ki']);
    var treo = [kt.rong / 2, 40];
    var day = p['chieu-dai'] / 2.0 * (kt.cao - 120);
    var vat = [treo[0] + day * Math.sin(goc), treo[1] + day * Math.cos(goc)];
    var banKinh = 10 + 14 * Math.sqrt(p['khoi-luong']);
    ctx.strokeStyle = K.MAU.net; ctx.lineWidth = 4;
    ctx.beginPath(); ctx.moveTo(treo[0] - 70, treo[1]); ctx.lineTo(treo[0] + 70, treo[1]); ctx.stroke();
    ctx.setLineDash([5, 5]); ctx.lineWidth = 1; ctx.strokeStyle = K.MAU.nhat;
    ctx.beginPath(); ctx.moveTo(treo[0], treo[1]); ctx.lineTo(treo[0], treo[1] + day + 30); ctx.stroke();
    ctx.setLineDash([]); ctx.lineWidth = 2; ctx.strokeStyle = K.MAU.net;
    ctx.beginPath(); ctx.moveTo(treo[0], treo[1]); ctx.lineTo(vat[0], vat[1]); ctx.stroke();
    ctx.fillStyle = K.MAU.chinh; ctx.beginPath(); ctx.arc(vat[0], vat[1], banKinh, 0, 2 * Math.PI); ctx.fill();
    ctx.fillStyle = K.MAU.net; ctx.font = '15px ' + K.PHONG;
    ctx.fillText('t = ' + K.dinhDang(t, 2) + ' s', 16, 24);
    ctx.fillText('Số dao động: ' + Math.floor(t / d['chu-ki']), 16, 46);
  }

  root.THI_NGHIEM_MO_HINH = { tinh: tinh, ve: ve };
})(typeof globalThis !== 'undefined' ? globalThis : this);
