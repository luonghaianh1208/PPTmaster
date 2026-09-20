(function (root) {
  'use strict';

  function thanhPhan(p) {
    var goc = p['goc'] * Math.PI / 180;
    return { vx: p['van-toc-dau'] * Math.cos(goc), vy: p['van-toc-dau'] * Math.sin(goc) };
  }

  function tinh(p) {
    var v = thanhPhan(p);
    var g = p['g'];
    var thoiGian = (v.vy + Math.sqrt(v.vy * v.vy + 2 * g * p['do-cao-dau'])) / g;
    return {
      'thoi-gian-bay': thoiGian,
      'tam-xa': v.vx * thoiGian,
      'do-cao-cuc-dai': p['do-cao-dau'] + v.vy * v.vy / (2 * g)
    };
  }

  function ve(ctx, p, t, kt, d) {
    var K = root.THI_NGHIEM_KHUNG;
    var v = thanhPhan(p);
    var le = 48;
    var tiLe = Math.min((kt.rong - 2 * le) / Math.max(d['tam-xa'], 1), (kt.cao - 2 * le) / Math.max(d['do-cao-cuc-dai'], 1));
    function diem(thoiDiem) {
      var x = v.vx * thoiDiem;
      var y = p['do-cao-dau'] + v.vy * thoiDiem - p['g'] * thoiDiem * thoiDiem / 2;
      return [le + x * tiLe, kt.cao - le - Math.max(y, 0) * tiLe];
    }
    ctx.strokeStyle = K.MAU.net; ctx.lineWidth = 2;
    ctx.beginPath(); ctx.moveTo(le - 20, kt.cao - le); ctx.lineTo(kt.rong - le + 20, kt.cao - le); ctx.stroke();
    if (p['do-cao-dau'] > 0) {
      ctx.fillStyle = K.MAU.nhat;
      ctx.fillRect(le - 20, kt.cao - le - p['do-cao-dau'] * tiLe, 20, p['do-cao-dau'] * tiLe);
    }
    ctx.setLineDash([6, 6]); ctx.strokeStyle = K.MAU.nhat; ctx.beginPath();
    for (var i = 0; i <= 60; i += 1) {
      var q = diem(d['thoi-gian-bay'] * i / 60);
      if (i === 0) { ctx.moveTo(q[0], q[1]); } else { ctx.lineTo(q[0], q[1]); }
    }
    ctx.stroke(); ctx.setLineDash([]);
    var vat = diem(Math.min(t, d['thoi-gian-bay']));
    ctx.fillStyle = K.MAU.phu; ctx.beginPath(); ctx.arc(vat[0], vat[1], 9, 0, 2 * Math.PI); ctx.fill();
    ctx.fillStyle = K.MAU.net; ctx.font = '15px ' + K.PHONG;
    ctx.fillText('t = ' + K.dinhDang(Math.min(t, d['thoi-gian-bay']), 2) + ' s', le, 24);
  }

  root.THI_NGHIEM_MO_HINH = { tinh: tinh, ve: ve, thoiLuong: function (p, d) { return d['thoi-gian-bay']; } };
})(typeof globalThis !== 'undefined' ? globalThis : this);
