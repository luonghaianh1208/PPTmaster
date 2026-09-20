(function (root) {
  'use strict';

  function giaTri(p, x) { return ((p['a'] * x + p['b']) * x + p['c']) * x + p['d']; }
  function daoHam(p, x) { return (3 * p['a'] * x + 2 * p['b']) * x + p['c']; }

  function tinh(p) {
    var a = p['a'], b = p['b'], c = p['c'];
    var cucDai = null, cucTieu = null;
    if (a !== 0) {
      var biet = 4 * b * b - 12 * a * c;
      if (biet > 0) {
        var x1 = (-2 * b - Math.sqrt(biet)) / (6 * a);
        var x2 = (-2 * b + Math.sqrt(biet)) / (6 * a);
        var nho = Math.min(x1, x2), lon = Math.max(x1, x2);
        if (a > 0) { cucDai = nho; cucTieu = lon; } else { cucTieu = nho; cucDai = lon; }
      }
    } else if (b !== 0) {
      if (b < 0) { cucDai = -c / (2 * b); } else { cucTieu = -c / (2 * b); }
    }
    return {
      'gia-tri': giaTri(p, p['x0']),
      'he-so-goc': daoHam(p, p['x0']),
      'hoanh-do-cuc-dai': cucDai === null ? null : cucDai + 0,
      'hoanh-do-cuc-tieu': cucTieu === null ? null : cucTieu + 0
    };
  }

  function ve(ctx, p, t, kt, d) {
    var K = root.THI_NGHIEM_KHUNG;
    var BIEN = 6;
    var tiLe = Math.min(kt.rong, kt.cao) / (2 * BIEN);
    var goc = [kt.rong / 2, kt.cao / 2];
    function diem(x, y) { return [goc[0] + x * tiLe, goc[1] - y * tiLe]; }
    ctx.font = '13px ' + K.PHONG; ctx.lineWidth = 1; ctx.strokeStyle = '#e2e8f0';
    for (var i = -BIEN; i <= BIEN; i += 1) {
      ctx.beginPath(); ctx.moveTo(diem(i, 0)[0], 0); ctx.lineTo(diem(i, 0)[0], kt.cao);
      ctx.moveTo(0, diem(0, i)[1]); ctx.lineTo(kt.rong, diem(0, i)[1]); ctx.stroke();
    }
    ctx.strokeStyle = K.MAU.net; ctx.lineWidth = 1.5;
    ctx.beginPath(); ctx.moveTo(0, goc[1]); ctx.lineTo(kt.rong, goc[1]); ctx.moveTo(goc[0], 0); ctx.lineTo(goc[0], kt.cao); ctx.stroke();
    ctx.fillStyle = K.MAU.net; ctx.fillText('x', kt.rong - 14, goc[1] - 6); ctx.fillText('y', goc[0] + 6, 14);
    ctx.fillText('1', diem(1, 0)[0] - 3, goc[1] + 15); ctx.fillText('1', goc[0] - 14, diem(0, 1)[1] + 4);
    ctx.strokeStyle = K.MAU.chinh; ctx.lineWidth = 2.5; ctx.beginPath();
    var dangVe = false;
    for (var buoc = 0; buoc <= 600; buoc += 1) {
      var x = -BIEN * kt.rong / kt.cao + buoc * (2 * BIEN * kt.rong / kt.cao) / 600;
      var q = diem(x, giaTri(p, x));
      if (q[1] < -2000 || q[1] > kt.cao + 2000) { dangVe = false; continue; }
      if (dangVe) { ctx.lineTo(q[0], q[1]); } else { ctx.moveTo(q[0], q[1]); dangVe = true; }
    }
    ctx.stroke();
    var tiep = diem(p['x0'], d['gia-tri']);
    ctx.strokeStyle = K.MAU.phu; ctx.lineWidth = 2; ctx.beginPath();
    ctx.moveTo(tiep[0] - 3 * tiLe, tiep[1] + 3 * tiLe * d['he-so-goc']);
    ctx.lineTo(tiep[0] + 3 * tiLe, tiep[1] - 3 * tiLe * d['he-so-goc']); ctx.stroke();
    ctx.fillStyle = K.MAU.phu; ctx.beginPath(); ctx.arc(tiep[0], tiep[1], 5, 0, 2 * Math.PI); ctx.fill();
    ['hoanh-do-cuc-dai', 'hoanh-do-cuc-tieu'].forEach(function (ma) {
      if (d[ma] === null) { return; }
      var ct = diem(d[ma], giaTri(p, d[ma]));
      ctx.fillStyle = K.MAU.xau; ctx.beginPath(); ctx.arc(ct[0], ct[1], 4, 0, 2 * Math.PI); ctx.fill();
    });
  }

  root.THI_NGHIEM_MO_HINH = { tinh: tinh, ve: ve };
})(typeof globalThis !== 'undefined' ? globalThis : this);
