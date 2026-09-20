(function (root) {
  'use strict';

  function tinh(p) {
    var u = p['suat-dien-dong'], r1 = p['dien-tro-1'], r2 = p['dien-tro-2'];
    if (p['kieu-mac'] === 'song-song') {
      return {
        'dien-tro-tuong-duong': r1 * r2 / (r1 + r2),
        'cuong-do-mach-chinh': u / r1 + u / r2,
        'cuong-do-1': u / r1, 'cuong-do-2': u / r2,
        'hieu-dien-the-1': u, 'hieu-dien-the-2': u
      };
    }
    var dong = u / (r1 + r2);
    return {
      'dien-tro-tuong-duong': r1 + r2,
      'cuong-do-mach-chinh': dong,
      'cuong-do-1': dong, 'cuong-do-2': dong,
      'hieu-dien-the-1': dong * r1, 'hieu-dien-the-2': dong * r2
    };
  }

  function dienTro(ctx, K, x, y, nhan) {
    ctx.fillStyle = K.MAU.nen; ctx.fillRect(x - 45, y - 14, 90, 28);
    ctx.strokeRect(x - 45, y - 14, 90, 28);
    ctx.fillStyle = K.MAU.net; ctx.textAlign = 'center'; ctx.fillText(nhan, x, y + 5); ctx.textAlign = 'left';
  }

  function ve(ctx, p, t, kt) {
    var K = root.THI_NGHIEM_KHUNG;
    var trai = 70, phai = kt.rong - 70, tren = 70, duoi = kt.cao - 70, giua = kt.rong / 2;
    ctx.strokeStyle = K.MAU.net; ctx.lineWidth = 2; ctx.font = '15px ' + K.PHONG;
    ctx.strokeRect(trai, tren, phai - trai, duoi - tren);
    ctx.fillStyle = K.MAU.nen; ctx.fillRect(giua - 14, duoi - 20, 28, 40);
    ctx.beginPath(); ctx.moveTo(giua - 8, duoi - 18); ctx.lineTo(giua - 8, duoi + 18);
    ctx.moveTo(giua + 8, duoi - 9); ctx.lineTo(giua + 8, duoi + 9); ctx.stroke();
    ctx.fillStyle = K.MAU.net;
    ctx.fillText('U = ' + K.dinhDang(p['suat-dien-dong'], 1) + ' V', giua + 22, duoi + 28);
    var nhan1 = 'R₁ = ' + p['dien-tro-1'] + ' Ω', nhan2 = 'R₂ = ' + p['dien-tro-2'] + ' Ω';
    if (p['kieu-mac'] === 'song-song') {
      ctx.beginPath(); ctx.moveTo(giua - 110, tren); ctx.lineTo(giua - 110, tren + 80); ctx.lineTo(giua + 110, tren + 80);
      ctx.lineTo(giua + 110, tren); ctx.stroke();
      dienTro(ctx, K, giua, tren, nhan1);
      dienTro(ctx, K, giua, tren + 80, nhan2);
    } else {
      dienTro(ctx, K, giua - 90, tren, nhan1);
      dienTro(ctx, K, giua + 90, tren, nhan2);
    }
    ctx.fillText(p['kieu-mac'] === 'song-song' ? 'Mắc song song' : 'Mắc nối tiếp', 16, 24);
  }

  root.THI_NGHIEM_MO_HINH = { tinh: tinh, ve: ve };
})(typeof globalThis !== 'undefined' ? globalThis : this);
