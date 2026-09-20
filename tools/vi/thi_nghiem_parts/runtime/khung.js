(function (root) {
  'use strict';

  // ---------- Logic thuần: chạy được cả trong trình duyệt lẫn Node ----------

  function taoNgauNhien(hatGiong) {
    var a = hatGiong | 0;
    return function () {
      a |= 0; a = a + 0x6D2B79F5 | 0;
      var t = Math.imul(a ^ a >>> 15, 1 | a);
      t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t;
      return ((t ^ t >>> 14) >>> 0) / 4294967296;
    };
  }

  function nhieuChuan(rng) {
    var u = 1 - rng();
    var v = rng();
    return Math.sqrt(-2 * Math.log(u)) * Math.cos(2 * Math.PI * v);
  }

  function thamSoMacDinh(khaiBao) {
    var p = {};
    khaiBao.thamSo.forEach(function (ts) { p[ts.ma] = ts.macDinh; });
    return p;
  }

  function gopThamSo(khaiBao, vao) {
    var p = thamSoMacDinh(khaiBao);
    Object.keys(vao || {}).forEach(function (ma) { p[ma] = vao[ma]; });
    return p;
  }

  function tuKiem(khaiBao, moHinh) {
    var truot = [];
    var dongTruot = {};
    khaiBao.bangKiem.forEach(function (dong, chiSo) {
      var ketQua;
      try {
        ketQua = moHinh.tinh(gopThamSo(khaiBao, dong.vao));
      } catch (loi) {
        truot.push({ dong: chiSo + 1, ma: '*', mong: null, duoc: String(loi) });
        dongTruot[chiSo] = true;
        return;
      }
      Object.keys(dong.ra).forEach(function (ma) {
        var mong = dong.ra[ma];
        var duoc = ketQua[ma];
        var dat = mong === null
          ? duoc === null
          : typeof duoc === 'number' && isFinite(duoc) && Math.abs(duoc - mong) <= dong.saiSo;
        if (!dat) {
          truot.push({ dong: chiSo + 1, ma: ma, mong: mong, duoc: duoc === undefined ? null : duoc });
          dongTruot[chiSo] = true;
        }
      });
    });
    var tong = khaiBao.bangKiem.length;
    return { tong: tong, dat: tong - Object.keys(dongTruot).length, truot: truot };
  }

  function apDungPhep(phep, x) {
    if (x === null || x === undefined || !isFinite(x)) { return null; }
    var y;
    if (phep === 'binh-phuong') { y = x * x; }
    else if (phep === 'nghich-dao') { y = x === 0 ? NaN : 1 / x; }
    else if (phep === 'ln') { y = x > 0 ? Math.log(x) : NaN; }
    else if (phep === 'can') { y = x >= 0 ? Math.sqrt(x) : NaN; }
    else { y = x; }
    return isFinite(y) ? y : null;
  }

  function khopTuyenTinh(diem) {
    var n = diem.length;
    if (n < 2) { return null; }
    var tx = 0, ty = 0;
    diem.forEach(function (d) { tx += d[0]; ty += d[1]; });
    tx /= n; ty /= n;
    var sxx = 0, sxy = 0, syy = 0;
    diem.forEach(function (d) {
      sxx += (d[0] - tx) * (d[0] - tx);
      sxy += (d[0] - tx) * (d[1] - ty);
      syy += (d[1] - ty) * (d[1] - ty);
    });
    if (sxx === 0) { return null; }
    var heSoGoc = sxy / sxx;
    return {
      heSoGoc: heSoGoc,
      tungDoGoc: ty - heSoGoc * tx,
      tuongQuan: syy === 0 ? 1 : sxy / Math.sqrt(sxx * syy)
    };
  }

  function dinhDang(x, chuSo) {
    if (x === null || x === undefined || typeof x !== 'number' || !isFinite(x)) { return '—'; }
    var tron = Number(x.toFixed(chuSo));
    return (tron === 0 ? 0 : tron).toFixed(chuSo).replace('.', ',');
  }

  function danhDau(chu) {
    var sach = String(chu)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    return sach
      .replace(/\*\*(.+?)\*\*/g, '<b>$1</b>')
      .replace(/~([^~]+)~/g, '<sub>$1</sub>')
      .replace(/\^([^\^]+)\^/g, '<sup>$1</sup>');
  }

  function mauDo(khaiBao, ketQua, rng, batSaiSo) {
    var mau = {};
    khaiBao.daiLuongDo.forEach(function (d) {
      var thuc = ketQua[d.ma];
      if (thuc === null || thuc === undefined || !batSaiSo || !d.saiSo) { mau[d.ma] = thuc === undefined ? null : thuc; }
      else { mau[d.ma] = thuc + d.saiSo * nhieuChuan(rng); }
    });
    return mau;
  }

  function taoNhiemVu(cauHinh) {
    var nv = {
      buoc: 'du-doan',
      duDoan: null,
      lanDo: [],
      giaoVien: cauHinh.nguoiThaoTac === 'giao-vien'
    };
    nv.duocThaoTac = function () { return nv.giaoVien || nv.buoc !== 'du-doan'; };
    nv.chonDuDoan = function (giaTri) {
      if (nv.buoc !== 'du-doan' || giaTri === null || giaTri === undefined || String(giaTri).trim() === '') { return false; }
      nv.duDoan = String(giaTri).trim();
      nv.buoc = 'quan-sat';
      return true;
    };
    nv.ghiLanDo = function (dong) {
      if (!nv.duocThaoTac()) { return false; }
      nv.lanDo.push(dong);
      if (nv.buoc === 'quan-sat' && nv.lanDo.length >= cauHinh.quanSat.soLanDo) { nv.buoc = 'giai-thich'; }
      return true;
    };
    nv.xoaLanDo = function (chiSo) {
      if (chiSo < 0 || chiSo >= nv.lanDo.length) { return false; }
      nv.lanDo.splice(chiSo, 1);
      return true;
    };
    nv.duDoanDung = function () {
      if (nv.buoc !== 'giai-thich' || !cauHinh.duDoan.dapAn) { return null; }
      return nv.duDoan === cauHinh.duDoan.dapAn;
    };
    nv.batGiaoVien = function () { nv.giaoVien = true; };
    return nv;
  }

  function diemDoThi(doThi, lanDo) {
    var diem = [];
    lanDo.forEach(function (dong) {
      var x = apDungPhep(doThi.hoanh.phep, dong[doThi.hoanh.ma]);
      var y = apDungPhep(doThi.tung.phep, dong[doThi.tung.ma]);
      if (x !== null && y !== null) { diem.push([x, y]); }
    });
    return diem;
  }

  root.THI_NGHIEM_KHUNG = {
    taoNgauNhien: taoNgauNhien, nhieuChuan: nhieuChuan,
    thamSoMacDinh: thamSoMacDinh, gopThamSo: gopThamSo, tuKiem: tuKiem,
    apDungPhep: apDungPhep, khopTuyenTinh: khopTuyenTinh, dinhDang: dinhDang, danhDau: danhDau,
    mauDo: mauDo, taoNhiemVu: taoNhiemVu, diemDoThi: diemDoThi
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
