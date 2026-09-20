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

  // ---------- Giao diện: chỉ chạy khi có document ----------

  var MAU = { nen: '#ffffff', net: '#1e293b', nhat: '#94a3b8', chinh: '#2563eb', phu: '#f59e0b', tot: '#16a34a', xau: '#dc2626' };
  var PHONG = '"Segoe UI", Arial, sans-serif';

  function tao(the, lop, html) {
    var nut = document.createElement(the);
    if (lop) { nut.className = lop; }
    if (html !== undefined) { nut.innerHTML = html; }
    return nut;
  }

  function soChuSo(buoc) {
    var chu = String(buoc);
    return chu.indexOf('.') < 0 ? 0 : chu.length - chu.indexOf('.') - 1;
  }

  function chuanBiKhung(canvas, tiLeCao) {
    canvas.style.height = 'auto';
    var rong = canvas.clientWidth || 640;
    var cao = Math.round(Math.min(rong * tiLeCao, 460));
    var mat = window.devicePixelRatio || 1;
    canvas.style.height = cao + 'px';
    canvas.width = Math.round(rong * mat); canvas.height = Math.round(cao * mat);
    var ctx = canvas.getContext('2d');
    ctx.setTransform(mat, 0, 0, mat, 0, 0);
    ctx.fillStyle = MAU.nen; ctx.fillRect(0, 0, rong, cao);
    return { ctx: ctx, kt: { rong: rong, cao: cao } };
  }

  function vachChia(nho, lon) {
    if (nho === lon) { nho -= 1; lon += 1; }
    var tho = (lon - nho) / 5;
    var bac = Math.pow(10, Math.floor(Math.log(tho) / Math.LN10));
    var buoc = [1, 2, 5, 10].map(function (m) { return m * bac; }).filter(function (b) { return b >= tho; })[0];
    var dau = Math.floor(nho / buoc) * buoc, vach = [];
    for (var v = dau; v <= lon + buoc * 0.5; v += buoc) { vach.push(Number(v.toPrecision(12))); }
    return vach;
  }

  function veDoThi(canvas, diem, khop, nhanHoanh, nhanTung) {
    var k = chuanBiKhung(canvas, 0.6), ctx = k.ctx, kt = k.kt;
    var trai = 64, phai = kt.rong - 20, tren = 34, day = kt.cao - 46;
    ctx.font = '13px ' + PHONG; ctx.fillStyle = MAU.net; ctx.strokeStyle = MAU.net; ctx.lineWidth = 1.5;
    ctx.beginPath(); ctx.moveTo(trai, tren); ctx.lineTo(trai, day); ctx.lineTo(phai, day); ctx.stroke();
    ctx.fillText(nhanTung, 8, 14); ctx.textAlign = 'right'; ctx.fillText(nhanHoanh, phai, kt.cao - 8); ctx.textAlign = 'left';
    if (!diem.length) { ctx.fillStyle = MAU.nhat; ctx.fillText('Chưa có số liệu', trai + 16, tren + 30); return; }
    var xs = diem.map(function (d) { return d[0]; }), ys = diem.map(function (d) { return d[1]; });
    var vx = vachChia(Math.min.apply(null, xs), Math.max.apply(null, xs));
    var vy = vachChia(Math.min.apply(null, ys), Math.max.apply(null, ys));
    function px(x) { return trai + (phai - trai) * (x - vx[0]) / (vx[vx.length - 1] - vx[0]); }
    function py(y) { return day - (day - tren) * (y - vy[0]) / (vy[vy.length - 1] - vy[0]); }
    ctx.strokeStyle = '#e2e8f0'; ctx.lineWidth = 1;
    vx.forEach(function (v) {
      ctx.beginPath(); ctx.moveTo(px(v), tren); ctx.lineTo(px(v), day); ctx.stroke();
      ctx.textAlign = 'center'; ctx.fillText(String(v).replace('.', ','), px(v), day + 16);
    });
    vy.forEach(function (v) {
      ctx.beginPath(); ctx.moveTo(trai, py(v)); ctx.lineTo(phai, py(v)); ctx.stroke();
      ctx.textAlign = 'right'; ctx.fillText(String(v).replace('.', ','), trai - 6, py(v) + 4);
    });
    ctx.textAlign = 'left';
    if (khop) {
      ctx.strokeStyle = MAU.phu; ctx.lineWidth = 2; ctx.beginPath();
      ctx.moveTo(px(vx[0]), py(khop.tungDoGoc + khop.heSoGoc * vx[0]));
      ctx.lineTo(px(vx[vx.length - 1]), py(khop.tungDoGoc + khop.heSoGoc * vx[vx.length - 1])); ctx.stroke();
    }
    ctx.fillStyle = MAU.chinh;
    diem.forEach(function (d) { ctx.beginPath(); ctx.arc(px(d[0]), py(d[1]), 5, 0, 2 * Math.PI); ctx.fill(); });
  }

  function khoiDong() {
    var duLieu = JSON.parse(document.getElementById('du-lieu').textContent);
    var cauHinh = duLieu.cauHinh, khaiBao = duLieu.khaiBao, moHinh = root.THI_NGHIEM_MO_HINH;
    var nv = taoNhiemVu(cauHinh);
    var rng = taoNgauNhien(Date.now() % 2147483647);
    var tt = { p: {}, t: 0, dangChay: false, daXong: khaiBao.hoatHinh !== 'mot-lan', d: null, mau: null, mocThoiGian: 0 };
    var theoMa = {};
    khaiBao.thamSo.forEach(function (ts) { theoMa[ts.ma] = ts; });
    khaiBao.daiLuongDo.forEach(function (dl) { theoMa[dl.ma] = dl; });
    Object.keys(cauHinh.thamSo).forEach(function (ma) {
      var ch = cauHinh.thamSo[ma];
      tt.p[ma] = ch.kieu === 'co-dinh' ? ch.giaTri : ch.macDinh;
    });

    document.getElementById('tieu-de').innerHTML = danhDau(cauHinh.tieuDe);
    document.getElementById('phu-de').innerHTML = danhDau(cauHinh.mon + ' ' + cauHinh.lop + ' · ' + khaiBao.ten);
    document.getElementById('cong-thuc').innerHTML = '<b>Mô hình:</b> ' + danhDau(khaiBao.congThuc.bieuThuc);
    document.getElementById('dieu-kien').innerHTML = '<b>Điều kiện lí tưởng hoá:</b> ' + danhDau(khaiBao.congThuc.dieuKien) +
      (khaiBao.congThuc.nguon ? ' <b>Nguồn số liệu:</b> ' + danhDau(khaiBao.congThuc.nguon) : '');

    var kiem = tuKiem(khaiBao, moHinh);
    var dongKiem = document.getElementById('tu-kiem');
    if (kiem.dat === kiem.tong) {
      dongKiem.textContent = 'Tự kiểm: ' + kiem.dat + '/' + kiem.tong + ' đạt.';
    } else {
      dongKiem.textContent = 'Tự kiểm: ' + kiem.dat + '/' + kiem.tong + ' đạt; trượt dòng ' +
        kiem.truot.map(function (m) { return m.dong; }).join(', ') + '.';
      var canhBao = tao('div', 'dai-do', 'Mô hình không qua tự kiểm — không dùng để dạy.');
      document.body.insertBefore(canhBao, document.body.firstChild);
    }

    var khungVe = document.getElementById('khung-ve');
    var oSoDo = document.getElementById('so-do');
    var oChay = document.getElementById('dieu-khien-chay');
    var oThamSo = document.getElementById('tham-so');
    var oNhiemVu = document.getElementById('nhiem-vu');
    var oBang = document.getElementById('bang');
    var oKhop = document.getElementById('khop');
    var khungDoThi = document.getElementById('do-thi');
    var cacNutKhoa = [];

    function nhan(ma) { var m = theoMa[ma]; return danhDau(m.ten) + (m.donVi ? ' (' + danhDau(m.donVi) + ')' : ''); }
    function nhanTho(ma) { var m = theoMa[ma]; return (m.ten + (m.donVi ? ' (' + m.donVi + ')' : '')).replace(/[~^*]/g, ''); }
    function chuSoCua(ma) { var m = theoMa[ma]; return m.chuSo !== undefined ? m.chuSo : soChuSo(m.buoc || 1); }
    function hienGiaTri(ma, giaTri) {
      var m = theoMa[ma];
      if (m.kieu === 'chon') { return danhDau(m.luaChon.filter(function (lc) { return lc.ma === giaTri; })[0].ten); }
      return dinhDang(giaTri, chuSoCua(ma));
    }
    function nhanBieuThuc(bt) {
      var goc = nhanTho(bt.ma);
      return { 'binh-phuong': '(' + goc + ')²', 'nghich-dao': '1/(' + goc + ')', 'ln': 'ln(' + goc + ')', 'can': '√(' + goc + ')' }[bt.phep] || goc;
    }

    function layMau() { tt.mau = mauDo(khaiBao, tt.d, rng, cauHinh.saiSo); }

    function veLai() {
      var k = chuanBiKhung(khungVe, 0.62);
      moHinh.ve(k.ctx, tt.p, tt.t, k.kt, tt.d);
    }

    function hienSoDo() {
      oSoDo.innerHTML = '';
      khaiBao.daiLuongDo.forEach(function (dl) {
        var o = tao('div', 'o-do');
        o.appendChild(tao('span', 'ten-do', nhan(dl.ma)));
        o.appendChild(tao('span', 'gia-tri-do', tt.daXong && nv.duocThaoTac() ? dinhDang(tt.mau[dl.ma], dl.chuSo) : '…'));
        oSoDo.appendChild(o);
      });
    }

    function tinhLai(giuThoiGian) {
      tt.d = moHinh.tinh(tt.p);
      if (!giuThoiGian) { tt.t = 0; tt.dangChay = false; tt.daXong = khaiBao.hoatHinh !== 'mot-lan'; }
      layMau(); veLai(); hienSoDo(); capNhatKhoa();
    }

    function nhip(moc) {
      if (!tt.dangChay) { return; }
      tt.t += (moc - tt.mocThoiGian) / 1000; tt.mocThoiGian = moc;
      if (khaiBao.hoatHinh === 'mot-lan' && tt.t >= moHinh.thoiLuong(tt.p, tt.d)) {
        tt.t = moHinh.thoiLuong(tt.p, tt.d); tt.dangChay = false; tt.daXong = true;
        layMau(); hienSoDo(); capNhatKhoa();
      }
      veLai();
      if (tt.dangChay) { window.requestAnimationFrame(nhip); }
    }

    function capNhatKhoa() {
      var mo = nv.duocThaoTac();
      cacNutKhoa.forEach(function (nut) { nut.disabled = !mo; });
      if (nutGhi) { nutGhi.disabled = !mo || !tt.daXong; }
      if (nutChay) { nutChay.textContent = tt.dangChay ? 'Dừng' : 'Chạy'; }
      document.body.className = mo ? '' : 'dang-khoa';
    }

    // Tham số
    var oCoDinh = tao('ul', 'co-dinh');
    khaiBao.thamSo.forEach(function (ts) {
      var ch = cauHinh.thamSo[ts.ma];
      if (ch.kieu === 'co-dinh') {
        oCoDinh.appendChild(tao('li', '', nhan(ts.ma) + ': <b>' + hienGiaTri(ts.ma, ch.giaTri) + '</b>'));
        return;
      }
      var hang = tao('label', 'hang-tham-so');
      var giaTri = tao('b', 'gia-tri', hienGiaTri(ts.ma, tt.p[ts.ma]));
      hang.appendChild(tao('span', '', nhan(ts.ma) + ': '));
      hang.appendChild(giaTri);
      var nhap;
      if (ch.kieu === 'chon') {
        nhap = tao('select');
        ch.luaChon.forEach(function (ma) {
          var muc = tao('option', '', hienGiaTri(ts.ma, ma)); muc.value = ma; nhap.appendChild(muc);
        });
        nhap.value = ch.macDinh;
      } else {
        nhap = tao('input'); nhap.type = 'range';
        nhap.min = ch.min; nhap.max = ch.max; nhap.step = ch.buoc; nhap.value = ch.macDinh;
      }
      nhap.addEventListener('input', function () {
        tt.p[ts.ma] = ch.kieu === 'chon' ? nhap.value : Number(nhap.value);
        giaTri.innerHTML = hienGiaTri(ts.ma, tt.p[ts.ma]);
        tinhLai(false);
      });
      cacNutKhoa.push(nhap);
      hang.appendChild(nhap);
      oThamSo.appendChild(hang);
    });
    if (oCoDinh.childNodes.length) {
      oThamSo.appendChild(tao('p', 'nho', 'Giữ cố định:'));
      oThamSo.appendChild(oCoDinh);
    }

    // Chạy, ghi lần đo
    var nutChay = null, nutGhi = tao('button', 'nut chinh', 'Ghi lần đo');
    if (khaiBao.hoatHinh !== 'khong') {
      nutChay = tao('button', 'nut', 'Chạy');
      nutChay.addEventListener('click', function () {
        if (tt.dangChay) { tt.dangChay = false; capNhatKhoa(); return; }
        if (khaiBao.hoatHinh === 'mot-lan' && tt.daXong) { tt.t = 0; tt.daXong = false; hienSoDo(); }
        tt.dangChay = true; tt.mocThoiGian = window.performance.now(); capNhatKhoa();
        window.requestAnimationFrame(nhip);
      });
      var nutDatLai = tao('button', 'nut', 'Đặt lại');
      nutDatLai.addEventListener('click', function () { tinhLai(false); });
      cacNutKhoa.push(nutChay, nutDatLai);
      oChay.appendChild(nutChay); oChay.appendChild(nutDatLai);
    }
    nutGhi.addEventListener('click', function () {
      var dong = {};
      cauHinh.quanSat.cot.forEach(function (ma) { dong[ma] = theoMa[ma].saiSo !== undefined ? tt.mau[ma] : tt.p[ma]; });
      if (nv.ghiLanDo(dong)) { layMau(); hienSoDo(); veBang(); veNhiemVu(); }
    });
    oChay.appendChild(nutGhi);
    if (cauHinh.saiSo) { oChay.appendChild(tao('span', 'nho', 'Đang bật sai số đo: mỗi lần đo lệch ngẫu nhiên một chút, như đo thật.')); }

    // Bảng số liệu và đồ thị
    function veBang() {
      var cot = cauHinh.quanSat.cot, doThi = cauHinh.quanSat.doThi;
      var html = '<table><thead><tr><th>Lần</th>' + cot.map(function (ma) { return '<th>' + nhan(ma) + '</th>'; }).join('') + '<th></th></tr></thead><tbody>';
      nv.lanDo.forEach(function (dong, chiSo) {
        html += '<tr><td>' + (chiSo + 1) + '</td>' + cot.map(function (ma) { return '<td>' + hienGiaTri(ma, dong[ma]) + '</td>'; }).join('') +
          '<td><button class="xoa" data-chi-so="' + chiSo + '" title="Xoá lần đo này">×</button></td></tr>';
      });
      oBang.innerHTML = html + '</tbody></table>';
      Array.prototype.forEach.call(oBang.querySelectorAll('.xoa'), function (nut) {
        nut.addEventListener('click', function () { nv.xoaLanDo(Number(nut.getAttribute('data-chi-so'))); veBang(); veNhiemVu(); });
      });
      var nutChep = tao('button', 'nut', 'Chép số liệu');
      nutChep.addEventListener('click', function () {
        var dongChu = [['Lần'].concat(cot.map(nhanTho)).join('\t')];
        nv.lanDo.forEach(function (dong, chiSo) {
          dongChu.push([chiSo + 1].concat(cot.map(function (ma) { return hienGiaTri(ma, dong[ma]).replace(/<[^>]+>/g, ''); })).join('\t'));
        });
        var vung = tao('textarea'); vung.value = dongChu.join('\n'); document.body.appendChild(vung);
        vung.select(); document.execCommand('copy'); document.body.removeChild(vung);
        nutChep.textContent = 'Đã chép, dán vào Excel';
      });
      oBang.appendChild(nutChep);
      if (!doThi) { khungDoThi.style.display = 'none'; oKhop.textContent = ''; return; }
      var diem = diemDoThi(doThi, nv.lanDo), khop = khopTuyenTinh(diem);
      veDoThi(khungDoThi, diem, khop, nhanBieuThuc(doThi.hoanh), nhanBieuThuc(doThi.tung));
      oKhop.textContent = khop
        ? 'Đường thẳng khớp: hệ số góc = ' + dinhDang(khop.heSoGoc, 4) + '; tung độ gốc = ' + dinhDang(khop.tungDoGoc, 4) +
          '; hệ số tương quan r = ' + dinhDang(khop.tuongQuan, 4)
        : 'Cần ít nhất hai lần đo khác nhau để vẽ đường thẳng khớp.';
    }

    // Ba bước nhiệm vụ
    function veNhiemVu() {
      oNhiemVu.innerHTML = '';
      var b1 = tao('div', 'buoc' + (nv.buoc === 'du-doan' ? ' dang-lam' : ''));
      b1.appendChild(tao('h2', '', '1. Dự đoán'));
      b1.appendChild(tao('p', '', danhDau(cauHinh.duDoan.cau)));
      if (nv.buoc === 'du-doan') {
        var oNhap;
        if (cauHinh.duDoan.luaChon.length) {
          oNhap = tao('div');
          cauHinh.duDoan.luaChon.forEach(function (lc) {
            var dong = tao('label', 'lua-chon');
            var o = tao('input'); o.type = 'radio'; o.name = 'du-doan'; o.value = lc.ma;
            dong.appendChild(o); dong.appendChild(tao('span', '', '<b>' + lc.ma + '.</b> ' + danhDau(lc.noiDung)));
            oNhap.appendChild(dong);
          });
        } else { oNhap = tao('textarea'); oNhap.rows = 3; oNhap.placeholder = 'Viết dự đoán của em'; }
        b1.appendChild(oNhap);
        var nutChot = tao('button', 'nut chinh', 'Chốt dự đoán');
        nutChot.addEventListener('click', function () {
          var chon = cauHinh.duDoan.luaChon.length ? (oNhap.querySelector('input:checked') || {}).value : oNhap.value;
          if (nv.chonDuDoan(chon)) { veNhiemVu(); hienSoDo(); capNhatKhoa(); }
        });
        b1.appendChild(nutChot);
        b1.appendChild(tao('p', 'nho', 'Chốt dự đoán xong mới làm được thí nghiệm. Dự đoán không sửa được.'));
      } else { b1.appendChild(tao('p', 'da-chon', 'Dự đoán của em: <b>' + danhDau(nv.duDoan || '(giáo viên trình diễn)') + '</b>')); }
      oNhiemVu.appendChild(b1);

      var b2 = tao('div', 'buoc' + (nv.buoc === 'quan-sat' ? ' dang-lam' : ''));
      b2.appendChild(tao('h2', '', '2. Quan sát'));
      b2.appendChild(tao('p', '', 'Thay đổi tham số, làm thí nghiệm và bấm <b>Ghi lần đo</b>. Đã ghi <b>' + nv.lanDo.length +
        '/' + cauHinh.quanSat.soLanDo + '</b> lần đo.'));
      oNhiemVu.appendChild(b2);

      var b3 = tao('div', 'buoc' + (nv.buoc === 'giai-thich' ? ' dang-lam' : ''));
      b3.appendChild(tao('h2', '', '3. Giải thích'));
      if (nv.buoc === 'giai-thich' || nv.giaoVien) {
        b3.appendChild(tao('p', '', danhDau(cauHinh.giaiThich.cau)));
        var dung = nv.duDoanDung();
        if (dung !== null) {
          b3.appendChild(tao('p', dung ? 'dung' : 'sai', dung ? 'Dự đoán của em khớp với kết quả thí nghiệm.'
            : 'Dự đoán của em chưa khớp với kết quả. Hãy dùng số liệu để giải thích vì sao.'));
        }
        var oGoiY = tao('div', 'goi-y'); oGoiY.style.display = 'none';
        oGoiY.innerHTML = '<p><b>Gợi ý đáp án:</b> ' + danhDau(cauHinh.giaiThich.goiY) + '</p><p><b>Kết luận:</b> ' + danhDau(cauHinh.ketLuan) + '</p>';
        var nutGoiY = tao('button', 'nut', 'Xem gợi ý đáp án và kết luận');
        nutGoiY.addEventListener('click', function () { oGoiY.style.display = 'block'; });
        b3.appendChild(nutGoiY); b3.appendChild(oGoiY);
      } else { b3.appendChild(tao('p', 'nho', 'Mở ra khi đã ghi đủ số lần đo.')); }
      oNhiemVu.appendChild(b3);

      if (!nv.giaoVien) {
        var nutGiaoVien = tao('button', 'nut nho-nut', 'Chế độ giáo viên (bỏ khoá)');
        nutGiaoVien.addEventListener('click', function () { nv.batGiaoVien(); veNhiemVu(); hienSoDo(); capNhatKhoa(); });
        oNhiemVu.appendChild(nutGiaoVien);
      }
    }

    window.addEventListener('resize', function () { veLai(); veBang(); });
    tinhLai(false); veBang(); veNhiemVu();
  }

  root.THI_NGHIEM_KHUNG = {
    taoNgauNhien: taoNgauNhien, nhieuChuan: nhieuChuan,
    thamSoMacDinh: thamSoMacDinh, gopThamSo: gopThamSo, tuKiem: tuKiem,
    apDungPhep: apDungPhep, khopTuyenTinh: khopTuyenTinh, dinhDang: dinhDang, danhDau: danhDau,
    mauDo: mauDo, taoNhiemVu: taoNhiemVu, diemDoThi: diemDoThi,
    MAU: MAU, PHONG: PHONG, khoiDong: khoiDong
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
