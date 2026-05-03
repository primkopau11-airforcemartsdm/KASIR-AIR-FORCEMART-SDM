const SPREADSHEET_ID = '1kCYBhss49iM22yXXSxjKSeuxZQRLaxCt1ZYfET-w6Q4';

const DB = {
  USERS: 'Users',
  SETTINGS: 'Settings',
  ANGGOTA: 'Anggota',
  BARANG: 'Barang',
  PEMBELIAN: 'Pembelian',
  DETAIL_BELI: 'DetailBeli',
  PENJUALAN: 'Penjualan',
  DETAIL_JUAL: 'DetailJual',
  STOK: 'Stok'
};

const DEFAULT_LOGIN_RFID = '000768567';

const HEADERS = {
  Users: ['id', 'username', 'password', 'role', 'nama', 'aktif', 'rfid_uid', 'created_at'],
  Settings: ['key', 'value'],
  Anggota: ['id', 'nama', 'nrp', 'pangkat', 'satker', 'rfid_uid', 'created_at'],
  Barang: ['id', 'kode', 'nama', 'kategori', 'supplier', 'satuan', 'harga_beli', 'harga_jual', 'stok', 'min_stok', 'created_at'],
  Pembelian: ['id_beli', 'tanggal', 'supplier', 'subtotal', 'diskon_total', 'ppn_persen', 'ppn_total', 'total', 'catatan'],
  DetailBeli: ['id_beli', 'kode', 'nama', 'qty', 'harga_beli', 'diskon_persen', 'diskon_nominal', 'subtotal'],
  Penjualan: ['id_jual', 'tanggal', 'user', 'anggota_nrp', 'anggota_nama', 'pelanggan', 'metode_bayar', 'total', 'bayar', 'kembali', 'catatan', 'anggota_rfid'],
  DetailJual: ['id_jual', 'kode', 'nama', 'qty', 'harga_beli', 'harga_jual', 'subtotal'],
  Stok: ['id', 'tanggal', 'kode', 'nama', 'tipe', 'qty', 'keterangan']
};

const DEFAULT_SETTINGS = {
  nama_toko: 'KOPERASI / KASIR ONLINE',
  subjudul: 'Struk Transaksi',
  alamat: 'Alamat toko belum diisi',
  telepon: '-',
  logo_url: '',
  footer_struk: 'Terima kasih atas transaksi Anda.',
  show_logo: 'YA'
};

function getSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Kasir Online RFID')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function now() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function makeId(prefix) {
  const t = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const r = Math.floor(Math.random() * 900 + 100);
  return prefix + '-' + t + '-' + r;
}

function angka(v) {
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}

function safeValue(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  if (v === null || v === undefined) return '';
  return v;
}

function normRfid(v) {
  return String(v || '').trim().replace(/^0+/, '');
}

function sh(name) {
  const sheet = getSS().getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" belum ada. Jalankan setupSheets dulu.');
  return sheet;
}

function setupSheets() {
  Object.keys(HEADERS).forEach(function(name) {
    ensureSheet(name);
  });
  seedUsers();
  seedSettings();
  seedDefaultRfidUser();
  return 'Database siap. Admin: admin/admin123. User: user/user123. RFID login default: ' + DEFAULT_LOGIN_RFID;
}

function ensureSheet(name) {
  const ss = getSS();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  const headers = HEADERS[name];
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), headers.length);

  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const oldHeaders = values[0].map(function(h) { return String(h).trim(); });
    const same = headers.length === oldHeaders.length && headers.every(function(h, i) { return h === oldHeaders[i]; });

    if (!same) {
      const newValues = [headers];
      for (let r = 1; r < values.length; r++) {
        const obj = {};
        oldHeaders.forEach(function(h, i) {
          if (h) obj[h] = values[r][i];
        });
        newValues.push(headers.map(function(h) {
          return obj[h] === undefined || obj[h] === null ? '' : obj[h];
        }));
      }
      sheet.clear();
      sheet.getRange(1, 1, newValues.length, headers.length).setValues(newValues);
    }
  }

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);

  const idxKode = headers.indexOf('kode');
  if (idxKode >= 0) sheet.getRange(1, idxKode + 1, Math.max(sheet.getMaxRows(), 1000), 1).setNumberFormat('@');
  const idxRfid = headers.indexOf('rfid_uid');
  if (idxRfid >= 0) sheet.getRange(1, idxRfid + 1, Math.max(sheet.getMaxRows(), 1000), 1).setNumberFormat('@');
  const idxAnggotaRfid = headers.indexOf('anggota_rfid');
  if (idxAnggotaRfid >= 0) sheet.getRange(1, idxAnggotaRfid + 1, Math.max(sheet.getMaxRows(), 1000), 1).setNumberFormat('@');
}

function resetDatabase() {
  const ss = getSS();
  Object.keys(HEADERS).forEach(function(name) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    sheet.clear();
    sheet.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, HEADERS[name].length);
  });
  seedUsers();
  seedSettings();
  seedDefaultRfidUser();
  return 'Database berhasil direset.';
}

function seedUsers() {
  const sheet = sh(DB.USERS);
  if (sheet.getLastRow() > 1) return;
  sheet.getRange(2, 1, 2, HEADERS.Users.length).setValues([
    [makeId('USR'), 'admin', 'admin123', 'ADMIN', 'Administrator', 'YA', '', now()],
    [makeId('USR'), 'user', 'user123', 'USER', 'Kasir/User', 'YA', DEFAULT_LOGIN_RFID, now()]
  ]);
}

function seedDefaultRfidUser() {
  const users = readTable(DB.USERS);
  const found = users.find(function(u) {
    return normRfid(u.rfid_uid) === normRfid(DEFAULT_LOGIN_RFID);
  });
  if (found) return;

  const user = users.find(function(u) {
    return String(u.username).toLowerCase() === 'user';
  });
  if (user) {
    user.rfid_uid = DEFAULT_LOGIN_RFID;
    updateRowByKey(DB.USERS, 'username', user.username, user);
  }
}

function seedSettings() {
  const sheet = sh(DB.SETTINGS);
  if (sheet.getLastRow() > 1) return;
  const rows = Object.keys(DEFAULT_SETTINGS).map(function(k) { return [k, DEFAULT_SETTINGS[k]]; });
  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
}

function readTable(name) {
  const sheet = sh(name);
  const headers = HEADERS[name];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values
    .filter(function(row) {
      return row.some(function(v) { return v !== '' && v !== null; });
    })
    .map(function(row) {
      const obj = {};
      headers.forEach(function(h, i) {
        obj[h] = safeValue(row[i]);
      });
      return obj;
    });
}

function appendData(name, obj) {
  const headers = HEADERS[name];
  const row = headers.map(function(h) {
    return obj[h] === undefined || obj[h] === null ? '' : obj[h];
  });
  sh(name).appendRow(row);
  SpreadsheetApp.flush();
}

function updateRowByKey(sheetName, keyName, keyValue, obj) {
  const sheet = sh(sheetName);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const keyIndex = headers.indexOf(keyName);
  if (keyIndex < 0) throw new Error('Kolom ' + keyName + ' tidak ditemukan.');

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][keyIndex]) === String(keyValue)) {
      const row = HEADERS[sheetName].map(function(h) {
        return obj[h] === undefined || obj[h] === null ? '' : obj[h];
      });
      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      SpreadsheetApp.flush();
      return true;
    }
  }
  return false;
}

function setupSheetsLight() {
  Object.keys(HEADERS).forEach(function(name) { ensureSheet(name); });
}

function getSettings() {
  setupSheetsLight();
  const out = {};
  Object.keys(DEFAULT_SETTINGS).forEach(function(k) { out[k] = DEFAULT_SETTINGS[k]; });
  readTable(DB.SETTINGS).forEach(function(r) {
    if (r.key) out[String(r.key)] = safeValue(r.value);
  });
  return out;
}

function setSettingValue(key, value) {
  const sheet = sh(DB.SETTINGS);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(key)) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

/* LOGIN */
function login(payload) {
  setupSheets();
  const username = String(payload.username || '').trim();
  const password = String(payload.password || '').trim();
  if (!username || !password) throw new Error('Username dan password wajib diisi.');

  const user = readTable(DB.USERS).find(function(u) {
    return String(u.username).toLowerCase() === username.toLowerCase() &&
      String(u.password) === password &&
      String(u.aktif || 'YA').toUpperCase() !== 'TIDAK';
  });

  if (!user) throw new Error('Login gagal. Username/password salah atau akun tidak aktif.');
  return createSessionFromUser(user);
}

function loginByRfid(payload) {
  setupSheets();
  const rfid = String(payload.rfid_uid || '').trim();
  if (!rfid) throw new Error('RFID UID kosong.');

  const user = readTable(DB.USERS).find(function(u) {
    return normRfid(u.rfid_uid) === normRfid(rfid) &&
      String(u.aktif || 'YA').toUpperCase() !== 'TIDAK';
  });

  if (!user) throw new Error('RFID ' + rfid + ' belum terdaftar sebagai user login.');
  return createSessionFromUser(user);
}

function createSessionFromUser(user) {
  const token = Utilities.getUuid();
  const session = {
    token: token,
    username: user.username,
    role: String(user.role || 'USER').toUpperCase(),
    nama: user.nama || user.username,
    rfid_uid: user.rfid_uid || '',
    login_at: now()
  };
  CacheService.getScriptCache().put('SESSION_' + token, JSON.stringify(session), 21600);
  return { ok: true, session: session };
}

function logout(payload) {
  if (payload && payload.token) CacheService.getScriptCache().remove('SESSION_' + payload.token);
  return { ok: true };
}

function getSession(token) {
  const raw = CacheService.getScriptCache().get('SESSION_' + token);
  if (!raw) throw new Error('Sesi login habis. Silakan login ulang.');
  return JSON.parse(raw);
}

function requireLogin(payload) {
  if (!payload || !payload.token) throw new Error('Belum login.');
  return getSession(payload.token);
}

function requireAdmin(payload) {
  const session = requireLogin(payload);
  if (session.role !== 'ADMIN') throw new Error('Akses ditolak. Khusus admin.');
  return session;
}

function getData(payload) {
  const session = requireLogin(payload);
  setupSheets();

  const data = {
    session: session,
    settings: getSettings(),
    barang: readTable(DB.BARANG),
    anggota: readTable(DB.ANGGOTA),
    pembelian: [],
    detailBeli: [],
    penjualan: [],
    detailJual: [],
    stok: [],
    users: [],
    debug: getDebugInfo()
  };

  if (session.role === 'ADMIN') {
    data.pembelian = readTable(DB.PEMBELIAN);
    data.detailBeli = readTable(DB.DETAIL_BELI);
    data.penjualan = readTable(DB.PENJUALAN);
    data.detailJual = readTable(DB.DETAIL_JUAL);
    data.stok = readTable(DB.STOK);
    data.users = readTable(DB.USERS).map(function(u) {
      u.password = '********';
      return u;
    });
  } else {
    // USER sekarang juga melihat semua transaksi untuk laporan.
    data.penjualan = readTable(DB.PENJUALAN);
    data.detailJual = readTable(DB.DETAIL_JUAL);
  }

  return data;
}

function getDebugInfo() {
  const ss = getSS();
  const barangSheet = ss.getSheetByName(DB.BARANG);
  return {
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    sheetNames: ss.getSheets().map(function(s) { return s.getName(); }),
    barangCount: barangSheet ? Math.max(barangSheet.getLastRow() - 1, 0) : 0,
    defaultLoginRfid: DEFAULT_LOGIN_RFID,
    time: now()
  };
}

/* ADMIN SETTINGS + USERS */
function saveSettings(payload) {
  requireAdmin(payload);
  const data = payload.data || {};
  Object.keys(DEFAULT_SETTINGS).forEach(function(k) {
    if (data[k] !== undefined) setSettingValue(k, data[k]);
  });
  SpreadsheetApp.flush();
  return getData({ token: payload.token });
}

function saveUser(payload) {
  requireAdmin(payload);
  const p = payload.data || {};
  const username = String(p.username || '').trim();
  const password = String(p.password || '').trim();
  const role = String(p.role || 'USER').toUpperCase();
  const rfid = String(p.rfid_uid || '').trim();

  if (!username) throw new Error('Username wajib diisi.');
  if (!password) throw new Error('Password wajib diisi.');
  if (role !== 'ADMIN' && role !== 'USER') throw new Error('Role harus ADMIN atau USER.');

  const users = readTable(DB.USERS);
  const existing = users.find(function(u) {
    return String(u.username).toLowerCase() === username.toLowerCase();
  });

  const rfidDipakai = rfid && users.some(function(u) {
    return normRfid(u.rfid_uid) === normRfid(rfid) &&
      String(u.username).toLowerCase() !== username.toLowerCase();
  });
  if (rfidDipakai) throw new Error('RFID UID sudah dipakai user lain.');

  const obj = {
    id: existing ? existing.id : makeId('USR'),
    username: username,
    password: password,
    role: role,
    nama: p.nama || username,
    aktif: p.aktif || 'YA',
    rfid_uid: rfid,
    created_at: existing ? existing.created_at : now()
  };

  if (existing) updateRowByKey(DB.USERS, 'username', username, obj);
  else appendData(DB.USERS, obj);

  return getData({ token: payload.token });
}

function deleteUser(payload) {
  requireAdmin(payload);
  const username = String(payload.username || '').trim();
  if (username.toLowerCase() === 'admin') throw new Error('User admin utama tidak boleh dihapus.');

  const sheet = sh(DB.USERS);
  const values = sheet.getDataRange().getValues();
  const idx = values[0].indexOf('username');

  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][idx]) === username) {
      sheet.deleteRow(i + 1);
      SpreadsheetApp.flush();
      return getData({ token: payload.token });
    }
  }
  throw new Error('User tidak ditemukan.');
}

/* MASTER ADMIN */
function simpanAnggota(payload) {
  requireAdmin(payload);
  const p = payload.data || {};
  const nama = String(p.nama || '').trim();
  const nrp = String(p.nrp || '').trim();
  const rfid = String(p.rfid_uid || '').trim();

  if (!nama) throw new Error('Nama anggota wajib diisi.');
  if (!nrp) throw new Error('NRP wajib diisi.');

  const all = readTable(DB.ANGGOTA);
  const rfidDipakai = rfid && all.some(function(x) {
    return normRfid(x.rfid_uid) === normRfid(rfid) && String(x.nrp) !== nrp;
  });
  if (rfidDipakai) throw new Error('RFID UID sudah dipakai anggota lain.');

  const existing = all.find(function(x) { return String(x.nrp) === nrp; });
  const obj = {
    id: existing ? existing.id : makeId('AGT'),
    nama: nama,
    nrp: nrp,
    pangkat: p.pangkat || '',
    satker: p.satker || '',
    rfid_uid: rfid,
    created_at: existing ? existing.created_at : now()
  };

  if (existing) updateRowByKey(DB.ANGGOTA, 'nrp', nrp, obj);
  else appendData(DB.ANGGOTA, obj);

  return getData({ token: payload.token });
}

function hapusAnggota(payload) {
  requireAdmin(payload);
  const nrp = String(payload.nrp || '').trim();
  const sheet = sh(DB.ANGGOTA);
  const values = sheet.getDataRange().getValues();
  const idx = values[0].indexOf('nrp');

  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][idx]) === nrp) {
      sheet.deleteRow(i + 1);
      SpreadsheetApp.flush();
      return getData({ token: payload.token });
    }
  }
  throw new Error('Anggota tidak ditemukan.');
}

function simpanBarang(payload) {
  requireAdmin(payload);
  const p = payload.data || {};
  const kode = String(p.kode || '').trim();
  const nama = String(p.nama || '').trim();

  if (!kode) throw new Error('Kode/barcode barang wajib diisi.');
  if (!nama) throw new Error('Nama barang wajib diisi.');

  const existing = readTable(DB.BARANG).find(function(x) {
    return String(x.kode) === kode;
  });

  const obj = {
    id: existing ? existing.id : makeId('BRG'),
    kode: kode,
    nama: nama,
    kategori: p.kategori || '',
    supplier: p.supplier || '',
    satuan: p.satuan || 'pcs',
    harga_beli: angka(p.harga_beli),
    harga_jual: angka(p.harga_jual),
    stok: angka(p.stok),
    min_stok: angka(p.min_stok),
    created_at: existing ? existing.created_at : now()
  };

  if (existing) updateRowByKey(DB.BARANG, 'kode', kode, obj);
  else appendData(DB.BARANG, obj);

  return getData({ token: payload.token });
}

function hapusBarang(payload) {
  requireAdmin(payload);
  const kode = String(payload.kode || '').trim();
  const sheet = sh(DB.BARANG);
  const values = sheet.getDataRange().getValues();
  const idx = values[0].indexOf('kode');

  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][idx]) === kode) {
      sheet.deleteRow(i + 1);
      SpreadsheetApp.flush();
      return getData({ token: payload.token });
    }
  }
  throw new Error('Barang tidak ditemukan.');
}

/* STOK + TRANSAKSI */
function ubahStok(kode, delta, tipe, keterangan) {
  const sheet = sh(DB.BARANG);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idxKode = headers.indexOf('kode');
  const idxNama = headers.indexOf('nama');
  const idxStok = headers.indexOf('stok');

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idxKode]) === String(kode)) {
      const stokBaru = angka(values[i][idxStok]) + angka(delta);
      if (stokBaru < 0) throw new Error('Stok ' + values[i][idxNama] + ' tidak cukup.');

      sheet.getRange(i + 1, idxStok + 1).setValue(stokBaru);

      appendData(DB.STOK, {
        id: makeId('STK'),
        tanggal: now(),
        kode: kode,
        nama: values[i][idxNama],
        tipe: tipe,
        qty: delta,
        keterangan: keterangan || ''
      });

      SpreadsheetApp.flush();
      return true;
    }
  }
  throw new Error('Barang dengan kode ' + kode + ' tidak ditemukan.');
}

function inputStok(payload) {
  requireAdmin(payload);
  const p = payload.data || {};
  const kode = String(p.kode || '').trim();
  const qty = angka(p.qty);
  const tipe = p.tipe || 'Masuk';

  if (!kode) throw new Error('Pilih barang terlebih dahulu.');
  if (qty <= 0) throw new Error('Qty harus lebih dari 0.');

  let delta = qty;
  if (tipe === 'Keluar' || tipe === 'Penyesuaian Minus') delta = -qty;

  ubahStok(kode, delta, tipe, p.keterangan || '');
  return getData({ token: payload.token });
}

function prosesPembelian(payload) {
  requireAdmin(payload);
  const p = payload.data || {};
  const items = p.items || [];
  if (!items.length) throw new Error('Keranjang pembelian stok kosong.');

  const id = makeId('BELI');
  let subtotalBarang = 0;

  items.forEach(function(item) {
    const qty = angka(item.qty);
    const harga = angka(item.harga_beli);
    const diskonPersen = angka(item.diskon_persen);
    const diskonNominal = angka(item.diskon_nominal);
    const bruto = qty * harga;
    const diskonItem = Math.min(bruto, ((bruto * diskonPersen) / 100) + diskonNominal);
    const subtotal = Math.max(0, bruto - diskonItem);

    if (!item.kode) throw new Error('Kode barang kosong.');
    if (qty <= 0) throw new Error('Qty pembelian tidak valid.');

    subtotalBarang += subtotal;

    appendData(DB.DETAIL_BELI, {
      id_beli: id,
      kode: item.kode,
      nama: item.nama,
      qty: qty,
      harga_beli: harga,
      diskon_persen: diskonPersen,
      diskon_nominal: diskonNominal,
      subtotal: subtotal
    });

    ubahStok(item.kode, qty, 'Pembelian Stok', 'Pembelian ' + id);
  });

  const diskonTotal = Math.min(subtotalBarang, angka(p.diskon_total));
  const dasarPpn = Math.max(0, subtotalBarang - diskonTotal);
  const ppnPersen = angka(p.ppn_persen);
  const ppnTotal = Math.round((dasarPpn * ppnPersen) / 100);
  const totalAkhir = dasarPpn + ppnTotal;

  appendData(DB.PEMBELIAN, {
    id_beli: id,
    tanggal: now(),
    supplier: p.supplier || '-',
    subtotal: subtotalBarang,
    diskon_total: diskonTotal,
    ppn_persen: ppnPersen,
    ppn_total: ppnTotal,
    total: totalAkhir,
    catatan: p.catatan || ''
  });

  return getData({ token: payload.token });
}

function prosesPenjualan(payload) {
  const session = requireLogin(payload);
  const p = payload.data || {};
  const items = p.items || [];
  if (!items.length) throw new Error('Keranjang transaksi kosong.');

  const id = makeId('JUAL');
  let total = 0;

  items.forEach(function(item) {
    const qty = angka(item.qty);
    const hargaBeli = angka(item.harga_beli);
    const hargaJual = angka(item.harga_jual);
    const subtotal = qty * hargaJual;

    if (!item.kode) throw new Error('Kode barang kosong.');
    if (qty <= 0) throw new Error('Qty transaksi tidak valid.');

    total += subtotal;

    appendData(DB.DETAIL_JUAL, {
      id_jual: id,
      kode: item.kode,
      nama: item.nama,
      qty: qty,
      harga_beli: hargaBeli,
      harga_jual: hargaJual,
      subtotal: subtotal
    });

    ubahStok(item.kode, -qty, 'Transaksi Kasir', 'Transaksi ' + id);
  });

  const metode = String(p.metode_bayar || 'TUNAI').toUpperCase();
  let bayar = angka(p.bayar);
  let kembali = bayar - total;

  if (metode === 'KREDIT') {
    bayar = 0;
    kembali = 0;
  } else {
    if (bayar < total) throw new Error('Uang bayar kurang.');
  }

  appendData(DB.PENJUALAN, {
    id_jual: id,
    tanggal: now(),
    user: session.username,
    anggota_nrp: p.anggota_nrp || '',
    anggota_nama: p.anggota_nama || '',
    pelanggan: p.pelanggan || 'Umum',
    metode_bayar: metode,
    total: total,
    bayar: bayar,
    kembali: kembali,
    catatan: p.catatan || '',
    anggota_rfid: p.anggota_rfid || ''
  });

  const receipt = {
    id: id,
    tanggal: now(),
    user: session.nama,
    anggota_nrp: p.anggota_nrp || '',
    anggota_nama: p.anggota_nama || '',
    anggota_rfid: p.anggota_rfid || '',
    pelanggan: p.pelanggan || 'Umum',
    metode_bayar: metode,
    items: items,
    total: total,
    bayar: bayar,
    kembali: kembali,
    catatan: p.catatan || '',
    settings: getSettings()
  };

  return {
    data: getData({ token: payload.token }),
    receipt: receipt
  };
}

function cekBarang() {
  setupSheets();
  const barang = readTable(DB.BARANG);
  Logger.log(JSON.stringify(barang));
  return barang;
}
