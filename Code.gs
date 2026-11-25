const SPREADSHEET_ID = '1YSmpFaCuEvCfS9hU2-xqRVOj194WuQYin7GU0b9Umxc';

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? String(e.parameter.action) : 'web_app';
    if (action === 'load_simptk') {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const payload = readSnapshotPayloadFromSheet(ss);
      if (payload) return jsonResponse({ status: 'ok', payload: stripPinsFromPayload(payload) });
      return jsonResponse({ status: 'empty' });
    }
    if (action === 'verify_login') {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const user = (e && e.parameter && e.parameter.user) ? String(e.parameter.user) : '';
      const pin = (e && e.parameter && e.parameter.pin) ? String(e.parameter.pin) : '';
      return jsonResponse(verifyLoginServer(ss, user, pin));
    }
    if (action === 'ping') {
      return jsonResponse({ status: 'ok', app: 'SIM-PTK MIN 4 BATANG HARI V15', spreadsheet: 'MADRASAH_ASSESSMENT_CORE_2027' });
    }
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('SIM-PTK MIN 4 Batang Hari — Pembagian Tugas GTK & Jadwal Full Day')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=5.0');
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e && e.postData && e.postData.contents ? e.postData.contents : '{}');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
    if (body.action === 'verify_login') {
      return jsonResponse(verifyLoginServer(ss, body.user, body.pin));
    }
    if (body.action === 'change_pin') {
      return jsonResponse(changePinServer(ss, body.kode, body.newPin, body.updatedBy));
    }
    if (body.action === 'save_simptk') {
      const res = writeSnapshotToSheet(ss, body.payload, body.updatedBy);
      return jsonResponse(res);
    }
    return jsonResponse({ status: 'ignored' });
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

const MAX_CELL_CHARS = 45000;

function readSnapshotPayloadFromSheet(ss) {
  const sh = ss.getSheetByName('SIMPTK_STATE_SNAPSHOT');
  if (!sh || sh.getLastRow() < 2) return null;
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim().toUpperCase());
  const payloadColIdx = headers.indexOf('PAYLOAD_JSON') >= 0 ? headers.indexOf('PAYLOAD_JSON') + 1 : 2;
  const chunkColIdx = headers.indexOf('CHUNK_INFO') >= 0 ? headers.indexOf('CHUNK_INFO') + 1 : -1;

  if (chunkColIdx >= 1) {
    const firstChunkInfo = String(sh.getRange(2, chunkColIdx).getValue() || '').trim();
    const match = firstChunkInfo.match(/^1\/(\d+)$/);
    if (match) {
      const totalChunks = parseInt(match[1], 10);
      const parts = [];
      let complete = true;
      for (let i = 0; i < totalChunks; i++) {
        const r = 2 + i;
        if (r > lastRow) { complete = false; break; }
        const cInfo = String(sh.getRange(r, chunkColIdx).getValue() || '').trim();
        if (cInfo !== `${i + 1}/${totalChunks}`) { complete = false; break; }
        parts.push(String(sh.getRange(r, payloadColIdx).getValue() || ''));
      }
      if (complete) {
        try {
          const joined = parts.join('');
          const parsed = JSON.parse(joined);
          if (parsed && Array.isArray(parsed.gtk) && parsed.gtk.length >= 1) {
            return parsed;
          }
        } catch (e) {}
      }
    }
  }

  for (let r = 2; r <= lastRow; r++) {
    const val = sh.getRange(r, payloadColIdx).getValue();
    if (val && typeof val === 'string' && val.trim().startsWith('{')) {
      try {
        const parsed = JSON.parse(val);
        if (parsed && Array.isArray(parsed.gtk) && parsed.gtk.length >= 1) {
          return parsed;
        }
      } catch (e) {}
    }
  }
  return null;
}

function getServerUsersList(ss, snapshotPayload) {
  const prop = PropertiesService.getScriptProperties();
  const savedPinsJson = prop.getProperty('SIMPTK_USER_PINS_JSON');
  let pinOverrides = {};
  if (savedPinsJson) {
    try { pinOverrides = JSON.parse(savedPinsJson); } catch (e) {}
  }

  let users = [];
  if (snapshotPayload && Array.isArray(snapshotPayload.users) && snapshotPayload.users.length > 0) {
    users = snapshotPayload.users.map(u => Object.assign({}, u));
  } else if (snapshotPayload && Array.isArray(snapshotPayload.gtk) && snapshotPayload.gtk.length > 0) {
    users = snapshotPayload.gtk.map(g => {
      let role = 'Guru';
      if (g.kode === 'RM') role = 'Kepala Madrasah';
      else if (g.kode === 'HP') role = 'Admin / Waka Kurikulum';
      else if (g.kategori === 'Tenaga Kependidikan') role = 'Tendik';
      return { gtkId: g.id, kode: g.kode, nama: g.nama, nip: g.nip || '', role: role, pin: '' };
    });
  } else {
    users = [
      { gtkId: 'KM-01', kode: 'RM', nama: 'RAMLINA, S.Pd.I', nip: '197406151999032003', role: 'Kepala Madrasah', pin: '' },
      { gtkId: 'GR-08', kode: 'HP', nama: 'HIPNI, S.Pd.I', nip: '197803082007101001', role: 'Admin / Waka Kurikulum', pin: '' }
    ];
  }

  const adminDefaultPin = prop.getProperty('DEFAULT_ADMIN_PIN') || '123456';
  const gtkDefaultPin = prop.getProperty('DEFAULT_GTK_PIN') || '1234';

  return users.map(u => {
    const k = String(u.kode || '').trim().toUpperCase();
    const isAdminRole = (k === 'RM' || k === 'HP' || u.role === 'Kepala Madrasah' || u.role === 'Admin / Waka Kurikulum');
    const effectivePin = pinOverrides[k] || (u.pin && String(u.pin).trim()) || (isAdminRole ? adminDefaultPin : gtkDefaultPin);
    return Object.assign({}, u, { pin: String(effectivePin) });
  });
}

function stripPinsFromPayload(payload) {
  if (!payload || typeof payload !== 'object') return payload;
  const copy = Object.assign({}, payload);
  if (Array.isArray(copy.users)) {
    copy.users = copy.users.map(u => ({
      gtkId: u.gtkId,
      kode: u.kode,
      nama: u.nama,
      nip: u.nip || '',
      role: u.role || 'Guru'
    }));
  }
  return copy;
}

function verifyLoginServer(ss, rawUser, rawPin) {
  const qUser = String(rawUser || '').trim().toUpperCase();
  const qPin = String(rawPin || '').trim();
  if (!qUser || !qPin) {
    return { status: 'fail', message: 'Username/Kode/NIP dan PIN wajib diisi.' };
  }
  const snapshot = readSnapshotPayloadFromSheet(ss);
  const users = getServerUsersList(ss, snapshot);
  const matched = users.find(u =>
    String(u.kode || '').trim().toUpperCase() === qUser ||
    (u.nip && String(u.nip).trim() === qUser) ||
    String(u.nama || '').trim().toUpperCase() === qUser
  );
  if (!matched || String(matched.pin || '').trim() !== qPin) {
    return { status: 'fail', message: 'Username / Kode / NIP atau PIN tidak sesuai.' };
  }
  return {
    status: 'ok',
    user: {
      gtkId: matched.gtkId,
      kode: matched.kode,
      nama: matched.nama,
      nip: matched.nip || '',
      role: matched.role || 'Guru'
    }
  };
}

function changePinServer(ss, kode, newPin, updatedBy) {
  const k = String(kode || '').trim().toUpperCase();
  const np = String(newPin || '').trim();
  if (!k || np.length < 4) {
    return { status: 'error', message: 'Kode akun tidak valid atau PIN kurang dari 4 karakter.' };
  }
  const prop = PropertiesService.getScriptProperties();
  let pinOverrides = {};
  try { pinOverrides = JSON.parse(prop.getProperty('SIMPTK_USER_PINS_JSON') || '{}'); } catch (e) {}
  pinOverrides[k] = np;
  prop.setProperty('SIMPTK_USER_PINS_JSON', JSON.stringify(pinOverrides));

  const snapshot = readSnapshotPayloadFromSheet(ss);
  if (snapshot && Array.isArray(snapshot.users)) {
    const idx = snapshot.users.findIndex(u => String(u.kode || '').trim().toUpperCase() === k);
    if (idx >= 0) {
      snapshot.users[idx].pin = np;
      writeSnapshotToSheet(ss, snapshot, updatedBy || k);
    }
  }
  return { status: 'ok', kode: k };
}

function verifyLoginFromClient(user, pin) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    return verifyLoginServer(ss, user, pin);
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function changePinFromClient(kode, newPin, updatedBy) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    return changePinServer(ss, kode, newPin, updatedBy);
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function loadSimptkForClient() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const payload = readSnapshotPayloadFromSheet(ss);
    if (payload) return { status: 'ok', payload: stripPinsFromPayload(payload) };
    return { status: 'empty' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function writeSnapshotToSheet(ss, payload, updatedBy) {
  const nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
  if (!payload || typeof payload !== 'object' || !Array.isArray(payload.gtk) || payload.gtk.length === 0) {
    return { status: 'ignored', reason: 'Payload tidak valid atau GTK kosong' };
  }
  const existingSnap = readSnapshotPayloadFromSheet(ss);
  const existingUsers = getServerUsersList(ss, existingSnap);
  const pinByKode = {};
  existingUsers.forEach(u => { pinByKode[String(u.kode || '').trim().toUpperCase()] = u.pin; });
  if (Array.isArray(payload.users)) {
    payload.users = payload.users.map(u => {
      const k = String(u.kode || '').trim().toUpperCase();
      const preservedPin = (u.pin && String(u.pin).trim()) ? String(u.pin).trim() : (pinByKode[k] || '');
      return Object.assign({}, u, { pin: preservedPin });
    });
  }
  let sh = ss.getSheetByName('SIMPTK_STATE_SNAPSHOT');
  if (!sh) {
    sh = ss.insertSheet('SIMPTK_STATE_SNAPSHOT');
  }
  const lastCol = sh.getLastColumn();
  let headers = [];
  if (lastCol >= 1) {
    headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim().toUpperCase());
  }
  if (!headers.includes('PAYLOAD_JSON')) {
    sh.clear();
    headers = ['UPDATED_AT', 'KODE_DB', 'TAHUN_PELAJARAN', 'SEMESTER', 'PAYLOAD_JSON', 'CHUNK_INFO'];
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else if (!headers.includes('CHUNK_INFO')) {
    headers.push('CHUNK_INFO');
    if (sh.getMaxColumns() < headers.length) {
      sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
    }
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  const jsonStr = JSON.stringify(payload);
  const chunks = [];
  for (let i = 0; i < jsonStr.length; i += MAX_CELL_CHARS) {
    chunks.push(jsonStr.slice(i, i + MAX_CELL_CHARS));
  }
  if (chunks.length === 0) chunks.push('{}');

  if (sh.getMaxColumns() < headers.length) {
    sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
  }

  const kodeDbBase = (payload.meta && payload.meta.kodeDatabase) || 'MAC-MIN4BH-2027';
  const tpVal = (payload.meta && payload.meta.tahunPelajaran) || '2026/2027';
  const semVal = (payload.meta && payload.meta.semester) || 'Ganjil';

  const rowsToWrite = chunks.map((chunkStr, idx) => {
    const chunkLabel = `${idx + 1}/${chunks.length}`;
    return headers.map(h => {
      if (h === 'UPDATED_AT') return nowStr;
      if (h === 'KODE_DB') return chunks.length > 1 ? `${kodeDbBase}#${idx + 1}` : kodeDbBase;
      if (h === 'TAHUN_PELAJARAN') return tpVal;
      if (h === 'SEMESTER') return semVal;
      if (h === 'PAYLOAD_JSON') return chunkStr;
      if (h === 'CHUNK_INFO') return chunkLabel;
      if (h === 'UPDATED_BY') return updatedBy || 'GAS Web App';
      return '';
    });
  });

  const currentLastRow = sh.getLastRow();
  const neededLastRow = 1 + rowsToWrite.length;
  if (currentLastRow > neededLastRow) {
    sh.getRange(neededLastRow + 1, 1, currentLastRow - neededLastRow, headers.length).clearContent();
  }

  sh.getRange(2, 1, rowsToWrite.length, headers.length).setValues(rowsToWrite);
  SpreadsheetApp.flush();
  return { status: 'ok', savedAt: nowStr, bytes: jsonStr.length, chunks: chunks.length };
}

function saveSimptkFromClient(payload, updatedBy) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    return writeSnapshotToSheet(ss, payload, updatedBy);
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}
