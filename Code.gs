// ============================================================
//  VidChat — Google Apps Script Backend
// ============================================================
const SPREADSHEET_ID = '1B2pYmLAqwsn_XrQCY9VB4Pd9Rf7qU4u3BjaVid2_xFM';
const FOLDER_ID      = '1Xu2Wyexv128LK6oWAH3R_W94gNhC_XUL';

// ── ROUTER ───────────────────────────────────────────────────
function doGet(e) {
  try {
    const page   = e.parameter.page || 'home';
    const roomId = e.parameter.room || '';
    const tmpl   = HtmlService.createTemplateFromFile('App');
    tmpl.initialPage   = page;
    tmpl.initialRoomId = roomId;
    return tmpl.evaluate()
      .setTitle('VidChat — Video Call & Chat')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch(err) {
    return HtmlService.createHtmlOutput(
      '<body style="font-family:Arial;display:flex;align-items:center;justify-content:center;height:100vh;background:#f0f4f8">' +
      '<div style="background:#fff;padding:40px;border-radius:12px;text-align:center;box-shadow:0 4px 20px rgba(0,0,0,.1)">' +
      '<h2 style="color:#ea4335">⚠️ Ralat / Error</h2><p style="color:#555;margin:12px 0">' + err.message + '</p>' +
      '<a href="javascript:history.back()" style="color:#1a73e8">← Kembali / Go Back</a></div></body>'
    ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

// ── SHEET HELPERS ────────────────────────────────────────────
function getOrCreateSheet(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  }
  return sheet;
}
function getSpreadsheet() { return SpreadsheetApp.openById(SPREADSHEET_ID); }

// ── REGISTRATION & LOGIN ─────────────────────────────────────
function registerUser(name, email, password) {
  try {
    email = email.toLowerCase().trim();
    const ss    = getSpreadsheet();
    const sheet = getOrCreateSheet(ss, 'Users',
      ['Name','Email','Password Hash','Registration Date','Status','Session Token']);
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === email)
        return { success: false, message: '⚠️ E-mel ini sudah berdaftar. / Email already registered.' };
    }
    const hash  = Utilities.base64Encode(email + ':' + password + ':vidchat2024');
    const token = Utilities.getUuid();
    sheet.appendRow([name, email, hash, new Date(), 'Active', token]);
    logActivity('DAFTAR', email, 'Pengguna baru berdaftar', '');
    return { success: true, message: '✅ Pendaftaran berjaya! Sila log masuk. / Registration successful!' };
  } catch(err) { return { success: false, message: 'Ralat: ' + err.message }; }
}

function loginUser(email, password) {
  try {
    email = email.toLowerCase().trim();
    const ss    = getSpreadsheet();
    const sheet = ss.getSheetByName('Users');
    if (!sheet) return { success: false, message: '⚠️ Tiada pengguna berdaftar lagi.' };
    const hash  = Utilities.base64Encode(email + ':' + password + ':vidchat2024');
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === email && data[i][2] === hash && data[i][4] === 'Active') {
        const token = Utilities.getUuid();
        sheet.getRange(i + 1, 6).setValue(token);
        logActivity('LOG_MASUK', email, 'Pengguna log masuk', '');
        return { success: true, token, name: data[i][0], email: data[i][1] };
      }
    }
    return { success: false, message: '❌ E-mel atau kata laluan salah. / Wrong email or password.' };
  } catch(err) { return { success: false, message: 'Ralat: ' + err.message }; }
}

function validateToken(email, token) {
  try {
    email = email.toLowerCase().trim();
    const sheet = getSpreadsheet().getSheetByName('Users');
    if (!sheet) return false;
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === email && data[i][5] === token) return true;
    }
    return false;
  } catch(e) { return false; }
}

// ── ROOM MANAGEMENT ──────────────────────────────────────────
function createRoom(email, token, roomName) {
  try {
    if (!validateToken(email, token))
      return { success: false, message: '⛔ Sesi tamat. Sila log masuk semula. / Session expired.' };
    const ss         = getSpreadsheet();
    const sheet      = getOrCreateSheet(ss, 'Rooms',
      ['Room ID','Room Name','Dicipta Oleh','Date-Time','Invite Link','Status','PDF Notes Link','Catatan']);
    const roomId     = generateRoomId();
    const inviteLink = getWebAppUrl() + '?page=join&room=' + roomId;
    sheet.appendRow([roomId, roomName, email, new Date(), inviteLink, 'Aktif', '', '']);
    logActivity('BUAT_BILIK', email, 'Bilik dibuat: ' + roomName, roomId);
    return { success: true, roomId, inviteLink, roomName,
             message: '✅ Bilik berjaya dibuat! / Room created successfully!' };
  } catch(err) { return { success: false, message: 'Ralat: ' + err.message }; }
}

function validateAndJoinRoom(email, token, roomId) {
  try {
    if (!validateToken(email, token))
      return { success: false, message: '⛔ Sila log masuk dahulu. / Please login first.' };
    const sheet = getSpreadsheet().getSheetByName('Rooms');
    if (!sheet) return { success: false, message: '⚠️ Bilik tidak dijumpai. / Room not found.' };
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === roomId && data[i][5] === 'Aktif') {
        logSession(roomId, email, 'MASUK');
        logActivity('MASUK_BILIK', email, 'Masuk bilik: ' + data[i][1], roomId);
        return { success: true, roomId, roomName: data[i][1] };
      }
    }
    return { success: false, message: '⚠️ Bilik tidak wujud atau tidak aktif. / Room not found or inactive.' };
  } catch(err) { return { success: false, message: 'Ralat: ' + err.message }; }
}

function getRoomInfo(roomId) {
  try {
    const sheet = getSpreadsheet().getSheetByName('Rooms');
    if (!sheet) return null;
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === roomId)
        return { roomId: data[i][0], roomName: data[i][1], createdBy: data[i][2], status: data[i][5] };
    }
    return null;
  } catch(e) { return null; }
}

function leaveRoom(email, token, roomId) {
  try {
    if (validateToken(email, token)) {
      logSession(roomId, email, 'KELUAR');
      logActivity('KELUAR_BILIK', email, 'Keluar bilik: ' + roomId, roomId);
    }
    return { success: true };
  } catch(e) { return { success: false }; }
}

// ── NOTES & PDF ───────────────────────────────────────────────
function saveNotes(email, token, roomId, manualNotes, transcript) {
  try {
    if (!validateToken(email, token))
      return { success: false, message: '⛔ Sesi tamat. / Session expired.' };
    const ss         = getSpreadsheet();
    const notesSheet = getOrCreateSheet(ss, 'Notes',
      ['Room ID','Disimpan Oleh','Masa Simpan','Ringkasan Nota','Transkrip','PDF Link']);
    const roomInfo   = getRoomInfo(roomId);
    const roomName   = roomInfo ? roomInfo.roomName : roomId;
    const pdfLink    = createNotesPDF(roomId, roomName, email, manualNotes, transcript);
    notesSheet.appendRow([roomId, email, new Date(),
      manualNotes.substring(0, 500), transcript.substring(0, 1000), pdfLink]);
    const roomsSheet = ss.getSheetByName('Rooms');
    if (roomsSheet) {
      const rData = roomsSheet.getDataRange().getValues();
      for (let i = 1; i < rData.length; i++) {
        if (rData[i][0] === roomId) { roomsSheet.getRange(i + 1, 7).setValue(pdfLink); break; }
      }
    }
    logActivity('SIMPAN_NOTA', email, 'Nota disimpan: ' + roomId, roomId);
    return { success: true, pdfLink, message: '✅ Nota berjaya disimpan sebagai PDF!' };
  } catch(err) { return { success: false, message: 'Ralat simpan nota: ' + err.message }; }
}

function createNotesPDF(roomId, roomName, email, manualNotes, transcript) {
  try {
    const folder  = DriveApp.getFolderById(FOLDER_ID);
    const tz      = Session.getScriptTimeZone();
    const dateStr = Utilities.formatDate(new Date(), tz, 'dd MMM yyyy, HH:mm');
    const fname   = 'Nota_' + roomName.replace(/[^a-zA-Z0-9]/g,'_') + '_' +
                    Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmm');
    const html = `<!DOCTYPE html><html><head><meta charset="UTF-8">
<style>*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Arial,sans-serif;padding:40px;color:#222;background:#fff}
.header{background:linear-gradient(135deg,#1a73e8,#0d47a1);color:white;padding:30px;border-radius:8px;margin-bottom:30px}
.header h1{font-size:24px;margin-bottom:8px}
.meta-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:12px;font-size:13px}
.meta-item{background:rgba(255,255,255,.15);padding:8px 12px;border-radius:4px}
.section{margin-bottom:28px}
.section-title{font-size:16px;font-weight:bold;color:#1a73e8;border-left:4px solid #1a73e8;padding-left:12px;margin-bottom:12px}
.content-box{background:#f8f9fa;border:1px solid #e0e0e0;border-radius:6px;padding:16px;white-space:pre-wrap;font-size:14px;line-height:1.7}
.transcript-box{background:#f0f7ff;border:1px solid #bbdefb;border-radius:6px;padding:16px;white-space:pre-wrap;font-size:13px;line-height:1.8}
.footer{margin-top:40px;padding-top:16px;border-top:1px solid #e0e0e0;font-size:11px;color:#999;text-align:center}
.badge{display:inline-block;background:#e8f5e9;color:#2e7d32;padding:2px 8px;border-radius:12px;font-size:12px;margin:2px}</style>
</head><body>
<div class="header"><h1>📋 Nota Mesyuarat / Meeting Notes</h1>
<p style="opacity:.85;margin-top:4px">VidChat — Video Call & Chat</p>
<div class="meta-grid">
<div class="meta-item"><strong>Bilik / Room:</strong> ${escapeHtml(roomName)}</div>
<div class="meta-item"><strong>Room ID:</strong> ${escapeHtml(roomId)}</div>
<div class="meta-item"><strong>Disimpan oleh:</strong> ${escapeHtml(email)}</div>
<div class="meta-item"><strong>Tarikh / Date:</strong> ${dateStr}</div>
</div></div>
<div class="section"><div class="section-title">📝 Nota Manual / Manual Notes</div>
<div class="content-box">${escapeHtml(manualNotes)||'<em style="color:#999">Tiada nota manual.</em>'}</div></div>
<div class="section"><div class="section-title">🎙️ Transkrip Perbualan
<span class="badge">Bahasa Melayu</span><span class="badge">English</span></div>
<div class="transcript-box">${escapeHtml(transcript)||'<em style="color:#999">Tiada transkrip.</em>'}</div></div>
<div class="footer">Dijana oleh VidChat • ${dateStr} • Bahasa Melayu & English</div>
</body></html>`;
    const htmlFile = folder.createFile(Utilities.newBlob(html, 'text/html', fname + '.html'));
    const pdfFile  = folder.createFile(htmlFile.getAs('application/pdf').setName(fname + '.pdf'));
    htmlFile.setTrashed(true);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return pdfFile.getUrl();
  } catch(err) { Logger.log('PDF error: ' + err.message); return ''; }
}

function escapeHtml(str) {
  if (!str) return '';
  return str.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
            .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

// ── LOGGING ───────────────────────────────────────────────────
function logActivity(action, email, description, roomId) {
  try {
    const sheet = getOrCreateSheet(getSpreadsheet(), 'Log Aktiviti',
      ['Masa','Tindakan','E-mel','Penerangan','Room ID']);
    sheet.appendRow([new Date(), action, email, description, roomId]);
  } catch(e) {}
}

function logSession(roomId, email, action) {
  try {
    const sheet = getOrCreateSheet(getSpreadsheet(), 'Sesi / Sessions',
      ['Room ID','E-mel','Tindakan','Masa']);
    sheet.appendRow([roomId, email, action, new Date()]);
  } catch(e) {}
}

// ── UTILITIES ─────────────────────────────────────────────────
function generateRoomId() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let id = '';
  for (let i = 0; i < 9; i++) {
    if (i === 3 || i === 6) id += '-';
    id += chars[Math.floor(Math.random() * chars.length)];
  }
  return id;
}
