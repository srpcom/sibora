/**
 * @fileoverview Script SIBORA (Sistem Booking Ruangan) FINAL
 * Fitur: Google SSO, Role Mgmt, Blokir User, Log Sistem, Email PIC, Statistik, Pembatalan User, Notifikasi Login + Device, AUTO ARSIP, REPORTING EXCEL.
 */

// --- FUNGSI PEMICU IZIN (Jalankan Sekali Saja) ---
function FIX_PERMISSIONS_RUN_ME() {
  try {
    const testFolder = DriveApp.getFolderById(FOLDER_FOTO_ID); 
    const testFile = testFolder.createFile("temp_permission_check.txt", "Izin OK");
    testFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    testFile.setTrashed(true);
  } catch(e) { console.log(e); }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  console.log("Spreadsheet OK: " + ss.getName());
  UrlFetchApp.fetch("https://www.google.com");
  
  // Memancing izin ScriptApp untuk trigger otomatis
  const triggers = ScriptApp.getProjectTriggers();
  
  Logger.log("SEMUA IZIN BERHASIL DIBERIKAN! SILAKAN DEPLOY ULANG.");
}

// --- KONFIGURASI ---
const SHEET_JADWAL = "Jadwal Ruangan";
const SHEET_ARSIP = "Arsip Booking"; 
const SHEET_LOG = "Log Sistem";
const SHEET_RUANGAN = "Data Ruangan";
const SHEET_USERS = "Data Pengguna"; 
const SUPER_ADMIN_EMAIL = "syamsul18782@gmail.com"; 
const APP_NAME = "SIBORA"; 

// Ganti dengan ID Folder Google Drive Anda untuk menyimpan foto profil
const FOLDER_FOTO_ID = "1S8uXxnArNB9AqNJbXAMN8v0nLpQzp4b9"; 

// --- KONFIGURASI API ---
const TELEGRAM_BOT_TOKEN = "8248296735:AAEgR9yBCVdfL1m6xA8q6hOyFR-dKYoQG88";
const TELEGRAM_GROUP_CHAT_ID = "@pinjamruangan"; 
const TELEGRAM_API_URL = "https://api.telegram.org/bot" + TELEGRAM_BOT_TOKEN;
const STARSENDER_API_KEY = "65162ade-7692-4ac9-9d20-8ff2ec097a08"; 
const ADMIN_WA_NUMBER = "6287834813776"; 

const DEFAULT_ROOMS = ["Gajah Mada", "Ronggolawe", "Aryo Tejo", "Diponegoro", "Komite Keperawatan"];

// --- LOGGER ---
function logActivity(level, message) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_LOG);
    if (!sheet) {
        sheet = ss.insertSheet(SHEET_LOG);
        sheet.appendRow(["Timestamp", "Level", "Pesan"]);
    }
    sheet.appendRow([new Date(), level, message]);
  } catch (e) { Logger.log("Log Error: " + e.toString()); }
}

// --- UTILS ---
function formatAndValidateWaNumber(number) {
    if (!number) return null;
    let str = String(number).replace(/[^\d]/g, '');
    if (str.startsWith('0')) str = '62' + str.substring(1);
    return str.length > 9 ? str : null;
}

function sendTelegram(chatId, text) {
  try {
    const payload = { 'chat_id': String(chatId), 'text': text, 'parse_mode': 'HTML' };
    UrlFetchApp.fetch(TELEGRAM_API_URL + '/sendMessage', { 'method': 'post', 'contentType': 'application/json', 'payload': JSON.stringify(payload), 'muteHttpExceptions': true });
  } catch (e) { logActivity("ERROR", "Telegram fail: " + e.toString()); }
}

function sendWhatsApp(number, message) {
    const fmtNum = formatAndValidateWaNumber(number);
    if (!fmtNum) return;
    const waMessage = message.replace(/<b>/g, '*').replace(/<\/b>/g, '*').replace(/<i>/g, '_').replace(/<\/i>/g, '_').replace(/<br>/g, '\n');
    try {
        UrlFetchApp.fetch('https://api.starsender.online/api/send', {
            'method': 'post', 'contentType': 'application/json',
            'headers': { 'Authorization': STARSENDER_API_KEY },
            'payload': JSON.stringify({ "messageType": "text", "to": fmtNum, "body": waMessage }),
            'muteHttpExceptions': true
        });
    } catch (e) { logActivity("ERROR", "WA fail: " + e.toString()); }
}

function createEmailTemplate(title, content, actionLink = null, actionText = null) {
    const year = new Date().getFullYear();
    const btnHtml = actionLink ? `<div style="text-align: center; margin-top: 25px;"><a href="${actionLink}" style="background-color: #4f46e5; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; font-weight: bold; font-size: 14px; display: inline-block;">${actionText}</a></div>` : '';
    return `<!DOCTYPE html><html><head><style>body{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background-color:#f3f4f6;margin:0;padding:0}.container{max-width:600px;margin:20px auto;background-color:#ffffff;border-radius:12px;overflow:hidden;box-shadow:0 4px 6px rgba(0,0,0,0.05);border:1px solid #e5e7eb}.header{background:linear-gradient(135deg,#4f46e5 0%,#4338ca 100%);color:white;padding:25px;text-align:center}.header h1{margin:0;font-size:22px;font-weight:700}.content{padding:30px;color:#374151;font-size:15px;line-height:1.6}.info-box{background-color:#f9fafb;border-left:4px solid #4f46e5;padding:15px;margin:20px 0;border-radius:4px}.info-row{margin-bottom:8px;display:flex}.info-label{font-weight:600;width:120px;color:#6b7280;flex-shrink:0}.info-val{color:#111827;font-weight:500}.status-badge{display:inline-block;padding:4px 12px;border-radius:50px;font-size:12px;font-weight:700;text-transform:uppercase}.status-approved{background-color:#d1fae5;color:#065f46}.status-rejected{background-color:#fee2e2;color:#991b1b}.status-pending{background-color:#fef3c7;color:#92400e}.reason-box{background-color:#fff1f2;color:#9f1239;padding:10px;border-radius:6px;border:1px solid #fecdd3;margin-top:15px;font-weight:500}.footer{background-color:#f9fafb;padding:20px;text-align:center;color:#9ca3af;font-size:12px;border-top:1px solid #e5e7eb}</style></head><body><div class="container"><div class="header"><h1>SIBORA</h1><div style="font-size:13px;opacity:0.9;margin-top:5px;">RSUD dr. R. Koesma Tuban</div></div><div class="content"><h2 style="color:#1f2937;margin-top:0;font-size:18px;">${title}</h2>${content}${btnHtml}</div><div class="footer"><p>&copy; ${year} RSUD dr. R. Koesma Tuban.<br>Sistem Booking Ruangan (SIBORA)</p></div></div></body></html>`;
}

function sendEmailNotification(to, subject, htmlBody) {
    if (!to || !to.includes("@")) return;
    try { MailApp.sendEmail({ to: to, subject: subject, htmlBody: htmlBody, name: APP_NAME }); } catch (e) { logActivity("ERROR", `Email fail: ${e.toString()}`); }
}

// --- API GATEWAY ---
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    let result;
    switch (payload.action) {
      case 'google_login': result = handleGoogleLogin(payload); break;
      case 'save_profile': result = handleSaveProfile(payload); break;
      case 'upload_foto': result = handleUploadFoto(payload); break; 
      case 'get_users': result = handleGetUsers(payload); break;
      case 'update_user_role': result = handleUpdateUserRole(payload); break;
      case 'addRoom': result = handleAddRoom(payload); break; 
      case 'editRoom': result = handleEditRoom(payload); break; 
      case 'deleteRoom': result = handleDeleteRoom(payload); break;
      case 'saveBackground': result = handleSaveBackground(payload); break;
      case 'deleteBackground': result = handleDeleteBackground(payload); break;
      case 'deleteBooking': result = handleDeleteBooking(payload); break;
      case 'cancelBooking': result = handleCancelBooking(payload); break; 
      case 'archiveOldBookings': result = handleArchiveOldBookings(payload); break; 
      case 'updateStatus': result = handleUpdateStatus(payload); break;
      case 'getStatistics': result = handleGetStatistics(); break;
      case 'getReportData': result = handleGetReportData(payload); break; // FITUR BARU: REPORT
      default: result = handleNewBooking(payload);
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_JADWAL);
    
    // Auto-fix header
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headers.includes("Catatan")) sheet.getRange(1, headers.length + 1).setValue("Catatan");
    const lastCol = sheet.getLastColumn();
    const headersV2 = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    if (!headersV2.includes("Kelengkapan")) sheet.getRange(1, lastCol + 1).setValue("Kelengkapan");

    const data = sheet.getRange(2, 1, Math.max(sheet.getLastRow()-1, 1), sheet.getLastColumn()).getValues();
    const headersUpdated = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const jsonArray = data.filter(r => r[0]).map(row => {
      let obj = {};
      headersUpdated.forEach((h, i) => {
        obj[h] = (row[i] instanceof Date) ? Utilities.formatDate(row[i], ss.getSpreadsheetTimeZone(), (h.includes("Tanggal") ? "yyyy-MM-dd" : "HH:mm")) : row[i];
      });
      return obj;
    });

    let sheetRuangan = ss.getSheetByName(SHEET_RUANGAN);
    let roomsList = [];
    if (sheetRuangan && sheetRuangan.getLastRow() > 1) {
        const numRows = sheetRuangan.getLastRow() - 1;
        const rawRooms = sheetRuangan.getRange(2, 1, numRows, 4).getValues(); 
        roomsList = rawRooms.map(r => ({
            name: r[0],
            pic: r[1] || "",
            wa: r[2] || "",
            email: r[3] || "" 
        }));
    } else {
        roomsList = DEFAULT_ROOMS.map(r => ({ name: r, pic: "", wa: "", email: "" }));
    }

    const bg = PropertiesService.getScriptProperties().getProperty('APP_BACKGROUND') || "";

    return ContentService.createTextOutput(JSON.stringify({ status: "success", data: jsonArray, rooms: roomsList, background: bg })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

// --- LOGIKA DATA RUANGAN (Update PIC & Email) ---
function handleAddRoom(p) { 
    const ss=SpreadsheetApp.getActiveSpreadsheet(); 
    const s=ss.getSheetByName(SHEET_RUANGAN); 
    const exist=s.getDataRange().getValues().map(r => r[0]); 
    if(exist.includes(p.roomName)) return {status:'error',message:'Ruangan sudah ada'}; 
    const picWa = formatAndValidateWaNumber(p.picWa) || p.picWa;
    s.appendRow([p.roomName, p.picName || "", picWa || "", p.picEmail || ""]); 
    return {status:'success', message:'Ruangan ditambahkan'}; 
}

function handleEditRoom(p) { 
    const ss=SpreadsheetApp.getActiveSpreadsheet(); 
    const s=ss.getSheetByName(SHEET_RUANGAN); 
    const data=s.getDataRange().getValues(); 
    for(let i=1; i<data.length; i++){ 
        if(data[i][0] === p.oldName) { 
            s.getRange(i+1, 1).setValue(p.newName); 
            s.getRange(i+1, 2).setValue(p.picName); 
            const picWa = formatAndValidateWaNumber(p.picWa) || p.picWa;
            s.getRange(i+1, 3).setValue(picWa);
            s.getRange(i+1, 4).setValue(p.picEmail || ""); 
            return {status:'success', message:'Data ruangan diperbarui'}; 
        } 
    } 
    return {status:'error', message:'Ruangan tidak ditemukan'}; 
}

function handleDeleteRoom(p) { 
    const ss=SpreadsheetApp.getActiveSpreadsheet(); 
    const s=ss.getSheetByName(SHEET_RUANGAN); 
    const d=s.getDataRange().getValues(); 
    for(let i=1;i<d.length;i++){ 
        if(d[i][0]===p.roomName){
            s.deleteRow(i+1); 
            return {status:'success', message:'Ruangan dihapus'};
        }
    } 
    return {status:'error', message:'Gagal menghapus'}; 
}

function getRoomPicEmail(roomName) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const s = ss.getSheetByName(SHEET_RUANGAN);
        const data = s.getDataRange().getValues();
        for(let i=1; i<data.length; i++) {
            if (data[i][0] === roomName) {
                return data[i][3]; 
            }
        }
    } catch(e) { logActivity("WARN", "Gagal fetch email PIC: " + e.toString()); }
    return null;
}

// --- LOGIKA BOOKING BARU ---
function handleNewBooking(data) { 
    const ss = SpreadsheetApp.getActiveSpreadsheet(); 
    const userSheet = ss.getSheetByName(SHEET_USERS); 
    const users = userSheet.getDataRange().getValues(); 
    let userData = null; 
    for(let i=1; i<users.length; i++){ 
        if(String(users[i][0]).toLowerCase() === String(data.email).toLowerCase()){ 
            userData = { nama: users[i][1], unit: users[i][2], wa: users[i][3] }; 
            break; 
        } 
    } 
    if(!userData) return { status: 'error', message: 'User tidak valid.' }; 
    
    // --- VALIDASI SERVER-SIDE (KEAMANAN) ---
    const reqStart = new Date(data.tanggalMulai + 'T' + data.waktuMulai);
    const reqEnd = new Date(data.tanggalSelesai + 'T' + data.waktuSelesai);
    const now = new Date();
    
    // 1. Cek Waktu Lampau (Toleransi 2 menit untuk latensi jaringan)
    if (reqStart < new Date(now.getTime() - 2 * 60 * 1000)) {
        return { status: 'error', message: 'Gagal: Waktu mulai sudah berlalu. Harap refresh halaman.' };
    }
    
    // 2. Cek Urutan Waktu
    if (reqEnd <= reqStart) {
        return { status: 'error', message: 'Gagal: Waktu selesai harus lebih akhir dari waktu mulai.' };
    }

    const sheet = ss.getSheetByName(SHEET_JADWAL); 
    const allData = sheet.getDataRange().getValues(); const headers = allData[0]; 
    const roomIdx = headers.indexOf("Nama Ruangan"); const dateStartIdx = headers.indexOf("Tanggal Mulai"); const timeStartIdx = headers.indexOf("Waktu Mulai"); const dateEndIdx = headers.indexOf("Tanggal Selesai"); const timeEndIdx = headers.indexOf("Waktu Selesai"); const statusIdx = headers.indexOf("Status");
    
    for (let i = 1; i < allData.length; i++) { 
        const row = allData[i]; 
        if (row[roomIdx] === data.namaRuangan && row[statusIdx] !== 'Ditolak' && row[statusIdx] !== 'Dibatalkan') { 
            const exStart = new Date(((row[dateStartIdx] instanceof Date)?Utilities.formatDate(row[dateStartIdx], ss.getSpreadsheetTimeZone(), "yyyy-MM-dd"):row[dateStartIdx]) + 'T' + ((row[timeStartIdx] instanceof Date)?Utilities.formatDate(row[timeStartIdx], ss.getSpreadsheetTimeZone(), "HH:mm"):row[timeStartIdx])); 
            const exEnd = new Date(((row[dateEndIdx] instanceof Date)?Utilities.formatDate(row[dateEndIdx], ss.getSpreadsheetTimeZone(), "yyyy-MM-dd"):row[dateEndIdx]) + 'T' + ((row[timeEndIdx] instanceof Date)?Utilities.formatDate(row[timeEndIdx], ss.getSpreadsheetTimeZone(), "HH:mm"):row[timeEndIdx])); 
            if (reqStart < exEnd && reqEnd > exStart) return { status: 'error', message: `Maaf, Ruangan ${data.namaRuangan} sudah terisi pada jam tersebut.` }; 
        } 
    }
    
    const tz = ss.getSpreadsheetTimeZone(); const idBooking = `BOOK-${Date.now()}`;
    let equipIdx = headers.indexOf("Kelengkapan"); if (equipIdx === -1) { equipIdx = headers.length; sheet.getRange(1, equipIdx + 1).setValue("Kelengkapan"); } let noteIdx = headers.indexOf("Catatan"); if (noteIdx === -1) { noteIdx = headers.length + (equipIdx === headers.length ? 1 : 0); sheet.getRange(1, noteIdx + 1).setValue("Catatan"); }
    const newRow = []; const headerMap = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    headerMap.forEach(h => { if(h === "ID Booking") newRow.push(idBooking); else if(h === "Timestamp") newRow.push(new Date()); else if(h === "Nama Ruangan") newRow.push(data.namaRuangan); else if(h === "Nama Peminjam") newRow.push(userData.nama); else if(h === "Unit/Bagian") newRow.push(userData.unit); else if(h === "Nomor WA") newRow.push(userData.wa); else if(h === "Tanggal Mulai") newRow.push(data.tanggalMulai); else if(h === "Waktu Mulai") newRow.push(data.waktuMulai); else if(h === "Tanggal Selesai") newRow.push(data.tanggalSelesai); else if(h === "Waktu Selesai") newRow.push(data.waktuSelesai); else if(h === "Keperluan") newRow.push(data.keperluan); else if(h === "Status") newRow.push("Belum Disetujui"); else if(h === "Catatan") newRow.push(""); else if(h === "Kelengkapan") newRow.push(data.kelengkapan || "-"); else newRow.push(""); });
    sheet.appendRow(newRow);
    
    const startDate = Utilities.formatDate(new Date(data.tanggalMulai), tz, "EEEE, dd MMMM yyyy"); const endDate = Utilities.formatDate(new Date(data.tanggalSelesai), tz, "EEEE, dd MMMM yyyy");
    const kelengkapanMsg = data.kelengkapan ? `\n<b>Kelengkapan:</b> ${data.kelengkapan}` : "";
    
    const adminMsg = `<b>🔔 PENGAJUAN BARU</b>\n\n<b>Oleh:</b> ${userData.nama} (${userData.unit})\n<b>Ruangan:</b> ${data.namaRuangan}\n<b>Keperluan:</b> ${data.keperluan}${kelengkapanMsg}\n<b>Waktu:</b> ${startDate} ${data.waktuMulai} s/d ${endDate} ${data.waktuSelesai}\n\n<i>Cek aplikasi untuk persetujuan.</i>`;
    sendTelegram(TELEGRAM_GROUP_CHAT_ID, adminMsg); sendWhatsApp(ADMIN_WA_NUMBER, adminMsg); sendWhatsApp(userData.wa, `Halo ${userData.nama}, pengajuan booking ruang ${data.namaRuangan} berhasil dikirim. Menunggu persetujuan.`);
    
    const adminEmailContent = `<p>Terdapat pengajuan booking ruangan baru.</p><div class="info-box"><div class="info-row"><span class="info-label">Peminjam</span><span class="info-val">${userData.nama}</span></div><div class="info-row"><span class="info-label">Unit/Bagian</span><span class="info-val">${userData.unit}</span></div><div class="info-row"><span class="info-label">Ruangan</span><span class="info-val">${data.namaRuangan}</span></div><div class="info-row"><span class="info-label">Waktu</span><span class="info-val">${startDate} ${data.waktuMulai} - ${data.waktuSelesai}</span></div><div class="info-row"><span class="info-label">Keperluan</span><span class="info-val">${data.keperluan}</span></div><div class="info-row"><span class="info-label">Kelengkapan</span><span class="info-val">${data.kelengkapan || "-"}</span></div></div>`; 
    sendEmailNotification(SUPER_ADMIN_EMAIL, "🔔 Pengajuan Booking Baru: " + data.namaRuangan, createEmailTemplate("Pengajuan Booking Baru", adminEmailContent, "https://srpcom.github.io/bookingruangan/#admin", "Buka Dashboard Admin"));
    
    const userEmailContent = `<p>Halo <b>${userData.nama}</b>,</p><p>Pengajuan booking ruangan Anda telah diterima dan <b>menunggu persetujuan admin</b>.</p><div class="info-box"><div class="info-row"><span class="info-label">Ruangan</span><span class="info-val">${data.namaRuangan}</span></div><div class="info-row"><span class="info-label">Waktu</span><span class="info-val">${startDate} ${data.waktuMulai}</span></div><div class="info-row"><span class="info-label">Keperluan</span><span class="info-val">${data.keperluan}</span></div></div>`; 
    sendEmailNotification(data.email, "⏳ Status Pengajuan: Menunggu Persetujuan", createEmailTemplate("Pengajuan Diterima", userEmailContent));
    
    const picEmail = getRoomPicEmail(data.namaRuangan);
    if (picEmail) {
        const picContent = `<p>Halo Petugas <b>${data.namaRuangan}</b>,</p>
        <p>Terdapat pengajuan booking baru untuk ruangan yang Anda kelola. Mohon monitor status persetujuan dari Admin.</p>
        <div class="info-box">
            <div class="info-row"><span class="info-label">Peminjam</span><span class="info-val">${userData.nama} (${userData.unit})</span></div>
            <div class="info-row"><span class="info-label">Waktu</span><span class="info-val">${startDate} ${data.waktuMulai} s.d ${data.waktuSelesai}</span></div>
            <div class="info-row"><span class="info-label">Acara</span><span class="info-val">${data.keperluan}</span></div>
            <div class="info-row"><span class="info-label">Kelengkapan</span><span class="info-val">${data.kelengkapan || "-"}</span></div>
        </div>
        <p style="font-size:12px; color:#666;">*Anda menerima email ini karena terdaftar sebagai PIC Ruangan.</p>`;
        sendEmailNotification(picEmail, `[INFO PIC] Booking Baru: ${data.namaRuangan}`, createEmailTemplate("Info Booking Baru", picContent));
    }

    return { status: "success", message: "Booking berhasil dikirim!" };
}

// --- LOGIKA UPDATE STATUS ---
function handleUpdateStatus(payload) { 
    const ss = SpreadsheetApp.getActiveSpreadsheet(); 
    const sheet = ss.getSheetByName(SHEET_JADWAL); 
    const data = sheet.getDataRange().getValues(); 
    const headers = data[0]; 
    const idIdx = headers.indexOf("ID Booking"); 
    const statusIdx = headers.indexOf("Status"); 
    let noteIdx = headers.indexOf("Catatan"); 
    if (noteIdx === -1) { noteIdx = headers.length; sheet.getRange(1, noteIdx + 1).setValue("Catatan"); } 
    
    for (let i = 1; i < data.length; i++) { 
        if (data[i][idIdx] === payload.bookingId) { 
            sheet.getRange(i + 1, statusIdx + 1).setValue(payload.newStatus); 
            if (payload.reason) sheet.getRange(i + 1, noteIdx + 1).setValue(payload.reason); else sheet.getRange(i + 1, noteIdx + 1).setValue(""); 
            
            const booking = {}; headers.forEach((h, k) => booking[h] = data[i][k]); 
            const userSheet = ss.getSheetByName(SHEET_USERS); 
            const users = userSheet.getDataRange().getValues(); 
            let userEmail = ""; 
            for(let u=1; u<users.length; u++) { if(users[u][1] === booking["Nama Peminjam"] && users[u][2] === booking["Unit/Bagian"]) { userEmail = users[u][0]; break; } } 
            
            const statusText = payload.newStatus.toUpperCase(); 
            const emoji = statusText === 'DISETUJUI' ? '✅' : '❌'; 
            
            const waMsg = `Halo ${booking["Nama Peminjam"]}, status booking *${booking["Nama Ruangan"]}* Anda: *${statusText}* ${emoji}`; 
            sendWhatsApp(booking["Nomor WA"], waMsg); 
            sendTelegram(TELEGRAM_GROUP_CHAT_ID, `Status Booking ${booking["Nama Peminjam"]} diubah menjadi ${statusText}`); 
            
            if(userEmail) { 
                const statusColorClass = statusText === 'DISETUJUI' ? 'status-approved' : 'status-rejected'; 
                const statusTitle = statusText === 'DISETUJUI' ? 'Pengajuan Disetujui' : 'Pengajuan Ditolak'; 
                let reasonHtml = (statusText === 'DITOLAK' && payload.reason) ? `<div class="reason-box">Alasan Penolakan:<br>${payload.reason}</div>` : ""; 
                const updateEmailContent = `<p>Halo <b>${booking["Nama Peminjam"]}</b>,</p><p>Status pengajuan booking ruangan Anda telah diperbarui.</p><div style="text-align:center; margin: 20px 0;"><span class="status-badge ${statusColorClass}" style="font-size: 16px; padding: 8px 20px;">${statusText}</span></div>${reasonHtml}<div class="info-box"><div class="info-row"><span class="info-label">Ruangan</span><span class="info-val">${booking["Nama Ruangan"]}</span></div><div class="info-row"><span class="info-label">Waktu</span><span class="info-val">${Utilities.formatDate(new Date(booking["Tanggal Mulai"]), ss.getSpreadsheetTimeZone(), "dd-MM-yyyy")} ${booking["Waktu Mulai"]}</span></div><div class="info-row"><span class="info-label">Keperluan</span><span class="info-val">${booking["Keperluan"]}</span></div></div>`; 
                sendEmailNotification(userEmail, `[SIBORA] Update Status: ${statusText}`, createEmailTemplate(statusTitle, updateEmailContent)); 
            } 

            const picEmail = getRoomPicEmail(booking["Nama Ruangan"]);
            if (picEmail) {
                const statusColor = statusText === 'DISETUJUI' ? '#d1fae5' : '#fee2e2';
                const statusTextColor = statusText === 'DISETUJUI' ? '#065f46' : '#991b1b';
                const picUpdateContent = `<p>Halo Petugas <b>${booking["Nama Ruangan"]}</b>,</p><p>Status booking untuk ruangan Anda telah diperbarui oleh Admin.</p><div style="text-align:center; margin: 20px 0;"><span style="background-color:${statusColor}; color:${statusTextColor}; padding: 8px 20px; border-radius: 50px; font-weight: bold; font-size: 14px;">${statusText}</span></div><div class="info-box"><div class="info-row"><span class="info-label">Peminjam</span><span class="info-val">${booking["Nama Peminjam"]} (${booking["Unit/Bagian"]})</span></div><div class="info-row"><span class="info-label">Waktu</span><span class="info-val">${Utilities.formatDate(new Date(booking["Tanggal Mulai"]), ss.getSpreadsheetTimeZone(), "dd-MM-yyyy")} ${booking["Waktu Mulai"]}</span></div><div class="info-row"><span class="info-label">Acara</span><span class="info-val">${booking["Keperluan"]}</span></div></div><p>Silakan siapkan ruangan jika disetujui, atau batalkan persiapan jika ditolak/dibatalkan.</p>`;
                sendEmailNotification(picEmail, `[INFO PIC] Status Booking ${statusText}: ${booking["Nama Ruangan"]}`, createEmailTemplate(`Update Status: ${statusText}`, picUpdateContent));
            }

            return { status: "success", message: `Status diubah ke ${payload.newStatus}` }; 
        } 
    } 
    return { status: "error", message: "ID tidak ditemukan" }; 
}

// --- LOGIKA PEMBATALAN USER ---
function handleCancelBooking(p) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_JADWAL);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIdx = headers.indexOf("ID Booking");
    const statusIdx = headers.indexOf("Status");
    const roomIdx = headers.indexOf("Nama Ruangan");
    const nameIdx = headers.indexOf("Nama Peminjam");
    const dateIdx = headers.indexOf("Tanggal Mulai");
    const timeIdx = headers.indexOf("Waktu Mulai");
    
    let noteIdx = headers.indexOf("Catatan");
    if (noteIdx === -1) { noteIdx = headers.length; sheet.getRange(1, noteIdx + 1).setValue("Catatan"); }

    for (let i = 1; i < data.length; i++) {
        if (data[i][idIdx] === p.bookingId && data[i][nameIdx] === p.userName) {
            
            // VALIDASI WAKTU: H-60 MENIT
            let bookingDate = data[i][dateIdx];
            let bookingTimeStr = data[i][timeIdx];
            
            if (bookingDate instanceof Date) {
                // Konversi string "HH:mm" ke angka jam & menit
                let parts = String(bookingTimeStr).split(':');
                if(parts.length >= 2) {
                    bookingDate.setHours(parseInt(parts[0]), parseInt(parts[1]), 0, 0);
                }
                
                const now = new Date();
                // Hitung selisih dalam menit
                const diffMs = bookingDate - now;
                const diffMins = diffMs / 1000 / 60;

                if (diffMins < 60) {
                    return { status: 'error', message: 'Gagal Membatalkan: Waktu acara kurang dari 60 menit.' };
                }
            }

            // Update Status
            sheet.getRange(i + 1, statusIdx + 1).setValue("Dibatalkan");
            const oldNote = data[i][noteIdx] || "";
            sheet.getRange(i + 1, noteIdx + 1).setValue(`[DIBATALKAN USER] Alasan: ${p.reason}. ${oldNote}`);

            const roomName = data[i][roomIdx];
            const bookingInfo = {
                room: roomName,
                user: data[i][nameIdx],
                unit: data[i][headers.indexOf("Unit/Bagian")],
                date: Utilities.formatDate(new Date(data[i][dateIdx]), ss.getSpreadsheetTimeZone(), "dd-MM-yyyy"),
                time: data[i][timeIdx]
            };

            // Notifikasi ke Telegram
            const msgTg = `<b>🚫 BOOKING DIBATALKAN (USER)</b>\n\n<b>Oleh:</b> ${bookingInfo.user}\n<b>Ruangan:</b> ${bookingInfo.room}\n<b>Waktu:</b> ${bookingInfo.date} ${bookingInfo.time}\n<b>Alasan:</b> ${p.reason}`;
            sendTelegram(TELEGRAM_GROUP_CHAT_ID, msgTg);

            // Notifikasi ke PIC
            const picEmail = getRoomPicEmail(roomName);
            if (picEmail) {
                const picHtml = `<div style="font-family:sans-serif; border:1px solid #fee2e2; border-left:5px solid #ef4444; padding:20px; border-radius:8px;"><h2 style="color:#991b1b; margin-top:0;">🚫 Booking Dibatalkan</h2><p>Halo PIC Ruangan <b>${roomName}</b>,</p><p>User telah membatalkan acara berikut secara mandiri. Mohon <b>JANGAN</b> siapkan ruangan.</p><ul><li><b>Peminjam:</b> ${bookingInfo.user} (${bookingInfo.unit})</li><li><b>Tanggal:</b> ${bookingInfo.date}</li><li><b>Jam:</b> ${bookingInfo.time}</li><li><b>Alasan:</b> ${p.reason}</li></ul></div>`;
                sendEmailNotification(picEmail, `[BATAL] Booking Ruang ${roomName}`, picHtml);
            }

            return { status: 'success', message: 'Booking berhasil dibatalkan.' };
        }
    }
    return { status: 'error', message: 'Data tidak ditemukan.' };
}

function handleGetStatistics() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_JADWAL);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const roomIdx = headers.indexOf("Nama Ruangan");
    const unitIdx = headers.indexOf("Unit/Bagian");
    const timeIdx = headers.indexOf("Waktu Mulai");
    const statusIdx = headers.indexOf("Status");
    let roomStats = {}; let unitStats = {}; let hourStats = {};
    for(let h=7; h<=16; h++) { hourStats[String(h).padStart(2, '0')+':00'] = 0; }
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const status = row[statusIdx];
        if (status !== 'Ditolak' && status !== 'Dibatalkan') {
            const room = row[roomIdx]; roomStats[room] = (roomStats[room] || 0) + 1;
            const unit = row[unitIdx]; unitStats[unit] = (unitStats[unit] || 0) + 1;
            let timeVal = row[timeIdx]; if (timeVal instanceof Date) { timeVal = Utilities.formatDate(timeVal, ss.getSpreadsheetTimeZone(), "HH:mm"); }
            const hourPrefix = String(timeVal).split(':')[0] + ':00'; hourStats[hourPrefix] = (hourStats[hourPrefix] || 0) + 1;
        }
    }
    return { status: 'success', roomStats: roomStats, unitStats: unitStats, hourStats: hourStats };
}

function handleGetReportData(p) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Ambil data dari Jadwal Ruangan (booking aktif)
    const sheetJadwal = ss.getSheetByName(SHEET_JADWAL);
    let dataJadwal = [];
    let headers = [];
    if (sheetJadwal) {
        const values = sheetJadwal.getDataRange().getValues();
        if (values.length > 0) {
            headers = values[0];
            if (values.length > 1) {
                dataJadwal = values.slice(1);
            }
        }
    }
    
    // Ambil data dari Arsip Booking
    const sheetArsip = ss.getSheetByName(SHEET_ARSIP);
    let dataArsip = [];
    if (sheetArsip) {
        const values = sheetArsip.getDataRange().getValues();
        if (values.length > 1) {
            dataArsip = values.slice(1);
        }
    }
    
    if (headers.length === 0) {
        return { status: 'error', message: 'Sheet Jadwal tidak ditemukan atau kosong' };
    }
    
    // Indeks Kolom
    const dateStartIdx = headers.indexOf("Tanggal Mulai");
    const roomIdx = headers.indexOf("Nama Ruangan");
    const nameIdx = headers.indexOf("Nama Peminjam");
    const unitIdx = headers.indexOf("Unit/Bagian");
    const activityIdx = headers.indexOf("Keperluan");
    const timeStartIdx = headers.indexOf("Waktu Mulai");
    const timeEndIdx = headers.indexOf("Waktu Selesai");
    const statusIdx = headers.indexOf("Status");

    // Gabungkan data aktif dan data arsip
    const allRows = dataJadwal.concat(dataArsip);
    let reportData = [];

    const tz = ss.getSpreadsheetTimeZone();

    for (let i = 0; i < allRows.length; i++) {
        const row = allRows[i];
        let rowDate = row[dateStartIdx];
        if (rowDate) {
            const rowDateStr = formatDateToYYYYMMDD(rowDate, tz);
            if (rowDateStr && rowDateStr >= p.startDate && rowDateStr <= p.endDate) {
                // Formatting Date & Time
                const dateParts = rowDateStr.split('-');
                const formattedDate = `${dateParts[2]}-${dateParts[1]}-${dateParts[0]}`; // dd-MM-yyyy
                
                let tStart = row[timeStartIdx];
                if (tStart instanceof Date) tStart = Utilities.formatDate(tStart, tz, "HH:mm");
                let tEnd = row[timeEndIdx];
                if (tEnd instanceof Date) tEnd = Utilities.formatDate(tEnd, tz, "HH:mm");
                
                reportData.push({
                    "Tanggal": formattedDate,
                    "Waktu": `${tStart} - ${tEnd}`,
                    "Ruangan": row[roomIdx],
                    "Peminjam": row[nameIdx],
                    "Unit": row[unitIdx],
                    "Acara": row[activityIdx],
                    "Status": row[statusIdx]
                });
            }
        }
    }

    // Urutkan laporan berdasarkan Tanggal (menaik/ascending)
    reportData.sort((a, b) => {
        const dateA = a.Tanggal.split('-').reverse().join('-');
        const dateB = b.Tanggal.split('-').reverse().join('-');
        if (dateA !== dateB) return dateA.localeCompare(dateB);
        // Jika tanggal sama, urutkan berdasarkan waktu mulai
        return a.Waktu.localeCompare(b.Waktu);
    });

    return { status: 'success', data: reportData };
}

// Fungsi pembantu untuk memformat tanggal ke format YYYY-MM-DD secara aman
function formatDateToYYYYMMDD(val, timeZone) {
    if (val instanceof Date) {
        return Utilities.formatDate(val, timeZone, "yyyy-MM-dd");
    }
    if (!val) return "";
    const str = String(val).trim();
    const parts = str.split(/[-/]/);
    if (parts.length === 3) {
        if (parts[0].length === 4) {
            // YYYY-MM-DD
            const y = parts[0];
            const m = ("0" + parts[1]).slice(-2);
            const d = ("0" + parts[2]).slice(-2);
            return y + "-" + m + "-" + d;
        } else {
            // DD-MM-YYYY atau DD/MM/YYYY
            const d = ("0" + parts[0]).slice(-2);
            const m = ("0" + parts[1]).slice(-2);
            const y = parts[2];
            return y + "-" + m + "-" + d;
        }
    }
    return "";
}


function handleUploadFoto(data) { try { if (!FOLDER_FOTO_ID) throw new Error("ID Folder Foto belum dikonfigurasi."); const folder = DriveApp.getFolderById(FOLDER_FOTO_ID); const decoded = Utilities.base64Decode(data.base64); const fileName = `FOTO_${data.email}_${Date.now()}.jpg`; const blob = Utilities.newBlob(decoded, data.mimeType, fileName); const file = folder.createFile(blob); try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){} return { status: 'success', url: "https://lh3.googleusercontent.com/d/" + file.getId() }; } catch (e) { return { status: 'error', message: "Gagal upload: " + e.toString() }; } }

function handleGoogleLogin(data) { 
    const ss = SpreadsheetApp.getActiveSpreadsheet(); 
    let sheet = ss.getSheetByName(SHEET_USERS); 
    if (!sheet) { setup(); sheet = ss.getSheetByName(SHEET_USERS); } 
    
    const email = String(data.email).toLowerCase().trim(); 
    const rows = sheet.getDataRange().getValues(); 
    const deviceInfo = data.deviceInfo || {};
    const deviceStr = `📱 <b>Device:</b> ${deviceInfo.userAgent || 'Unknown'}\n🌐 <b>IP:</b> ${deviceInfo.ip || 'Unknown'}`;

    // 1. Cek Existing User
    for (let i = 1; i < rows.length; i++) { 
        if (String(rows[i][0]).toLowerCase() === email) { 
            // Case A: User Terblokir
            if (rows[i][4] === 'BLOCKED') {
                logActivity("WARNING", `Login DITOLAK (BLOCKED): ${email}. IP: ${deviceInfo.ip}`); 
                sendTelegram(TELEGRAM_GROUP_CHAT_ID, `⚠️ <b>LOGIN DITOLAK (BLOCKED)</b>\nUser: ${email}\nNama: ${rows[i][1]}\n${deviceStr}`);
                return { status: 'blocked' }; 
            }
            // Case B: Login Berhasil
            logActivity("INFO", `Login Berhasil: ${email}. IP: ${deviceInfo.ip}`);
            sendTelegram(TELEGRAM_GROUP_CHAT_ID, `✅ <b>LOGIN BERHASIL</b>\nUser: ${rows[i][1]} (${rows[i][2]})\nEmail: ${email}\n${deviceStr}`);
            return { status: 'success', user: { email: rows[i][0], nama: rows[i][1], unit: rows[i][2], wa: rows[i][3], role: rows[i][4], foto: rows[i][5] } }; 
        } 
    } 
    
    // 2. Cek Super Admin Baru (Hardcoded)
    if (email === SUPER_ADMIN_EMAIL.toLowerCase()) { 
        sheet.appendRow([email, "Super Admin", "IT", ADMIN_WA_NUMBER, "ADMIN", data.photo || ""]); 
        logActivity("INFO", `Super Admin dibuat/login: ${email}. IP: ${deviceInfo.ip}`);
        sendTelegram(TELEGRAM_GROUP_CHAT_ID, `👑 <b>SUPER ADMIN LOGIN</b>\nEmail: ${email}\n${deviceStr}`);
        return { status: 'success', user: { email: email, nama: "Super Admin", unit: "IT", wa: ADMIN_WA_NUMBER, role: "ADMIN", foto: data.photo } }; 
    } 
    
    // 3. User Baru (Try Login / Signup)
    logActivity("INFO", `User Baru Mencoba Login (Setup): ${email}. IP: ${deviceInfo.ip}`);
    sendTelegram(TELEGRAM_GROUP_CHAT_ID, `🆕 <b>USER BARU MENCOBA LOGIN</b>\nEmail: ${email}\nStatus: Menunggu Setup Profil\n${deviceStr}`);
    return { status: 'new_user', email: email, photo: data.photo }; 
}

function handleSaveProfile(data) { const ss = SpreadsheetApp.getActiveSpreadsheet(); const sheet = ss.getSheetByName(SHEET_USERS); const email = String(data.email).toLowerCase().trim(); const rows = sheet.getDataRange().getValues(); let rowIndex = -1; for (let i = 1; i < rows.length; i++) { if (String(rows[i][0]).toLowerCase() === email) { rowIndex = i + 1; break; } } const wa = formatAndValidateWaNumber(data.wa) || data.wa; if (rowIndex === -1) { const role = (email === SUPER_ADMIN_EMAIL.toLowerCase()) ? "ADMIN" : "USER"; sheet.appendRow([email, data.nama, data.unit, wa, role, data.foto]); } else { sheet.getRange(rowIndex, 2).setValue(data.nama); sheet.getRange(rowIndex, 3).setValue(data.unit); sheet.getRange(rowIndex, 4).setValue(wa); if (data.foto && data.foto.length > 5) sheet.getRange(rowIndex, 6).setValue(data.foto); } const currentFoto = (rowIndex !== -1) ? sheet.getRange(rowIndex, 6).getValue() : data.foto; return { status: 'success', user: { email, nama: data.nama, unit: data.unit, wa, role: (rowIndex!==-1 ? rows[rowIndex-1][4] : "USER"), foto: currentFoto } }; }

function handleGetUsers(payload) { const ss = SpreadsheetApp.getActiveSpreadsheet(); const sheet = ss.getSheetByName(SHEET_USERS); const data = sheet.getRange(2, 1, sheet.getLastRow()-1, 6).getValues(); return { status: 'success', users: data.map(r => ({ email: r[0], nama: r[1], unit: r[2], wa: r[3], role: r[4], foto: r[5] })) }; }

function handleUpdateUserRole(payload) { const ss = SpreadsheetApp.getActiveSpreadsheet(); const sheet = ss.getSheetByName(SHEET_USERS); const rows = sheet.getDataRange().getValues(); for (let i = 1; i < rows.length; i++) { if (String(rows[i][0]).toLowerCase() === String(payload.targetEmail).toLowerCase()) { if (payload.targetEmail.toLowerCase() === SUPER_ADMIN_EMAIL.toLowerCase()) return { status: 'error', message: "Tidak bisa ubah Super Admin." }; sheet.getRange(i + 1, 5).setValue(payload.newRole); return { status: 'success', message: `Role ${payload.targetEmail} diubah menjadi ${payload.newRole}` }; } } return { status: 'error', message: "User tidak ditemukan." }; }

function handleDeleteBooking(p) { const ss=SpreadsheetApp.getActiveSpreadsheet(); const s=ss.getSheetByName(SHEET_JADWAL); const d=s.getDataRange().getValues(); const idx=d[0].indexOf("ID Booking"); for (let i=d.length-1; i>=1; i--) { if (d[i][idx] === p.bookingId) { s.deleteRow(i + 1); return { status: 'success', message: 'Data dihapus.' }; } } return { status: 'error', message: 'ID tidak ditemukan.' }; }
function handleSaveBackground(p) { PropertiesService.getScriptProperties().setProperty('APP_BACKGROUND', p.url); return {status:'success', message:'Background saved'}; }
function handleDeleteBackground(p) { PropertiesService.getScriptProperties().deleteProperty('APP_BACKGROUND'); return {status:'success', message:'Background reset'}; }

// --- FIXED FUNCTION: ARCHIVE OLD BOOKINGS ---
function handleArchiveOldBookings(p) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sJadwal = ss.getSheetByName(SHEET_JADWAL);
    let sArsip = ss.getSheetByName(SHEET_ARSIP);
    
    // Buat sheet arsip jika belum ada
    if (!sArsip) {
        sArsip = ss.insertSheet(SHEET_ARSIP);
        sArsip.appendRow(sJadwal.getRange(1, 1, 1, sJadwal.getLastColumn()).getValues()[0]);
    }
    
    const data = sJadwal.getDataRange().getValues();
    const headers = data[0];
    const dateIdx = headers.indexOf("Tanggal Selesai"); // MENGGUNAKAN TANGGAL SELESAI ACARA
    const timeIdx = headers.indexOf("Waktu Selesai");
    
    if (dateIdx === -1 || timeIdx === -1) {
        return { status: 'error', message: 'Kolom Tanggal Selesai atau Waktu Selesai tidak ditemukan.' };
    }
    
    // Waktu sekarang untuk perbandingan
    const now = new Date();
    let movedCount = 0;
    
    // Loop dari bawah ke atas agar deleteRow aman
    for (let i = data.length - 1; i >= 1; i--) {
        const rowData = data[i];
        
        try {
            // 1. AMBIL TANGGAL SELESAI ACARA (BUKAN TGL BOOKING)
            let endDatePart = rowData[dateIdx];
            let year, month, day;
            let isValidDate = false;
            
            // Logika Parsing Tanggal (Anti Salah Baca Format US/Indo)
            if (endDatePart instanceof Date) {
                year = endDatePart.getFullYear();
                month = endDatePart.getMonth(); // 0-11
                day = endDatePart.getDate();
                isValidDate = true;
            } else {
                // Jika format string (misal: "25-12-2025" atau "25/12/2025")
                // Kita paksa baca sebagai DD-MM-YYYY agar tidak tertukar dengan MM-DD-YYYY
                const dateString = String(endDatePart).trim();
                
                // Cek format DD-MM-YYYY atau DD/MM/YYYY
                const parts = dateString.split(/[-/]/); 
                if (parts.length === 3) {
                    // Asumsi format Indonesia: [Hari, Bulan, Tahun]
                    // Jika tahun ada di depan (YYYY-MM-DD), parts[0] akan 4 digit
                    if (parts[0].length === 4) {
                        year = parseInt(parts[0]);
                        month = parseInt(parts[1]) - 1;
                        day = parseInt(parts[2]);
                    } else {
                        // Asumsi DD-MM-YYYY
                        day = parseInt(parts[0]);
                        month = parseInt(parts[1]) - 1;
                        year = parseInt(parts[2]);
                    }
                    isValidDate = true;
                }
            }
            
            if (!isValidDate) continue; // Skip jika tanggal tidak valid
            
            // 2. AMBIL WAKTU SELESAI
            let timePart = rowData[timeIdx];
            let hour = 23, minute = 59; // Default akhir hari
            
            if (timePart instanceof Date) {
                hour = timePart.getHours();
                minute = timePart.getMinutes();
            } else {
                const tStr = String(timePart).trim();
                const parts = tStr.split(':');
                if (parts.length >= 2) {
                    hour = parseInt(parts[0], 10);
                    minute = parseInt(parts[1], 10);
                }
            }
            
            // 3. GABUNGKAN MENJADI TIMESTAMP ACARA
            const eventEndTime = new Date(year, month, day, hour, minute);
            
            // 4. BANDINGKAN: APAKAH ACARA SUDAH LEWAT DARI SEKARANG?
            // "now" adalah waktu saat tombol diklik.
            if (eventEndTime < now) {
                sArsip.appendRow(rowData);
                sJadwal.deleteRow(i + 1);
                movedCount++;
            }
            
        } catch (e) {
            console.error("Error processing row " + (i+1) + ": " + e.toString());
        }
    }
    
    return { status: 'success', message: `${movedCount} data berhasil diarsipkan.` };
}

function setup() { const ss = SpreadsheetApp.getActiveSpreadsheet(); let sheetUsers = ss.getSheetByName(SHEET_USERS); if (!sheetUsers) { sheetUsers = ss.insertSheet(SHEET_USERS); sheetUsers.appendRow(["Email", "Nama Lengkap", "Unit/Bagian", "No WA", "Role", "Foto"]); sheetUsers.appendRow([SUPER_ADMIN_EMAIL, "Super Admin", "IT", ADMIN_WA_NUMBER, "ADMIN", ""]); } let sheetJadwal = ss.getSheetByName(SHEET_JADWAL); if (!sheetJadwal) { sheetJadwal = ss.insertSheet(SHEET_JADWAL); sheetJadwal.appendRow(["ID Booking", "Timestamp", "Nama Ruangan", "Nama Peminjam", "Unit/Bagian", "Nomor WA", "Tanggal Mulai", "Waktu Mulai", "Tanggal Selesai", "Waktu Selesai", "Keperluan", "Status", "Catatan", "Kelengkapan"]); } let sheetRuangan = ss.getSheetByName(SHEET_RUANGAN); if (!sheetRuangan) { sheetRuangan = ss.insertSheet(SHEET_RUANGAN); sheetRuangan.appendRow(["Nama Ruangan", "Nama PIC", "Nomor WA PIC", "Email PIC"]); DEFAULT_ROOMS.forEach(r => sheetRuangan.appendRow([r, "", "", ""])); } let sheetArsip = ss.getSheetByName(SHEET_ARSIP); if(!sheetArsip) { sheetArsip = ss.insertSheet(SHEET_ARSIP); sheetArsip.appendRow(["ID Booking", "Timestamp", "Nama Ruangan", "Nama Peminjam", "Unit/Bagian", "Nomor WA", "Tanggal Mulai", "Waktu Mulai", "Tanggal Selesai", "Waktu Selesai", "Keperluan", "Status", "Catatan", "Kelengkapan"]);} try { DriveApp.getFolderById(FOLDER_FOTO_ID); } catch(e) { Logger.log("Drive init check: " + e.toString()); } }

// --- AUTOMATION & TRIGGERS ---
function INSTALL_TRIGGERS() {
    // Delete existing to prevent duplicates
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => ScriptApp.deleteTrigger(t));
    
    // Create new daily trigger
    ScriptApp.newTrigger('runDailyArchive')
        .timeBased()
        .everyDays(1)
        .atHour(2) // 2:00 AM
        .create();
        
    Logger.log("Trigger Harian Berhasil Diinstal.");
}

function runDailyArchive() {
    const result = handleArchiveOldBookings({});
    logActivity("INFO", "AUTO ARCHIVE: " + result.message);
}