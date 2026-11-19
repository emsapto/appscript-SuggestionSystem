// --- KONFIGURASI ---
const CONFIG = {
  SHEET_NAME: "Submissions", 
  TARGET_SHEET: "TargetTahunan", // 👈 Nambahin conf target dinamis
  TARGET_QTY_DEFAULT: 100 // Target jaga-jaga kalo ga ada di database
};

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('🚀 SUGESSTION SYSTEM DASHBOARD ')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDashboardData() {

  // const cache = CacheService.getScriptCache();
  // const cachedData = cache.get("DASHBOARD_DATA");
  // if (cachedData) return cachedData;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) throw new Error("Sheet 'Submissions' gak ketemu!");

  const data = sheet.getDataRange().getDisplayValues(); 
  
  // Ambil Target Dinamis dari sheet TargetTahunan (kalau ada)
  const targetSheet = ss.getSheetByName(CONFIG.TARGET_SHEET);
  let targetData = {};
  if (targetSheet) {
    const tData = targetSheet.getDataRange().getValues();
    // Asumsi baris 1 header (Tahun, Target), baris 2 dst isinya
    for (let i = 1; i < tData.length; i++) {
        let thn = String(tData[i][0]).trim();
        let tgt = parseInt(tData[i][1]) || CONFIG.TARGET_QTY_DEFAULT;
        if(thn) targetData[thn] = tgt;
    }
  }

  // ANTI-TYPO: Kita bersihin nama kolom dari spasi-spasi gaib
  const headers = data[0].map(h => String(h).trim()); 
  const rows = data.slice(1);

  const colIdx = {
    tgl: headers.indexOf("Tanggal Form"),
    npk: headers.indexOf("NPK"), 
    nama: headers.indexOf("Nama"),
    dept: headers.indexOf("Department"),
    kategori: headers.indexOf("Kategori"),
    score: headers.indexOf("Final Score"), 
    status: headers.indexOf("Status"), 
    reward: headers.indexOf("Reward"), 
    judul: headers.indexOf("Judul"),
    id: headers.indexOf("ID") //    
  };

  // --- SIHIR PENCARI KOLOM OTOMATIS (ANTI TYPO) ---
  const jCols = [];
  headers.forEach((h, idx) => {
    let namaKolom = String(h).toLowerCase();
    // 🚨 KUNCI: Jangan pakai kata "nilai", ganti jadi "tgl" atau "date" 🚨
    if ((namaKolom.includes('j1') || namaKolom.includes('j2') || namaKolom.includes('j3') || namaKolom.includes('j4') || namaKolom.includes('j5')) 
        && (namaKolom.includes('waktu') || namaKolom.includes('tgl') || namaKolom.includes('date') || namaKolom.includes('app'))) {
      jCols.push(idx);
    }
  });

  let monthlyData = {};

  rows.forEach(row => {
    
    // 1. Tarik status duluan biar gampang misahin logikanya
    let st = String(row[colIdx.status] || "").toLowerCase();
    let isSelesai = st.includes("selesai");
    let isMenunggu = st.includes("menunggu");

    let validMaxObj = null;
    let tglFormAsli = parseDate(row[colIdx.tgl]); // Simpan tanggal asli buat nampil di tabel

    if (isSelesai) {
      // --- LOGIKA A (SELESAI): Patokan mutlak pakai Tanggal Juri Terakhir ---
      let maxTime = 0;
      jCols.forEach(idx => {
        let jDateStr = row[idx];
        if (jDateStr) {
          let jDateObj = parseDate(jDateStr);
          if (jDateObj && !isNaN(jDateObj.getTime())) {
            if (jDateObj.getTime() > maxTime) {
              maxTime = jDateObj.getTime();
              validMaxObj = jDateObj;
            }
          }
        }
      });
      // Cadangan kalau juri kosong
      if (!validMaxObj) validMaxObj = tglFormAsli;

    } else if (isMenunggu) {
      // --- LOGIKA B (MENUNGGU): Sihir teleportasi paksa ke bulan saat ini ---
      validMaxObj = new Date(); 
    }

    // Kalau form error / kosong total, lewatin
    if (!validMaxObj || isNaN(validMaxObj.getTime())) return; 

    // Pembentuk Kotak Bulan
    let monthKey = Utilities.formatDate(validMaxObj, Session.getScriptTimeZone(), "yyyy-MM");
    let monthLabel = Utilities.formatDate(validMaxObj, Session.getScriptTimeZone(), "MMMM yyyy");

    if (!monthlyData[monthKey]) {
      monthlyData[monthKey] = {
        label: monthLabel,
        totalQty: 0,
        users: {},
        categories: {},
        statusCount: { selesai: 0, menunggu: 0 }, 
        depts: {}, 
        allSubs: [] 
      };
    }

    // 2. Isi Progress Bar Sesuai Kotak Waktu
    if (isSelesai) {
      monthlyData[monthKey].statusCount.selesai++;
    } else if (isMenunggu) {
      monthlyData[monthKey].statusCount.menunggu++;
    }

    let deptName = row[colIdx.dept] || "Lainnya";
    
    // 3. KUNCI AMAN GRAFIK: Cuma nambah Qty & Grafik Dept kalau statusnya SELESAI
    if (isSelesai) {
      monthlyData[monthKey].totalQty++; 
      monthlyData[monthKey].depts[deptName] = (monthlyData[monthKey].depts[deptName] || 0) + 1; 
    }

    let rawReward = String(row[colIdx.reward] || "0").replace(/[^\d]/g, ''); 
    let rawCatName = String(row[colIdx.kategori] || "").trim(); 
    let rawScoreList = String(row[colIdx.score] || "0").replace(',', '.'); 

    // KUNCI AMAN TABEL: Pakai Tanggal Form Asli biar histori submit nggak ketimpa
    let displayDateObj = (tglFormAsli && !isNaN(tglFormAsli.getTime())) ? tglFormAsli : validMaxObj;

    // Masukin ke gerbong AllSubs (List Tabel)
    monthlyData[monthKey].allSubs.push({
      tgl: Utilities.formatDate(displayDateObj, Session.getScriptTimeZone(), "dd/MM/yyyy"), 
      npk: row[colIdx.npk],
      nama: row[colIdx.nama] || "Anonim",
      dept: deptName,
      judul: row[colIdx.judul] || "Improvement SS",
      reward: parseFloat(rawReward) || 0,
      status: st,
      kategori: rawCatName || "Lainnya",
      score: parseFloat(rawScoreList) || 0,
      id: row[colIdx.id] || ""   
    });

    // 4. KUNCI AMAN LEADERBOARD: HANYA dihitung kalau SELESAI
    if (isSelesai) {
      let userKey = `${row[colIdx.npk] || 'N/A'}||${row[colIdx.nama] || 'Tanpa Nama'}||${row[colIdx.dept] || '-'}`;
      if (!monthlyData[monthKey].users[userKey]) {
        monthlyData[monthKey].users[userKey] = 0;
      }
      monthlyData[monthKey].users[userKey]++;
    }

    const KATEGORI_RESMI = ["5S", "Quality", "Cost", "Productivity", "Safety", "Moral", "Environment", "Delivery"];
    let validCatName = KATEGORI_RESMI.find(k => k.toLowerCase() === rawCatName.toLowerCase());

    if (validCatName) {
      let rawScore = String(row[colIdx.score]).replace(',', '.'); 
      let scoreVal = parseFloat(rawScore) || 0;

      if (!monthlyData[monthKey].categories[validCatName]) {
        monthlyData[monthKey].categories[validCatName] = { score: -1, winner: { id: "", judul: "", cr: 0 } };
      }

      // KUNCI AMAN HIGHEST SCORE: Harus berstatus SELESAI
      if (isSelesai && scoreVal > monthlyData[monthKey].categories[validCatName].score) {
        let rawCR = String(row[headers.indexOf("Cost Reduction")] || "0").replace(/[^\d]/g, '');
        
        monthlyData[monthKey].categories[validCatName] = {
          score: scoreVal,
          winner: { 
            npk: row[colIdx.npk], 
            nama: row[colIdx.nama], 
            dept: row[colIdx.dept],
            id: row[headers.indexOf("ID")] || "N/A", 
            judul: row[colIdx.judul] || "Tidak ada judul", 
            cr: parseFloat(rawCR) || 0 
          }
        };
      }
    }
  });

  let result = [];
  let sortedKeys = Object.keys(monthlyData).sort();
  
  sortedKeys.forEach((key, index) => {
    let d = monthlyData[key];

    // ... (KODE LEADERBOARD SAMA PERSIS) ...
    let leaderboard = [];
    for (let u in d.users) {
      let [npk, nama, dept] = u.split("||");
      leaderboard.push({ npk, nama, dept, qty: d.users[u] });
    }
    leaderboard.sort((a, b) => b.qty - a.qty);
    
    let prevMonthData = index > 0 ? monthlyData[sortedKeys[index - 1]] : null;
    let prevLeaderboard = [];
    if (prevMonthData && prevMonthData.users) {
      for (let u in prevMonthData.users) {
        let [npk, nama, dept] = u.split("||");
        prevLeaderboard.push({ npk, nama, dept, qty: prevMonthData.users[u] });
      }
      prevLeaderboard.sort((a, b) => b.qty - a.qty);
    }

    let catWinners = [];
    const KATEGORI_RESMI = ["5S", "Quality", "Cost", "Productivity", "Safety", "Moral", "Environment", "Delivery"];
    
    KATEGORI_RESMI.forEach(k => {
      // 👇 UPDATE BAGIAN INI: Tambahin id, judul, cr biar aman kalau kosong 👇
      let currentWinner = { score: -1, npk: "", nama: "", dept: "", id: "", judul: "", cr: 0 };
      if (d.categories[k] && d.categories[k].score > -1) {
        currentWinner = { score: d.categories[k].score, ...d.categories[k].winner };
      }

      let prevWinnerName = "Belum ada";
      if (prevMonthData && prevMonthData.categories[k] && prevMonthData.categories[k].score > -1) {
        prevWinnerName = prevMonthData.categories[k].winner.nama;
      }

      catWinners.push({
        kategori: k,
        ...currentWinner,
        prevWinner: prevWinnerName 
      });
    });

    let deptArr = Object.keys(d.depts).map(k => ({ nama: k, qty: d.depts[k] })).sort((a,b) => b.qty - a.qty);
    let finalDepts = [];
    if (deptArr.length > 6) {
      finalDepts = deptArr.slice(0, 5);
      let othersQty = deptArr.slice(5).reduce((sum, item) => sum + item.qty, 0);
      finalDepts.push({ nama: "Others", qty: othersQty });
    } else {
      finalDepts = deptArr;
    }

    let recentFeed = d.allSubs.slice(-5).reverse();

    // 👈 LOGIKA BARU TARGET DINAMIS: Ambil target sesuai tahun bulan ini
    let currentYear = String(key).substring(0, 4);
    let dynamicTarget = targetData[currentYear] || CONFIG.TARGET_QTY_DEFAULT;

    result.push({
      id: key, 
      label: d.label,
      stats: { total: d.totalQty, target: dynamicTarget, status: d.statusCount },
      depts: finalDepts,
      allSubs: d.allSubs,
      recent: recentFeed, 
      top3: leaderboard, 
      winners: catWinners,
      prevLeaderboard: prevLeaderboard
    });
  });

    let finalJSON = JSON.stringify(result);
  // cache.put("DASHBOARD_DATA", finalJSON, 300); // 300 = 5 menit
  return finalJSON;
}

// ==========================================
// RUMUS DEWA BACA TANGGAL (ANTI ERROR & ANTI DESEMBER GAIB)
// ==========================================
function parseDate(dateStr) {
  if (!dateStr || String(dateStr).trim() === "") return null;
  let cleanStr = String(dateStr).trim();
  
  let parts = cleanStr.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})/);
  
  if (parts) {
    let p1 = parseInt(parts[1], 10);
    let p2 = parseInt(parts[2], 10);
    let p3 = parseInt(parts[3], 10); 
    
    // TENDANG JAUH JAUH kalau tahunnya ga masuk akal (dibawah thn 2000)
    if (p3 < 2000 || p3 > 2100) return null; 

    // SIHIR JAM 12 SIANG: Biar gak geser bulan karena timezone
    if (p2 > 12) {
      return new Date(p3, p1 - 1, p2, 12, 0, 0); 
    } 
    else {
      return new Date(p3, p2 - 1, p1, 12, 0, 0); 
    }
  }

  // FALLBACK: Kalau pakai format lain, cek validitasnya!
  let d = new Date(cleanStr);
  if (!isNaN(d.getTime())) {
    d.setHours(12); // Kunci juga jamnya di sini
    if (d.getFullYear() >= 2000 && d.getFullYear() <= 2100) {
       return d;
    }
  }
  
  return null;
}

// ==========================================
// FUNGSI BARU: AMBIL FOTO PEMENANG DARI SHEET
// ==========================================
function getWinnersPhotosBase64(npkList) {
  let result = {};
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Database Foto");
    if (!sheet) return result;
    
    const data = sheet.getDataRange().getValues(); // Isinya [NPK, File ID]
    
    // Bikin kamus pencarian biar cepat
    let mapFoto = {};
    for (let i = 1; i < data.length; i++) {
       let sheetNpk = String(data[i][0]).trim();
       let fileId = data[i][1];
       if(sheetNpk) mapFoto[sheetNpk] = fileId;
    }
    
    // Loop NPK pemenang yang direquest Dashboard
    npkList.forEach(npk => {
      let cleanNpk = String(npk).trim();
      
      // Kalau NPK gak ada atau gak terdaftar di Database Foto, lewati
      if (!cleanNpk || !mapFoto[cleanNpk]) {
        result[cleanNpk] = null;
        return;
      }
      
      // Kalau NPK ada ID-nya, ambil Blob filenya via ID (Hitungan detik!)
      try {
         let file = DriveApp.getFileById(mapFoto[cleanNpk]);
         let blob = file.getBlob();
         let base64 = Utilities.base64Encode(blob.getBytes());
         let mimeType = blob.getContentType();
         result[cleanNpk] = "data:" + mimeType + ";base64," + base64;
      } catch(e) {
         result[cleanNpk] = null; // Failsafe kalau file aslinya kehapus di Drive
      }
    });
  } catch (e) {
    // Return kosong kalau ada error tak terduga
  }
  
  return result;
}

// ==========================================
// FITUR BARU: EXPORT DATA SS VIA EMAIL
// ==========================================
function verifyJuriForExport(npk, email) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Mapping Juri");
  if (!sheet) return { success: false, message: "Sheet Mapping Juri tidak ditemukan" };
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: false, message: "Data Juri kosong" };
  
  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const npkIdx = headers.indexOf("npk");
  const emailIdx = headers.indexOf("email");
  const namaIdx = headers.indexOf("nama");
  const deptIdx = headers.findIndex(h => h.includes("department") || h.includes("dept"));
  
  if (npkIdx === -1 || emailIdx === -1) return { success: false, message: "Format kolom Mapping Juri salah" };
  
  let searchNpk = String(npk).trim().toLowerCase();
  let searchEmail = String(email).trim().toLowerCase();
  
  let foundDepts = [];
  let juriName = "Juri";
  
  for (let i = 1; i < data.length; i++) {
    let rowNpk = String(data[i][npkIdx]).trim().toLowerCase();
    let rowEmail = String(data[i][emailIdx]).trim().toLowerCase();
    
    // 💡 SIHIR BARU: Kumpulkan semua departemennya, jangan disetop!
    if (rowNpk === searchNpk && rowEmail === searchEmail) {
      if (namaIdx !== -1) juriName = data[i][namaIdx];
      let dpt = deptIdx !== -1 ? data[i][deptIdx] : "-";
      
      // Masukin ke keranjang kalau belum ada
      if (!foundDepts.includes(dpt)) {
        foundDepts.push(dpt);
      }
    }
  }
  
  if (foundDepts.length > 0) {
    return { 
      success: true, 
      data: { 
        nama: juriName, 
        depts: foundDepts // <-- Berubah jadi Array (List)
      } 
    };
  }
  
  return { success: false, message: "NPK atau Email tidak terdaftar sebagai Juri" };
}

function processExportAndEmail(npk, email, filters) {
  const verify = verifyJuriForExport(npk, email);
  if (!verify.success) return { success: false, message: "Verifikasi Juri Gagal!" };
  
  // 🔥 FIX 1: Tangkap array dari multi-departemen juri!
  const juriDepts = verify.data.depts.map(d => String(d).trim().toLowerCase());
  
  // Tangkap pilihan departemen (opsi 'all' kalau dia gak milih spesifik)
  const selectedDept = String(filters.targetDept || "all").trim().toLowerCase();
  
  if (selectedDept !== "all" && !juriDepts.includes(selectedDept)) {
    return { success: false, message: "Akses ke departemen ditolak!" };
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return { success: false, message: "Sheet Submissions tidak ditemukan" };
  
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length < 2) return { success: false, message: "Data Submissions kosong" };
  
  const headers = data[0].map(h => String(h).trim());
  const colIdx = {
    tglForm: headers.indexOf("Tanggal Form"),
    npk: headers.indexOf("NPK"), 
    nama: headers.indexOf("Nama"),
    dept: headers.indexOf("Department"),
    judul: headers.indexOf("Judul"),
    kategori: headers.indexOf("Kategori"),
    status: headers.indexOf("Status"), 
    score: headers.indexOf("Final Score"), 
    reward: headers.indexOf("Reward")
  };

  // 🚨 FIX 2: SIHIR PENCARI KOLOM JURI (ANTI DESEMBER GAIB) 🚨
  // Sama persis kayak di getDashboardData, kata "nilai" WAJIB dihapus!
  const jCols = [];
  headers.forEach((h, idx) => {
    let namaKolom = String(h).toLowerCase();
    if ((namaKolom.includes('j1') || namaKolom.includes('j2') || namaKolom.includes('j3') || namaKolom.includes('j4') || namaKolom.includes('j5')) 
        && (namaKolom.includes('waktu') || namaKolom.includes('tgl') || namaKolom.includes('date') || namaKolom.includes('app'))) {
      jCols.push(idx);
    }
  });
  
  let startM = String(filters.startDate).substring(0, 7);
  let endM = String(filters.endDate).substring(0, 7);
  
  // 👇 HEADER KOLOM EXCEL SESUAI PERMINTAAN 👇
  let csvContent = "Tanggal Submit,Tgl Penilaian Terakhir,NPK,Nama,Dept,Judul SS,Kategori,Status,Score,Reward\n";
  let count = 0;
  
  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    let statusRaw = String(row[colIdx.status] || "");
    let isSelesai = statusRaw.toLowerCase().includes("selesai");
    
    let tglFormAsli = parseDate(row[colIdx.tglForm]);
    let maxTime = 0;
    let validMaxObj = null;

    // 🔥 LOGIKA SINKRONISASI 100% DENGAN DASHBOARD YTD 🔥
    if (isSelesai) {
      // 1. Kalau SELESAI, patokan filter periode pakai Tanggal Juri Terakhir!
      jCols.forEach(idx => {
        let jDateStr = row[idx];
        if (jDateStr) {
          let jDateObj = parseDate(jDateStr);
          if (jDateObj && !isNaN(jDateObj.getTime())) {
            if (jDateObj.getTime() > maxTime) {
              maxTime = jDateObj.getTime();
              validMaxObj = jDateObj;
            }
          }
        }
      });
      // Failsafe: kalau ternyata juri beneran kosong padahal status selesai
      if (!validMaxObj) validMaxObj = tglFormAsli;

    } else {
      // 2. Kalau MENUNGGU, tetap masuk Excel tapi patokan filter pakai Tanggal Submit
      validMaxObj = tglFormAsli;
    }

    // Kalau bener-bener kosong dari awal input, lewatin aja barisnya
    if (!validMaxObj || isNaN(validMaxObj.getTime())) continue;
    
    // FILTER BULAN berdasarkan validMaxObj yang udah dipisah logikanya di atas
    let ymDate = Utilities.formatDate(validMaxObj, Session.getScriptTimeZone(), "yyyy-MM");
    if (ymDate < startM || ymDate > endM) continue;
    
    let rowDept = String(row[colIdx.dept] || "").trim().toLowerCase();
    let rowNpk = String(row[colIdx.npk] || "").trim().toLowerCase();
    
    // FILTER DEPARTEMEN
    if (selectedDept === "all") {
       if (!juriDepts.includes(rowDept)) continue;
    } else {
       if (rowDept !== selectedDept) continue;
    }
    
    // FILTER NPK
    if (filters.filterType === 'npk') {
      let allowedNpks = String(filters.filterValue).toLowerCase().split(",").map(n => n.trim());
      if (!allowedNpks.includes(rowNpk)) continue;
    }

    let finalStatus = statusRaw;
    let finalScore = row[colIdx.score] || "";
    let finalReward = row[colIdx.reward] || "";
    
    let strTglAkhir = "-"; // Default string kalau SS belum Selesai

    if (isSelesai) {
       // Isi kolom baru dengan tanggal juri terakhir kalau udah beres
       if (maxTime > 0) strTglAkhir = Utilities.formatDate(validMaxObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else {
       // Kosongkan detail score/reward kalau masih OTW Penjurian
       finalStatus = "Menunggu Penjurian";
       finalScore = "";
       finalReward = "";
    }
    
    let strTglSubmit = tglFormAsli ? Utilities.formatDate(tglFormAsli, Session.getScriptTimeZone(), "yyyy-MM-dd") : "";

    // Membangun baris Excel (Sesuai header yg lu minta)
    let rowArray = [
      strTglSubmit,
      strTglAkhir,
      row[colIdx.npk] || "",
      row[colIdx.nama] || "",
      row[colIdx.dept] || "",
      `"${String(row[colIdx.judul] || "").replace(/"/g, '""')}"`,
      row[colIdx.kategori] || "",
      finalStatus,
      finalScore,
      finalReward
    ];
    csvContent += rowArray.join(",") + "\n";
    count++;
  }
  
  if (count === 0) return { success: false, message: "Tidak ada ide SS pada periode tanggal tersebut di Departemen Anda. Silakan coba cek periode lain." };
  
  try {
    let blob = Utilities.newBlob(csvContent, "text/csv", "Data_SS.csv");
    let subj = `[Export] Data SS (${filters.startDate} s/d ${filters.endDate})`;
    let targetInfo = filters.filterType === 'dept' ? 'All Member' : 'Custom NPK';
    let deptText = selectedDept === 'all' ? verify.data.depts.join(", ") : filters.targetDept;

    let body = `Halo Juri,\n\nBerikut terlampir Data SS sesuai dengan filter:\n- Departemen Tertuju: ${deptText}\n- Tipe Filter: ${targetInfo}\n- Total Data: ${count} SS.\n\nTerima kasih,\nSekretariat Improvement`;
    
    MailApp.sendEmail({
      to: email,
      subject: subj,
      body: body,
      attachments: [blob],
      name: "Sekretariat SS MAI"
    });
    
    return { success: true, message: "Excel berhasil dikirim ke " + email };
  } catch (e) {
    return { success: false, message: "Gagal email: " + e.toString() };
  }
}

function pancingIzinAkses() {
  // 1. Cek Folder PDF
  try {
    DriveApp.getFolderById("1Wsukt7SHQi7HUE63pWj5fDs8o5REY7U8");
    Logger.log("✅ Akses Folder PDF: SUKSES");
  } catch(e) { Logger.log("❌ Gagal Akses Folder PDF: " + e.message); }
  
  // 2. Cek Template Docs
  try {
    DocumentApp.openById("1O9Alq-b73Vz71Tmw0HiHImfIOR6fmXyWylWUC6tQMKs");
    Logger.log("✅ Akses Template Docs: SUKSES");
  } catch(e) { Logger.log("❌ Gagal Akses Template Docs: " + e.message); }
  
  // 3. Cek Sheet Database
  try {
    SpreadsheetApp.openById("1lvuZLVGR9OVsf6Kz1fo_TFsffkrKI0YkFsZkfvlAiCQ");
    Logger.log("✅ Akses Sheet Database: SUKSES");
  } catch(e) { Logger.log("❌ Gagal Akses Sheet Database: " + e.message); }
}

// ==========================================
// FITUR KEAMANAN: VERIFIKASI PASSWORD
// ==========================================
function verifyPassword(inputPwd) {
  // Password rahasia lu sekarang aman terkunci di server!
  const SECRET = "Metalart2026";
  return inputPwd === SECRET;
}

// ==========================================
// FITUR BARU: AUTO-SYNC DATABASE FOTO (ANTI-CRASH)
// ==========================================
function syncFotoKaryawan() {
  const FOLDER_ID = "1n0QyH8fKSwbyjgGhAcG6GqaUQI9Cy4Y1";
  const SHEET_NAME = "Database Foto";
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["NPK", "File ID"]);
    sheet.getRange("A1:B1").setFontWeight("bold");
  }
  
  let folder;
  try {
    folder = DriveApp.getFolderById(FOLDER_ID);
  } catch(e) {
    Logger.log("❌ Gagal akses folder, cek ID Folder lagi!");
    return;
  }
  
  let files = folder.getFiles();
  let dataToWrite = [];
  let count = 0;
  
  while (files.hasNext()) {
    try {
      let file = files.next();
      let name = file.getName(); 
      let npk = name.replace(/[^0-9]/g, ''); // Ambil angkanya aja buat NPK
      
      if (npk) {
        dataToWrite.push([npk, file.getId()]);
      }
      
      count++;
      // 🔥 OBAT PENENANG SERVER: 
      // Kasih napas ke server Google tiap baca 50 foto biar ga keselek
      if (count % 50 === 0) {
        Utilities.sleep(1000); // Istirahat 1 detik
      }
      
    } catch(e) {
      // Kalau ketemu file hantu/rusak, lewatin aja jangan sampai script mati
      Logger.log("⚠️ Skip 1 file error: " + e.message);
    }
  }
  
  // Hapus data lama, timpa dengan update data foto terbaru
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
  }
  
  if (dataToWrite.length > 0) {
    sheet.getRange(2, 1, dataToWrite.length, 2).setValues(dataToWrite);
  }
  
  Logger.log("✅ Sync selesai sukses! Total foto dilist: " + dataToWrite.length);
}
// ==========================================
// FITUR REKAP OTOMATIS KE GOOGLE SHEET
// ==========================================
function autoGenerateRekapBulanan() {
  // 1. Nyontek data dari Dashboard lu yang udah mateng
  let rawJson = getDashboardData();
  let allData = JSON.parse(rawJson);

  // 2. Tentukan "Bulan Kemarin" (Karena robot jalan tgl 1 subuh)
  let today = new Date();
  let lastMonth = new Date(today.getFullYear(), today.getMonth(), 0);
  let targetPeriode = Utilities.formatDate(lastMonth, Session.getScriptTimeZone(), "yyyy-MM");

  // 3. Cari data khusus bulan kemarin
  let monthData = allData.find(d => d.id === targetPeriode);
  if (!monthData) {
    Logger.log("Data bulan " + targetPeriode + " tidak ditemukan!");
    return;
  }

  // 4. Siapin wadah buat ngerekap per NPK
  let rekap = {};

  // Hitung Qty & Reward Dasar khusus yang statusnya "Selesai"
  monthData.allSubs.forEach(s => {
     if (String(s.status).toLowerCase().includes('selesai')) {
         let npk = String(s.npk).trim();
         if (!rekap[npk]) {
             rekap[npk] = {
                 npk: npk, nama: s.nama, email: "", dept: s.dept,
                 qty: 0, rewardDasar: 0, bonus: 0, ket: []
             };
         }
         rekap[npk].qty += 1;
         rekap[npk].rewardDasar += (parseFloat(s.reward) || 0);
     }
  });

  // Tambahin Bonus 100rb kalau dia Juara Kategori
  monthData.winners.forEach(w => {
     if (w.score > -1 && w.npk) {
         let npk = String(w.npk).trim();
         if (rekap[npk]) {
             rekap[npk].bonus += 100000;
             rekap[npk].ket.push("Juara " + w.kategori);
         }
     }
  });

  // 5. Susun jadi format baris Excel
  let rowsToWrite = [];
  for (let npk in rekap) {
     let emp = rekap[npk];
     let idRekap = targetPeriode + "-" + emp.npk; // Contoh: 2026-10-12345
     let totalSemua = emp.rewardDasar + emp.bonus;
     let ketText = emp.ket.length > 0 ? emp.ket.join(", ") : "-";

     // Urutan wajib: ID Rekap | Periode | NPK | Nama | Email | Dept | Qty | Reward | Bonus | Total | Keterangan
     rowsToWrite.push([
        idRekap, targetPeriode, emp.npk, emp.nama, emp.email, emp.dept,
        emp.qty, emp.rewardDasar, emp.bonus, totalSemua, ketText
     ]);
  }

  if (rowsToWrite.length === 0) return;

  // 6. Tulis ke Sheet "Rekap Reward Bulanan"
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetRekap = ss.getSheetByName("Rekap Reward Bulanan");
  if (!sheetRekap) return;

  // Fitur Anti-Double: Ngecek ID biar nggak kecetak 2x kalau robot jalan berulang
  let existingData = sheetRekap.getDataRange().getValues();
  let existingIDs = existingData.map(r => String(r[0])); 
  let finalRows = rowsToWrite.filter(r => !existingIDs.includes(r[0]));

  if (finalRows.length > 0) {
     sheetRekap.getRange(sheetRekap.getLastRow() + 1, 1, finalRows.length, finalRows[0].length).setValues(finalRows);
  }
}