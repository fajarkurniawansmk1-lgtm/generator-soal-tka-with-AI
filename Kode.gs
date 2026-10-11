/**
 * AI TKA English Generator v3.8 - Standar Pusmendik Kemendikdasmen 2026
 * Features: Strict Prompting Validation, HTML-Based Visual Stimulus, Separate Database per Packet, Token System, LockService, Flexible Timer, Student Whitelist, & Anti-Cheat Tracker
 */

function getApiKey() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  if (!sheet) {
    throw new Error("Sheet 'Config' tidak ditemukan! Buat tab bernama Config.");
  }
  var apiKey = sheet.getRange("B1").getValue();
  if (!apiKey || apiKey.toString().trim() === "") {
    throw new Error("API Key Gemini belum diisi di sel B1 sheet Config!");
  }
  return apiKey.toString().trim();
}

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page.toLowerCase() : '';

  if (page === 'ujian') {
    return HtmlService.createTemplateFromFile('Ujian')
      .evaluate()
      .setTitle('TKA English Test - Murid')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  if (page === 'dashboard') {
    return HtmlService.createTemplateFromFile('Dashboard')
      .evaluate()
      .setTitle('Dashboard Evaluasi Guru - CBT TKA')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('AI TKA English Generator v3.8')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Helper: Generator Token Acak Unik
function generateRandomToken(prefix) {
  var cleanPrefix = (prefix || 'TKA').substring(0, 3).toUpperCase();
  var randomNum = Math.floor(100 + Math.random() * 900);
  return cleanPrefix + randomNum;
}

// FUNGSI: Mengambil Daftar Siswa Berdasarkan Kelas & Jurusan dari Sheet 'DataSiswa'
function getDaftarSiswa(kelasFilter, jurusanFilter) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DataSiswa");
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var cleanKelas = String(kelasFilter || '').trim().toLowerCase();
  var cleanJurusan = String(jurusanFilter || '').trim().replace(/\s+/g, '').toLowerCase();

  var listSiswa = [];
  for (var i = 1; i < data.length; i++) {
    var nama = String(data[i][1]).trim();
    var kelas = String(data[i][2]).trim().toLowerCase();
    var jurusan = String(data[i][3]).trim().replace(/\s+/g, '').toLowerCase();

    var matchKelas = (!cleanKelas || kelas === cleanKelas);
    var matchJurusan = (!cleanJurusan || jurusan === cleanJurusan || jurusan.indexOf(cleanJurusan) !== -1 || cleanJurusan.indexOf(jurusan) !== -1);

    if (nama !== "" && matchKelas && matchJurusan) {
      listSiswa.push(nama);
    }
  }

  return listSiswa.sort();
}

// FUNGSI: Mengambil Daftar Jurusan Unik dari Sheet 'DataSiswa'
function getDaftarJurusanUnique() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("DataSiswa");
  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var mapJurusan = {};
  for (var i = 1; i < data.length; i++) {
    var jur = String(data[i][3]).trim();
    if (jur !== "") {
      mapJurusan[jur] = true;
    }
  }

  var listJurusan = Object.keys(mapJurusan).sort();
  return listJurusan;
}

// 1. SISI GURU: Generate Soal HOTS + Simpan ke Sheet Database Baru Khusus
function generateTKASoal(params) {
  var apiKey = getApiKey();
  var endpoint = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash-lite:generateContent?key=" + apiKey;
  var totalSoal = params.jumlahSoal || 20;
  var kelas = params.kelas || "XII";
  var jurusan = params.jurusan || "Umum";
  var namaPaketInput = params.namaPaket || ("Paket TKA Bahasa Inggris - " + jurusan);

  var nTekstual = Math.round(totalSoal * 0.30);
  var nInferensial = Math.round(totalSoal * 0.50);
  var nEvaluasi = totalSoal - (nTekstual + nInferensial);
  var nStimulus = Math.max(3, Math.ceil(totalSoal / 4));

  var prompt = "";
  prompt += "Anda adalah pakar penyusun, pengembang, validator, dan reviewer soal Tes Kemampuan Akademik (TKA) Bahasa Inggris SMA/SMK 2026 yang mengacu pada Peraturan Kepala BSKAP No. 045/H/AN/2025 dan Kerangka Asesmen Resmi Pusmendik.\n";
  prompt += "Tugas Anda adalah menghasilkan soal Bahasa Inggris presisi tinggi yang mengukur KETERAMPILAN MEMBACA (READING COMPREHENSION), bukan sekadar soal Bahasa Inggris sekolah biasa.\n";
  prompt += "Jangan mengubah asesmen menjadi tes grammar, vocabulary, atau hafalan kosakata secara terpisah kecuali kemampuan tersebut memang diperlukan untuk memahami teks.\n\n";

  prompt += "============================================================\n";
  prompt += "JUMLAH SOAL WAJIB PRESISI (STRICT COUNT CONSTRAINT)\n";
  prompt += "============================================================\n";
  prompt += "Buat TEPAT PERSIS " + totalSoal + " BUTIR SOAL untuk murid kelas " + kelas + " SMK Jurusan " + jurusan + ".\n";
  prompt += "DILARANG KERAS menghasilkan kurang dari " + totalSoal + " soal (misal hanya 28) atau lebih dari " + totalSoal + " soal! Output JSON Array HARUS berisi tepat " + totalSoal + " elemen objek soal.\n";
  prompt += "Nama paket: " + namaPaketInput + ".\n";
  prompt += "Output akhir HARUS berupa JSON ARRAY MURNI tanpa markdown fence, tanpa teks pembuka/penutup.\n\n";

  prompt += "============================================================\n";
  prompt += "ATURAN WAJIB VARIATIF TIPE SOAL (STRICT MANDATORY DISTRIBUTION)\n";
  prompt += "============================================================\n";
  prompt += "DILARANG KERAS membuat seluruh soal berjenis Pilihan Ganda biasa (PG)!\n";
  prompt += "Dari total " + totalSoal + " soal, Anda WAJIB membaginya secara presisi ke dalam 3 jenis tipe soal berikut:\n";
  prompt += "1. PG (Pilihan Ganda Biasa) : Sekitar 50% dari total soal.\n";
  prompt += "2. PGK-MCMA (Pilihan Ganda Kompleks - Jawaban Lebih dari Satu) : Sekitar 30% dari total soal.\n";
  prompt += "3. PGK-Kategori (Pilihan Ganda Kompleks - Tabel Benar/Salah) : Sekitar 20% dari total soal.\n\n";

  prompt += "============================================================\n";
  prompt += "KARAKTER & TINGKAT KOMPLEKSITAS TEKS (CEFR A2 & B1)\n";
  prompt += "============================================================\n";
  prompt += "1. PROPORSI KONTEKS:\n";
  prompt += "   - 40% Everyday communication / kehidupan sehari-hari.\n";
  prompt += "   - 30% Vocational / dunia kerja (simulasi komunikasi profesional jurusan " + jurusan + ").\n";
  prompt += "   - 30% Academic / akademik dasar.\n";
  prompt += "   - Jangan membuat seluruh soal menjadi vokasional atau akademik.\n";
  prompt += "   - Jangan memaksakan istilah teknis kejuruan " + jurusan + " ke semua teks. Informasi kejuruan teknis yang kompleks TIDAK BOLEH diuji; fokus murni pada pemahaman bacaan.\n\n";
  prompt += "2. PANJANG TEKS & CEFR LEVEL:\n";
  prompt += "   - Teks Utama (Level B1): Sepanjang 250 HINGGA 400 KATA. Menggunakan kosakata hingga 3000 kata umum dan kalimat bertingkat.\n";
  prompt += "   - Teks Pendukung (Level A2): Sepanjang 150 HINGGA 250 KATA. Menggunakan kosakata hingga 2000 kata umum.\n";
  prompt += "   - Jangan menggunakan kosakata B2/C1 yang sangat langka hanya untuk memersulit soal. Kesulitan WAJIB berasal dari penalaran, inferensi, dan evaluasi isi teks.\n\n";

  prompt += "============================================================\n";
  prompt += "ATURAN WAJIB JUDUL STIMULUS TEKS\n";
  prompt += "============================================================\n";
  prompt += "- Setiap stimulus teks WAJIB DIAWALI DENGAN JUDUL TEKS bercetak tebal/heading di baris pertama.\n";
  prompt += "- Contoh di awal string stimulus: \"<h4><b>The Metro Youth Learning Hub</b></h4><br>\"\n\n";

  prompt += "============================================================\n";
  prompt += "JENIS TEKS & STIMULUS INFOGRAFIS VISUAL\n";
  prompt += "============================================================\n";
  prompt += "Gunakan variasi genre teks secara seimbang:\n";
  prompt += "1. Descriptive (Deskripsi objek, tempat, atau fenomena)\n";
  prompt += "2. Recount (Pengalaman kerja/magang, peristiwa bersejarah)\n";
  prompt += "3. Narrative (Cerita moral/kisah inspiratif)\n";
  prompt += "4. Procedure (Langkah kerja, panduan, tutorial, SOP)\n";
  prompt += "5. Analytical Exposition (Argumen isu sosial, teknologi, kesehatan)\n";
  prompt += "6. Infografis Visual / HTML Component (SOP, aturan kerja, brosur, poster, tabel perbandingan)\n\n";

  prompt += "- PANDUAN UTAMA INFOGRAFIS: Anda WAJIB menyertakan minimal 1 atau 2 kelompok stimulus berupa tabel data, kartu peringatan (card), atau daftar langkah SOP menggunakan tag HTML asli (seperti <table class='table table-bordered'>, <div class='card p-3'>, atau <ul class='list-group'>).\n";
  prompt += "- PENTING (MENCEGAH ERROR JSON): Gunakan KUTIP TUNGGAL (single quote) untuk seluruh atribut HTML di dalam string stimulus! (Contoh: <span class='badge bg-primary'>, BUKAN kutip dua).\n";
  prompt += "ATURAN KHUSUS STIMULUS VISUAL / INFOGRAFIS HTML:\n";
  prompt += "- Sekitar 20-30% soal WAJIB ber-stimulus visual infografis yang ditulis dalam kode HTML + Bootstrap 5 (menggunakan card, badge, list-group, atau table).\n";
  prompt += "- Set field 'tipeStimulus' ke 'hybrid-html' jika ber-stimulus HTML.\n";
  prompt += "- STRICT NEGATIVE CONSTRAINT (DILARANG KERAS): DILARANG MEMBUAT SOAL TRIVIA DESAIN MENTAH seperti menanyakan warna badge, nomor ikon, bentuk fisik, atau posisi elemen HTML!\n";
  prompt += "- Pertanyaan pada stimulus visual WAJIB MURNI menguji PEMAHAMAN BACAAN & INFORMASI TERKANDUNG.\n\n";

  prompt += "============================================================\n";
  prompt += "TIGA KOMPETENSI UTAMA & TARGET BLUEPRINT\n";
  prompt += "============================================================\n";
  prompt += "Gunakan target blueprint berikut sebagai panduan distribusi " + totalSoal + " soal:\n";
  prompt += "- " + nTekstual + " soal Pemahaman Tekstual (~30%)\n";
  prompt += "- " + nInferensial + " soal Pemahaman Inferensial (~50%)\n";
  prompt += "- " + nEvaluasi + " soal Evaluasi dan Apresiasi (~20%)\n\n";

  prompt += "============================================================\n";
  prompt += "STIMULUS BERKELOMPOK (PASSAGE-BASED)\n";
  prompt += "============================================================\n";
  prompt += "DILARANG membuat 1 teks baru untuk setiap 1 soal!\n";
  prompt += "Gunakan sistem berkelompok (Passage-Based). 1 Stimulus digunakan untuk 3 HINGGA 4 SOAL.\n";
  prompt += "Bagi secara rata menjadi persis " + nStimulus + " kelompok stimulus yang berbeda.\n\n";

  prompt += "============================================================\n";
  prompt += "ATURAN SEBARAN KUNCI JAWABAN (DISTRIBUSI PROPOSIONAL)\n";
  prompt += "============================================================\n";
  prompt += "- DILARANG KERAS menumpuk kunci jawaban hanya pada pilihan A, B, atau C!\n";
  prompt += "- Sebarkan kunci jawaban untuk soal PG secara MERATA dan ACAK ke seluruh pilihan (A, B, C, D, E).\n";
  prompt += "- Pastikan opsi D dan E mendapat porsi kunci jawaban yang seimbang dengan opsi A, B, dan C.\n\n";

  prompt += "============================================================\n";
  prompt += "BENTUK SOAL & FORMAT KUNCI JAWABAN (RIGID PROPERTIES)\n";
  prompt += "============================================================\n";
  prompt += "Setiap objek soal di dalam JSON WAJIB memiliki properti persis berikut:\n";
  prompt += "- \"stimulus\" : String teks/HTML stimulus (diawali judul).\n";
  prompt += "- \"kompetensi\" : String (\"Pemahaman Tekstual\", \"Pemahaman Inferensial\", atau \"Evaluasi dan Apresiasi\").\n";
  prompt += "- \"indikator\" : String indikator pencapaian soal.\n";
  prompt += "- \"levelCEFR\" : String (\"A2\" atau \"B1\").\n";
  prompt += "- \"jenisTeks\" : String genre teks.\n";
  prompt += "- \"tipeSoal\" : String WAJIB BERVARIASI TERDIRI DARI (\"PG\", \"PGK-MCMA\", atau \"PGK-Kategori\").\n";
  prompt += "- \"tipeStimulus\" : String (\"text\" atau \"hybrid-html\").\n";
  prompt += "- \"pertanyaan\" : String kalimat pertanyaan.\n";
  prompt += "- \"pilihanJawaban\" : Objek 5 pilihan A-E.\n";
  prompt += "- \"kunci\" : String Kunci Jawaban (WAJIB GUNAKAN NAMA PROPERTI 'kunci', DILARANG MENGGUNAKAN 'kunciJawaban').\n";
  prompt += "   1. Jika tipeSoal PG: \"A\" (Sebarkan acak A-E)\n";
  prompt += "   2. Jika tipeSoal PGK-MCMA: \"A,C\"\n";
  prompt += "   3. Jika tipeSoal PGK-Kategori: \"A:Benar, B:Salah, C:Benar, D:Salah, E:Benar\" (atau \"True/False\"). DILARANG KERAS mengganti label 'Benar/Salah' atau 'True/False' dengan istilah/kategori kontekstual lain apapun topik bacaannya!\n";
  prompt += "- \"pembahasan\" : String penjelasan pembahasan.\n\n";

  prompt += "HASIL AKHIR HANYA JSON ARRAY VALID DENGAN TOTAL " + totalSoal + " OBJEK SOAL. DILARANG MEMBERIKAN TEKS TAMBAHAN APAPUN DI LUAR KODE JSON!";

  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.2,
      responseMimeType: "application/json"
    }
  };

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch(endpoint, options);
    var rawText = response.getContentText();
    var jsonResponse = JSON.parse(rawText);

    logToSheet("LOG", JSON.stringify(payload), rawText);

    if (jsonResponse.error) throw new Error("API Gemini Error: " + (jsonResponse.error.message || JSON.stringify(jsonResponse.error)));
    if (!jsonResponse.candidates || !jsonResponse.candidates[0]) throw new Error("Respons Gemini kosong.");

    var resultText = jsonResponse.candidates[0].content.parts[0].text.trim();
    if (resultText.indexOf("```json") === 0) resultText = resultText.substring(7);
    if (resultText.indexOf("```") === 0) resultText = resultText.substring(3);
    if (resultText.lastIndexOf("```") === resultText.length - 3) resultText = resultText.substring(0, resultText.length - 3);
    resultText = resultText.trim();

    var soalArray = JSON.parse(resultText);
    var infoPaket = simpanKeDatabaseTerpisah(soalArray, params, namaPaketInput);

    return { status: "success", total: soalArray.length, paket: infoPaket, data: soalArray };
  } catch (error) {
    logToSheet("ERROR", JSON.stringify(payload), error.toString());
    throw new Error(error.message);
  }
}

// Helper: Membuat Sheet Baru per Paket & Registrasi ke Master 'DaftarPaket'
function simpanKeDatabaseTerpisah(soalArray, params, namaPaket) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var timeStamp = new Date();
  var idPaket = "PKT-" + Date.now().toString().slice(-6);
  var token = generateRandomToken(params.jurusan);
  var sheetName = "SOAL_" + idPaket;
  var durasi = params.durasi || 60;

  var newSheet = ss.getSheetByName(sheetName);
  if (!newSheet) {
    newSheet = ss.insertSheet(sheetName);
  } else {
    newSheet.clear();
  }
  
  newSheet.appendRow([
    "ID Soal", "Timestamp", "Jurusan", "Tema", "Kompetensi", 
    "Level CEFR", "Jenis Teks", "Tipe Soal", "Tipe Stimulus", "Stimulus (Text/HTML)", 
    "Pertanyaan", "Pilihan Jawaban", "Kunci", "Pembahasan", "Indikator"
  ]);

  soalArray.forEach(function(item, index) {
    var idSoal = idPaket + "-" + (index + 1);
    var strPilihan = (typeof item.pilihanJawaban === 'object') ? JSON.stringify(item.pilihanJawaban) : item.pilihanJawaban;

    var valKunci = item.kunci || item.kunciJawaban || item.answer || "A";
    if (typeof valKunci === 'object' && valKunci !== null) {
      if (Array.isArray(valKunci)) {
        valKunci = valKunci.join(",");
      } else {
        var keys = Object.keys(valKunci).sort();
        valKunci = keys.map(function(k) { return k + ":" + valKunci[k]; }).join(", ");
      }
    }

    var valKompetensi = item.kompetensi || item.indikator || "Pemahaman Tekstual";
    var valIndikator = item.indikator || item.kompetensi || "Siswa dapat menganalisis informasi pada teks.";
    var valStimulus = item.stimulus || "";

    if (valStimulus && !valStimulus.includes("<h4>") && !valStimulus.includes("<b>") && !valStimulus.includes("<strong>")) {
      var autoTitle = "<h4><b>Reading Passage: " + (item.jenisTeks || "General Text") + "</b></h4><br>";
      valStimulus = autoTitle + valStimulus;
    }

    item.kunci = valKunci;
    item.kunciJawaban = valKunci;
    item.stimulus = valStimulus;
    item.indikator = valIndikator;

    newSheet.appendRow([
      idSoal,
      timeStamp,
      params.jurusan,
      params.tema || item.jenisTeks || "Umum",
      valKompetensi,
      item.levelCEFR || "B1",
      item.jenisTeks || "Descriptive",
      item.tipeSoal || "PG",
      item.tipeStimulus || "text",
      valStimulus,
      item.pertanyaan,
      strPilihan,
      valKunci,
      item.pembahasan || "-",
      valIndikator
    ]);
  });

  var masterSheet = ss.getSheetByName("DaftarPaket");
  if (!masterSheet) {
    masterSheet = ss.insertSheet("DaftarPaket");
    masterSheet.appendRow(["ID Paket", "Nama Paket", "Nama Sheet Database", "Token", "Jumlah Soal", "Tanggal Dibuat", "Status", "Durasi (Menit)"]);
  }

  masterSheet.appendRow([
    idPaket, namaPaket, sheetName, token, soalArray.length, timeStamp, "Aktif", durasi
  ]);

  return { idPaket: idPaket, namaPaket: namaPaket, token: token, sheetName: sheetName, durasi: durasi };
}

// 2. SISI SISWA: Mengambil Daftar Paket Aktif
function getDaftarPaketAktif() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("DaftarPaket");
  if (!masterSheet) return [];

  var data = masterSheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][6]).trim().toLowerCase() === 'aktif') {
      list.push({
        idPaket: data[i][0],
        namaPaket: data[i][1],
        sheetName: data[i][2],
        jumlahSoal: data[i][4],
        durasi: data[i][7] || 60
      });
    }
  }
  return list;
}

function cekApakahSudahMengerjakan(namaSiswa, idPaket) {
  var sheetJawaban = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("JawabanMurid");
  if (!sheetJawaban) return false;

  var data = sheetJawaban.getDataRange().getValues();
  if (data.length <= 1) return false;

  var cleanNamaInput = String(namaSiswa).trim().toLowerCase();
  var cleanPaketInput = String(idPaket).trim().toLowerCase();

  for (var i = 1; i < data.length; i++) {
    var namaDiSheet = String(data[i][1]).trim().toLowerCase();
    var paketDiSheet = String(data[i][4]).trim().toLowerCase();

    if (namaDiSheet === cleanNamaInput && paketDiSheet === cleanPaketInput) {
      return true;
    }
  }
  return false;
}

function getSoalDenganToken(payload) {
  var idPaket = payload.idPaket;
  var inputToken = payload.token;
  var namaSiswa = payload.nama;

  if (namaSiswa && cekApakahSudahMengerjakan(namaSiswa, idPaket)) {
    throw new Error("MAAF, " + namaSiswa.toUpperCase() + " TERCATAT SUDAH PERNAH MENGERJAKAN PAKET UJIAN INI!");
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("DaftarPaket");
  if (!masterSheet) throw new Error("Daftar Paket tidak ditemukan!");

  var dataMaster = masterSheet.getDataRange().getValues();
  var targetSheetName = "";
  var tokenValid = false;
  var durasiPaket = 60;

  for (var i = 1; i < dataMaster.length; i++) {
    if (dataMaster[i][0] === idPaket) {
      var tokenResmi = String(dataMaster[i][3]).trim().toUpperCase();
      if (tokenResmi === String(inputToken).trim().toUpperCase()) {
        tokenValid = true;
        targetSheetName = dataMaster[i][2];
        durasiPaket = dataMaster[i][7] || 60;
      }
      break;
    }
  }

  if (!tokenValid) {
    throw new Error("TOKEN UJIAN SALAH! Silakan minta token yang benar kepada Guru Anda.");
  }

  var sheetPaket = ss.getSheetByName(targetSheetName);
  if (!sheetPaket) throw new Error("Database paket soal (" + targetSheetName + ") tidak ditemukan!");

  var dataSoal = sheetPaket.getDataRange().getValues();
  if (dataSoal.length <= 1) throw new Error("Paket ini belum memiliki butir soal.");

  var rows = dataSoal.slice(1);

  var daftarSoal = rows.map(function(row, index) {
    var pilihan = {};
    try {
      var rawPilihan = row[11];
      if (typeof rawPilihan === 'object' && rawPilihan !== null) {
        pilihan = rawPilihan;
      } else if (typeof rawPilihan === 'string' && rawPilihan.trim() !== '') {
        pilihan = JSON.parse(rawPilihan);
      }
    } catch(e) {
      pilihan = {};
    }

    return {
      idSoal: row[0],
      nomor: index + 1,
      kompetensi: row[4],
      levelCEFR: row[5],
      jenisTeks: row[6],
      tipeSoal: row[7],
      tipeStimulus: row[8],
      stimulus: row[9],
      pertanyaan: row[10],
      pilihanJawaban: pilihan
    };
  });

  return { soal: daftarSoal, durasi: durasiPaket };
}

function submitJawabanSiswa(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    throw new Error("Sistem sedang sibuk memproses jawaban Anda. Mohon tunggu sebentar lalu coba lagi.");
  }

  try {
    if (cekApakahSudahMengerjakan(payload.nama, payload.idPaket)) {
      throw new Error("MAAF, Siswa atas nama " + payload.nama.toUpperCase() + " sudah pernah mengirimkan jawaban untuk paket ini.");
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName("DaftarPaket");
    var dataMaster = masterSheet.getDataRange().getValues();
    
    var targetSheetName = "";
    var namaPaket = payload.idPaket;

    for (var i = 1; i < dataMaster.length; i++) {
      if (dataMaster[i][0] === payload.idPaket) {
        targetSheetName = dataMaster[i][2];
        namaPaket = dataMaster[i][1];
        break;
      }
    }

    var sheetPaket = ss.getSheetByName(targetSheetName);
    var dataBank = sheetPaket.getDataRange().getValues();
    
    var kunciMap = {};
    var skorPerKompetensi = {
      "Pemahaman Tekstual": { benar: 0, total: 0 },
      "Pemahaman Inferensial": { benar: 0, total: 0 },
      "Evaluasi & Apresiasi": { benar: 0, total: 0 },
      "Evaluasi dan Apresiasi": { benar: 0, total: 0 }
    };

    var totalSoalPaket = dataBank.length - 1;

    for (var i = 1; i < dataBank.length; i++) {
      var idSoal = dataBank[i][0];
      var kom = String(dataBank[i][4]).trim();
      var tipe = String(dataBank[i][7]).trim();
      var kunci = String(dataBank[i][12]).trim();

      kunciMap[idSoal] = {
        kunci: kunci,
        kompetensi: kom,
        tipeSoal: tipe
      };

      if (!skorPerKompetensi[kom]) {
        skorPerKompetensi[kom] = { benar: 0, total: 0 };
      }
      skorPerKompetensi[kom].total++;
    }

    var totalBenar = 0;
    var jawabanSiswa = payload.jawaban || {};
    
    for (var idSoal in kunciMap) {
      var infoSoal = kunciMap[idSoal];
      var ansUser = jawabanSiswa[idSoal];

      if (ansUser !== undefined && ansUser !== null && ansUser !== "") {
        var isCorrect = false;

        if (infoSoal.tipeSoal === 'PGK-MCMA') {
          var kunciArr = infoSoal.kunci.split(',').map(function(s) { return s.trim().toUpperCase(); }).sort().join(',');
          var userArr = Array.isArray(ansUser) ? ansUser.map(function(s) { return s.trim().toUpperCase(); }).sort().join(',') : String(ansUser).trim().toUpperCase();
          if (kunciArr === userArr && userArr !== "") isCorrect = true;
        } 
        else if (infoSoal.tipeSoal === 'PGK-Kategori') {
          var formattedUserAns = "";
          if (typeof ansUser === 'object' && ansUser !== null && !Array.isArray(ansUser)) {
            var keys = Object.keys(ansUser).sort();
            formattedUserAns = keys.map(function(k) { return k + ":" + ansUser[k]; }).join(", ");
          } else {
            formattedUserAns = String(ansUser).trim();
          }
          if (formattedUserAns.toLowerCase() === infoSoal.kunci.toLowerCase()) isCorrect = true;
        } 
        else {
          if (String(ansUser).trim().toUpperCase() === infoSoal.kunci.toUpperCase()) isCorrect = true;
        }

        if (isCorrect) {
          totalBenar++;
          if (skorPerKompetensi[infoSoal.kompetensi]) {
            skorPerKompetensi[infoSoal.kompetensi].benar++;
          }
        }
      }
    }

    var skorAkhir = totalSoalPaket > 0 ? Math.round((totalBenar / totalSoalPaket) * 100) : 0;
    
    var pctTekstual = skorPerKompetensi["Pemahaman Tekstual"].total > 0 ? Math.round((skorPerKompetensi["Pemahaman Tekstual"].benar / skorPerKompetensi["Pemahaman Tekstual"].total) * 100) : 0;
    var pctInferensial = skorPerKompetensi["Pemahaman Inferensial"].total > 0 ? Math.round((skorPerKompetensi["Pemahaman Inferensial"].benar / skorPerKompetensi["Pemahaman Inferensial"].total) * 100) : 0;
    
    var evalTotal = (skorPerKompetensi["Evaluasi & Apresiasi"].total || 0) + (skorPerKompetensi["Evaluasi dan Apresiasi"].total || 0);
    var evalBenar = (skorPerKompetensi["Evaluasi & Apresiasi"].benar || 0) + (skorPerKompetensi["Evaluasi dan Apresiasi"].benar || 0);
    var pctEvaluasi = evalTotal > 0 ? Math.round((evalBenar / evalTotal) * 100) : 0;

    var diagnosa = [];
    if (pctTekstual < 70) diagnosa.push("Perlu peningkatan pemahaman informasi eksplisit.");
    if (pctInferensial < 70) diagnosa.push("Perlu latihan menarik kesimpulan tersirat.");
    if (pctEvaluasi < 70) diagnosa.push("Perlu memperdalam evaluasi argumen/opini.");
    if (diagnosa.length === 0) diagnosa.push("Sangat Baik! Semua kompetensi dikuasai.");

    var timestp = new Date();
    var sheetJawaban = ss.getSheetByName("JawabanMurid");
    var sheetAnalisis = ss.getSheetByName("AnalisisMurid");
    
    var jmlPelanggaran = payload.pelanggaran || 0;

    if (sheetJawaban) {
      sheetJawaban.appendRow([
        timestp, 
        payload.nama, 
        payload.kelas, 
        payload.jurusan, 
        payload.idPaket, 
        JSON.stringify(jawabanSiswa),
        jmlPelanggaran
      ]);
    }

    if (sheetAnalisis) {
      sheetAnalisis.appendRow([
        timestp, 
        payload.nama, 
        payload.kelas, 
        payload.jurusan, 
        payload.idPaket, 
        skorAkhir, 
        (pctTekstual / 100), 
        (pctInferensial / 100), 
        (pctEvaluasi / 100), 
        diagnosa.join(" "),
        jmlPelanggaran
      ]);
    }

    return {
      status: "success",
      nama: payload.nama,
      skorAkhir: skorAkhir,
      detailSkor: { tekstual: pctTekstual, inferensial: pctInferensial, evaluasi: pctEvaluasi },
      diagnosa: diagnosa.join(" ")
    };
  } finally {
    lock.releaseLock();
  }
}

function logToSheet(status, requestText, responseText) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("LogAI");
  if (sheet) sheet.appendRow([new Date(), status, requestText, responseText]);
}

function getPaketByToken(inputQuery) {
  try {
    if (!inputQuery) {
      throw new Error("Masukkan Token atau ID Paket terlebih dahulu.");
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName("DaftarPaket"); 
    
    if (!masterSheet) {
      throw new Error("Sheet 'DaftarPaket' tidak ditemukan!");
    }

    var masterData = masterSheet.getDataRange().getValues();
    var queryClean = inputQuery.toString().trim().toUpperCase();
    
    var targetSheetName = "";
    var namaPaket = "";
    var tokenPaket = "";
    var idPaket = "";

    for (var i = 1; i < masterData.length; i++) {
      var rowId = masterData[i][0] ? masterData[i][0].toString().trim().toUpperCase() : "";
      var rowNama = masterData[i][1] ? masterData[i][1].toString() : "";
      var rowSheet = masterData[i][2] ? masterData[i][2].toString().trim() : "";
      var rowToken = masterData[i][3] ? masterData[i][3].toString().trim().toUpperCase() : "";

      if (queryClean === rowToken || queryClean === rowId || queryClean === rowSheet.toUpperCase()) {
        idPaket = masterData[i][0];
        namaPaket = rowNama;
        targetSheetName = rowSheet;
        tokenPaket = rowToken;
        break;
      }
    }

    if (!targetSheetName) {
      throw new Error("Token atau ID Paket '" + inputQuery + "' tidak ditemukan di sheet DaftarPaket!");
    }

    var targetSheet = ss.getSheetByName(targetSheetName);
    if (!targetSheet) {
      throw new Error("Sheet database '" + targetSheetName + "' tidak ditemukan.");
    }

    var rows = targetSheet.getDataRange().getValues();
    if (rows.length < 2) {
      throw new Error("Sheet database '" + targetSheetName + "' masih kosong.");
    }

    var soalArray = [];

    for (var j = 1; j < rows.length; j++) {
      var row = rows[j];
      var item = {};
      
      item["idSoal"] = row[0];
      item["kompetensi"] = row[4];
      item["levelCEFR"] = row[5];
      item["jenisTeks"] = row[6];
      item["tipeSoal"] = row[7];
      item["tipeStimulus"] = row[8] || "text";
      item["stimulus"] = row[9];
      item["pertanyaan"] = row[10];
      item["pilihanJawaban"] = row[11];
      
      var kunciVal = row[12] ? String(row[12]).trim() : "A";
      item["kunci"] = kunciVal;
      item["kunciJawaban"] = kunciVal;

      item["pembahasan"] = row[13];
      item["indikator"] = row[14];

      soalArray.push(item);
    }

    return {
      status: "success",
      success: true,
      paket: {
        token: tokenPaket || idPaket,
        namaPaket: namaPaket,
        sheetName: targetSheetName
      },
      data: soalArray
    };

  } catch (err) {
    return { status: "error", success: false, message: err.message };
  }
}

/**
 * FUNGSI DASHBOARD GURU (v3.8) - Pencocokan Presisi Peta Jawaban & Pelanggaran
 */
function getDashboardData(idPaket, kelasFilter, jurusanFilter) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName("DaftarPaket");
    if (!masterSheet) throw new Error("Sheet DaftarPaket tidak ditemukan!");

    var masterData = masterSheet.getDataRange().getValues();
    var targetSheetName = "";
    var namaPaket = "";
    
    for (var i = 1; i < masterData.length; i++) {
      if (String(masterData[i][0]).trim() === String(idPaket).trim()) {
        namaPaket = masterData[i][1];
        targetSheetName = masterData[i][2];
        break;
      }
    }

    if (!targetSheetName) throw new Error("Paket Ujian tidak ditemukan!");

    var bankSheet = ss.getSheetByName(targetSheetName);
    var bankData = bankSheet ? bankSheet.getDataRange().getValues() : [];
    var totalSoalPaket = bankData.length > 1 ? bankData.length - 1 : 0;

    var mapSoal = {};
    for (var b = 1; b < bankData.length; b++) {
      var idSoal = bankData[b][0];
      mapSoal[idSoal] = {
        nomor: b,
        kompetensi: bankData[b][4],
        tipeSoal: bankData[b][7],
        pertanyaan: bankData[b][10],
        kunci: bankData[b][12],
        salahCount: 0,
        totalJawab: 0
      };
    }

    var analisisSheet = ss.getSheetByName("AnalisisMurid");
    var jawabanSheet = ss.getSheetByName("JawabanMurid");
    
    var mapJawabanSiswa = {};
    if (jawabanSheet) {
      var dataJawaban = jawabanSheet.getDataRange().getValues();
      for (var j = 1; j < dataJawaban.length; j++) {
        var keyNama = String(dataJawaban[j][1]).trim().toLowerCase();
        var keyPaket = String(dataJawaban[j][4]).trim();
        var keyMap = keyNama + "_" + keyPaket;
        mapJawabanSiswa[keyMap] = dataJawaban[j][5];
      }
    }

    var listSiswa = [];
    var totalNilaiSum = 0;
    var maxNilai = 0;
    var minNilai = 100;

    var sumTekstual = 0;
    var sumInferensial = 0;
    var sumEvaluasi = 0;
    var countSiswa = 0;

    function parsePersen(val) {
      if (val === undefined || val === null || val === "") return 0;
      var num = parseFloat(String(val).replace('%', '').trim());
      if (isNaN(num)) return 0;
      if (num > 0 && num <= 1) {
        num = Math.round(num * 100);
      } else {
        num = Math.round(num);
      }
      return num;
    }

    var cleanKelasFilter = String(kelasFilter || '').trim().toLowerCase();
    var cleanJurusanFilter = String(jurusanFilter || '').trim().toLowerCase();
    var selectedPaketStr = String(idPaket).trim();

    if (analisisSheet) {
      var dataAnalisis = analisisSheet.getDataRange().getValues();

      for (var a = 1; a < dataAnalisis.length; a++) {
        var namaSiswaRow = String(dataAnalisis[a][1]).trim();
        var rowKelas = String(dataAnalisis[a][2]).trim().toLowerCase();
        var rowJurusan = String(dataAnalisis[a][3]).trim().toLowerCase();
        
        var rowPaketVal = dataAnalisis[a][4] ? String(dataAnalisis[a][4]).trim() : "";
        var isNewFormat = (rowPaketVal.indexOf("PKT-") === 0);

        var matchPaket = false;
        if (isNewFormat) {
          matchPaket = (rowPaketVal === selectedPaketStr);
        } else {
          var lookupKeyTest = namaSiswaRow.toLowerCase() + "_" + selectedPaketStr;
          matchPaket = (mapJawabanSiswa[lookupKeyTest] !== undefined);
        }

        var matchKelas = (!cleanKelasFilter || cleanKelasFilter === 'semua' || rowKelas === cleanKelasFilter);
        var matchJurusan = (!cleanJurusanFilter || cleanJurusanFilter === 'semua' || rowJurusan.indexOf(cleanJurusanFilter) !== -1);

        if (matchPaket && matchKelas && matchJurusan) {
          countSiswa++;
          
          var nilai = isNewFormat ? (Number(dataAnalisis[a][5]) || 0) : (Number(dataAnalisis[a][4]) || 0);
          var pTekstual = isNewFormat ? parsePersen(dataAnalisis[a][6]) : parsePersen(dataAnalisis[a][5]);
          var pInferensial = isNewFormat ? parsePersen(dataAnalisis[a][7]) : parsePersen(dataAnalisis[a][6]);
          var pEvaluasi = isNewFormat ? parsePersen(dataAnalisis[a][8]) : parsePersen(dataAnalisis[a][7]);
          var diagnosa = isNewFormat ? (dataAnalisis[a][9] || "") : (dataAnalisis[a][8] || "");
          var pPelanggaran = isNewFormat ? (Number(dataAnalisis[a][10]) || 0) : (Number(dataAnalisis[a][9]) || 0);

          totalNilaiSum += nilai;
          if (nilai > maxNilai) maxNilai = nilai;
          if (nilai < minNilai) minNilai = nilai;

          sumTekstual += pTekstual;
          sumInferensial += pInferensial;
          sumEvaluasi += pEvaluasi;

          listSiswa.push({
            nama: namaSiswaRow,
            kelas: dataAnalisis[a][2] + " " + dataAnalisis[a][3],
            nilai: nilai,
            tekstual: pTekstual,
            inferensial: pInferensial,
            evaluasi: pEvaluasi,
            diagnosa: diagnosa,
            pelanggaran: pPelanggaran
          });

          var lookupKey = namaSiswaRow.toLowerCase() + "_" + selectedPaketStr;
          var rawJawabanJson = mapJawabanSiswa[lookupKey];

          if (rawJawabanJson) {
            try {
              var userAnsObj = JSON.parse(rawJawabanJson);
              for (var sId in mapSoal) {
                mapSoal[sId].totalJawab++;
                var uAns = userAnsObj[sId];
                var infoSoal = mapSoal[sId];
                var isCorrect = false;

                if (uAns !== undefined && uAns !== null && uAns !== "") {
                  if (infoSoal.tipeSoal === 'PGK-MCMA') {
                    var kunciArr = infoSoal.kunci.split(',').map(function(s) { return s.trim().toUpperCase(); }).sort().join(',');
                    var userArr = Array.isArray(uAns) ? uAns.map(function(s) { return s.trim().toUpperCase(); }).sort().join(',') : String(uAns).trim().toUpperCase();
                    if (kunciArr === userArr && userArr !== "") isCorrect = true;
                  } 
                  else if (infoSoal.tipeSoal === 'PGK-Kategori') {
                    var formattedUserAns = "";
                    if (typeof uAns === 'object' && uAns !== null && !Array.isArray(uAns)) {
                      var keys = Object.keys(uAns).sort();
                      formattedUserAns = keys.map(function(k) { return k + ":" + uAns[k]; }).join(", ");
                    } else {
                      formattedUserAns = String(uAns).trim();
                    }
                    if (formattedUserAns.toLowerCase() === String(infoSoal.kunci).trim().toLowerCase()) isCorrect = true;
                  } 
                  else {
                    if (String(uAns).trim().toUpperCase() === String(infoSoal.kunci).trim().toUpperCase()) isCorrect = true;
                  }
                }

                if (!isCorrect) {
                  mapSoal[sId].salahCount++;
                }
              }
            } catch(e) {}
          }
        }
      }
    }

    if (countSiswa === 0) minNilai = 0;

    var avgNilai = countSiswa > 0 ? (totalNilaiSum / countSiswa).toFixed(1) : 0;
    var avgTekstual = countSiswa > 0 ? Math.round(sumTekstual / countSiswa) : 0;
    var avgInferensial = countSiswa > 0 ? Math.round(sumInferensial / countSiswa) : 0;
    var avgEvaluasi = countSiswa > 0 ? Math.round(sumEvaluasi / countSiswa) : 0;

    var listAnalisisSoal = [];
    for (var key in mapSoal) {
      var item = mapSoal[key];
      var pctSalah = item.totalJawab > 0 ? Math.round((item.salahCount / item.totalJawab) * 100) : 0;
      listAnalisisSoal.push({
        nomor: item.nomor,
        kompetensi: item.kompetensi,
        pertanyaan: item.pertanyaan,
        kunci: item.kunci,
        salahCount: item.salahCount,
        pctSalah: pctSalah
      });
    }

    listAnalisisSoal.sort(function(a, b) { return b.pctSalah - a.pctSalah; });

    return {
      status: "success",
      namaPaket: namaPaket,
      totalSoal: totalSoalPaket,
      totalPeserta: countSiswa,
      metrics: {
        avgNilai: avgNilai,
        maxNilai: maxNilai,
        minNilai: minNilai,
        avgTekstual: avgTekstual,
        avgInferensial: avgInferensial,
        avgEvaluasi: avgEvaluasi
      },
      soalSulit: listAnalisisSoal.slice(0, 5),
      listSiswa: listSiswa
    };

  } catch (err) {
    return { status: "error", message: err.message };
  }
}