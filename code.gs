const SHEET_SOAL = "BankSoal";
const SHEET_HASIL = "HasilUjian";

/**
 * Jalankan fungsi ini SATU KALI dari editor Apps Script
 * untuk membuat sheet otomatis dan mengisi 64 butir soal.
 */
function setupDanIsiBankSoal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let sheetSoal = ss.getSheetByName(SHEET_SOAL);
  if (!sheetSoal) {
    sheetSoal = ss.insertSheet(SHEET_SOAL);
  }
  sheetSoal.clear();
  
  const headerSoal = [
    "No", "Tipe", "Kategori", "Wacana", "Pertanyaan", 
    "OpsiA", "OpsiB", "OpsiC", "OpsiD", "Kunci", "Pembahasan"
  ];
  sheetSoal.appendRow(headerSoal);
  sheetSoal.getRange(1, 1, 1, headerSoal.length).setFontWeight("bold").setBackground("#e2e8f0");

  const dataSoal = getData64Soal();
  const rows = dataSoal.map(item => [
    item.no, item.tipe, item.kategori, item.wacana || "", item.pertanyaan,
    item.a, item.b, item.c, item.d, item.kunci, item.pembahasan
  ]);
  
  sheetSoal.getRange(2, 1, rows.length, headerSoal.length).setValues(rows);

  let sheetHasil = ss.getSheetByName(SHEET_HASIL);
  if (!sheetHasil) {
    sheetHasil = ss.insertSheet(SHEET_HASIL);
    const headerHasil = [
      "WaktuSubmit", "NamaPeserta", "Instansi", "Jenjang", 
      "SkorAkhir", "BenarSJT", "BenarStudiKasus", "TotalBenar", "RincianJawaban"
    ];
    sheetHasil.appendRow(headerHasil);
    sheetHasil.getRange(1, 1, 1, headerHasil.length).setFontWeight("bold").setBackground("#e2e8f0");
  }
}

/**
 * Endpoint GET: Mengambil seluruh bank soal untuk ditampilkan di GitHub Pages
 */
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_SOAL);
    if (!sheet) {
      return responseJSON({ status: "error", message: "Sheet BankSoal belum dibuat. Jalankan fungsi setupDanIsiBankSoal terlebih dahulu." });
    }

    const values = sheet.getDataRange().getValues();
    const questions = [];

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (!row[0]) continue;
      questions.push({
        no: Number(row[0]),
        tipe: String(row[1]),
        kategori: String(row[2]),
        wacana: String(row[3]),
        pertanyaan: String(row[4]),
        opsi: {
          A: String(row[5]),
          B: String(row[6]),
          C: String(row[7]),
          D: String(row[8])
        },
        kunci: String(row[9]).trim().toUpperCase(),
        pembahasan: String(row[10])
      });
    }

    return responseJSON({ status: "success", data: questions });
  } catch (err) {
    return responseJSON({ status: "error", message: err.toString() });
  }
}

/**
 * Endpoint POST: Menyimpan hasil pengerjaan peserta ke sheet HasilUjian
 */
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_HASIL);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_HASIL);
      sheet.appendRow([
        "WaktuSubmit", "NamaPeserta", "Instansi", "Jenjang", 
        "SkorAkhir", "BenarSJT", "BenarStudiKasus", "TotalBenar", "RincianJawaban"
      ]);
    }

    const payload = JSON.parse(e.postData.contents);
    const waktu = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");

    sheet.appendRow([
      waktu,
      payload.nama || "-",
      payload.instansi || "-",
      payload.jenjang || "-",
      payload.skorAkhir || 0,
      payload.benarSJT || 0,
      payload.benarKasus || 0,
      payload.totalBenar || 0,
      payload.rincian || ""
    ]);

    return responseJSON({ status: "success", message: "Hasil ujian berhasil disimpan." });
  } catch (err) {
    return responseJSON({ status: "error", message: err.toString() });
  }
}

function responseJSON(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Database 64 Butir Soal Uji Kompetensi JF Pengawas Sekolah 2026
 */
function getData64Soal() {
  const kasus1 = "Anda mendampingi tiga SMP binaan yang baru saja menyelaraskan Kurikulum Satuan Pendidikan (KSP) tahun ajaran 2025/2026 merujuk Permendikdasmen Nomor 13 Tahun 2025 serta Keputusan BSKAP Nomor 058/H/KR/2025. Hasil pemantauan awal menunjukkan tiga persoalan krusial. Pertama, sejumlah guru menggunakan perangkat Kecerdasan Artifisial (KA) generatif secara mentah untuk menyusun modul ajar dan membiarkan murid menyalin jawaban tugas dari KA tanpa proses berpikir kritis, sebuah risiko yang juga diidentifikasi dalam kajian kebijakan literasi digital BSKAP. Kedua, program Kokurikuler dan Gerakan 7 Kebiasaan Anak Indonesia Hebat (7KAIH) diterapkan secara kaku melalui sistem poin hukuman bagi murid yang terlambat bangun atau lupa mengisi jurnal. Ketiga, dua guru senior menyatakan keberatan atas pengenalan pembelajaran Koding dan KA serta menolak mengubah metode ceramah karena menganggap beban administrasi semakin berat. Sementara itu, sebagian orang tua juga mengeluhkan banyaknya tugas pengamatan proyek di rumah yang dianggap tidak jelas kaitannya dengan nilai akademik.";
  
  const kasus2 = "Pada pertengahan tahun 2026, Anda mendampingi empat satuan pendidikan (satu satuan PAUD, satu SD, dan dua SMP) dalam menyusun Perencanaan Berbasis Data (PBD) menggunakan Rapor Pendidikan Tahun Data 2025 yang dirilis pada pembaruan sistem tahun 2026. Saat eksplorasi dasbor, terdapat tiga temuan penting: (1) Satu SD yang telah memperbaiki program numerasi sejak awal 2026 heran karena skor numerasi pada Rapor Pendidikan rilis 2026 masih merah, sementara satu satuan PAUD tidak memunculkan label warna capaian indikator maupun skor sama sekali karena partisipasi pengisian Survei Lingkungan Belajar (Sulingjar) PTK belum mencapai 60%; (2) Pada dua SMP, tim sekolah menyalin seluruh daftar contoh kegiatan ARKAS dari berkas unduhan tanpa menganalisis akar masalah sehingga anggaran terserap pada belanja rutin yang tidak menyentuh mutu kelas; dan (3) Menjelang pelaksanaan Tes Kemampuan Akademik (TKA) jenjang SD dan SMP tahun 2026 sesuai Permendikdasmen Nomor 9 Tahun 2025, salah satu kepala sekolah berencana mewajibkan bimbingan tes berbayar bagi seluruh murid kelas 6 serta mempertimbangkan mengecualikan murid berkebutuhan khusus non-disabilitas intelektual agar rata-rata capaian sekolah tetap tinggi.";

  const kasus3 = "Anda bertugas memantau tata kelola manajerial pada enam sekolah binaan di sebuah kabupaten. Menjelang periode penataan jabatan tahun 2026, dua kepala sekolah definitif memasuki masa pensiun dalam enam bulan ke depan, tetapi belum banyak guru yang mendaftar karena menganggap harus memiliki sertifikat Guru Penggerak. Di saat bersamaan, seorang kepala sekolah berkinerja sangat baik baru menjabat selama 15 bulan di sekolah pinggiran berprestasi, namun diusulkan oleh oknum pejabat daerah untuk segera dirotasi ke sekolah inti di kota. Pada aspek tenaga kependidikan lainnya, hasil pemantauan Anda menunjukkan dua sekolah memiliki tenaga administrasi dan operator data yang baru lulus pendidikan menengah tanpa sertifikat kompetensi sesuai bidang tugas, dan masih sering terjadi ketidaksesuaian sinkronisasi Dapodik dengan data Ruang GTK yang menghambat verifikasi data guru.";

  const kasus4 = "Merujuk ketentuan PermenPANRB Nomor 7 Tahun 2026 sebagaimana dilaporkan situs analisis kebijakan pendidikan dan media berita pendidikan, JF Pengawas Sekolah kembali ditetapkan sebagai jabatan fungsional tersendiri (Ahli Muda, Ahli Madya, dan Ahli Utama) yang terpisah dari JF Guru. Menyambut penataan baru ini, timbul dua dinamika di forum pengawas kabupaten Anda: (1) Sebagian pengawas beranggapan bahwa dengan terpisahnya JF Pengawas Sekolah dari JF Guru, tugas pengawas kembali menjadi 'inspektur penilai kesalahan administratif' seperti pola lama dan tidak lagi perlu mendampingi kepala sekolah; (2) Pembagian wilayah tugas pengawas belum membedakan fokus peran antara jenjang Ahli Muda, Ahli Madya, dan Ahli Utama di tengah kondisi geografis kepulauan atau pegunungan yang sulit. Selain itu, banyak guru dan tenaga kependidikan di daerah Anda masih bingung membedakan fungsi Ruang GTK (dalam super-app Rumah Pendidikan) dengan Info GTK, sehingga kerap mengadukan persoalan status tunjangan profesi ke layanan pusat bantuan Ruang GTK.";

  const kasus5 = "Pelaksanaan Sistem Penerimaan Murid Baru (SPMB) berdasarkan Permendikdasmen Nomor 3 Tahun 2025 di salah satu SMA favorit binaan Anda memicu polemik masyarakat. Setelah pengumuman hasil seleksi, puluhan orang tua calon murid berunjuk rasa di depan sekolah karena mempertanyakan transparansi pengisian kuota jalur seleksi dan mengeluhkan kurang jelasnya informasi alur verifikasi berkas. Di sisi lain, hasil penelusuran awal Anda menemukan bahwa panitia sekolah memang kurang rapi dalam mengomunikasikan informasi daya tampung dan publikasi prosedur verifikasi secara terbuka. Menjelang kunjungan pemantauan dari pemerintah daerah, seorang pejabat lokal meminta Anda membuat laporan singkat bahwa seluruh proses SPMB di sekolah tersebut telah berjalan '100% tanpa kendala dan tanpa keluhan' demi menjaga citra baik daerah.";

  return [
    {
      no: 1, tipe: "SJT", kategori: "Manajerial - Integritas", wacana: "",
      pertanyaan: "Sebagai Pengawas Sekolah, Anda sering mendengar beberapa Kepala Sekolah mengeluh tentang kinerja pengawas sekolah lainnya. Beberapa Kepala Sekolah merasa kecewa dan ingin pengawas sekolah yang mendampingi mereka lebih memahami kondisi dan kebutuhan sekolah. Dalam situasi ini, upaya yang seharusnya dilakukan sebagai Pengawas Sekolah adalah ....",
      a: "Menyampaikan keresahannya dengan apa yang terjadi di lapangan dalam forum pengawas sekolah, agar pengawas sekolah lain sadar dan berubah untuk berperilaku sesuai kode etik dan norma pengawas sekolah.",
      b: "Melakukan refleksi diri setiap saat, agar tetap berpegang teguh pada kode etik dan norma sebagai pengawas sekolah, dan sesekali meminta umpan balik dari pengawas sekolah lainnya.",
      c: "Memberikan usul untuk membuat sesi rutin refleksi dan umpan balik dalam forum pengawas sekolah, sehingga hal-hal yang melanggar kode etik akan terus direfleksikan dan diperbaiki bersama.",
      d: "Mengingatkan pengawas sekolah lain untuk tetap berpedoman pada kode etik dan norma sebagai seorang pengawas sekolah.",
      kunci: "C",
      pembahasan: "Pengawas sekolah berjenjang tinggi dituntut membangun sistem dan budaya organisasi secara kolektif, bukan sekadar mengingatkan secara individual. Mengusulkan sesi rutin refleksi dan umpan balik dalam forum pengawas (Opsi C) melembagakan mekanisme perbaikan berkelanjutan guna menjaga kode etik profesi tanpa menimbulkan kesan menggurui rekan sejawat."
    },
    {
      no: 2, tipe: "SJT", kategori: "Teknis - Profesional", wacana: "",
      pertanyaan: "Dalam masa transisi penerapan kebijakan Pembelajaran Mendalam dan penyelarasan Kurikulum Satuan Pendidikan (KSP) sesuai Permendikdasmen Nomor 13 Tahun 2025, sebagian kepala sekolah dan guru di sekolah binaan Anda masih merasa kebingungan serta terbebani dengan pola pendampingan baru yang diterapkan. Sebagai Pengawas Sekolah, tindakan yang paling tepat Anda lakukan untuk membangun pemahaman bersama dan mendorong komitmen perubahan di sekolah binaan adalah ....",
      a: "Menginstruksikan kepala sekolah dan guru untuk segera menyesuaikan diri dengan kebijakan baru karena aturan tersebut bersifat wajib dari pusat.",
      b: "Menyampaikan rincian tugas administratif yang harus dipenuhi sekolah secara tertulis agar tidak terjadi kesalahan dalam pelaporan.",
      c: "Membuka ruang dialog bersama kepala sekolah dan guru untuk mendengarkan kendala mereka, lalu menyepakati tahapan pendampingan yang sesuai dengan kesiapan masing-masing sekolah.",
      d: "Menunggu sekolah binaan siap secara mandiri sebelum memberikan pendampingan lebih lanjut agar tidak menambah beban kerja guru.",
      kunci: "C",
      pembahasan: "Paradigma Perdirjen GTK Nomor 4831/B/HK.03.01/2023 menempatkan pengawas sebagai pendamping yang membangun komitmen perubahan secara partisipatif dan berdiferensiasi. Membuka ruang dialog untuk mendengarkan kendala satuan pendidikan terlebih dahulu sebelum menyepakati tahapan pendampingan sesuai tingkat kesiapan sekolah (Opsi C) jauh lebih efektif daripada instruksi sepihak (A) maupun menunggu pasif (D)."
    },
    {
      no: 3, tipe: "SJT", kategori: "Teknis - Profesional", wacana: "",
      pertanyaan: "Dalam pelaksanaan penjaminan mutu dan penyesuaian kurikulum sesuai Permendikdasmen Nomor 13 Tahun 2025, banyak sekali informasi baru yang perlu disampaikan kepada sekolah binaan, mulai dari Pembelajaran Mendalam, kokurikuler 8 Dimensi Profil Lulusan, asesmen, hingga tata kelola administrasi. Tindakan yang seharusnya Anda lakukan sebagai Pengawas Sekolah agar informasi tersebut dipahami dan dijalankan sesuai kesiapan sekolah di daerah Anda adalah ....",
      a: "Menggunakan gaya komunikasi informal agar sekolah binaan merasa nyaman saat kegiatan sosialisasi informasi tersebut.",
      b: "Membuat media penyampaian informasi yang memudahkan sekolah binaan memahami kebijakan kurikulum terbaru.",
      c: "Membuat sesi sosialisasi ke dalam beberapa tahapan, agar sekolah binaan bisa memahami sedikit demi sedikit.",
      d: "Mengintegrasikan informasi dengan membuat pemahaman konsep yang kontekstual dan mudah dipahami oleh sekolah binaan.",
      kunci: "D",
      pembahasan: "Hambatan utama penerapan regulasi baru bukan sekadar kurangnya media (B) atau pemecahan jadwal sosialisasi (C), melainkan kesenjangan antara teori regulasi dan realitas sekolah. Mengintegrasikan informasi menjadi pemahaman konsep yang kontekstual (Opsi D) memampukan sekolah menerjemahkan kebijakan sesuai karakteristik dan kesiapannya."
    },
    {
      no: 4, tipe: "SJT", kategori: "Teknis - Profesional", wacana: "",
      pertanyaan: "Selama menjabat sebagai pengawas sekolah, Anda telah menemukan banyak praktik baik yang efektif dalam mendampingi sekolah binaan. Namun, Anda menyadari bahwa tidak semua pengawas sekolah melakukan hal yang sama. Sesuai peran organisasi profesi pengawas sekolah, langkah paling strategis untuk membagikan praktik baik teruji dan mendorong pengawas lain mengikutinya adalah ....",
      a: "Mengapresiasi pengawas sekolah lain yang sudah melakukan praktik baik agar terus mempertahankannya.",
      b: "Dalam organisasi profesi pengawas sekolah yang Anda ikuti, mengusulkan beberapa praktik baik yang teruji menjadi kebijakan dalam organisasi untuk bisa diikuti oleh pengawas sekolah lainnya.",
      c: "Memberikan masukan kepada pengawas sekolah lain untuk terus memperbaiki peranannya.",
      d: "Menyampaikan praktik baik yang teruji dalam forum pengawas sekolah.",
      kunci: "B",
      pembahasan: "Pengawas sekolah wajib menjadi anggota organisasi profesi, dan jenjang madya hingga utama menuntut dampak pengembangan mutu sistem pengawasan. Melembagakan praktik baik yang sudah teruji menjadi kebijakan atau standar bersama dalam organisasi profesi (Opsi B) menciptakan dampak yang terstruktur dan berkelanjutan dibandingkan sekadar berbagi secara lisan (D)."
    },
    {
      no: 5, tipe: "SJT", kategori: "Teknis - Profesional", wacana: "",
      pertanyaan: "Sekolah binaan Anda memiliki beberapa kelas paralel di tingkatan yang sama yang diajar oleh guru berbeda di setiap kelasnya. Setelah Anda amati secara saksama melalui pemantauan akademik, terdapat perbedaan metode pengajaran yang signifikan antarkelas pada tingkatan yang sama. Sebagai upaya menjamin pemerataan mutu pembelajaran di setiap kelas paralel, yang Anda lakukan adalah ....",
      a: "Melakukan pelatihan untuk menyamakan kompetensi dan metode pengajaran setiap guru yang ada serta melakukan pemantauan secara berkala.",
      b: "Perbedaan metode pengajaran tidak selalu membuat prestasi siswa berbeda sehingga Anda membiarkan saja perbedaan yang ada.",
      c: "Anda mengumpulkan guru-guru terkait dan menyampaikan apa yang harus mereka lakukan secara langsung.",
      d: "Melakukan koordinasi dengan Kepala Sekolah agar Kepala Sekolah melakukan penyetaraan di setiap kelas.",
      kunci: "A",
      pembahasan: "Sesuai Pasal 16 ayat (3) Permendikdasmen Nomor 23 Tahun 2026, kompetensi profesional pengawas meliputi pemantauan, penilaian, dan pembinaan akademik. Melakukan pelatihan penyetaraan kompetensi guru yang diikuti pemantauan berkala (Opsi A) menjamin pemerataan mutu layanan bagi murid lintas kelas paralel secara lebih terukur daripada instruksi satu arah (C) atau pendelegasian pasif (D)."
    },
    {
      no: 6, tipe: "SJT", kategori: "Manajerial - Pelayanan Publik", wacana: "",
      pertanyaan: "Salah satu sekolah binaan Anda memiliki gedung serbaguna yang cukup luas dan sering digunakan oleh masyarakat sekitar untuk acara-acara nonakademik. Setelah beberapa kali penggunaan, Anda menemukan beberapa kerusakan dan masalah kebersihan setelah gedung tersebut digunakan oleh masyarakat. Agar hubungan antara pihak sekolah dan masyarakat tetap terjalin baik namun tidak ada pihak yang dirugikan, tindakan yang Anda lakukan adalah ....",
      a: "Mengadakan uang iuran gedung bagi masyarakat sekitar sehingga masyarakat memiliki tanggung jawab untuk merawat gedung, dan iuran digunakan untuk pemeliharaan fasilitas secara berkala.",
      b: "Menerapkan biaya sewa gedung per jam yang tinggi untuk menurunkan keinginan warga sekitar menggunakan gedung serbaguna milik sekolah.",
      c: "Membuat kebijakan baru dengan mengizinkan penggunaan gedung serbaguna oleh semua pihak namun akan dikenakan denda jika ditemukan kerusakan setelah penggunaan gedung.",
      d: "Mengadakan peraturan baru yang disosialisasikan ke masyarakat setempat bahwa gedung serbaguna hanya dapat dipakai di jam sekolah dan hanya untuk kepentingan sekolah saja.",
      kunci: "A",
      pembahasan: "Kompetensi Pelayanan Publik dan kemitraan sekolah-masyarakat mengutamakan solusi saling menguntungkan tanpa pembatasan kaku (D) ataupun pendekatan punitif (C). Kesepakatan iuran pemeliharaan fasilitas secara berkala (Opsi A) menumbuhkan rasa kepemilikan bersama sekaligus menjaga keberlangsungan perawatan gedung sekolah."
    },
    {
      no: 7, tipe: "SJT", kategori: "Teknis - Profesional", wacana: "",
      pertanyaan: "Menurut pengamatan Anda, salah satu sekolah binaan Anda cukup tertinggal dari segi proses pembelajaran dan variasi metode pengajaran guru. Tindakan pembinaan akademik yang paling tepat Anda lakukan adalah ....",
      a: "Berdiskusi dengan kepala sekolah mengenai kemungkinan-kemungkinan inovasi baru yang dapat dilakukan.",
      b: "Menyampaikan pandangan Anda kepada Kepala Sekolah untuk dipikirkan oleh Kepala Sekolah.",
      c: "Membiarkan sekolah binaan Anda berjalan seperti biasa karena mengubah tradisi terlalu berisiko.",
      d: "Perlahan-lahan mengubah metode yang lama, sambil mendampingi dan memberikan pelatihan yang sesuai kepada guru.",
      kunci: "D",
      pembahasan: "Transformasi pembelajaran pada satuan pendidikan yang tertinggal memerlukan manajemen perubahan bertahap agar tidak memicu resistensi. Mengubah metode lama secara perlahan disertai pendampingan dan pelatihan kontekstual kepada guru (Opsi D) merealisasikan fungsi pembinaan akademik dalam Permendikdasmen Nomor 23 Tahun 2026."
    },
    {
      no: 8, tipe: "SJT", kategori: "Teknis - Profesional", wacana: "",
      pertanyaan: "Saat mendampingi sekolah binaan dalam pelaksanaan kegiatan kokurikuler penguatan 8 Dimensi Profil Lulusan sesuai Permendikdasmen Nomor 13 Tahun 2025 dan Keputusan BSKAP Nomor 058/H/KR/2025, ditemukan bahwa banyak sekolah menganggap produk pameran sebagai tujuan utama kegiatan sehingga penguatan delapan dimensi profil lulusan tidak tercapai. Fenomena tersebut juga terjadi di banyak sekolah di daerah Anda. Langkah yang bisa dilakukan agar masalah tersebut diatasi secara efektif adalah ....",
      a: "Menyampaikan temuan di lapangan tentang pelaksanaan kokurikuler beserta data-data tidak tercapainya tujuan dimensi yang diharapkan, mengajak pengawas lain melakukan refleksi, dan mendiskusikan strategi yang lebih baik.",
      b: "Mengusulkan cara baru dalam sosialisasi kokurikuler agar sekolah tidak lagi fokus pada produk, melainkan pada delapan dimensi profil lulusan.",
      c: "Mengajak sekolah binaan untuk merefleksikan pelaksanaan kokurikuler dan memperbaiki prosesnya ke depan agar tidak fokus pada produk lagi.",
      d: "Mengajak pengawas lain untuk merefleksikan pelaksanaan kokurikuler di sekolah-sekolah binaan, apakah sudah sesuai dengan tujuan esensial yang diharapkan atau masih fokus pada produk.",
      kunci: "A",
      pembahasan: "Karena miskonsepsi terjadi secara luas lintas sekolah di daerah, intervensi pengawas perlu berbasis bukti: menyajikan data temuan lapangan, mengajak pengawas lain berefleksi, sekaligus merumuskan strategi perbaikan bersama (Opsi A), lebih lengkap daripada Opsi D yang belum mencakup data lapangan maupun strategi tindak lanjut."
    },
    {
      no: 9, tipe: "SJT", kategori: "Teknis - Kepribadian", wacana: "",
      pertanyaan: "Meskipun Pengawas Sekolah seharusnya berperan sebagai pendamping kepala satuan pendidikan sesuai Perdirjen GTK Nomor 4831/B/HK.03.01/2023, sering kali terlihat banyak pengawas yang masih terlalu mengendalikan situasi di sekolah binaannya. Upaya yang paling efektif agar lebih banyak pengawas mulai memahami peran dan fungsinya sebagai pendamping adalah ....",
      a: "Memberikan masukan pada forum pengawas sekolah di daerah untuk membuat kebijakan peran pengawas agar tidak lagi melakukan miskonsepsi.",
      b: "Dalam forum pengawas sekolah, meminta pengawas sekolah lain untuk kembali merefleksikan peran mereka di sekolah.",
      c: "Memberitahu pengawas sekolah lain untuk menjadi pengawas sekolah yang bukan mengendalikan sekolah binaan, melainkan mendampingi dan menemani.",
      d: "Dalam organisasi profesi pengawas sekolah yang Anda ikuti, terus menunjukkan peran-peran pengawas sekolah sebagai pendamping dalam setiap interaksi dengan pengawas sekolah lain.",
      kunci: "D",
      pembahasan: "Kompetensi kepribadian pengawas menuntut keteladanan perilaku dan pola pikir bertumbuh. Menunjukkan secara konsisten peran pengawas sebagai pendamping dalam setiap interaksi profesional di organisasi profesi (Opsi D) menggeser paradigma sejawat lebih kuat daripada imbauan lisan (C) atau teguran terbuka (B)."
    },
    {
      no: 10, tipe: "SJT", kategori: "Manajerial - Pelayanan Publik", wacana: "",
      pertanyaan: "Sebagai Pengawas Sekolah, Anda mendampingi beberapa sekolah dengan karakteristik dan kondisi geografis yang beragam. Isu kesenjangan sempat muncul karena beberapa sekolah merasa Anda lebih memperhatikan salah satu sekolah dibandingkan yang lain. Agar bersikap adil dan setara terhadap seluruh sekolah binaan, upaya yang seharusnya Anda lakukan adalah ....",
      a: "Membuat forum sekolah binaan untuk mendengarkan kebutuhan dan tantangan yang dihadapi sekolah, serta membuat kesepakatan pendampingan yang diketahui bersama.",
      b: "Memperlakukan cara pendampingan yang sama persis pada setiap sekolah binaan.",
      c: "Membagi waktu pendampingan dengan adil dan disamaratakan antar sekolah agar tidak terjadi kecemburuan.",
      d: "Tidak membanding-bandingkan sekolah binaan satu dan lainnya karena setiap sekolah memiliki kelebihan dan kebutuhan masing-masing.",
      kunci: "A",
      pembahasan: "Beban kerja kepengawasan mempertimbangkan jumlah, bentuk satuan pendidikan, dan kondisi geografis. Prinsip keadilan layanan publik bukanlah penyamarataan waktu secara kaku (B, C), melainkan keadilan proporsional berbasis pemetaan kebutuhan dan tantangan melalui kesepakatan terbuka dalam forum sekolah binaan (Opsi A)."
    },
    {
      no: 11, tipe: "SJT", kategori: "Manajerial - Pengambilan Keputusan", wacana: "",
      pertanyaan: "Kepala Sekolah di salah satu sekolah binaan Anda melakukan pelanggaran cukup berat terkait Hak Asasi Manusia (HAM) sehingga membuat suasana di sekolah tersebut tidak kondusif dan guru-guru menunjukkan ketidaksukaan secara terbuka. Untuk mengembalikan kondusivitas sekolah sekaligus menegakkan akuntabilitas, yang Anda lakukan adalah ....",
      a: "Tidak melakukan apa-apa agar kegiatan belajar mengajar berlangsung seperti biasa sambil membawa masalah ini ke Kepala Dinas dan menunggu keputusan Kepala Dinas.",
      b: "Membawa masalah kepada Kepala Dinas untuk ditindaklanjuti dan mengimbau warga sekolah mengawasi tingkah laku Kepala Sekolah lalu melaporkannya kepada Anda.",
      c: "Memberikan peringatan keras bagi Kepala Sekolah dan mengawasi Kepala Sekolah tersebut secara ketat selama beberapa bulan.",
      d: "Memberikan peringatan kepada Kepala Sekolah dan membawa masalah ini ke Kepala Dinas untuk ditindaklanjuti serta memberikan pembinaan kepada Kepala Sekolah.",
      kunci: "D",
      pembahasan: "Pengawas sekolah bertanggung jawab langsung kepada pejabat pimpinan tinggi pratama atau pejabat administrator terkait di Dinas Pendidikan. Menghadapi pelanggaran berat HAM, pengawas bertindak sesuai kewenangannya: memberi peringatan dan pembinaan untuk memulihkan suasana sekolah, sekaligus melaporkan ke Kepala Dinas guna penindakan kepegawaian lebih lanjut (Opsi D)."
    },
    {
      no: 12, tipe: "SJT", kategori: "Teknis - Sosial", wacana: "",
      pertanyaan: "Sekolah binaan Anda terletak di wilayah rawan banjir yang hampir setiap tahun berdampak besar terhadap aktivitas wilayah sekitar sekolah. Anda mengamati salah satu penyebab terjadinya banjir ialah kebiasaan membuang sampah sembarangan. Tindakan yang paling tepat Anda lakukan sebagai Pengawas Sekolah adalah ....",
      a: "Sebulan sekali melakukan kegiatan gotong royong bersih-bersih oleh seluruh warga sekolah.",
      b: "Mengimbau para guru untuk menerapkan pentingnya menjaga kebersihan lingkungan kepada siswa guna mengurangi risiko banjir di masa depan.",
      c: "Bersama warga sekolah dan masyarakat setempat melakukan kegiatan gotong royong bersih-bersih sebulan sekali.",
      d: "Menggerakkan setiap bagian dari sekolah binaan Anda untuk menerapkan kebiasaan membuang sampah yang benar.",
      kunci: "D",
      pembahasan: "Pengawas berperan menggerakkan budaya sekolah sehat dan penguatan karakter melalui kolaborasi seluruh warga satuan pendidikan. Menggerakkan seluruh unsur sekolah membangun pembiasaan perilaku membuang sampah secara benar dalam keseharian (Opsi D) lebih berkelanjutan dibandingkan kerja bakti bulanan (A, C) atau imbauan pasif (B)."
    },
    {
      no: 13, tipe: "SJT", kategori: "Teknis - Profesional", wacana: "",
      pertanyaan: "Ketika membaca Rapor Pendidikan Tahun Data 2025 pada sekolah binaan, Anda menemukan bahwa setiap sekolah memiliki kompetensi yang masih rendah. Anda merasa perlu mengajak setiap sekolah melakukan refleksi dan mengambil tindakan perbaikan dalam siklus Perencanaan Berbasis Data (PBD). Tindakan yang seharusnya Anda lakukan untuk membantu sekolah binaan meningkatkan kompetensinya adalah ....",
      a: "Memberikan solusi-solusi sesuai kebutuhan sekolah agar bisa langsung diimplementasikan pada masing-masing sekolah.",
      b: "Melakukan refleksi masing-masing sekolah, memetakan kebutuhan belajar bersama, memberikan umpan balik, dan mengapresiasi ide-ide yang disampaikan sekolah.",
      c: "Mendengarkan apa yang ingin dilakukan masing-masing sekolah dari hasil Rapor Pendidikan tersebut, lalu memberikan beberapa solusi yang wajib dilakukan sekolah.",
      d: "Mengajak refleksi kepala sekolah saja dan mendiskusikan masalah serta solusi untuk meningkatkan poin-poin yang rendah dalam Rapor Pendidikan.",
      kunci: "B",
      pembahasan: "Siklus Perencanaan Berbasis Data (Identifikasi-Refleksi-Benahi) dan pendekatan pendampingan menuntut pengawas memfasilitasi refleksi, memetakan kebutuhan belajar bersama, memberi umpan balik, serta mengapresiasi gagasan sekolah (Opsi B), bukan menyuapi solusi siap pakai (A)."
    },
    {
      no: 14, tipe: "SJT", kategori: "Manajerial - Integritas", wacana: "",
      pertanyaan: "Keluarga dekat Anda mengetahui bahwa Anda adalah Pengawas Sekolah dan meminta Anda mengosongkan satu kursi pada pelaksanaan Sistem Penerimaan Murid Baru (SPMB) tahun ajaran berikutnya untuk anaknya di salah satu sekolah binaan favorit di daerah tersebut. Tindakan yang Anda lakukan adalah ....",
      a: "Menyarankan mereka mengikuti tata cara yang seharusnya, tetapi tanpa sepengetahuan keluarga Anda sudah menyiapkan satu tempat agar tetap terlihat taat peraturan.",
      b: "Menolak memasukkan ke sekolah terfavorit, tetapi memasukkan ke sekolah terfavorit kedua di daerah tersebut agar tidak terlihat mencurigakan oleh masyarakat sekitar.",
      c: "Langsung menyetujui permintaan keluarga tersebut dan memberikan jaminan anaknya otomatis diterima pada tahun ajaran berikutnya.",
      d: "Menolak permintaan tersebut karena tidak sesuai dengan prosedur serta merugikan banyak pihak, dan memintanya mengikuti alur pendaftaran yang seharusnya.",
      kunci: "D",
      pembahasan: "Integritas ASN dan kompetensi kepribadian pengawas menuntut penolakan tegas terhadap segala bentuk nepotisme dan penyalahgunaan wewenang. Menolak permintaan keluarga secara terbuka dan mengarahkan mereka mengikuti prosedur SPMB sesuai Permendikdasmen Nomor 3 Tahun 2025 (Opsi D) merupakan pilihan mutlak."
    },
    {
      no: 15, tipe: "SJT", kategori: "Teknis - Sosial", wacana: "",
      pertanyaan: "Salah satu guru di sekolah binaan Anda memiliki masalah dengan masyarakat setempat sehingga terjadi permusuhan antara guru dan masyarakat sekitar, yang membuat suasana sekolah mulai tidak kondusif. Agar keadaan kembali kondusif, tindakan yang Anda lakukan adalah ....",
      a: "Mempertemukan kedua pihak yang bermasalah dan merundingkan apa yang diinginkan kedua pihak agar tercapai perdamaian.",
      b: "Memberikan peringatan tegas kepada guru yang bermasalah serta meminta maaf langsung kepada masyarakat sekitar.",
      c: "Menegur baik guru maupun kepala sekolah bersangkutan atas masalah yang ditimbulkan.",
      d: "Memindahkan guru yang bermasalah ke sekolah binaan lain yang tidak jauh dari sekolah awal namun masyarakatnya belum mengenal guru tersebut.",
      kunci: "A",
      pembahasan: "Kompetensi sosial pengawas mencakup komunikasi efektif serta penguatan kolaborasi antara satuan pendidikan dan masyarakat. Mediasi dialogis yang mempertemukan kedua pihak (Opsi A) menyelesaikan akar konflik secara bermartabat, berbeda dari penghindaran masalah lewat pemindahan sepihak (D)."
    },
    {
      no: 16, tipe: "SJT", kategori: "Teknis - Kepribadian", wacana: "",
      pertanyaan: "Dalam pendampingan transformasi kurikulum dan Pembelajaran Mendalam, Anda melihat bahwa beberapa rekan pengawas sekolah menerapkan cara-cara instan yang tidak sesuai dengan esensi tujuan pembelajaran. Dalam situasi ini, tindakan yang seharusnya Anda lakukan sebagai Pengawas Sekolah adalah ....",
      a: "Menyampaikan apa yang dilihatnya kepada pengawas sekolah tersebut dan memberikan masukan apa yang seharusnya dilakukan dalam mendorong sekolah binaan.",
      b: "Menyampaikan masalah yang dilihatnya kepada pemangku kebijakan agar pemangku kebijakan melakukan penindakan ke depannya.",
      c: "Membuat sesi refleksi dalam forum pengawas sekolah dan memberikan rekomendasi mana yang perlu dihentikan, dilanjutkan, dan dikembangkan dalam pendampingan di sekolah binaan.",
      d: "Mengingatkan pengawas sekolah tersebut bahwa cara-cara instan tersebut kurang relevan dalam implementasi kebijakan kurikulum.",
      kunci: "C",
      pembahasan: "Menggunakan forum refleksi sejawat untuk memetakan praktik pendampingan yang perlu dihentikan, dilanjutkan, dan dikembangkan (Opsi C) menjaga kolektivitas profesi sekaligus meluruskan miskonsepsi secara konstruktif tanpa konfrontasi."
    },
    {
      no: 17, tipe: "SJT", kategori: "Manajerial - Pengembangan Diri dan Orang Lain", wacana: "",
      pertanyaan: "Dalam mendampingi sekolah binaan, Anda mendapatkan sebuah praktik baik yang berpotensi berdampak positif pula apabila diterapkan oleh pengawas sekolah lain di sekolah binaan mereka. Tindakan yang seharusnya Anda lakukan sebagai rekan sejawat agar praktik baik tersebut dapat diimplementasikan lebih luas adalah ....",
      a: "Mengajak pengawas sekolah lain menggunakan konsep pemikiran Anda karena hasilnya sudah terlihat baik di sekolah binaan.",
      b: "Dalam sesi pertemuan antarpengawas di daerah, menyampaikan secara lisan keberhasilan yang sudah dilakukan dalam mendampingi sekolah binaan.",
      c: "Memberitahu bahwa praktik pengawas lain masih perlu direfleksikan sambil menyampaikan praktik baik yang sudah terbukti berhasil.",
      d: "Menuliskan langkah-langkah praktik baik dalam mendampingi sekolah binaan untuk dipublikasikan agar dibaca pengawas lain sebagai sumber inspirasi yang dapat direplikasi dan dimodifikasi di sekolah binaannya.",
      kunci: "D",
      pembahasan: "Mendokumentasikan dan memublikasikan panduan langkah-langkah praktik baik secara tertulis (Opsi D) menciptakan pengetahuan terstruktur yang mudah dipelajari, direplikasi, dan dimodifikasi sesuai konteks sekolah lain."
    },
    {
      no: 18, tipe: "SJT", kategori: "Manajerial - Pengambilan Keputusan", wacana: "",
      pertanyaan: "Sekolah binaan Anda beberapa kali mengalami perselisihan di antara warga sekolah. Pada beberapa sekolah masalah dapat diselesaikan antar-guru, tetapi di sekolah lain persoalan tingkat guru langsung dibawa ke Kepala Sekolah hingga memakan waktu lama dan mengganggu kinerja guru. Solusi tata kelola yang tepat Anda lakukan untuk masalah ini adalah ....",
      a: "Memberikan peringatan bagi guru yang bermasalah dengan harapan memberi efek jera dan contoh bagi guru lain.",
      b: "Menyerahkan semua masalah di bawah Kepala Sekolah kepada Kepala Sekolah sehingga Anda hanya mengurusi masalah pada tingkatan Kepala Sekolah.",
      c: "Membuat sistem penyelesaian masalah yang berjenjang sesuai jenjang jabatan, sehingga eskalasi masalah hanya akan melibatkan satu jenjang jabatan di atas jenjang pihak yang bermasalah.",
      d: "Menangani langsung masalah yang terjadi pada warga sekolah jika masalah sudah terlalu lama.",
      kunci: "C",
      pembahasan: "Tata kelola organisasi yang efektif membangun sistem penyelesaian masalah berjenjang (Opsi C) sehingga persoalan diselesaikan di tingkat terdekat terlebih dahulu, mencegah penumpukan di pimpinan sekolah, dan menjaga kelancaran tugas guru."
    },
    {
      no: 19, tipe: "SJT", kategori: "Teknis - Profesional", wacana: "",
      pertanyaan: "Anda bersama Kepala Sekolah akan memasuki tahun ajaran baru dengan beberapa guru yang baru bergabung pada tahun ajaran tersebut. Agar kompetensi guru dapat berkembang secara terarah, yang Anda lakukan adalah ....",
      a: "Berdiskusi bersama semua warga sekolah untuk menetapkan pengembangan kompetensi yang dibutuhkan guru.",
      b: "Mengimbau semua guru mengikuti pelatihan agar setiap guru memiliki keterampilan minimal yang setara sebelum tahun ajaran dimulai.",
      c: "Merancang sendiri program pengembangan dan pelatihan guru lalu membagikannya kepada Kepala Sekolah untuk dilaksanakan.",
      d: "Mengadakan pertemuan sebelum tahun ajaran baru dimulai untuk merancang perbaikan kompetensi guru berdasarkan evaluasi tahun ajaran sebelumnya.",
      kunci: "D",
      pembahasan: "Perencanaan pengembangan kompetensi guru pada awal tahun ajaran harus berpijak pada data hasil evaluasi kinerja dan pembelajaran tahun sebelumnya (Opsi D) agar program peningkatan mutu tepat sasaran."
    },
    {
      no: 20, tipe: "SJT", kategori: "Manajerial - Mengelola Perubahan", wacana: "",
      pertanyaan: "Dinas Pendidikan melakukan sosialisasi strategi pemanfaatan Ruang GTK (hasil transformasi Platform Merdeka Mengajar dalam Rumah Pendidikan). Namun saat sosialisasi, banyak rincian teknis yang menurut Anda kurang sesuai apabila dijalankan seragam di lapangan. Tindakan yang seharusnya Anda lakukan adalah ....",
      a: "Melakukan terlebih dahulu apa yang sudah dirancang Dinas Pendidikan, dan di lapangan dimodifikasi sesuai kesiapan belajar masing-masing sekolah binaan.",
      b: "Memberikan umpan balik tentang Ruang GTK dan meminta Dinas Pendidikan meninjau ulang serta menunda pelaksanaannya.",
      c: "Mengikuti terlebih dahulu arahan Dinas Pendidikan secara kaku, lalu jika di lapangan tidak sesuai baru direfleksikan di kemudian hari.",
      d: "Menyampaikan keberatan pada seluruh poin yang disampaikan agar Dinas Pendidikan segera mengganti strateginya.",
      kunci: "A",
      pembahasan: "Kepemimpinan adaptif dalam mengelola perubahan menyeimbangkan kepatuhan terhadap arah kebijakan Dinas Pendidikan dengan penyesuaian teknis di lapangan sesuai kesiapan belajar masing-masing sekolah binaan (Opsi A)."
    },
    {
      no: 21, tipe: "SJT", kategori: "Teknis - Profesional", wacana: "",
      pertanyaan: "Salah satu indikator kompetensi dalam Rapor Pendidikan yang perlu dikuasai peserta didik ialah kompetensi literasi dan numerasi. Upaya pendampingan yang paling tepat Anda lakukan untuk mendorong sekolah binaan meningkatkan kompetensi literasi dan numerasi murid adalah ....",
      a: "Menyampaikan target skor literasi dan numerasi peserta didik dan mendiskusikan cara melampaui target tersebut.",
      b: "Menyampaikan target yang perlu dikejar sekolah dalam hal literasi dan numerasi dan mendorong sekolah melampauinya dalam kurun waktu tertentu.",
      c: "Memberikan rekomendasi program literasi dan numerasi siap pakai yang sudah terbukti berhasil di tempat lain untuk dilaksanakan di sekolah binaan.",
      d: "Menyampaikan indikator kompetensi literasi dan numerasi sesuai fase sekolah binaan, dan mengajak guru mengintegrasikan literasi dan numerasi pada setiap mata pelajaran.",
      kunci: "D",
      pembahasan: "Peningkatan literasi dan numerasi bertumpu pada pemahaman indikator capaian fase dan integrasi literasi-numerasi ke seluruh mata pelajaran (Opsi D), bukan tekanan target skor semata (A, B)."
    },
    {
      no: 22, tipe: "SJT", kategori: "Teknis - Profesional", wacana: "",
      pertanyaan: "Kemendikdasmen menerbitkan penguatan kebijakan kurikulum tahun ajaran baru yang harus diikuti satuan pendidikan sesuai ketentuan nasional. Sebagai Pengawas Sekolah, untuk memastikan sekolah binaan Anda menerapkan penyesuaian kurikulum tersebut secara tepat, langkah yang Anda lakukan adalah ....",
      a: "Mengingatkan Kepala Sekolah untuk mengimbau para guru mengenai perubahan kurikulum.",
      b: "Mengumpulkan Kepala Sekolah dan guru untuk memberikan arahan secara langsung, serta melakukan pemantauan langsung pada interval waktu tertentu.",
      c: "Mengimbau secara langsung kepada guru dan meminta Kepala Sekolah mengirimkan laporan kepada Anda setiap pertengahan semester.",
      d: "Menugaskan Kepala Sekolah memberitahukan ke guru dan melaporkannya kepada Anda secara berkala.",
      kunci: "B",
      pembahasan: "Memastikan keberhasilan transisi kurikulum membutuhkan penyamaan persepsi langsung bersama kepala sekolah dan guru yang disertai pemantauan berkala pada interval waktu tertentu (Opsi B), bukan pendelegasian pasif."
    },
    {
      no: 23, tipe: "SJT", kategori: "Teknis - Sosial", wacana: "",
      pertanyaan: "Sekolah binaan Anda mengalami kendala dalam menjalankan program pengembangan sekolah karena belum memiliki jaringan mitra untuk diajak bekerja sama. Dalam situasi tersebut, tindakan pendampingan yang seharusnya Anda lakukan sebagai Pengawas Sekolah adalah ....",
      a: "Menghubungi jaringan mitra yang dibutuhkan sekolah dan meminta sekolah menerima mitra tersebut agar program berjalan baik.",
      b: "Memberikan alternatif jaringan mitra dan mendiskusikan kekurangan serta kelebihan dari masing-masing jaringan mitra tersebut bersama sekolah.",
      c: "Memberikan kontak beberapa jaringan mitra yang dimiliki kepada Kepala Sekolah agar sekolah menjajakinya sendiri.",
      d: "Meminta jaringan mitra yang Anda kenal untuk langsung berkolaborasi dengan sekolah tanpa melibatkan pertimbangan sekolah.",
      kunci: "B",
      pembahasan: "Pengawas berperan memfasilitasi kolaborasi berkualitas dengan menyajikan alternatif mitra sekaligus mendiskusikan kelebihan dan kekurangannya (Opsi B) agar sekolah tetap mandiri mengambil keputusan."
    },
    {
      no: 24, tipe: "SJT", kategori: "Teknis - Profesional", wacana: "",
      pertanyaan: "Tugas pengawas mencakup pemantauan dan evaluasi hasil kerja sekolah binaan terutama dalam meningkatkan kompetensi literasi dan numerasi peserta didik. Tindakan berikut yang paling tepat dilakukan oleh Pengawas Sekolah terkait aktivitas memantau dan mengevaluasi praktik peningkatan kompetensi tersebut adalah ....",
      a: "Mengajak kepala sekolah memiliki alternatif cara untuk mencapai tujuan peningkatan nilai literasi dan numerasi.",
      b: "Mengajak kepala sekolah dan guru di sekolah binaan melakukan refleksi dan mendiskusikan rencana aksi yang akan dilakukan.",
      c: "Melakukan monitoring dan pembinaan administratif kepada kepala sekolah dan guru.",
      d: "Melakukan mentoring searah kepada guru tentang pembelajaran literasi dan numerasi yang kontekstual.",
      kunci: "B",
      pembahasan: "Evaluasi yang memberdayakan selalu melibatkan kepala sekolah dan guru secara aktif untuk merefleksikan praktik pembelajaran yang sudah berjalan dan merumuskan rencana aksi perbaikan bersama (Opsi B)."
    },
    {
      no: 25, tipe: "SJT", kategori: "Teknis - Profesional", wacana: "",
      pertanyaan: "Salah satu prinsip utama dalam Pembelajaran Mendalam adalah pembelajaran yang berkesadaran, bermakna, dan berdiferensiasi sesuai kebutuhan belajar murid. Saat melakukan asesmen awal, ternyata banyak sekolah binaan Anda belum mengetahui praktik pembelajaran berdiferensiasi. Upaya yang Anda lakukan sebagai Pengawas Sekolah ketika melihat masalah tersebut adalah ....",
      a: "Mencarikan narasumber eksternal untuk menjadi pelatih ahli pada setiap sekolah binaan.",
      b: "Mengadakan pelatihan atau workshop tentang praktik pembelajaran berdiferensiasi bagi guru-guru sekolah binaan, yang menjelaskan konsep dan manfaat sekaligus memberikan contoh praktik yang dapat diaplikasikan di kelas.",
      c: "Memetakan praktik baik pembelajaran berdiferensiasi pada sekolah binaan dan menginisiasi sesi saling berbagi antarsekolah binaan.",
      d: "Menjadi narasumber ceramah umum pembelajaran diferensiasi bagi sekolah-sekolah binaan agar memiliki pemahaman yang seragam.",
      kunci: "B",
      pembahasan: "Karena asesmen awal menunjukkan mayoritas sekolah belum memahami praktik diferensiasi sehingga sesi berbagi antarsekolah (C) belum memiliki basis praktik memadai, workshop aplikatif yang memadukan konsep, manfaat, dan contoh kelas (Opsi B) merupakan intervensi paling relevan."
    },
    {
      no: 26, tipe: "SJT", kategori: "Teknis - Sosial", wacana: "",
      pertanyaan: "Daerah di sekitar salah satu sekolah binaan Anda sedang rawan terhadap ancaman penculikan anak, sehingga keselamatan siswa terancam terutama pada jam pulang sekolah. Untuk melindungi siswa di sekolah binaan, tindakan yang Anda lakukan adalah ....",
      a: "Anda dan seluruh warga sekolah mensosialisasikan kepada siswa langkah-langkah pencegahan yang sudah disiapkan dan disetujui oleh seluruh warga sekolah.",
      b: "Mengimbau siswa agar lebih awas dan berhati-hati saat pulang guna menurunkan risiko penculikan.",
      c: "Mengumpulkan warga sekolah dan mendiskusikan tindakan berjaga-jaga tanpa menyepakati protokol tertulis.",
      d: "Meminta partisipasi warga sekolah agar lebih teliti mengawasi siswa selama jam sekolah terutama jam pulang.",
      kunci: "A",
      pembahasan: "Mitigasi risiko keselamatan peserta didik memerlukan penyepakatan protokol pencegahan bersama seluruh warga sekolah yang langsung disosialisasikan kepada siswa agar siap dijalankan segera (Opsi A)."
    },
    {
      no: 27, tipe: "SJT", kategori: "Teknis - Profesional", wacana: "",
      pertanyaan: "Salah satu sekolah binaan Anda berada di daerah pedalaman yang jauh dari perkotaan, sehingga pengetahuan, pola pikir, dan metode pengajaran guru di sekolah tersebut masih tertinggal dan sangat konvensional. Tindakan Anda untuk mengembangkan sekolah ini adalah ....",
      a: "Melakukan kunjungan belajar ke sekolah lain untuk mempelajari metode pengajaran langsung dari sesama guru di sekolah lain.",
      b: "Tidak melakukan apa-apa dan melanjutkan saja praktik yang sudah berjalan.",
      c: "Melaksanakan seminar satu arah untuk membagikan pengetahuan terbaru kepada sekolah binaan.",
      d: "Mengatur jadwal dan menyelenggarakan pendidikan dan latihan rutin teoritis yang diikuti semua guru.",
      kunci: "A",
      pembahasan: "Untuk menggeser pola pikir guru di daerah pedalaman, pembelajaran berbasis pengalaman konkret melalui kunjungan belajar antarsejawat (peer learning) di sekolah lain (Opsi A) jauh lebih meyakinkan dibandingkan seminar satu arah (C)."
    },
    {
      no: 28, tipe: "SJT", kategori: "Sosial Kultural - Perekat Kebangsaan", wacana: "",
      pertanyaan: "Dinas Pendidikan dan Kebudayaan mengadakan perlombaan antarsekolah dengan memilih dekorasi sekolah terbaik per wilayah bertema Bhinneka Tunggal Ika. Sekolah binaan Anda memiliki guru dan warga sekolah dari latar belakang yang berbeda-beda, dan Anda ingin sekolah menampilkan nilai kebangsaan autentik sekaligus berpeluang meraih penghargaan. Tindakan yang Anda ambil adalah ....",
      a: "Berusaha menampilkan semua unsur suku dan agama secara adil pada dekorasi sekolah.",
      b: "Menganjurkan penggunaan tema dekorasi kelompok mayoritas dengan alasan kemudahan dan ketersediaan pernak-pernik dekorasi.",
      c: "Berdiskusi dengan Kepala Sekolah untuk menyerahkan sepenuhnya tema dekorasi yang diinginkan.",
      d: "Memilih dekorasi kelompok minoritas saja untuk menunjukkan kepedulian terhadap kelompok minoritas.",
      kunci: "A",
      pembahasan: "Kompetensi Perekat Kebangsaan diwujudkan secara adil dengan memastikan seluruh unsur suku dan agama terwakili proporsional (Opsi A), menghindari bias mayoritas (B) maupun simbolisasi sepihak (D)."
    },
    {
      no: 29, tipe: "SJT", kategori: "Sosial Kultural - Perekat Kebangsaan", wacana: "",
      pertanyaan: "Selama satu tahun ajaran terdapat beberapa hari peringatan nasional seperti Hari Kemerdekaan Indonesia dan Hari Kartini. Untuk menanamkan nilai kebangsaan pada setiap warga sekolah selain melalui upacara bendera, hal yang akan Anda dorong adalah ....",
      a: "Mengadakan aktivitas kerja sama tim (team building activity) secara berkala dengan tujuan meningkatkan kesadaran akan keberagaman serta toleransi satu sama lain.",
      b: "Mengimbau seluruh warga sekolah menggunakan atribut adat salah satu provinsi tertentu dua kali sebulan pada hari Kamis minggu genap.",
      c: "Pada hari nasional, memberikan ketentuan berbusana daerah yang seragam bagi seluruh warga sekolah.",
      d: "Pada setiap hari kerja yang bertepatan dengan hari nasional, meminta warga sekolah menggunakan atribut penunjuk identitas suku asli masing-masing.",
      kunci: "A",
      pembahasan: "Internalisasi nilai kebangsaan dan toleransi jauh lebih mendalam melalui interaksi kolaboratif (team building) yang menumbuhkan empati lintas latar belakang (Opsi A) dibandingkan simbolisasi atribut fisik (B, C, D)."
    },
    {
      no: 30, tipe: "SJT", kategori: "Teknis - Sosial", wacana: "",
      pertanyaan: "Sekolah-sekolah di bawah binaan Anda semakin sering membuat program yang seragam dan identik dengan identitas sosial tertentu sehingga tampak tidak memberi ruang bagi identitas sosial lainnya, dengan alasan program tersebut lebih diminati masyarakat. Anda mendorong sekolah memiliki program yang lebih plural dengan ....",
      a: "Menyediakan saran berupa daftar program standar sehingga sekolah tinggal memilih menyelenggarakan program baru.",
      b: "Menyarankan sekolah-sekolah binaan terus merancang program baru yang sekadar berbeda dengan sekolah lainnya.",
      c: "Mengajak sekolah menganalisis kebutuhan masyarakat serta minat dan bakat murid demi menghasilkan beragam program sekolah yang inklusif.",
      d: "Mengajak sekolah merefleksikan profil lulusan masing-masing tanpa memetakan kebutuhan masyarakat sekitar.",
      kunci: "C",
      pembahasan: "Mendorong inklusivitas dan keragaman program sekolah paling tepat bertumpu pada analisis objektif kebutuhan seluruh lapisan masyarakat serta minat dan bakat murid (Opsi C)."
    },
    {
      no: 31, tipe: "SJT", kategori: "Sosial Kultural - Perekat Kebangsaan", wacana: "",
      pertanyaan: "Muncul anggapan negatif di masyarakat bahwa semua sekolah binaan Anda hanya bisa dikepalai oleh orang dengan identitas sosial tertentu. Agar isu tersebut diatasi sesuai ketentuan seleksi Kepala Sekolah dalam Permendikdasmen Nomor 7 Tahun 2025, hal yang akan Anda lakukan adalah ....",
      a: "Menanggapinya dengan memberikan dorongan agar segera muncul calon kepala sekolah baru yang berasal dari latar belakang sosial berbeda.",
      b: "Menepis anggapan tersebut dengan menunjukkan kelebihan-kelebihan para kepala sekolah yang sekarang sedang menjabat di sekolah binaan.",
      c: "Membuat sebuah gerakan khusus yang mendukung pendidik dengan identitas sosial berbeda untuk menjadi kepala sekolah.",
      d: "Mengomunikasikan secara transparan kriteria kepala sekolah yang digunakan sesuai regulasi, lalu mengajak pendidik yang memenuhi syarat dari berbagai kalangan mengembangkan karier dan mengecek kualifikasinya di Ruang GTK.",
      kunci: "D",
      pembahasan: "Pengangkatan kepala sekolah dalam Permendikdasmen Nomor 7 Tahun 2025 tunduk pada sistem merit dan kualifikasi objektif (S-1/D-IV, sertifikat pendidik, PNS minimal III/c, pengalaman mengajar 8 tahun, kinerja minimal Baik 2 tahun terakhir) yang dapat dicek melalui Ruang GTK—bukan kuota atau afirmasi identitas (A, C)."
    },
    {
      no: 32, tipe: "SJT", kategori: "Teknis - Sosial", wacana: "",
      pertanyaan: "Sekolah binaan Anda terdiri atas guru yang berasal dari berbagai latar belakang suku, agama, dan ras. Keberagaman ini sudah berlangsung bertahun-tahun tanpa konflik signifikan. Untuk terus meningkatkan kedekatan di antara guru, hal yang Anda lakukan adalah ....",
      a: "Menganjurkan Kepala Sekolah mengadakan program pelatihan khusus setiap semester guna meningkatkan ikatan kerja sama antarguru.",
      b: "Senantiasa menanamkan nilai Bhinneka Tunggal Ika di setiap kesempatan interaksi yang ada.",
      c: "Mendorong Kepala Sekolah melakukan lokakarya formal bertema Bhinneka Tunggal Ika untuk meningkatkan toleransi antarguru.",
      d: "Mendorong Kepala Sekolah menyelenggarakan program khusus bagi kaum minoritas agar mereka merasa dihargai.",
      kunci: "B",
      pembahasan: "Memelihara harmoni sosial yang sudah kondusif paling efektif dilakukan melalui pembiasaan dan penanaman nilai Bhinneka Tunggal Ika secara alami di setiap kesempatan interaksi sehari-hari (Opsi B)."
    },
    {
      no: 33, tipe: "SJT", kategori: "Teknis - Sosial", wacana: "",
      pertanyaan: "Terdapat persaingan antarsekolah yang menempatkan sekolah dengan identitas minoritas menjadi pemenang sehingga muncul anggapan sinis terhadap sekolah tersebut, padahal kompetisi berlangsung secara adil. Tindakan yang Anda lakukan adalah ....",
      a: "Menyelenggarakan kompetisi baru yang dirancang secara khusus hingga mampu menempatkan sekolah dengan identitas mayoritas sebagai pemenang.",
      b: "Mengajak sekolah berefleksi terhadap capaian sekaligus mengupayakan penghapusan kompetisi yang menimbulkan sinisme pada identitas tertentu.",
      c: "Mengajak sekolah berefleksi terhadap capaian hingga memperbanyak ragam program agar sekolah dengan beragam identitas berkesempatan tampil.",
      d: "Menyelenggarakan program khusus terpisah bagi masing-masing identitas sosial agar tampil di ruang masing-masing.",
      kunci: "C",
      pembahasan: "Refleksi capaian yang diikuti perluasan ragam program (Opsi C) membuka ruang aktualisasi setara bagi seluruh sekolah tanpa mencederai keadilan kompetisi (A) ataupun menciptakan segregasi (D)."
    },
    {
      no: 34, tipe: "SJT", kategori: "Manajerial - Mengelola Perubahan", wacana: "",
      pertanyaan: "Para kepala sekolah mulai membanding-bandingkan kualitas antarpengawas sekolah dan menganggap Anda lebih baik daripada pengawas lainnya. Sebagian rekan pengawas tidak nyaman dengan kondisi ini terlebih karena Anda dianggap lebih baik. Tindakan bijaksana yang Anda lakukan adalah ....",
      a: "Memposisikan diri sebagai rekan kepala sekolah untuk ikut menyuarakan kekecewaan pada kinerja pengawas lain.",
      b: "Muncul di sesi pelatihan maupun seminar sebagai narasumber utama agar menjadi teladan pengawas sekolah lain.",
      c: "Mencari beberapa pengawas sekolah yang paling mudah diajak berefleksi agar mereka menularkan praktik baik kepada yang lain.",
      d: "Meminta arahan langsung kepada Kepala Dinas agar segera melakukan penindakan peningkatan kapasitas pengawas lain.",
      kunci: "C",
      pembahasan: "Menggandeng beberapa rekan pengawas yang paling terbuka untuk berefleksi sebagai penggerak awal (early adopters) (Opsi C) menularkan perubahan secara organik tanpa menonjolkan ego pribadi (B)."
    },
    {
      no: 35, tipe: "SJT", kategori: "Manajerial - Kerjasama", wacana: "",
      pertanyaan: "Anda sangat menginginkan terjadinya kolaborasi antarpengawas sekolah agar pendampingan satuan pendidikan di daerah Anda semakin baik. Setiap kali Anda mengajak berkolaborasi, muncul pertanyaan balik, 'Apa yang bisa dikolaborasikan?' Anda menanggapinya dengan ....",
      a: "Menggunakan berkas kerja bersama untuk berbagi proses pendampingan, sehingga mempermudah menemukan kesempatan kolaborasi.",
      b: "Menggunakan grup percakapan bersama untuk saling berinteraksi antarpengawas sekolah agar dapat berkolaborasi.",
      c: "Membuat daftar berbagai program Anda untuk ditawarkan kepada Pengawas Sekolah lainnya sebagai pintu awal kolaborasi.",
      d: "Membuat lomba yang berisi aktivitas kolaborasi antarsekolah sehingga diharapkan pengawas sekolahnya ikut berkolaborasi.",
      kunci: "A",
      pembahasan: "Menjawab pertanyaan konkret tentang ruang kolaborasi paling efektif dilakukan dengan menyediakan berkas kerja bersama yang memetakan proses, tantangan, dan irisan pendampingan tiap pengawas (Opsi A), lebih terarah daripada sekadar grup percakapan umum (B)."
    },
    {
      no: 36, tipe: "SJT", kategori: "Manajerial - Kerjasama", wacana: "",
      pertanyaan: "Semakin banyak sekolah yang membutuhkan narasumber intensif terkait penerapan Pembelajaran Mendalam dan Kurikulum Satuan Pendidikan. Anda berinisiatif mengajak pengawas lain bergantian menjadi narasumber, tetapi pengawas yang sudah menguasai topik terlalu sibuk sedangkan pengawas lainnya merasa belum siap. Jika kebutuhan tersebut tidak segera dipenuhi, semangat belajar sekolah dapat menurun. Tindakan Anda adalah ....",
      a: "Mengirim para kepala sekolah ke daerah lain untuk belajar dengan biaya mandiri.",
      b: "Mengakomodasi kebutuhan banyak sekolah sekaligus dengan memaksakan diri menjadi narasumber tunggal di MKKS/K3S.",
      c: "Melibatkan kepala sekolah yang sudah berhasil menerapkannya sebagai narasumber berbagi praktik baik.",
      d: "Membuat rekaman penjelasan kebijakan untuk dipelajari sendiri oleh sekolah tanpa pendampingan.",
      kunci: "C",
      pembahasan: "Pemberdayaan aset lokal (asset-based thinking) dan kerja sama ekosistem pendidikan dilakukan dengan melibatkan kepala sekolah praktisi yang telah berhasil sebagai narasumber berbagi praktik baik (Opsi C)."
    },
    {
      no: 37, tipe: "SJT", kategori: "Teknis - Kepribadian", wacana: "",
      pertanyaan: "Sering kali arahan pendampingan yang Anda berikan tidak dilaksanakan oleh pihak sekolah, padahal arahan tersebut sudah sesuai dengan pengalaman Anda sebelumnya yang terbukti berhasil. Tindakan reflektif Anda adalah ....",
      a: "Menghentikan penggunaan arahan, mempelajari, kemudian menerapkan cara baru dalam memberi instruksi tertulis.",
      b: "Memberi umpan balik terus-menerus agar sekolah yang belum menerapkan arahan segera melaksanakannya.",
      c: "Merefleksikan pendampingan sekolah sebelumnya dan mencari cara yang lebih sesuai dalam memberi umpan balik.",
      d: "Membuat media seperti kartu kendali kepatuhan pendampingan sekolah untuk terus memonitor perkembangan.",
      kunci: "C",
      pembahasan: "Pengawas dengan pola pikir bertumbuh merefleksikan kembali pendekatan pendampingannya dan menyesuaikan gaya penyampaian umpan balik dengan kesiapan serta konteks spesifik sekolah tersebut (Opsi C)."
    },
    {
      no: 38, tipe: "SJT", kategori: "Teknis - Profesional", wacana: "",
      pertanyaan: "Seorang rekan pengawas sekolah ingin memberikan pendampingan yang lebih baik ke sekolah binaannya. Ia sudah mempelajari program yang berhasil di sekolah lain, tetapi kesulitan mereplikasinya dan khawatir dianggap tidak inovatif apabila sekadar meniru. Upaya dukungan yang Anda lakukan adalah ....",
      a: "Mengajak rekan-rekan pengawas membuat sesi refleksi bersama untuk mengidentifikasi kelebihan dan kelemahan diri masing-masing.",
      b: "Membuat bank data program sekolah yang hanya memuat daftar kegiatan tanpa menjelaskan proses adaptasinya.",
      c: "Memberikan saran kepada rekan pengawas untuk langsung melaksanakan program Anda yang paling mudah diterapkan.",
      d: "Mengajak rekan-rekan pengawas membuat sesi berbagi cerita inovasi program yang berhasil di masing-masing sekolah binaan, berisi pemaparan proses penyesuaian hingga kemajuan yang dicapai sekolah.",
      kunci: "D",
      pembahasan: "Sesi berbagi cerita proses inovasi di tiap sekolah binaan (Opsi D) menunjukkan bagaimana suatu gagasan dimodifikasi secara kreatif sesuai konteks lokal, sehingga menghilangkan kekhawatiran rekan sejawat bahwa adaptasi praktik baik berarti tidak inovatif."
    },
    {
      no: 39, tipe: "SJT", kategori: "Manajerial - Pengembangan Diri dan Orang Lain", wacana: "",
      pertanyaan: "Anda merasa guru-guru dan kepala sekolah di daerah semakin sering mengikuti pelatihan melalui Ruang GTK sehingga dikhawatirkan pengawas akan bingung jika ditanya topik pedagogik terkini. Anda sempat mengajak pengawas lain ikut belajar bersama guru, namun mereka menolak. Tindakan Anda berikutnya adalah ....",
      a: "Menunjukkan poster-poster pelatihan guru kepada para Pengawas Sekolah agar merasa tertinggal dan mau ikut.",
      b: "Mengajak organisasi profesi Pengawas Sekolah untuk menyelenggarakan sesi-sesi peningkatan kompetensi khusus bagi pengawas.",
      c: "Memberi nasihat kepada Pengawas Sekolah lain bahwa mereka pasti ketinggalan jika tidak mau belajar bersama guru.",
      d: "Menyampaikan keluhan kepada Kepala Dinas agar mewajibkan seluruh pengawas mengikuti pelatihan bersama guru.",
      kunci: "B",
      pembahasan: "Menghargai kenyamanan psikologis rekan sejawat sambil tetap mendorong peningkatan kompetensi profesional dilakukan dengan menggerakkan organisasi profesi Pengawas Sekolah menyelenggarakan sesi peningkatan kompetensi (Opsi B)."
    },
    {
      no: 40, tipe: "SJT", kategori: "Teknis - Profesional", wacana: "",
      pertanyaan: "Sekolah di bawah binaan Anda sudah sering menyelenggarakan pelatihan bagi guru berdasarkan rekomendasi Anda, namun Kepala Sekolah merasa belum ada perkembangan berarti pada kualitas SDM guru di kelas. Tindakan pendampingan yang Anda ambil adalah ....",
      a: "Mencari pelatihan dengan topik dan pelatih berbeda secara terus-menerus hingga menemukan yang berdampak.",
      b: "Membuat refleksi rutin terkait kesulitan yang dialami guru pada proses pembelajaran untuk mengidentifikasi kebutuhan pelatihan yang tepat.",
      c: "Memberi saran kepada kepala sekolah untuk tetap mengikuti rekomendasi pelatihan sebelumnya dengan mengganti narasumbernya saja.",
      d: "Menggali alasan administratif dari kepala sekolah tanpa mengamati kesulitan guru di ruang kelas.",
      kunci: "B",
      pembahasan: "Pelatihan gagal berdampak apabila tidak berangkat dari kesulitan nyata guru di kelas. Membangun refleksi rutin atas tantangan pembelajaran sebagai dasar analisis kebutuhan pelatihan (Opsi B) menjamin program pengembangan SDM tepat sasaran."
    },
    {
      no: 41, tipe: "SJT", kategori: "Manajerial - Komunikasi", wacana: "",
      pertanyaan: "Salah satu sekolah binaan Anda sudah jauh lebih maju menerapkan pembelajaran berdiferensiasi. Ketika Anda mengajak sekolah binaan lain turut menerapkannya, mereka beranggapan hal itu mustahil dilakukan karena jumlah murid jauh lebih banyak dengan latar belakang sosial dan persoalan yang lebih kompleks. Cara komunikasi persuasif Anda adalah ....",
      a: "Mensosialisasikan teori dan dampak positif jangka panjang setelah murid mengalami diferensiasi pembelajaran.",
      b: "Membangun proses diskusi antarsekolah terkait permasalahan nyata murid yang berhasil teratasi berkat diferensiasi pembelajaran.",
      c: "Menyelenggarakan lomba publikasi video penerapan diferensiasi pembelajaran berhadiah menarik.",
      d: "Memberi penghargaan khusus kepada sekolah yang lebih dulu menerapkan diferensiasi agar sekolah lain termotivasi.",
      kunci: "B",
      pembahasan: "Hambatan mental ('tidak mungkin diterapkan di sekolah kami') paling efektif diatasi melalui dialog antarsekolah yang membedah bukti nyata pemecahan masalah murid lewat pembelajaran berdiferensiasi (Opsi B)."
    },
    {
      no: 42, tipe: "SJT", kategori: "Manajerial - Orientasi Pada Hasil", wacana: "",
      pertanyaan: "Permasalahan kualitas pembelajaran terjadi bersamaan di beberapa sekolah binaan, dan masing-masing kepala sekolah merasa membutuhkan pendampingan khusus untuk mengatasinya. Pilihan strategi pendampingan yang paling efektif dan efisien adalah ....",
      a: "Membagi waktu ke satu sekolah terlebih dahulu hingga tuntas sebelum berpindah menyelesaikan masalah sekolah lainnya.",
      b: "Mencari permasalahan administratif paling ringan yang dimiliki sekolah untuk diselesaikan lebih dahulu.",
      c: "Memilih satu sekolah yang paling berpengaruh saja agar hasilnya cepat terlihat di daerah.",
      d: "Mengidentifikasi kesamaan dari permasalahan dan kebutuhan beberapa sekolah tersebut agar perbaikan dapat dirancang dan dilakukan sekaligus secara kolektif.",
      kunci: "D",
      pembahasan: "Memetakan irisan kesamaan permasalahan dan kebutuhan beberapa sekolah binaan untuk merancang intervensi kolektif secara serentak (Opsi D) memaksimalkan efektivitas dan efisiensi capaian mutu."
    },
    {
      no: 43, tipe: "SJT", kategori: "Manajerial - Komunikasi", wacana: "",
      pertanyaan: "Anda sudah sering melakukan pendampingan dan menghasilkan perkembangan nyata pada sekolah binaan, tetapi perkembangan tersebut belum dirasakan atau terlihat oleh rekan kerja pengawas lainnya. Anda pun mengurungkan niat meminta umpan balik karena khawatir dianggap tidak memiliki capaian berarti. Cara yang sebaiknya dilakukan adalah ....",
      a: "Mengirimkan berkas catatan pribadi pendampingan sekolah ke seluruh rekan kerja agar mereka memberi penilaian.",
      b: "Meminta umpan balik pada rekan kerja secara berkala dalam setiap rapat tanpa bukti nyata capaian sekolah.",
      c: "Mendorong sekolah binaan menyelenggarakan pameran karya sebagai bukti nyata pendampingan sekaligus media dialog untuk meminta masukan.",
      d: "Menyediakan kotak saran atau formulir umpan balik anonim mengenai diri Anda kepada rekan kerja.",
      kunci: "C",
      pembahasan: "Pameran karya sekolah binaan (Opsi C) menampilkan bukti autentik dampak pendampingan di ruang publik tanpa kesan menyombongkan diri, sekaligus membuka ruang dialog konstruktif untuk memperoleh umpan balik sejawat."
    },
    {
      no: 44, tipe: "SJT", kategori: "Teknis - Kepribadian", wacana: "",
      pertanyaan: "Seorang rekan Pengawas Sekolah mudah tersinggung saat diajak menerapkan inovasi program pendampingan yang sudah Anda lakukan, karena menganggap Anda ingin tampil sebagai pengawas yang ditiru oleh pengawas lain. Tindakan yang Anda lakukan adalah ....",
      a: "Mengomunikasikan secara langsung kepada pengawas sekolah tersebut bahwa yang Anda lakukan bukan untuk mencari panggung.",
      b: "Menjaga jarak dan berupaya tidak bekerja bersama pengawas sekolah tersebut untuk menghindari konflik.",
      c: "Meminta masukan kepada rekan kerja terkait sikap dan komunikasi Anda, lalu berupaya mengenal pengawas sekolah tersebut secara personal.",
      d: "Mengomunikasikan permintaan maaf secara formal sambil menegaskan tujuan Anda sebenarnya di forum dinas.",
      kunci: "C",
      pembahasan: "Menyatakan secara langsung bahwa Anda 'bukan mencari panggung' (A) bernada defensif dan berisiko memperlebar jarak emosional; meminta masukan rekan kerja atas cara komunikasi diri lalu membangun kedekatan personal (Opsi C) menunjukkan kematangan emosional dan kerendahan hati sesuai Pasal 16 ayat (1) Permendikdasmen Nomor 23 Tahun 2026."
    },
    {
      no: 45, tipe: "SJT", kategori: "Teknis - Kepribadian", wacana: "",
      pertanyaan: "Dalam pendampingan kebijakan kurikulum terbaru, beberapa Pengawas Sekolah saling bersaing memamerkan sekolah binaannya yang langsung mengadopsi skema lanjutan tanpa bertahap, bahkan memaksa sekolah binaan mengambil pilihan tersebut meski belum siap. Padahal tujuan pilihan bertahap adalah menyesuaikan dengan kondisi dan kesiapan masing-masing sekolah. Terhadap fenomena ini tindakan Anda adalah ....",
      a: "Menganggap kompetisi antarpengawas sekolah ini sebagai hal baik agar adopsi kebijakan cepat merata di banyak sekolah.",
      b: "Mengajak para pengawas mempelajari ulang dokumen kebijakan lalu menjelaskan bahwa sekolah tidak boleh dipaksa.",
      c: "Mengajak para pengawas menghindari persaingan melalui lomba kolaborasi antarsekolah.",
      d: "Menginisiasi forum yang membahas fungsi serta kode etik kepengawasan dan manajemen diri pemimpin, lalu mengajak diskusi tujuan dan proses implementasi bertahap sesuai kesiapan sekolah.",
      kunci: "D",
      pembahasan: "Akar masalah pemaksaan target akibat gengsi antarpengawas adalah integritas peran dan pengendalian ego pemimpin. Membuka forum yang menyentuh kode etik kepengawasan dan manajemen diri pemimpin terlebih dahulu sebelum meluruskan esensi implementasi bertahap (Opsi D) menyasar akar persoalan secara menyeluruh."
    },
    {
      no: 46, tipe: "Studi Kasus", kategori: "Teknis - Profesional", wacana: kasus1,
      pertanyaan: "Melihat praktik pemanfaatan Kecerdasan Artifisial (KA) oleh guru dan murid yang belum disertai literasi digital maupun proses berpikir mendalam, tindakan pembinaan akademik yang paling tepat Anda lakukan bersama kepala sekolah adalah ....",
      a: "Melarang seluruh guru dan murid menggunakan aplikasi KA selama satu semester penuh agar kemampuan menulis manual kembali pulih.",
      b: "Memfasilitasi pendampingan penyusunan desain Pembelajaran Mendalam yang memadukan aktivitas eksplorasi nyata dengan etika dan literasi digital KA, sehingga KA berfungsi sebagai alat bantu berpikir kritis bukan pengganti proses belajar.",
      c: "Menyerahkan daftar modul ajar siap pakai kepada guru agar tidak perlu lagi menggunakan KA dalam merancang pembelajaran.",
      d: "Meminta kepala sekolah memberikan teguran tertulis kepada guru yang terdeteksi menggunakan KA dalam menyusun perangkat ajar.",
      kunci: "B",
      pembahasan: "Kebijakan Pembelajaran Mendalam dan Koding/KA dalam Permendikdasmen Nomor 13 Tahun 2025 serta rekomendasi kajian kebijakan BSKAP menekankan penguatan literasi digital guru dan murid serta pengombinasian metode pembelajaran agar tidak bergantung semata pada KA. Memadukan eksplorasi nyata dan literasi etis KA (Opsi B) menjawab kebutuhan kompetensi abad ke-21 tanpa mematikan inovasi teknologi. (Catatan coaching Ahli Utama: ajukan pertanyaan reflektif kepada guru mengenai bukti berpikir mandiri murid di balik tugas yang dikumpulkan)."
    },
    {
      no: 47, tipe: "Studi Kasus", kategori: "Teknis - Profesional", wacana: kasus1,
      pertanyaan: "Menanggapi penerapan Gerakan 7 Kebiasaan Anak Indonesia Hebat (7KAIH) dan kegiatan Kokurikuler yang berubah menjadi mekanisme hukuman poin di sekolah binaan, rekomendasi perbaikan yang paling selaras dengan prinsip kebijakan Kemendikdasmen adalah ....",
      a: "Menghapus seluruh kegiatan kokurikuler 7KAIH karena dinilai menyulitkan guru dalam memberikan penilaian objektif.",
      b: "Mengganti sanksi poin dengan kewajiban merangkum buku bagi murid yang belum konsisten menjalankan pembiasaan harian.",
      c: "Membimbing sekolah merancang ulang Kokurikuler Lintas Disiplin Ilmu dan 7KAIH berbasis prinsip memuliakan, berkesadaran, dan bermakna dengan pendekatan motivasional guna memperkuat 8 Dimensi Profil Lulusan.",
      d: "Meminta guru hanya menilai murid yang secara sukarela mengumpulkan jurnal pembiasaan tanpa perlu mendampingi murid lainnya.",
      kunci: "C",
      pembahasan: "Panduan kokurikuler Kemendikdasmen menetapkan tiga bentuk utama pelaksanaan (pembelajaran kolaboratif lintas disiplin ilmu, 7KAIH, dan cara lainnya) dengan pendekatan Pembelajaran Mendalam berprinsip memuliakan (tanpa hukuman bagi yang gagal, melainkan dorongan motivasi), berkesadaran, dan bermakna untuk memperkuat 8 Dimensi Profil Lulusan. Mengembalikan orientasi pembiasaan ke pendekatan motivasional (Opsi C) memastikan internalisasi karakter berjalan autentik."
    },
    {
      no: 48, tipe: "Studi Kasus", kategori: "Teknis - Kepribadian", wacana: kasus1,
      pertanyaan: "Dalam pertemuan refleksi bersama guru, dua guru senior menyampaikan penolakan emosional terhadap pembelajaran Koding dan KA serta perubahan metode mengajar di hadapan guru lain. Sikap dan respons Anda yang mencerminkan kematangan emosional serta pola pikir bertumbuh sesuai Pasal 16 ayat (1) Permendikdasmen Nomor 23 Tahun 2026 adalah ....",
      a: "Mendengarkan kekhawatiran mereka dengan empati tanpa menyela, mengakui pengalaman mengajar mereka, lalu mengajak memetakan langkah adaptasi teknologi yang paling sederhana dan relevan bagi mata pelajaran mereka.",
      b: "Menegaskan secara tegas di depan forum bahwa aturan Permendikdasmen bersifat mengikat sehingga guru senior wajib mengikuti tanpa alasan usia.",
      c: "Mengabaikan keluhan guru senior tersebut dan memilih hanya membina guru-guru muda yang lebih melek teknologi digital.",
      d: "Menyarankan kepala sekolah memindahkan kedua guru senior tersebut menjadi tenaga administrasi agar tidak menghambat inovasi sekolah.",
      kunci: "A",
      pembahasan: "Kompetensi kepribadian pengawas menuntut kematangan spiritual, moral, dan emosional serta pola pikir bertumbuh. Merespons penolakan secara empatik sambil tetap memandu langkah adaptasi bertahap yang realistis (Opsi A) menjaga martabat guru senior dan menurunkan resistensi terhadap perubahan."
    },
    {
      no: 49, tipe: "Studi Kasus", kategori: "Teknis - Sosial", wacana: kasus1,
      pertanyaan: "Menyikapi keluhan sebagian orang tua yang merasa terbebani oleh tugas proyek kokurikuler di rumah dan meragukan relevansinya dengan capaian akademik anak, langkah komunikasi dan kolaborasi yang Anda fasilitasi bagi kepala sekolah dan guru adalah ....",
      a: "Meminta sekolah menghentikan semua proyek rumah dan menggantinya dengan latihan soal pilihan ganda di kelas agar orang tua puas.",
      b: "Mengeluarkan surat edaran wajib bagi orang tua untuk menandatangani surat dukungan terhadap seluruh kebijakan kokurikuler sekolah.",
      c: "Memfasilitasi forum komunikasi sekolah dan komite atau orang tua untuk menjelaskan keterkaitan 8 Dimensi Profil Lulusan dengan masa depan murid, menyederhanakan peran pendampingan orang tua di rumah, dan memamerkan proses belajar murid secara terbuka.",
      d: "Menyerahkan urusan penjelasan sepenuhnya kepada wali kelas masing-masing melalui pesan singkat grup kelas.",
      kunci: "C",
      pembahasan: "Pasal 16 ayat (2) Permendikdasmen Nomor 23 Tahun 2026 mengarahkan pengawas memfasilitasi kolaborasi berkualitas dengan lingkungan masyarakat dan orang tua. Dialog terbuka yang meluruskan esensi 8 Dimensi Profil Lulusan serta menyederhanakan porsi tugas rumah (Opsi C) membangun kemitraan sekolah dan keluarga secara sehat."
    },
    {
      no: 50, tipe: "Studi Kasus", kategori: "Teknis - Profesional", wacana: kasus2,
      pertanyaan: "Sebagai pengawas yang memahami pembaruan Rapor Pendidikan 2026 (Kepmendikdasmen Nomor 63/2026 dan SE Sesjen Nomor 6/2026), penjelasan teknis dan langkah tindak lanjut yang paling akurat Anda berikan kepada SD dan satuan PAUD tersebut adalah ....",
      a: "Menyarankan SD mengajukan revisi skor Rapor Pendidikan kepada Dinas Pendidikan berdasarkan hasil ujian internal awal 2026, dan meminta satuan PAUD menyalin profil rapor satuan terdekat.",
      b: "Menjelaskan bahwa penamaan Rapor Pendidikan mengikuti tahun data (data 2025 dirilis 2026) sehingga perbaikan awal 2026 dapat dianalisis sebagai data internal pendamping tanpa mengubah skor resmi, serta membina satuan PAUD agar memenuhi minimal 60% pengisian Sulingjar pada periode berikutnya sembari memanfaatkan data Dapodik dan evaluasi mandiri sekolah/guru/orang tua untuk PBD saat ini.",
      c: "Meminta kedua satuan pendidikan menunda penyusunan RKT dan ARKAS hingga Rapor Pendidikan tahun berikutnya terbit agar datanya mutakhir.",
      d: "Menyatakan bahwa Rapor Pendidikan tidak perlu dipakai lagi sehingga sekolah cukup menyusun program berdasarkan usulan rapat guru saja.",
      kunci: "B",
      pembahasan: "Pembaruan Maret 2026 menetapkan penamaan Rapor Pendidikan mengikuti tahun data (Tahun Data 2025), di mana sekolah boleh melengkapi analisis dengan asesmen internal dan observasi terbaru tanpa mengubah skor resmi secara sepihak. Bagi satuan PAUD yang tingkat pengisian Sulingjar-nya di bawah minimal 60% sehingga label warna dan skor tidak tampil, panduan resmi Kemendikdasmen mengarahkan PBD tetap disusun menggunakan data Dapodik serta pengumpulan data mandiri dari murid, guru, dan orang tua sembari membenahi partisipasi Sulingjar (Opsi B)."
    },
    {
      no: 51, tipe: "Studi Kasus", kategori: "Teknis - Profesional", wacana: kasus2,
      pertanyaan: "Menyikapi dua SMP yang menyalin seluruh daftar contoh kegiatan ARKAS tanpa menganalisis akar masalah, pendampingan Perencanaan Berbasis Data (PBD) yang paling tepat Anda fasilitasi bersama tim pengembang sekolah adalah ....",
      a: "Membiarkan draf ARKAS tersebut disahkan agar serapan anggaran BOSP Reguler mencapai 100% tepat waktu.",
      b: "Membimbing tim manajemen sekolah (kepala sekolah, wakil kurikulum, guru, bendahara, operator) menelusuri siklus Identifikasi-Refleksi-Benahi (IRB) dari akar masalah level 1 dan level 2, memilih inspirasi benahi yang berdampak langsung pada kualitas pembelajaran, baru memetakan kebutuhan riil ke lembar kerja ARKAS.",
      c: "Mengambil alih penyusunan RKT dan lembar kerja ARKAS kedua SMP tersebut agar tidak terjadi kesalahan penginputan kode kegiatan.",
      d: "Meminta sekolah menghapus seluruh kegiatan peningkatan kompetensi guru di ARKAS dan memfokuskan anggaran hanya pada pembelian sarana fisik.",
      kunci: "B",
      pembahasan: "Esensi PBD adalah alur logis Identifikasi masalah prioritas, Refleksi akar masalah (termasuk indikator level 1 dan level 2), pemilihan inspirasi Benahi di RKT, serta penganggaran yang relevan di lembar kerja ARKAS (Opsi B). Mengambil alih penyusunan (C) menghambat kemandirian tata kelola sekolah."
    },
    {
      no: 52, tipe: "Studi Kasus", kategori: "Manajerial - Orientasi Pada Hasil", wacana: kasus2,
      pertanyaan: "Menghadapi rencana kepala sekolah SD binaan yang ingin mewajibkan bimbingan tes TKA berbayar demi mengejar hasil sertifikat TKA tinggi bagi jalur prestasi SPMB, langkah pengendalian mutu yang berorientasi pada hasil autentik sesuai Permendikdasmen Nomor 9 Tahun 2025 adalah ....",
      a: "Mendukung bimbingan tes berbayar tersebut asalkan pihak komite sekolah menandatangani persetujuan tertulis.",
      b: "Mengarahkan kepala sekolah mengganti bimbingan berbayar dengan pembenahan pembelajaran literasi (Bahasa Indonesia) dan numerasi (Matematika) reguler yang bermakna serta asesmen diagnostik gratis di jam sekolah guna memetakan kesiapan murid secara objektif.",
      c: "Meminta sekolah tetap mewajibkan les tambahan setiap sore bagi seluruh murid kelas 6 namun digratiskan dengan memotong dana kegiatan kokurikuler.",
      d: "Mengimbau guru memberikan latihan soal TKA secara berulang-ulang setiap hari menggantikan jam pelajaran lainnya.",
      kunci: "B",
      pembahasan: "Permendikdasmen Nomor 9 Tahun 2025 menetapkan tujuan TKA untuk mengukur capaian akademik terstandar, meningkatkan kapasitas pendidik dalam penilaian, serta menjadi acuan pengendalian dan penjaminan mutu pendidikan. Membebani orang tua dengan bimbingan berbayar (A) atau mengorbankan mata pelajaran lain (C, D) merusak orientasi mutu; penguatan pembelajaran reguler Bahasa Indonesia dan Matematika sebagai mata uji SD (Opsi B) menghasilkan capaian belajar yang autentik."
    },
    {
      no: 53, tipe: "Studi Kasus", kategori: "Teknis - Kepribadian", wacana: kasus2,
      pertanyaan: "Ketika mengetahui adanya pertimbangan kepala sekolah untuk tidak mengikutsertakan sebagian murid berkebutuhan khusus (non-disabilitas intelektual) dalam TKA demi mendongkrak rata-rata nilai sekolah, sikap tegas berbasis integritas dan regulasi yang Anda ambil adalah ....",
      a: "Memaklumi pertimbangan tersebut karena sekolah perlu menjaga citra daya saing di mata masyarakat sekitar.",
      b: "Mengingatkan kepala sekolah secara tegas dan persuasif bahwa pengecualian peserta TKA menurut Permendikdasmen Nomor 9 Tahun 2025 hanya berlaku bagi murid berkebutuhan khusus penyandang disabilitas intelektual, sehingga seluruh murid lainnya berhak mengikuti TKA dan memperoleh sertifikat hasil belajar.",
      c: "Menyarankan kepala sekolah menyurati orang tua murid tersebut agar menandatangani surat pengunduran diri sukarela dari TKA.",
      d: "Tidak mencampuri daftar peserta TKA sekolah karena pendaftaran peserta merupakan kewenangan penuh operator Dapodik sekolah.",
      kunci: "B",
      pembahasan: "Permendikdasmen Nomor 9 Tahun 2025 secara jelas menjamin akses murid formal, nonformal, dan informal serta hanya mengecualikan penyandang disabilitas intelektual. Menjaga hak seluruh peserta didik lainnya untuk memperoleh asesmen terstandar dan sertifikat TKA (Opsi B) menegakkan integritas dan keadilan sesuai Pasal 16 ayat (1) Permendikdasmen Nomor 23 Tahun 2026."
    },
    {
      no: 54, tipe: "Studi Kasus", kategori: "Teknis - Profesional", wacana: kasus3,
      pertanyaan: "Menghadapi kelangkaan bakal calon kepala sekolah akibat kekeliruan pemahaman guru mengenai syarat Sertifikat Guru Penggerak, tindakan pembinaan manajerial yang paling tepat Anda lakukan sesuai Permendikdasmen Nomor 7 Tahun 2025 adalah ....",
      a: "Meminta Dinas Pendidikan menunjuk langsung guru senior tanpa memperhatikan syarat pangkat atau pengalaman mengajar agar tidak terjadi kekosongan jabatan.",
      b: "Mensosialisasikan kepada guru-guru potensial bahwa Permendikdasmen Nomor 7 Tahun 2025 membuka peluang seleksi bagi guru berkualifikasi S-1/D-IV, bersertifikat pendidik, PNS minimal III/c, pengalaman mengajar minimal 8 tahun, dan kinerja minimal Baik 2 tahun terakhir tanpa syarat wajib sertifikat Guru Penggerak, serta membimbing mereka memakai fitur Cek Kualifikasi di Ruang GTK.",
      c: "Menunggu hingga kedua kepala sekolah pensiun baru meminta Dinas menunjuk pelaksana tugas (Plt) dari luar kecamatan.",
      d: "Menganjurkan guru menunggu dibukanya kembali program Guru Penggerak sebelum memikirkan karier sebagai kepala sekolah.",
      kunci: "B",
      pembahasan: "Permendikdasmen Nomor 7 Tahun 2025 mencabut Permendikbudristek Nomor 40 Tahun 2021 dan Nomor 26 Tahun 2022 tentang Pendidikan Guru Penggerak. Guru potensial yang memenuhi syarat kualifikasi dapat mengecek kelayakan melalui menu Seleksi Kepala Sekolah (tombol Cek Kualifikasi) di Ruang GTK tanpa harus memiliki sertifikat Guru Penggerak (Opsi B). Penyiapan kader sejak enam bulan sebelum masa pensiun juga mencegah kekosongan jabatan."
    },
    {
      no: 55, tipe: "Studi Kasus", kategori: "Teknis - Profesional", wacana: kasus3,
      pertanyaan: "Dalam menilai dan membina kinerja kepala sekolah yang sedang menjabat agar selaras dengan standar kompetensi terbaru pada Permendikdasmen Nomor 23 Tahun 2026, fokus utama instrumen pembinaan Anda diarahkan pada kemampuan kepala sekolah untuk ....",
      a: "Menguasai seluruh pengisian aplikasi administrasi keuangan secara pribadi tanpa melibatkan bendahara maupun guru.",
      b: "Mengembangkan visi budaya mutu, memimpin kurikulum dan pembelajaran berpusat pada murid, melaksanakan supervisi dan pendampingan berbasis data, mengelola sumber daya secara akuntabel, membangun kemitraan orang tua/masyarakat, serta mendorong inovasi dan kepemimpinan transformasional.",
      c: "Menjaga ketertiban upacara dan kedisiplinan jam masuk sekolah sebagai indikator tunggal keberhasilan manajemen sekolah.",
      d: "Menyusun laporan pertanggungjawaban seremonial kepada pengawas sekolah setiap akhir bulan.",
      kunci: "B",
      pembahasan: "Permendikdasmen Nomor 23 Tahun 2026 merumuskan kompetensi profesional Kepala Satuan Pendidikan mencakup pengembangan visi-budaya belajar, lingkungan inklusif-aman, kepemimpinan kurikulum berpusat pada murid, supervisi/evaluasi/refleksi berbasis data, pengelolaan sumber daya secara transparan dan akuntabel, kerja sama orang tua/masyarakat, serta inovasi, entrepreneurship, dan kepemimpinan transformasional (Opsi B)."
    },
    {
      no: 56, tipe: "Studi Kasus", kategori: "Teknis - Profesional", wacana: kasus3,
      pertanyaan: "Menyikapi kondisi tenaga administrasi dan operator data di dua sekolah binaan yang berijazah pendidikan menengah tanpa sertifikat kompetensi serta mengalami kendala sinkronisasi Dapodik dengan Ruang GTK, langkah pemantauan dan pembinaan yang sesuai dengan Permendikdasmen Nomor 23 Tahun 2026 adalah ....",
      a: "Merekomendasikan pemberhentian segera tenaga administrasi dan operator data tersebut karena melanggar standar kualifikasi Pasal 14.",
      b: "Membiarkan saja ketidaksesuaian data Dapodik karena hal itu merupakan tanggung jawab penuh pemerintah pusat.",
      c: "Membina kepala sekolah untuk memfasilitasi pemenuhan sertifikat kompetensi dan peningkatan keterampilan pengelolaan data serta pemanfaatan TIK bagi tenaga administrasi dan operator dalam masa transisi 5 tahun sesuai Pasal 24, sekaligus menjadwalkan audit rutin pemutakhiran dan sinkronisasi Dapodik-Ruang GTK.",
      d: "Menugaskan guru mata pelajaran mengambil alih seluruh pekerjaan operator data agar sinkronisasi Dapodik lebih cepat.",
      kunci: "C",
      pembahasan: "Permendikdasmen Nomor 23 Tahun 2026 Pasal 14 menetapkan kualifikasi minimal lulusan pendidikan menengah disertai Sertifikat Kompetensi bagi tenaga administrasi dan operator data, sedangkan ketentuan peralihan Pasal 24 memberi batas waktu pemenuhan kualifikasi paling lambat 5 tahun sejak peraturan diundangkan bagi yang telah diangkat sebelumnya. Pembinaan kapasitas teknis pengelolaan data secara berkala dalam masa transisi tersebut (Opsi C) tepat secara hukum dan operasional."
    },
    {
      no: 57, tipe: "Studi Kasus", kategori: "Teknis - Kepribadian", wacana: kasus3,
      pertanyaan: "Ketika dimintai rekomendasi oleh pejabat daerah mengenai usulan rotasi cepat kepala sekolah berprestasi yang baru bertugas 15 bulan di sekolah pinggiran menuju sekolah favorit di kota, sikap berintegritas yang wajib Anda pegang sesuai Permendikdasmen Nomor 7 Tahun 2025 adalah ....",
      a: "Menyetujui usulan rotasi tersebut demi menjaga hubungan baik dengan pejabat daerah yang menjadi atasan langsung.",
      b: "Menyampaikan pertimbangan objektif berbasis regulasi bahwa Permendikdasmen Nomor 7 Tahun 2025 melarang rotasi kepala sekolah sebelum genap dua tahun memimpin pada satuan pendidikan yang sama, agar keberlanjutan transformasi sekolah pinggiran tetap terjaga.",
      c: "Memberikan rekomendasi netral tanpa menyinggung batas waktu dua tahun agar keputusan diserahkan kepada preferensi pimpinan daerah.",
      d: "Menyarankan kepala sekolah tersebut mengundurkan diri sementara lalu diangkat kembali di sekolah kota agar terhindar dari aturan rotasi.",
      kunci: "B",
      pembahasan: "Permendikdasmen Nomor 7 Tahun 2025 menetapkan ketentuan bahwa rotasi kepala sekolah tidak dilakukan sebelum masa penugasan dua tahun terlampaui agar kepala sekolah dapat menjalankan tanggung jawab transformasi secara maksimal. Ketegasan moral menyampaikan dasar hukum secara objektif (Opsi B) merupakan wujud integritas Pasal 16 ayat (1) Permendikdasmen Nomor 23 Tahun 2026. (Catatan wawancara Ahli Utama: keteguhan menjaga aturan di hadapan tekanan eksternal menjadi indikator kunci integritas strategis)."
    },
    {
      no: 58, tipe: "Studi Kasus", kategori: "Teknis - Profesional", wacana: kasus4,
      pertanyaan: "Sebagai bagian dari penataan sistem pengawasan mutu berbasis PermenPANRB Nomor 7 Tahun 2026 dan kondisi geografis daerah yang menantang, rancangan pembagian peran berjenjang yang paling tepat Anda usulkan dalam tim kerja pengawas adalah ....",
      a: "Membagi tugas secara acak berdasarkan kedekatan domisili pengawas saja tanpa membedakan jenjang jabatan fungsional.",
      b: "Menugaskan pengawas Ahli Utama hanya mengurus pekerjaan administratif kantor dinas agar tidak perlu menjangkau sekolah di wilayah sulit.",
      c: "Menyelaraskan pembagian kerja sesuai ruang lingkup penjenjangan: Ahli Muda berfokus pada analisis mutu satuan pendidikan, Ahli Madya pada pengendalian dan pengembangan mutu, sedangkan Ahli Utama memimpin kegiatan strategis dan inovasi sistem pengawasan (seperti desain pendampingan hibrida bagi wilayah sulit) sambil memperhitungkan indikator jumlah, bentuk sekolah, dan kondisi geografis.",
      d: "Menyeragamkan seluruh target kerja dan instrumen laporan Ahli Muda hingga Ahli Utama agar mudah dinilai oleh atasan.",
      kunci: "C",
      pembahasan: "Ringkasan PermenPANRB Nomor 7 Tahun 2026 membedakan ruang lingkup tugas tiap jenjang: Ahli Muda melakukan analisis mutu, Ahli Madya melakukan pengendalian dan pengembangan mutu, serta Ahli Utama melaksanakan kegiatan strategis dan inovasi pengembangan sistem pengawasan manajerial dan akademik pada satuan pendidikan formal, sementara penetapan kebutuhan pengawas memperhitungkan jumlah satuan pendidikan, bentuk satuan pendidikan, dan kondisi geografis (Opsi C)."
    },
    {
      no: 59, tipe: "Studi Kasus", kategori: "Teknis - Profesional", wacana: kasus4,
      pertanyaan: "Menyikapi kebingungan guru dan kepala sekolah dalam membedakan ekosistem digital Kemendikdasmen, penjelasan teknis dan pembinaan literasi digital yang paling akurat Anda berikan di sekolah binaan adalah ....",
      a: "Meminta guru mengunduh ulang aplikasi Platform Merdeka Mengajar (PMM) versi lama karena fiturnya dianggap lebih lengkap daripada Ruang GTK.",
      b: "Menjelaskan bahwa Ruang GTK merupakan transformasi PMM di dalam super-app Rumah Pendidikan (login menggunakan akun belajar.id) untuk belajar berkelanjutan, inspirasi pembelajaran, refleksi kompetensi, dan pengelolaan kinerja, sedangkan verifikasi data individu dan status tunjangan (SKTP/TPG) tetap dicek melalui Info GTK.",
      c: "Menyampaikan bahwa seluruh proses verifikasi dan pencairan tunjangan sertifikasi kini otomatis dilakukan melalui menu Pelatihan Mandiri di Ruang GTK.",
      d: "Melarang guru mengakses Ruang GTK selama jam kerja dan menyerahkan urusan pengelolaan kinerja sepenuhnya kepada operator sekolah.",
      kunci: "B",
      pembahasan: "Kepsesjen Kemendikdasmen Nomor 47/O/2025 serta panduan Rumah Pendidikan menegaskan bahwa PMM telah bertransformasi sejak Januari 2025 menjadi Ruang GTK untuk pengembangan kompetensi, perangkat ajar, pengelolaan kinerja, dan seleksi kepala sekolah menggunakan akun belajar.id, sedangkan verifikasi data guru dan status tunjangan sertifikasi (SKTP/TPG) diakses melalui Info GTK (Opsi B)."
    },
    {
      no: 60, tipe: "Studi Kasus", kategori: "Teknis - Kepribadian", wacana: kasus4,
      pertanyaan: "Menanggapi kecenderungan sejumlah rekan pengawas yang ingin kembali menerapkan pola inspeksi otoriter setelah JF Pengawas Sekolah dipisah kembali dari JF Guru, sikap profesional dan reflektif yang paling tepat Anda tunjukkan adalah ....",
      a: "Ikut mendukung kembalinya pola inspeksi ketat agar kepala sekolah lebih takut dan patuh terhadap pengawas sekolah.",
      b: "Memilih diam saja dalam diskusi kelompok pengawas agar tidak dianggap menentang arus mayoritas senior.",
      c: "Menegaskan bahwa pemisahan kembali JF Pengawas Sekolah dalam PermenPANRB Nomor 7 Tahun 2026 dan Permendikdasmen Nomor 23 Tahun 2026 justru menguatkan profesionalisme pengawasan mutu melalui pemantauan, penilaian, dan pembinaan termasuk pendampingan pada satuan pendidikan formal, bukan mengembalikan pola menghakimi.",
      d: "Mengusulkan penghapusan seluruh kegiatan pendampingan kepala sekolah karena dianggap sebagai sisa regulasi lama.",
      kunci: "C",
      pembahasan: "Pengembalian jabatan pengawas sekolah bertujuan menjaga layanan pendidikan bermutu untuk semua tanpa menggeser paradigma pendampingan. Pasal 16 ayat (3) huruf c Permendikdasmen Nomor 23 Tahun 2026 (diundangkan September 2026) secara eksplisit tetap mencantumkan 'melakukan pembinaan manajerial, akademik, dan penguatan pendidikan karakter termasuk pendampingan pada Satuan Pendidikan formal' sebagai inti kompetensi profesional pengawas (Opsi C)."
    },
    {
      no: 61, tipe: "Studi Kasus", kategori: "Teknis - Sosial", wacana: kasus4,
      pertanyaan: "Untuk menjembatani kesenjangan kapasitas antar-pengawas di wilayah kepulauan/pegunungan sekaligus menyelaraskan program kerja bersama MKKS/K3S, inisiatif kolaboratif berbasis organisasi profesi pengawas sekolah yang paling berdampak adalah ....",
      a: "Membentuk komunitas belajar internal pengawas lintas jenjang (Muda-Madya-Utama) melalui organisasi profesi untuk berbagi strategi pemecahan masalah wilayah terpencil serta menyinkronkan jadwal pendampingan dengan MKKS/K3S dan komunitas belajar sekolah.",
      b: "Mewajibkan setiap pengawas membayar iuran tambahan untuk menyewa konsultan swasta dari luar daerah.",
      c: "Memisahkan forum pengawas sekolah kota dan pengawas sekolah pedalaman agar standar kerjanya tidak saling dibandingkan.",
      d: "Mengirimkan surat edaran organisasi profesi kepada MKKS agar mengundang pengawas hanya pada acara seremonial pembukaan rapat.",
      kunci: "A",
      pembahasan: "Kompetensi sosial Pasal 16 ayat (2) Permendikdasmen Nomor 23 Tahun 2026 menekankan fasilitasi kolaborasi berkualitas. Mengaktifkan organisasi profesi sebagai komunitas belajar sejawat lintas jenjang yang terhubung dengan MKKS/K3S (Opsi A) memperkuat efektivitas layanan pengawasan di daerah dengan tantangan geografis."
    },
    {
      no: 62, tipe: "Studi Kasus", kategori: "Teknis - Profesional", wacana: kasus5,
      pertanyaan: "Sebagai Pengawas Sekolah yang menjalankan fungsi pemantauan dan pembinaan manajerial pada satuan pendidikan formal sesuai Pasal 16 ayat (3) Permendikdasmen Nomor 23 Tahun 2026, tindakan awal paling tepat terhadap penyelenggaraan SPMB sekolah tersebut adalah ....",
      a: "Membatalkan seluruh hasil pengumuman SPMB secara sepihak tanpa memeriksa dokumen verifikasi panitia sekolah.",
      b: "Melakukan pemantauan dan penilaian objektif terhadap dokumen penetapan daya tampung, kesesuaian pelaksanaan jalur seleksi berdasarkan Permendikdasmen Nomor 3 Tahun 2025, serta arsip verifikasi berkas panitia, lalu membina sekolah memperbaiki keterbukaan informasi dan kanal pengaduan resmi.",
      c: "Meminta sekolah menghapus papan pengumuman hasil seleksi dan menggantinya dengan pengumuman tertutup agar tidak memicu kerumunan orang tua.",
      d: "Menyerahkan penanganan keluhan sepenuhnya kepada petugas keamanan sekolah sampai unjuk rasa berakhir dengan sendirinya.",
      kunci: "B",
      pembahasan: "Pemantauan dan penilaian manajerial dalam Pasal 16 ayat (3) Permendikdasmen Nomor 23 Tahun 2026 menuntut pengawas memeriksa bukti pelaksanaan layanan secara objektif sebelum memberikan pembinaan perbaikan (Opsi B). Pemeriksaan kesesuaian prosedural SPMB sesuai Permendikdasmen Nomor 3 Tahun 2025 disertai pembenahan keterbukaan informasi menjaga akuntabilitas publik satuan pendidikan."
    },
    {
      no: 63, tipe: "Studi Kasus", kategori: "Teknis - Sosial", wacana: kasus5,
      pertanyaan: "Untuk meredam ketegangan sosial antara masyarakat/orang tua calon murid dan pihak sekolah sekaligus memperkuat kepercayaan publik terhadap proses SPMB, langkah komunikasi sosial yang paling bijaksana Anda fasilitasi adalah ....",
      a: "Menghindari pertemuan tatap muka dengan perwakilan orang tua dan hanya merilis pernyataan pers pembelaan diri sekolah.",
      b: "Memfasilitasi dialog terbuka yang tertib antara kepala sekolah, ketua panitia SPMB, perwakilan orang tua pengadu, dan komite sekolah untuk memaparkan data daya tampung serta alur verifikasi secara transparan sekaligus menyediakan posko layanan klarifikasi.",
      c: "Meminta aparat keamanan membubarkan perwakilan orang tua agar aktivitas sekolah tidak terganggu.",
      d: "Menyalahkan sistem aplikasi daerah di hadapan orang tua agar pihak sekolah terlihat tidak bersalah.",
      kunci: "B",
      pembahasan: "Kompetensi sosial Pasal 16 ayat (2) Permendikdasmen Nomor 23 Tahun 2026 mengedepankan komunikasi efektif, kolaborasi berkualitas dengan masyarakat, dan penguatan jiwa kewargaan. Ruang dialog terbuka berbasis data dengan layanan klarifikasi (Opsi B) meredam prasangka publik secara bermartabat."
    },
    {
      no: 64, tipe: "Studi Kasus", kategori: "Teknis - Kepribadian", wacana: kasus5,
      pertanyaan: "Menghadapi permintaan pejabat lokal agar Anda menulis laporan pengawasan SPMB dengan klaim '100% tanpa kendala dan tanpa keluhan' demi citra pemerintah daerah, keputusan etis yang mencerminkan integritas seorang Pengawas Sekolah adalah ....",
      a: "Menuruti permintaan tersebut karena laporan pengawas merupakan dokumen internal yang tidak perlu diketahui masyarakat.",
      b: "Membuat dua versi laporan: versi 'tanpa kendala' untuk pejabat lokal dan versi sebenarnya hanya untuk arsip pribadi Anda.",
      c: "Menolak merekayasa fakta laporan, lalu menyajikan laporan pemantauan yang jujur, objektif, dan berimbang—mencantumkan aspek yang sudah sesuai aturan maupun temuan kelemahan komunikasi panitia lengkap dengan langkah perbaikan yang sedang dijalankan sekolah.",
      d: "Menghapus bagian temuan kendala dari laporan tertulis dan hanya menyampaikannya secara lisan apabila ditanya.",
      kunci: "C",
      pembahasan: "Perilaku berintegritas dan bertanggung jawab dalam Pasal 16 ayat (1) Permendikdasmen Nomor 23 Tahun 2026 serta nilai dasar ASN Undang-Undang Nomor 20 Tahun 2023 melarang manipulasi fakta laporan pengawasan. Laporan objektif yang dilengkapi langkah perbaikan konkret (Opsi C) membantu pimpinan daerah mengambil kebijakan korektif yang tepat sekaligus menjaga kredibilitas profesi pengawas sekolah."
    }
  ];
}
