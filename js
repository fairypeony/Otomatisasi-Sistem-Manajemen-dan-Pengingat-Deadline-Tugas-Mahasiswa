/**
 * Fungsi utama untuk membuat event di Google Calendar
 */
function buatEventDeadline() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1');  
  if (!sheet) {
    Logger.log('Sheet tidak ditemukan: Form Responses 1');
    return;  // Keluar jika sheet tidak ditemukan
  }

  const data = sheet.getDataRange().getValues();  // Ambil semua data di sheet
  const calendar = CalendarApp.getCalendarById('ira.arifa31@gmail.com');  // Menggunakan email sebagai Calendar ID

  // Loop melalui data dan buat event untuk setiap tugas
  data.forEach((row, index) => {
    if (index === 0 || row[5]) return;  // Lewati header atau jika event sudah ada di kolom 'Progress'

    const [timestamp, email, tugas, nama, deadline, progress] = row;
    const waktuDeadline = new Date(deadline);
    
    // Ubah jam 00:00 menjadi 11:59
    if (waktuDeadline.getHours() === 0 && waktuDeadline.getMinutes() === 0) {
      waktuDeadline.setHours(11);
      waktuDeadline.setMinutes(59);
    }

    try {
      // Buat event di Google Calendar
      const event = calendar.createEvent(
        `Deadline: ${tugas} - ${nama}`, 
        waktuDeadline, 
        new Date(waktuDeadline.getTime() + 30 * 60 * 1000),  // Durasi 30 menit
        { guests: email, sendInvites: true }
      );

      // Simpan ID Event di kolom 'Progress' untuk menandakan bahwa event telah dibuat
      sheet.getRange(index + 1, 6).setValue('Event Created');  
      Logger.log(`Event dibuat untuk: ${tugas} - ${nama}`);
    } catch (error) {
      Logger.log(`Gagal membuat event untuk ${tugas}: ${error}`);
    }
  });
}

/**
 * Fungsi untuk update status event secara otomatis.
 * Digunakan untuk memberi notifikasi jika deadline sudah dekat.
 */
function cekDanKirimNotifikasi() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1');
  if (!sheet) {
    Logger.log('Sheet tidak ditemukan: Form Responses 1');
    return;  // Keluar jika sheet tidak ditemukan
  }

  const data = sheet.getDataRange().getValues();

  data.forEach((row, index) => {
    if (index === 0) return;  // Lewati header

    const [timestamp, email, tugas, nama, deadline, progress] = row;
    const waktuDeadline = new Date(deadline);
    const waktuSekarang = new Date();

    // Ubah jam 00:00 menjadi 11:59 jika perlu
    if (waktuDeadline.getHours() === 0 && waktuDeadline.getMinutes() === 0) {
      waktuDeadline.setHours(11);
      waktuDeadline.setMinutes(59);
    }

    // Format tanggal untuk ditampilkan tanpa zona waktu
    const tanggalDeadline = waktuDeadline.toLocaleString('id-ID', { 
      year: 'numeric', 
      month: 'long', 
      day: 'numeric', 
      hour: 'numeric', 
      minute: 'numeric',
      hour12: false 
    });

    // Kirim notifikasi jika deadline dalam 1 hari dan tugas belum selesai
    if (progress !== "Selesai" && waktuSekarang >= waktuDeadline - 24 * 60 * 60 * 1000) {
      MailApp.sendEmail({
        to: email,
        subject: `Pengingat Deadline: ${tugas}`,
        body: `Hai ${nama}! 🌟\n\n` +
              `Aku tahu, hidup itu berat, tapi tugas yang satu ini lebih berat lagi kalau nggak kamu kerjakan! 😄\n\n` +
              `Ini reminder super penting buat kamu:\n` +
              `Tugas "${tugas}" punya deadline yang sangat dekat, yaitu pada ${tanggalDeadline} ⏳. ` +
              `Jangan biarkan dia mendekat tanpa kamu bersiap ya, atau dia bisa jadi seperti mantan yang datang tiba-tiba dan bikin pusing! 😆\n\n` +
              `Tapi tenang, kamu bukan mahasiswa biasa, kamu superhero tugas 🦸‍♀️🦸‍♂️! Dan tugas ini? ` +
              `Cuma batu kecil di jalanmu menuju kesuksesan besar 🎓.\n\n` +
              `So, yuk mulai sekarang! Kalau butuh jeda, ambil segelas kopi ☕, putar musik favoritmu 🎶, dan gas! ` +
              `Kamu pasti bisa menyelesaikannya.\n\n` +
              `Salam semangat,\nTim Pengingat Tugas yang Selalu Mendukungmu 💪`
      });
      Logger.log(`Notifikasi dikirim ke ${email} untuk ${tugas}`);
    }
  });
}

/**
 * Fungsi untuk menandai tugas sebagai selesai di Sheets.
 */
function tandaiSelesai(tugas) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1');
  
  // Periksa apakah sheet berhasil ditemukan
  if (!sheet) {
    Logger.log('Sheet tidak ditemukan: Form Responses 1');
    return; // Keluar dari fungsi jika sheet tidak ditemukan
  }

  const data = sheet.getDataRange().getValues();

  data.forEach((row, index) => {
    if (index === 0) return;  // Lewati header

    const [timestamp, email, tugasForm, nama, deadline, progress] = row;
    if (tugasForm === tugas) {
      sheet.getRange(index + 1, 6).setValue('Selesai');  // Update status jadi "Selesai"
      Logger.log(`Tugas ${tugas} ditandai sebagai selesai.`);
    }
  });
}

/**
 * Fungsi untuk mengirim laporan rekap tugas mahasiswa
 */
function kirimRekapTugas() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1');
  if (!sheet) {
    Logger.log('Sheet tidak ditemukan: Form Responses 1');
    return;  // Keluar jika sheet tidak ditemukan
  }

  const data = sheet.getDataRange().getValues();
  let rekapSelesai = [];
  let rekapBelumSelesai = [];

  // Loop melalui data untuk mengumpulkan informasi tugas
  data.forEach((row, index) => {
    if (index === 0) return;  // Lewati header

    const [timestamp, email, tugas, nama, deadline, progress] = row;
    const waktuDeadline = new Date(deadline);
    
    // Format tanggal deadline
    const tanggalDeadline = waktuDeadline.toLocaleString('id-ID', { 
      year: 'numeric', 
      month: 'long', 
      day: 'numeric' 
    });

    if (progress === 'Selesai') {
      rekapSelesai.push(`- ${tugas} (Selesai pada ${tanggalDeadline})`);
    } else {
      rekapBelumSelesai.push(`- ${tugas} (Deadline: ${tanggalDeadline})`);
    }
  });

  // Membuat isi email
  const body = `
    Laporan Rekap Tugas:
    
    Tugas yang sudah selesai:
    ${rekapSelesai.length > 0 ? rekapSelesai.join('\n') : 'Tidak ada tugas yang selesai.'}
    
    Tugas yang belum selesai:
    ${rekapBelumSelesai.length > 0 ? rekapBelumSelesai.join('\n') : 'Tidak ada tugas yang belum selesai.'}
    
    Terima kasih,
    Tim Manajemen Tugas
  `;

  // Kirim email ke dosen
  MailApp.sendEmail({
    to: 'ira.arifa31@gmail.com', // Ganti dengan email dosen
    subject: 'Laporan Rekap Tugas',
    body: body
  });

  Logger.log('Rekap tugas telah dikirim ke dosen.');
}
