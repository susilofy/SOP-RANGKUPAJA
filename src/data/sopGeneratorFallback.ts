import { SopDocument, SchoolProfile } from "../types";

export function buildFallbackSopDocument(
  namaSop: string,
  kategori: string,
  schoolProfile: SchoolProfile,
  interviewAnswers?: Record<string, string>
): SopDocument {
  const currentYear = new Date().getFullYear();
  const rawCode = (kategori || "A").slice(0, 1).toUpperCase();

  // Extract interview answers if provided
  const ansPelaksana = interviewAnswers?.["q1"] || "Kepala Sekolah, Guru Kelas, Tenaga Administrasi (TAS), Komite";
  const ansDokumen = interviewAnswers?.["q2"] || "SK Penetapan Tim, Panduan Teknis, Format Blanko Instrumen";
  const ansWaktu = interviewAnswers?.["q4"] || "3 - 5 hari kerja";
  const ansOutput = interviewAnswers?.["q5"] || "Dokumen final bertanda tangan, berita acara, dan arsip digital";

  // Parse pelaksana roles
  const pelaksanaRoles = ansPelaksana
    .split(/[,;\n]+/)
    .map((r) => r.trim())
    .filter((r) => r.length > 0);

  const pelaksanaList = [
    pelaksanaRoles[0] || "Kepala Sekolah",
    pelaksanaRoles[1] || "Guru / Tim Pelaksana",
    pelaksanaRoles[2] || "Tenaga Administrasi (TAS)",
    pelaksanaRoles[3] || "Komite / Orang Tua",
    pelaksanaRoles[4] || "Pengawas / Dinas Pendidikan",
  ].slice(0, 5);

  while (pelaksanaList.length < 3) {
    pelaksanaList.push(`Pelaksana ${pelaksanaList.length + 1}`);
  }

  return {
    id: `sop-${Date.now()}`,
    identitas: {
      namaSop: namaSop || "SOP Standar Operasional Sekolah",
      nomorSop: `SOP/SD/${rawCode}/${String(Math.floor(100 + Math.random() * 900))}/${currentYear}`,
      tanggalPembuatan: `02 Januari ${currentYear}`,
      tanggalRevisi: "00",
      tanggalPengesahan: `10 Januari ${currentYear}`,
      disahkanOleh: `Kepala ${schoolProfile?.namaSekolah || "Sekolah Dasar"}`,
      namaKepalaSekolah: schoolProfile?.namaKepalaSekolah || "Kepala Sekolah",
      nip: schoolProfile?.nip || "-",
      unitKerja: schoolProfile?.namaSekolah || "Satuan Pendidikan Dasar",
    },
    dasarHukum: [
      {
        id: "dh-1",
        namaRegulasi: "Undang-Undang Nomor 20 Tahun 2003",
        nomor: "20",
        tahun: "2003",
        tentang: "Sistem Pendidikan Nasional",
        statusVerifikasi: "Terverifikasi Resmi",
        sumber: "JDIH Kemendikbudristek",
      },
      {
        id: "dh-2",
        namaRegulasi: "Peraturan Pemerintah Nomor 57 Tahun 2021 jo PP Nomor 4 Tahun 2022",
        nomor: "57 jo 4",
        tahun: "2021/2022",
        tentang: "Standar Nasional Pendidikan",
        statusVerifikasi: "Terverifikasi Resmi",
        sumber: "JDIH BPK RI",
      },
      {
        id: "dh-3",
        namaRegulasi: "Permendikbudristek Nomor 47 Tahun 2023",
        nomor: "47",
        tahun: "2023",
        tentang: "Standar Pengelolaan pada PAUD, Dikdas, dan Dikmen",
        statusVerifikasi: "Terverifikasi Resmi",
        sumber: "JDIH Kemendikbudristek",
      },
    ],
    kualifikasiPelaksana: [
      "Memahami tugas pokok dan fungsi (tupoksi) di lingkungan sekolah dasar",
      "Memiliki keterampilan administrasi dan komunikasi efektif dengan warga sekolah",
      "Mampu mendokumentasikan setiap tahapan kegiatan secara transparan dan akuntabel",
    ],
    keterkaitan: [
      "SOP Manajemen Tata Kelola Sekolah Dasar",
      "SOP Administrasi dan Pelaporan Satuan Pendidikan",
    ],
    peralatanPerlengkapan: [
      "Komputer / Laptop dan Jaringan Internet",
      "Perangkat ATK, Lembar Verifikasi, dan Format Formulir",
      "Buku Agenda Kerja dan Buku Ekspedisi Persuratan",
    ],
    peringatan: [
      "Pelaksanaan prosedur wajib mematuhi tenggat waktu untuk menghindari kendala operasional",
      "Setiap kelalaian prosedur berpotensi menimbulkan kendala koordinasi atau temuan audit administrasi",
    ],
    pencatatanPendataan: {
      dokumenBukti: [
        "Lembar disposisi pimpinan",
        "Daftar hadir dan notulen koordinasi",
        "Berita acara / lembar pengesahan",
        "Arsip berkas pendukung fisik dan digital",
      ],
      penanggungJawabArsip: "Tenaga Administrasi Sekolah & Kepala Sekolah",
      mediaPenyimpanan: "Map Ordner Ruang Tata Usaha & Google Drive Sekolah",
      periodePenyimpanan: "Minimal 3 sampai 5 tahun ajaran",
    },
    pelaksanaList,
    tabelPelaksanaMutuBaku: [
      {
        id: "step-1",
        no: 1,
        uraianProsedur: `Kepala Sekolah memberikan arahan kebijakan dan mandat pembentukan tim kerja terkait ${namaSop}.`,
        pelaksanaChecks: {
          [pelaksanaList[0]]: "start",
          [pelaksanaList[1]]: false,
          [pelaksanaList[2]]: false,
        },
        persyaratan: "Juknis / Regulasi Terkait, Agenda Rapat",
        waktu: "1 hari kerja",
        output: "Instruksi Pelaksanaan & SK Tim Kerja",
        flowType: "start",
      },
      {
        id: "step-2",
        no: 2,
        uraianProsedur: "Tim pelaksana mengumpulkan data awal, menyusun draf dokumen, serta menyiapkan kelengkapan persyaratan.",
        pelaksanaChecks: {
          [pelaksanaList[0]]: false,
          [pelaksanaList[1]]: "process",
          [pelaksanaList[2]]: pelaksanaList[2] ? "process" : false,
        },
        persyaratan: ansDokumen,
        waktu: ansWaktu || "2 hari kerja",
        output: "Draf berkas awal & rekapitulasi data",
        flowType: "process",
      },
      {
        id: "step-3",
        no: 3,
        uraianProsedur: "Tenaga administrasi dan tim memeriksa kelengkapan berkas serta memvalidasi keabsahan data.",
        pelaksanaChecks: {
          [pelaksanaList[0]]: false,
          [pelaksanaList[1]]: false,
          [pelaksanaList[2]]: "decision",
        },
        persyaratan: "Checklist verifikasi data",
        waktu: "1 hari kerja",
        output: "Lembar verifikasi terparaf",
        flowType: "decision",
      },
      {
        id: "step-4",
        no: 4,
        uraianProsedur: "Kepala Sekolah menelaah hasil verifikasi dan memberikan persetujuan pengesahan resmi dokumen.",
        pelaksanaChecks: {
          [pelaksanaList[0]]: "decision",
          [pelaksanaList[1]]: false,
          [pelaksanaList[2]]: false,
        },
        persyaratan: "Draf terverifikasi & berkas usulan",
        waktu: "1 hari kerja",
        output: "Dokumen disetujui & ditandatangani",
        flowType: "decision",
      },
      {
        id: "step-5",
        no: 5,
        uraianProsedur: "Sosialisasi hasil pelaksanaan kepada pihak berkepentingan serta pengarsipan dokumen resmi secara tertib.",
        pelaksanaChecks: {
          [pelaksanaList[0]]: false,
          [pelaksanaList[1]]: "end",
          [pelaksanaList[2]]: "end",
        },
        persyaratan: "Dokumen final bertanda tangan, buku ekspedisi",
        waktu: "1 hari kerja",
        output: ansOutput,
        flowType: "end",
      },
    ],
    checklistKelengkapan: {
      namaSopTersedia: true,
      nomorSopTersedia: true,
      tanggalTersedia: true,
      kepalaSekolahTersedia: true,
      nipTersedia: true,
      dasarHukumTersedia: true,
      dasarHukumTerverifikasi: true,
      kualifikasiPelaksanaTersedia: true,
      keterkaitanTersedia: true,
      peralatanTersedia: true,
      peringatanTersedia: true,
      pencatatanTersedia: true,
      alurProsedurJelas: true,
      semuaLangkahMemilikiPelaksana: true,
      semuaLangkahMemilikiOutput: true,
      waktuTersedia: true,
      dokumenPendukungTersedia: true,
      tidakAdaLangkahAmbigu: true,
      tidakAdaPengulangan: true,
      skorAdministratif: 100,
      catatanPerbaikan: [],
    },
    versi: "1.0",
    status: "DRAFT",
    kategori: kategori || "A. Manajemen Sekolah",
    tanggalReviewBerikutnya: `${currentYear + 1}-01-10`,
    penanggungJawab: schoolProfile?.namaKepalaSekolah || "Kepala Sekolah",
    riwayatRevisi: [],
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };
}
