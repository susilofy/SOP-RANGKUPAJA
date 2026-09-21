import {
  Document,
  Packer,
  Paragraph,
  Table,
  TableRow,
  TableCell,
  TextRun,
  WidthType,
  AlignmentType,
  PageOrientation,
  BorderStyle,
  Header,
  Footer,
  PageNumber,
} from "docx";
import { SopDocument, SchoolProfile } from "../types";

export async function exportSopToDocx(sop: SopDocument, schoolProfile: SchoolProfile): Promise<void> {
  const thinBorder = {
    top: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
    bottom: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
    left: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
    right: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
  };

  // Helper for cell creation
  const createCell = (text: string | Paragraph[], opts?: { widthPct?: number; bold?: boolean; bg?: string; colSpan?: number; rowSpan?: number; align?: (typeof AlignmentType)[keyof typeof AlignmentType] | any }) => {
    const paragraphs = Array.isArray(text)
      ? text
      : [
          new Paragraph({
            alignment: opts?.align || AlignmentType.LEFT,
            spacing: { before: 40, after: 40, line: 220 },
            children: [
              new TextRun({
                text: text || "-",
                bold: opts?.bold || false,
                size: 18, // 9pt
                font: "Arial",
              }),
            ],
          }),
        ];

    return new TableCell({
      width: opts?.widthPct ? { size: opts.widthPct, type: WidthType.PERCENTAGE } : undefined,
      columnSpan: opts?.colSpan,
      rowSpan: opts?.rowSpan,
      shading: opts?.bg ? { fill: opts.bg } : undefined,
      borders: thinBorder,
      margins: { top: 60, bottom: 60, left: 100, right: 100 },
      children: paragraphs,
    });
  };

  // Build Identitas Table
  const identitasRows: TableRow[] = [
    // Row 1: Header Title & Meta
    new TableRow({
      children: [
        new TableCell({
          width: { size: 55, type: WidthType.PERCENTAGE },
          borders: thinBorder,
          margins: { top: 80, bottom: 80, left: 100, right: 100 },
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 60, after: 60 },
              children: [
                new TextRun({
                  text: `${schoolProfile.namaSekolah.toUpperCase()}`,
                  bold: true,
                  size: 20,
                  font: "Arial",
                }),
              ],
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: `STANDAR OPERASIONAL PROSEDUR (SOP)`,
                  bold: true,
                  size: 22,
                  font: "Arial",
                }),
              ],
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: `${sop.identitas.namaSop.toUpperCase()}`,
                  bold: true,
                  size: 20,
                  font: "Arial",
                }),
              ],
            }),
          ],
        }),
        new TableCell({
          width: { size: 45, type: WidthType.PERCENTAGE },
          borders: thinBorder,
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          children: [
            new Paragraph({
              children: [
                new TextRun({ text: "Nomor SOP            : ", bold: true, size: 18, font: "Arial" }),
                new TextRun({ text: sop.identitas.nomorSop || "-", size: 18, font: "Arial" }),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "Tanggal Pembuatan : ", bold: true, size: 18, font: "Arial" }),
                new TextRun({ text: sop.identitas.tanggalPembuatan || "-", size: 18, font: "Arial" }),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "Tanggal Revisi        : ", bold: true, size: 18, font: "Arial" }),
                new TextRun({ text: sop.identitas.tanggalRevisi || "0", size: 18, font: "Arial" }),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "Tanggal Pengesahan: ", bold: true, size: 18, font: "Arial" }),
                new TextRun({ text: sop.identitas.tanggalPengesahan || "-", size: 18, font: "Arial" }),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({ text: "Disahkan Oleh        : ", bold: true, size: 18, font: "Arial" }),
                new TextRun({ text: sop.identitas.disahkanOleh || `Kepala ${schoolProfile.namaSekolah}`, size: 18, font: "Arial" }),
              ],
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 720, after: 40 }, // 4 lines of space (4 spasi) for signature and official seal
              children: [
                new TextRun({
                  text: sop.identitas.namaKepalaSekolah || schoolProfile.namaKepalaSekolah,
                  bold: true,
                  underline: {},
                  size: 18,
                  font: "Arial",
                }),
              ],
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: `NIP. ${sop.identitas.nip || schoolProfile.nip}`,
                  size: 18,
                  font: "Arial",
                }),
              ],
            }),
          ],
        }),
      ],
    }),
    // Row 2: Nama SOP Title
    new TableRow({
      children: [
        createCell("NAMA SOP", { widthPct: 55, bold: true, bg: "F3F4F6" }),
        createCell(sop.identitas.namaSop, { widthPct: 45, bold: true }),
      ],
    }),
    // Row 3: Dasar Hukum vs Kualifikasi Pelaksana
    new TableRow({
      children: [
        new TableCell({
          width: { size: 55, type: WidthType.PERCENTAGE },
          borders: thinBorder,
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          children: [
            new Paragraph({
              children: [new TextRun({ text: "DASAR HUKUM:", bold: true, size: 18, font: "Arial" })],
            }),
            ...sop.dasarHukum.map(
              (dh, idx) =>
                new Paragraph({
                  spacing: { before: 20, after: 20 },
                  children: [
                    new TextRun({
                      text: `${idx + 1}. ${dh.namaRegulasi} tentang ${dh.tentang}`,
                      size: 17,
                      font: "Arial",
                    }),
                  ],
                })
            ),
          ],
        }),
        new TableCell({
          width: { size: 45, type: WidthType.PERCENTAGE },
          borders: thinBorder,
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          children: [
            new Paragraph({
              children: [new TextRun({ text: "KUALIFIKASI PELAKSANA:", bold: true, size: 18, font: "Arial" })],
            }),
            ...sop.kualifikasiPelaksana.map(
              (kp, idx) =>
                new Paragraph({
                  spacing: { before: 20, after: 20 },
                  children: [new TextRun({ text: `${idx + 1}. ${kp}`, size: 17, font: "Arial" })],
                })
            ),
          ],
        }),
      ],
    }),
    // Row 4: Keterkaitan vs Peralatan/Perlengkapan
    new TableRow({
      children: [
        new TableCell({
          width: { size: 55, type: WidthType.PERCENTAGE },
          borders: thinBorder,
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          children: [
            new Paragraph({
              children: [new TextRun({ text: "KETERKAITAN:", bold: true, size: 18, font: "Arial" })],
            }),
            ...sop.keterkaitan.map(
              (item, idx) =>
                new Paragraph({
                  spacing: { before: 20, after: 20 },
                  children: [new TextRun({ text: `${idx + 1}. ${item}`, size: 17, font: "Arial" })],
                })
            ),
          ],
        }),
        new TableCell({
          width: { size: 45, type: WidthType.PERCENTAGE },
          borders: thinBorder,
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          children: [
            new Paragraph({
              children: [new TextRun({ text: "PERALATAN / PERLENGKAPAN:", bold: true, size: 18, font: "Arial" })],
            }),
            ...sop.peralatanPerlengkapan.map(
              (item, idx) =>
                new Paragraph({
                  spacing: { before: 20, after: 20 },
                  children: [new TextRun({ text: `${idx + 1}. ${item}`, size: 17, font: "Arial" })],
                })
            ),
          ],
        }),
      ],
    }),
    // Row 5: Peringatan vs Pencatatan dan Pendataan
    new TableRow({
      children: [
        new TableCell({
          width: { size: 55, type: WidthType.PERCENTAGE },
          borders: thinBorder,
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          children: [
            new Paragraph({
              children: [new TextRun({ text: "PERINGATAN:", bold: true, size: 18, font: "Arial" })],
            }),
            ...sop.peringatan.map(
              (item, idx) =>
                new Paragraph({
                  spacing: { before: 20, after: 20 },
                  children: [new TextRun({ text: `${idx + 1}. ${item}`, size: 17, font: "Arial" })],
                })
            ),
          ],
        }),
        new TableCell({
          width: { size: 45, type: WidthType.PERCENTAGE },
          borders: thinBorder,
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
          children: [
            new Paragraph({
              children: [new TextRun({ text: "PENCATATAN DAN PENDATAAN:", bold: true, size: 18, font: "Arial" })],
            }),
            ...(sop.pencatatanPendataan?.dokumenBukti || []).map(
              (item, idx) =>
                new Paragraph({
                  spacing: { before: 20, after: 20 },
                  children: [new TextRun({ text: `${idx + 1}. ${item}`, size: 17, font: "Arial" })],
                })
            ),
            new Paragraph({
              spacing: { before: 40 },
              children: [
                new TextRun({
                  text: `Arsip: ${sop.pencatatanPendataan?.penanggungJawabArsip || "-"} | ${sop.pencatatanPendataan?.mediaPenyimpanan || "-"} | Retensi: ${sop.pencatatanPendataan?.periodePenyimpanan || "-"}`,
                  italics: true,
                  size: 16,
                  font: "Arial",
                }),
              ],
            }),
          ],
        }),
      ],
    }),
  ];

  const identitasTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: identitasRows,
  });

  // Build Pelaksana Mutu Baku Table
  const pelaksanaCols = sop.pelaksanaList || ["Kepala Sekolah", "Guru", "Tendik"];
  const totalPelaksana = pelaksanaCols.length;

  const headerRow1 = new TableRow({
    children: [
      createCell("NO", { bold: true, bg: "D1D5DB", align: AlignmentType.CENTER, rowSpan: 2 }),
      createCell("URAIAN PROSEDUR", { bold: true, bg: "D1D5DB", align: AlignmentType.CENTER, rowSpan: 2 }),
      createCell("PELAKSANA", { bold: true, bg: "D1D5DB", align: AlignmentType.CENTER, colSpan: totalPelaksana }),
      createCell("MUTU BAKU", { bold: true, bg: "D1D5DB", align: AlignmentType.CENTER, colSpan: 3 }),
    ],
  });

  const headerRow2 = new TableRow({
    children: [
      ...pelaksanaCols.map((p) => createCell(p, { bold: true, bg: "E5E7EB", align: AlignmentType.CENTER })),
      createCell("Persyaratan / Perlengkapan", { bold: true, bg: "E5E7EB", align: AlignmentType.CENTER }),
      createCell("Waktu", { bold: true, bg: "E5E7EB", align: AlignmentType.CENTER }),
      createCell("Output", { bold: true, bg: "E5E7EB", align: AlignmentType.CENTER }),
    ],
  });

  const numberIndexRow = new TableRow({
    children: [
      createCell("1", { bg: "F3F4F6", align: AlignmentType.CENTER }),
      createCell("2", { bg: "F3F4F6", align: AlignmentType.CENTER }),
      ...pelaksanaCols.map((_, i) => createCell(String(3 + i), { bg: "F3F4F6", align: AlignmentType.CENTER })),
      createCell(String(3 + totalPelaksana), { bg: "F3F4F6", align: AlignmentType.CENTER }),
      createCell(String(4 + totalPelaksana), { bg: "F3F4F6", align: AlignmentType.CENTER }),
      createCell(String(5 + totalPelaksana), { bg: "F3F4F6", align: AlignmentType.CENTER }),
    ],
  });

  const totalSteps = (sop.tabelPelaksanaMutuBaku || []).length;
  const stepRows: TableRow[] = (sop.tabelPelaksanaMutuBaku || []).map((step, idx) => {
    const isEven = idx % 2 === 1;
    const rowBg = isEven ? "F9FAFB" : "FFFFFF";
    const isLast = idx === totalSteps - 1;
    const flow = step.flowType || (idx === 0 ? "start" : isLast ? "end" : "process");

    const checkedCols = pelaksanaCols
      .map((colName, i) => (step.pelaksanaChecks?.[colName] ? i : -1))
      .filter((i) => i !== -1);
    const primaryCol =
      step.activePelaksanaIndex !== undefined && checkedCols.includes(step.activePelaksanaIndex)
        ? step.activePelaksanaIndex
        : checkedCols.length > 0
        ? checkedCols[0]
        : 0;

    const nextStep = !isLast ? (sop.tabelPelaksanaMutuBaku || [])[idx + 1] : null;
    const nextCheckedCols = nextStep
      ? pelaksanaCols.map((colName, i) => (nextStep.pelaksanaChecks?.[colName] ? i : -1)).filter((i) => i !== -1)
      : [];
    const nextPrimaryCol = nextStep
      ? (nextStep.activePelaksanaIndex !== undefined && nextCheckedCols.includes(nextStep.activePelaksanaIndex)
          ? nextStep.activePelaksanaIndex
          : nextCheckedCols[0] ?? 0)
      : -1;

    const pelaksanaCells = pelaksanaCols.map((colName, pIdx) => {
      const isChecked = !!step.pelaksanaChecks?.[colName];
      let mark = "";
      if (!isChecked) {
        if (!isLast && flow !== "end") {
          const minCol = Math.min(primaryCol, nextPrimaryCol);
          const maxCol = Math.max(primaryCol, nextPrimaryCol);
          if (pIdx > minCol && pIdx < maxCol) {
            mark = "────►";
          } else if (pIdx === nextPrimaryCol && primaryCol !== nextPrimaryCol) {
            mark = "│ ▼";
          }
        }
      } else if (pIdx === primaryCol) {
        let arrow = "▼";
        if (!isLast && primaryCol !== nextPrimaryCol) {
          arrow = primaryCol < nextPrimaryCol ? "►" : "◄";
        }
        if (flow === "start") mark = `[MULAI] ${arrow}`;
        else if (flow === "decision") mark = `[KEPUTUSAN] ${arrow}`;
        else if (flow === "end") mark = "[SELESAI]";
        else if (flow === "check") mark = "✓";
        else mark = isLast ? "[PROSES]" : `[PROSES] ${arrow}`;
      } else {
        mark = "── [PROSES] ──";
      }
      return createCell(mark, { bg: rowBg, align: AlignmentType.CENTER, bold: !!mark });
    });

    return new TableRow({
      children: [
        createCell(String(step.no || idx + 1), { bg: rowBg, align: AlignmentType.CENTER }),
        createCell(step.uraianProsedur || "", { bg: rowBg }),
        ...pelaksanaCells,
        createCell(step.persyaratan || "-", { bg: rowBg }),
        createCell(step.waktu || "-", { bg: rowBg, align: AlignmentType.CENTER }),
        createCell(step.output || "-", { bg: rowBg }),
      ],
    });
  });

  const mutuBakuTable = new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [headerRow1, headerRow2, numberIndexRow, ...stepRows],
  });

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            size: {
              orientation: PageOrientation.LANDSCAPE,
              width: 16838, // A4 Landscape width in twips (297mm)
              height: 11906, // A4 Landscape height in twips (210mm)
            },
            margin: {
              top: 720, // 0.5 inch
              bottom: 720,
              left: 720,
              right: 720,
            },
          },
        },
        headers: {
          default: new Header({
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [
                  new TextRun({
                    text: `DOKUMEN RESMI SOP - ${schoolProfile.namaSekolah.toUpperCase()}`,
                    size: 16,
                    color: "666666",
                    font: "Arial",
                  }),
                ],
              }),
            ],
          }),
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [
                  new TextRun({ text: "Halaman ", size: 16, font: "Arial" }),
                  new TextRun({ children: [PageNumber.CURRENT], size: 16, font: "Arial" }),
                  new TextRun({ text: " dari ", size: 16, font: "Arial" }),
                  new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, font: "Arial" }),
                  new TextRun({ text: ` | SOP SMART SCHOOL SD`, italics: true, size: 16, font: "Arial", color: "888888" }),
                ],
              }),
            ],
          }),
        },
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { after: 120 },
            children: [
              new TextRun({
                text: `SOP : ${sop.identitas.namaSop.toUpperCase()}`,
                bold: true,
                size: 22,
                font: "Arial",
              }),
            ],
          }),
          identitasTable,
          new Paragraph({
            pageBreakBefore: true,
            alignment: AlignmentType.CENTER,
            spacing: { before: 120, after: 140 },
            children: [
              new TextRun({
                text: `PELAKSANA MUTU BAKU : ${sop.identitas.namaSop.toUpperCase()}`,
                bold: true,
                size: 22,
                font: "Arial",
              }),
            ],
          }),
          mutuBakuTable,
          new Paragraph({
            spacing: { before: 80, after: 140 },
            children: [
              new TextRun({
                text: "Keterangan Alur: (MULAI)/(SELESAI) = Titik Awal/Akhir  |  [PROSES] = Aktivitas Pelaksanaan  |  <KEPUTUSAN> = Pengambilan Keputusan  |  | v = Garis Alur Panah\n",
                size: 15,
                color: "555555",
                font: "Arial",
              }),
              new TextRun({
                text: "Format Baku Permenpan RB No. 35 / Kemendikbudristek",
                italics: true,
                size: 15,
                color: "777777",
                font: "Arial",
              }),
            ],
          }),
          new Paragraph({
            spacing: { before: 160 },
            alignment: AlignmentType.RIGHT,
            children: [
              new TextRun({ text: `${schoolProfile.kabupatenKota}, ${sop.identitas.tanggalPengesahan || ".................. 2026"}\n`, size: 18, font: "Arial" }),
              new TextRun({ text: `Kepala ${schoolProfile.namaSekolah}\n`, bold: true, size: 18, font: "Arial" }),
              new TextRun({
                text: "\n\n\n\n", // 4 spaces for manual signature and seal
                size: 18,
                font: "Arial",
              }),
              new TextRun({ text: `${sop.identitas.namaKepalaSekolah || schoolProfile.namaKepalaSekolah}\n`, bold: true, underline: {}, size: 18, font: "Arial" }),
              new TextRun({ text: `NIP. ${sop.identitas.nip || schoolProfile.nip}`, size: 18, font: "Arial" }),
            ],
          }),
        ],
      },
    ],
  });

  const blob = await Packer.toBlob(doc);
  const cleanFilename = `SOP_${sop.identitas.namaSop.replace(/[^a-zA-Z0-9]/g, "_")}_A4_Landscape.docx`;
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = cleanFilename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
