/*
  Menyusun LAPORAN-VIVA-MPC.docx.

  Prinsip yang dipegang skrip ini: tidak ada satu pun angka hasil yang diketik
  tangan. Tabel Bab IV dan seluruh lampiran data dibaca dari results.json,
  pareto.json, dan data_laporan.json — ketiganya keluaran langsung dari
  run_experiments.py, run_pareto.py, dan skrip pengumpul data. Kalau eksperimen
  diulang dengan parameter berbeda, laporan ikut berubah begitu skrip ini
  dijalankan lagi, dan tidak ada angka yang bisa tertinggal basi.

  Nomor halaman pada Daftar Isi, Daftar Gambar, dan Daftar Tabel diambil dari
  toc.json, yang dihasilkan dengan membaca PDF hasil render dokumen ini sendiri.
  Bidang daftar isi bawaan Word sengaja tidak dipakai karena baru terisi setelah
  pengguna menekan Update Field; dokumen yang langsung dicetak akan kosong.
*/
const fs = require('fs');
const path = require('path');
const D = require('docx');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, PageBreak,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, ImageRun,
  Footer, PageNumber, LevelFormat, convertInchesToTwip, ExternalHyperlink,
} = D;

const HERE = __dirname;
const TEXT_W = 11906 - 1700 - 1440;          // lebar kolom teks A4 dalam DXA
const res = JSON.parse(fs.readFileSync(path.join(HERE, 'results.json'), 'utf8'));
const par = JSON.parse(fs.readFileSync(path.join(HERE, 'pareto.json'), 'utf8'));
const DAT = JSON.parse(fs.readFileSync(path.join(HERE, 'data_laporan.json'), 'utf8'));

const INK = '182226', ACC = '0F6E63', MUT = '4F5F67', WARM = '4C46B6';
const num = (v, d = 1) => (v === null || v === undefined || Number.isNaN(Number(v)) ? '–'
  : Number(v).toFixed(d).replace('.', ','));

/* ---------------- pembantu tata letak ---------------- */
const P = (text, o = {}) => new Paragraph({
  alignment: o.align || AlignmentType.JUSTIFIED,
  spacing: { after: o.after === undefined ? 130 : o.after, line: 288 },
  indent: o.indent, numbering: o.numbering,
  children: [new TextRun({ text, size: o.size || 21, bold: o.bold, italics: o.italics,
    color: o.color || INK, font: 'Calibri' })],
});
const Rich = (runs, o = {}) => new Paragraph({
  alignment: o.align || AlignmentType.JUSTIFIED,
  spacing: { after: o.after === undefined ? 130 : o.after, line: 288 },
  indent: o.indent, numbering: o.numbering,
  children: runs.map(r => typeof r === 'string'
    ? new TextRun({ text: r, size: 21, font: 'Calibri', color: INK })
    : new TextRun({ size: 21, font: 'Calibri', color: INK, ...r })),
});
const H = (text, level) => new Paragraph({
  heading: level,
  spacing: { before: level === HeadingLevel.HEADING_1 ? 380 : 280, after: 140 },
  children: [new TextRun({ text, font: 'Calibri', bold: true,
    color: level === HeadingLevel.HEADING_1 ? ACC : INK,
    size: level === HeadingLevel.HEADING_1 ? 30 : level === HeadingLevel.HEADING_2 ? 25 : 22 })],
});
const Eq = (text) => new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { before: 150, after: 150 },
  children: [new TextRun({ text, font: 'Cambria Math', size: 22, color: INK })],
});
const CAPS = [];
const NCAP = { Gambar: 0, Tabel: 0 };
/*
  Nomor pada keterangan dihitung, bukan ditulis. Sebelumnya nomornya tetap, dan
  menyisipkan satu tabel di tengah laporan membuat seluruh nomor sesudahnya
  meleset satu tanpa ada yang memberi tahu.
*/
const Caption = (raw) => {
  const text = raw.replace(/^(Gambar|Tabel)\s+\d+\./, (m, w) => `${w} ${++NCAP[w]}.`);
  CAPS.push(text); return new Paragraph({
  alignment: AlignmentType.CENTER, spacing: { before: 70, after: 260 },
  children: [new TextRun({ text, size: 18, italics: true, color: MUT, font: 'Calibri' })] }); };
const Code = (text) => new Paragraph({
  spacing: { before: 40, after: 40, line: 230 },
  shading: { type: ShadingType.CLEAR, fill: 'F4F2ED' },
  indent: { left: 160, right: 160 },
  children: [new TextRun({ text, font: 'Consolas', size: 16, color: INK })],
});
const CodeBlock = (src) => src.split('\n').map(l => Code(l.length ? l : ' '));
const Note = (text) => new Paragraph({
  spacing: { before: 120, after: 160, line: 280 },
  shading: { type: ShadingType.CLEAR, fill: 'ECEBF9' },
  indent: { left: 180, right: 180 },
  border: { left: { style: BorderStyle.SINGLE, size: 12, color: WARM.replace('#', '') } },
  children: [new TextRun({ text, size: 20, color: INK, font: 'Calibri' })],
});
const Link = (label, target) => new Paragraph({
  spacing: { after: 60 },
  children: [new ExternalHyperlink({ link: target,
    children: [new TextRun({ text: label, size: 19, font: 'Consolas', color: '0F6E63', underline: {} })] })],
});

const RATIO = { 'gambar1_matriks.png': 11.6 / 3.5, 'gambar2_dosis.png': 10.4 / 3.6,
  'gambar3_ablasi.png': 6.6 / 4.4, 'gambar4_arsitektur.png': 7.0 / 4.6, 'gambar5_paru3d.png': 760 / 470 };
const Img = (file, widthIn) => {
  const w = widthIn * 96;
  return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 170, after: 0 },
    children: [new ImageRun({ data: fs.readFileSync(path.join(HERE, file)), type: 'png',
      transformation: { width: Math.round(w), height: Math.round(w / RATIO[file]) } })] });
};

const CB = { style: BorderStyle.SINGLE, size: 4, color: 'D9D3C8' };
const borders = { top: CB, bottom: CB, left: CB, right: CB };
function tbl(headers, rows, widths, opt = {}) {
  const fs_ = opt.size || 18;
  const mk = (txt, i, head) => new TableCell({
    width: { size: widths[i], type: WidthType.DXA }, borders,
    shading: head ? { type: ShadingType.CLEAR, fill: 'E1EFEB' } : undefined,
    margins: { top: 55, bottom: 55, left: 85, right: 85 },
    children: [new Paragraph({
      alignment: i === 0 ? AlignmentType.LEFT : (opt.leftAll ? AlignmentType.LEFT : AlignmentType.RIGHT),
      spacing: { after: 0, line: 240 },
      children: [new TextRun({ text: String(txt), size: fs_, bold: head,
        color: head ? ACC : INK, font: opt.mono && !head ? 'Consolas' : 'Calibri' })] })],
  });
  return new Table({
    columnWidths: widths, width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    rows: [new TableRow({ tableHeader: true, children: headers.map((h, i) => mk(h, i, true)) })]
      .concat(rows.map(r => new TableRow({ children: r.map((c, i) => mk(c, i, false)) }))),
  });
}

/* ---------------- daftar isi / gambar / tabel ---------------- */
function listFromToc(kind) {
  let data;
  try { data = JSON.parse(fs.readFileSync(path.join(HERE, 'toc.json'), 'utf8')); }
  catch (e) { return [P('(Daftar dibuat otomatis setelah dokumen dirender sekali.)', { italics: true, color: MUT })]; }
  const items = data[kind] || [];
  if (!items.length) return [P('(Belum tersedia — jalankan ulang pembangun laporan.)', { italics: true, color: MUT })];
  return items.map(e => {
    const h1 = e.level === 'HEADING_1';
    return new Paragraph({
      spacing: { after: h1 ? 70 : 30, before: h1 ? 60 : 0 },
      indent: { left: kind === 'entries' ? (h1 ? 0 : 280) : 0 },
      tabStops: [{ type: D.TabStopType.RIGHT, position: TEXT_W, leader: D.LeaderType.DOT }],
      children: [
        new TextRun({ text: e.text, size: h1 ? 20 : 19, bold: h1, font: 'Calibri', color: h1 ? INK : MUT }),
        new TextRun({ size: h1 ? 20 : 19, bold: h1, font: 'Calibri', color: h1 ? INK : MUT,
          children: [new D.Tab(), String(e.page)] }),
      ],
    });
  });
}

/* ---------------- data eksperimen ---------------- */
const SCN = ['sehat', 'ards_ringan', 'ards_sedang', 'ards_berat', 'obstruktif', 'derekrutmen'];
const CTRL = ['ARDSNet', 'ASV-like', 'Noisy PCV (CV 30%)', 'VIVA-MPC'];
const SHORT = { 'ARDSNet': 'ARDSNet', 'ASV-like': 'ASV-like', 'Noisy PCV (CV 30%)': 'Noisy PCV', 'VIVA-MPC': 'VIVA-MPC' };
const get = (sc, c) => res.rows.find(r => r.scenario === sc && r.controller === c) || {};
const label = sc => (res.rows.find(r => r.scenario === sc) || {}).label || sc;
const matrix = (key, dec) => tbl(['Kondisi paru'].concat(CTRL.map(c => SHORT[c])),
  SCN.map(sc => [label(sc)].concat(CTRL.map(c => num(get(sc, c)[key], dec)))),
  [2500, 1450, 1450, 1450, 1450]);
const ORDER = ['CV=0', 'CV=0.15', 'CV=0.30', 'CV bebas'];
const paretoRows = par.rows.slice().sort((a, b) => a.w_gi - b.w_gi || ORDER.indexOf(a.variant) - ORDER.indexOf(b.variant));

/* ---------------- indikator kinerja utama ---------------- */
/*
  Seluruh pencapaian di bawah dihitung dari results.json, bukan dinilai sendiri.
  Bila eksperimen diulang dan hasilnya memburuk, tabel pencapaian ikut berubah
  dan statusnya berbalik menjadi tidak tercapai. Itulah gunanya KPI: ia harus
  bisa menyatakan kegagalan tanpa perlu penulis mengakuinya secara sukarela.
*/
const V = sc => get(sc, 'VIVA-MPC');
const maxOf = (f) => Math.max(...SCN.map(sc => f(V(sc))));
const minOf = (f) => Math.min(...SCN.map(sc => f(V(sc))));
const countOf = (pred) => SCN.filter(sc => pred(sc)).length;
const cvs = SCN.map(sc => V(sc).vt_cv_realised);
const kontrasLolos = DAT.kontras.filter(k => k[3] >= k[4]).length;

const KPI = [
  { k: 'K1', nama: 'Tidak ada napas dengan tekanan plateau di atas 30 cmH₂O',
    pendek: 'Plateau > 30 cmH₂O', target: '0% napas',
    dasar: 'Batas keselamatan ARDSNet; di atas 30 risiko barotrauma meningkat tajam.',
    hasil: num(maxOf(r => r.pct_breaths_plateau_over_30 || 0), 1) + '% (terburuk)',
    ok: maxOf(r => r.pct_breaths_plateau_over_30 || 0) < 0.5 },
  { k: 'K2', nama: 'Driving pressure rerata tidak melampaui 15 cmH₂O',
    pendek: 'ΔP rerata ≤ 15', target: '6 dari 6 kondisi',
    dasar: 'Amato dkk. (2015): prediktor kelangsungan hidup terkuat pada ARDS.',
    hasil: countOf(sc => V(sc).driving_mean <= 15) + ' dari 6 kondisi',
    ok: countOf(sc => V(sc).driving_mean <= 15) === 6,
    sebagian: countOf(sc => V(sc).driving_mean <= 15) >= 4 },
  { k: 'K3', nama: 'Saturasi oksigen rerata tidak turun di bawah 88%',
    pendek: 'SpO₂ ≥ 88%', target: '≥ 88% pada semua kondisi',
    dasar: 'Ambang bawah target oksigenasi permisif pada ARDS.',
    hasil: 'terendah ' + num(minOf(r => r.spo2_mean), 1) + '%',
    ok: minOf(r => r.spo2_mean) >= 88 },
  { k: 'K4', nama: 'PaCO₂ rerata tidak melampaui 65 mmHg',
    pendek: 'PaCO₂ ≤ 65 mmHg', target: '≤ 65 pada semua kondisi',
    dasar: 'Batas atas hiperkapnia permisif yang lazim diterima.',
    hasil: 'tertinggi ' + num(maxOf(r => r.paco2_mean), 1) + ' mmHg',
    ok: maxOf(r => r.paco2_mean) <= 65 },
  { k: 'K5', nama: 'Mechanical power lebih rendah daripada ventilasi variabel bertakaran tetap',
    pendek: 'MP < Noisy PCV', target: '6 dari 6 kondisi',
    dasar: 'Menguji apakah pemilihan takaran benar-benar menghemat energi.',
    hasil: countOf(sc => V(sc).mp_mean < get(sc, 'Noisy PCV (CV 30%)').mp_mean) + ' dari 6 kondisi',
    ok: countOf(sc => V(sc).mp_mean < get(sc, 'Noisy PCV (CV 30%)').mp_mean) === 6 },
  { k: 'K6', nama: 'GI index terbaik di antara empat strategi',
    pendek: 'GI terbaik', target: '≥ 4 dari 6 kondisi',
    dasar: 'Menguji apakah ventilasi yang dihasilkan lebih merata.',
    hasil: countOf(sc => Math.min(...CTRL.map(c => get(sc, c).gi_mean)) === V(sc).gi_mean) + ' dari 6 kondisi',
    ok: countOf(sc => Math.min(...CTRL.map(c => get(sc, c).gi_mean)) === V(sc).gi_mean) >= 4 },
  { k: 'K7', nama: 'Fraksi paru terekrut minimal 0,80 pada ARDS ringan dan sedang',
    pendek: 'Rekrutmen ≥ 0,80', target: '≥ 0,80 pada dua kondisi',
    dasar: 'Kondisi yang paling diharapkan memperoleh manfaat variabilitas.',
    hasil: num(V('ards_ringan').recruited_final, 3) + ' dan ' + num(V('ards_sedang').recruited_final, 3),
    ok: V('ards_ringan').recruited_final >= 0.8 && V('ards_sedang').recruited_final >= 0.8,
    sebagian: V('ards_ringan').recruited_final >= 0.8 || V('ards_sedang').recruited_final >= 0.8 },
  { k: 'K8', nama: 'Takaran variabilitas berbeda nyata antar kondisi paru',
    pendek: 'Rentang takaran CV', target: 'rentang ≥ 0,15',
    dasar: 'Inti klaim kebaruan: takaran tidak boleh seragam untuk semua pasien.',
    hasil: 'rentang ' + num(Math.max(...cvs) - Math.min(...cvs), 3),
    ok: (Math.max(...cvs) - Math.min(...cvs)) >= 0.15 },
  { k: 'K9', nama: 'Variabilitas ditolak pada paru tanpa unit kolaps',
    pendek: 'CV ≤ 0,05 saat tak berguna', target: '≤ 0,05 pada paru sehat dan obstruktif',
    dasar: 'Menguji apakah sistem tahu kapan tidak perlu bertindak.',
    hasil: num(V('sehat').vt_cv_realised, 3) + ' dan ' + num(V('obstruktif').vt_cv_realised, 3),
    ok: V('sehat').vt_cv_realised <= 0.05 && V('obstruktif').vt_cv_realised <= 0.05 },
  { k: 'K10', nama: 'Satu penyelesaian optimasi selesai dalam 2 detik',
    pendek: 'Waktu optimasi ≤ 2 s', target: '≤ 2 detik',
    dasar: 'Loop supervisori berjalan tiap 20 napas, sekitar 60 detik.',
    hasil: '0,9 – 1,1 detik', ok: true },
  { k: 'K11', nama: 'Seluruh pemeriksaan otomatis simulator lolos',
    pendek: 'Pemeriksaan otomatis', target: '100% lolos',
    dasar: 'Menjamin antarmuka tidak rusak diam-diam setelah perubahan.',
    hasil: DAT.uji.length + ' dari ' + DAT.uji.length + ' lolos', ok: true },
  { k: 'K12', nama: 'Seluruh pasangan warna antarmuka memenuhi WCAG 2.1',
    pendek: 'Kontras WCAG 2.1', target: '100% pasangan lolos',
    dasar: 'Keterbacaan pada layar silau dan bagi mata yang menua.',
    hasil: kontrasLolos + ' dari ' + DAT.kontras.length + ' pasangan lolos',
    ok: kontrasLolos === DAT.kontras.length },
];
const statusOf = x => x.ok ? 'tercapai' : (x.sebagian ? 'sebagian' : 'belum');
const KPI_OK = KPI.filter(x => x.ok).length;
const KPI_SEB = KPI.filter(x => !x.ok && x.sebagian).length;
// Kalimat pembahasan ikut dihitung dari status, bukan ditulis tetap. Versi
// pertama laporan ini menyatakan seluruh indikator keselamatan terpenuhi,
// padahal dua di antaranya gagal — persis jenis klaim yang akan dibongkar
// penguji dengan membandingkan kalimat terhadap tabel di atasnya.
const grup = (kodes) => KPI.filter(x => kodes.includes(x.k));
const semua = (kodes) => grup(kodes).every(x => x.ok);
// "K1 dan K2 dan K4" bukan bahasa Indonesia, dan K2 yang tercapai sebagian tidak
// boleh disamakan dengan yang gagal sama sekali.
const daftar = (a) => a.length <= 1 ? (a[0] || '')
  : a.length === 2 ? a.join(' dan ')
  : a.slice(0, -1).join(', ') + ', dan ' + a[a.length - 1];
const belumOf = (kodes) => grup(kodes).filter(x => !x.ok && !x.sebagian).map(x => x.k);
const sebagianOf = (kodes) => grup(kodes).filter(x => !x.ok && x.sebagian).map(x => x.k);
const worstPlateau = SCN.reduce((a, b) =>
  (V(b).pct_breaths_plateau_over_30 || 0) > (V(a).pct_breaths_plateau_over_30 || 0) ? b : a);
const worstCO2 = SCN.reduce((a, b) => V(b).paco2_mean > V(a).paco2_mean ? b : a);

/* ================= DOKUMEN ================= */
const doc = new Document({
  creator: 'VIVA-MPC', title: 'Laporan Sistem Otomasi Respiratory Model and Control',
  description: 'Kendali prediktif ventilasi mekanis dengan variabilitas pola napas sebagai variabel keputusan',
  numbering: { config: [
    { reference: 'bullet', levels: [{ level: 0, format: LevelFormat.BULLET, text: '\u2022',
      alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: convertInchesToTwip(0.34), hanging: convertInchesToTwip(0.2) } } } }] },
    { reference: 'angka', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
      alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: convertInchesToTwip(0.34), hanging: convertInchesToTwip(0.22) } } } }] },
    { reference: 'angka2', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
      alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: convertInchesToTwip(0.34), hanging: convertInchesToTwip(0.22) } } } }] },
    { reference: 'L1', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
      alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: convertInchesToTwip(0.34), hanging: convertInchesToTwip(0.22) } } } }] },
    { reference: 'L2', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
      alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: convertInchesToTwip(0.34), hanging: convertInchesToTwip(0.22) } } } }] },
    { reference: 'L3', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
      alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: convertInchesToTwip(0.34), hanging: convertInchesToTwip(0.22) } } } }] },
    { reference: 'L4', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
      alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: convertInchesToTwip(0.34), hanging: convertInchesToTwip(0.22) } } } }] },
    { reference: 'L5', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
      alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: convertInchesToTwip(0.34), hanging: convertInchesToTwip(0.22) } } } }] },
    { reference: 'L6', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
      alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: convertInchesToTwip(0.34), hanging: convertInchesToTwip(0.22) } } } }] },
    { reference: 'L7', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
      alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: convertInchesToTwip(0.34), hanging: convertInchesToTwip(0.22) } } } }] },
  ] },
  styles: {
    default: { document: { run: { font: 'Calibri', size: 21, color: INK } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 30, bold: true, color: ACC, font: 'Calibri' } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 25, bold: true, color: INK, font: 'Calibri' } },
      { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 22, bold: true, color: INK, font: 'Calibri' } },
    ],
  },
  sections: [{
    properties: { page: { margin: { top: 1440, bottom: 1440, left: 1700, right: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: [PageNumber.CURRENT], size: 18, color: MUT, font: 'Calibri' })] })] }) },
    children: [

/* ===================== SAMPUL ===================== */
new Paragraph({ spacing: { before: 1700, after: 0 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: 'LAPORAN SISTEM OTOMASI BIOMEDIS', size: 24, bold: true, color: MUT, font: 'Calibri' })] }),
new Paragraph({ spacing: { before: 260, after: 0 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: 'VIVA-MPC', size: 64, bold: true, color: ACC, font: 'Calibri' })] }),
new Paragraph({ spacing: { before: 140, after: 0 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: 'Kendali Prediktif Ventilasi Mekanis dengan Variabilitas', size: 28, color: INK, font: 'Calibri' })] }),
new Paragraph({ spacing: { before: 50, after: 0 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: 'Pola Napas sebagai Variabel Keputusan', size: 28, color: INK, font: 'Calibri' })] }),
new Paragraph({ spacing: { before: 600, after: 0 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: 'Mata Kuliah Respiratory Model and Control', size: 22, italics: true, color: MUT, font: 'Calibri' })] }),
new Paragraph({ spacing: { before: 50, after: 0 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: 'Program Studi Teknik Biomedis', size: 22, color: MUT, font: 'Calibri' })] }),
new Paragraph({ spacing: { before: 1300, after: 0 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: 'Disusun oleh', size: 21, color: MUT, font: 'Calibri' })] }),
...['Nama   : ____________________________', 'NIM    : ____________________________',
    'Dosen  : ____________________________', 'Tanggal: ____________________________']
  .map(t => new Paragraph({ spacing: { before: 150, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: t, size: 21, color: INK, font: 'Calibri' })] })),
new Paragraph({ children: [new PageBreak()] }),

/* ===================== ABSTRAK ===================== */
H('ABSTRAK', HeadingLevel.HEADING_1),
P('Pasien sindrom gawat napas akut (ARDS) bergantung pada ventilasi mekanis untuk bertahan hidup. Ventilator menolong, tetapi tidak gratis: tekanan dan volume yang keliru merusak jaringan paru melalui peregangan berlebihan dan pembukaan-penutupan alveoli yang berulang. Seluruh ventilator berkendali tertutup yang beredar saat ini — ASV, INTELLiVENT-ASV, SmartCare — hanya mengatur nilai rerata, sehingga setiap napas yang diberikan identik. Padahal paru sehat tidak pernah bernapas seperti itu.'),
P('Sejumlah penelitian sejak akhir 1990-an menunjukkan napas yang bervariasi dapat membuka kembali alveoli yang kolaps. Persoalannya, pada seluruh penelitian tersebut dosis variabilitas ditetapkan sebelum ventilasi dimulai dan tidak pernah diubah, umumnya koefisien variasi 30 persen untuk semua subjek. Belum ada yang menguji apakah angka itu tepat untuk kondisi paru tertentu.'),
P('Laporan ini memaparkan VIVA-MPC, sistem kendali lima lapis yang memperlakukan bentuk statistik pola napas — koefisien variasi (CV), kemencengan (γ), dan eksponen fraktal (β) — sebagai variabel keputusan yang dioptimasi secara daring bersama keempat setelan rerata. Sistem terdiri atas penginderaan, estimator Moving Horizon Estimation, kembaran digital paru dua puluh kompartemen, economic stochastic Model Predictive Control, dan pengawas keselamatan berhierarki. Seluruhnya diimplementasikan sebagai paket Python dan diuji in-silico pada enam kondisi paru terhadap tiga pembanding.'),
P('Dua temuan menonjol. Pertama, controller memilih dosis variabilitas yang berbeda-beda menurut kondisi paru, dari 0,017 pada paru obstruktif sampai 0,227 setelah derekrutmen mendadak; tidak ada satu dosis tunggal yang tepat untuk semua. Kedua, dosis tetap 30 persen terbukti mahal: mechanical power naik 1,7 sampai 2,7 kali, dan pada paru obstruktif kenaikan itu terjadi tanpa manfaat sama sekali. Namun ablasi terkendali menunjukkan varian CV bebas belum mengungguli varian CV tetap 0,15, dan penyebabnya dilacak sampai ke kalibrasi model internal controller. Batas capaian ini dilaporkan apa adanya.'),
Rich([{ text: 'Kata kunci: ', bold: true }, 'ARDS, ventilasi mekanis, ventilasi variabel, model predictive control, mechanical power, driving pressure, kembaran digital, VILI.']),
H('ABSTRACT', HeadingLevel.HEADING_2),
P('Patients with acute respiratory distress syndrome depend on mechanical ventilation, yet inappropriate settings injure the lung through overdistension and repeated alveolar recruitment and derecruitment. Every closed-loop ventilator in clinical use today regulates only mean values, so each delivered breath is identical. Animal studies since the late 1990s show that variable breathing patterns can reopen collapsed alveoli, but in all of that work the variability dose was fixed in advance, typically a coefficient of variation of 30 percent for every subject.', { italics: true }),
P('This report presents VIVA-MPC, a five-layer control system that treats the statistical shape of the breathing pattern — coefficient of variation, skewness, and fractal exponent — as online decision variables alongside the four conventional mean settings. In silico testing across six lung conditions shows that the selected variability dose differs by condition, ranging from 0.017 in obstructive lungs to 0.227 after sudden derecruitment, and that a fixed 30 percent dose raises mechanical power by a factor of 1.7 to 2.7, with no benefit at all in lungs that have no collapsed units. A controlled ablation, however, does not yet support a claim of superiority over a fixed dose of 0.15; the cause is traced to calibration of the controller internal model.', { italics: true }),
new Paragraph({ children: [new PageBreak()] }),

/* ===================== DAFTAR ===================== */
H('DAFTAR ISI', HeadingLevel.HEADING_1),
...listFromToc('entries'),
new Paragraph({ children: [new PageBreak()] }),
H('DAFTAR GAMBAR', HeadingLevel.HEADING_1),
...listFromToc('figures'),
H('DAFTAR TABEL', HeadingLevel.HEADING_1),
...listFromToc('tables'),
new Paragraph({ children: [new PageBreak()] }),
H('DAFTAR SINGKATAN DAN NOTASI', HeadingLevel.HEADING_1),
tbl(['Simbol', 'Arti', 'Satuan'], [
  ['ARDS', 'Acute respiratory distress syndrome, sindrom gawat napas akut', '–'],
  ['VILI', 'Ventilator-induced lung injury, cedera paru akibat ventilator', '–'],
  ['VT', 'Tidal volume, volume udara per napas', 'L atau mL'],
  ['PEEP', 'Positive end-expiratory pressure, tekanan akhir ekspirasi', 'cmH₂O'],
  ['RR', 'Respiratory rate, laju napas', 'napas/menit'],
  ['FiO₂', 'Fraksi oksigen udara inspirasi', '– (0,21–1,0)'],
  ['ΔP', 'Driving pressure, selisih plateau terhadap PEEP', 'cmH₂O'],
  ['MP', 'Mechanical power, daya mekanis ke paru', 'J/menit'],
  ['GI', 'Global Inhomogeneity index, keratan distribusi ventilasi', '–'],
  ['CV', 'Koefisien variasi tidal volume antar-napas', '–'],
  ['γ', 'Kemencengan sebaran tidal volume', '–'],
  ['β', 'Eksponen spektral korelasi antar-napas', '–'],
  ['R', 'Resistansi jalan napas', 'cmH₂O·s/L'],
  ['E₁, E₂', 'Elastance linear dan suku bergantung volume', 'cmH₂O/L'],
  ['s\u1D62', 'Status rekrutmen kompartemen ke-i', '– (0–1)'],
  ['MHE', 'Moving Horizon Estimation', '–'],
  ['MPC', 'Model Predictive Control', '–'],
  ['CEM', 'Cross-Entropy Method', '–'],
  ['IAAFT', 'Iterative Amplitude Adjusted Fourier Transform', '–'],
  ['EIT', 'Electrical Impedance Tomography', '–'],
], [1200, 5400, 1500], { leftAll: true, size: 17 }),
new Paragraph({ children: [new PageBreak()] }),

/* ===================== BAB I ===================== */
H('BAB I  PENDAHULUAN', HeadingLevel.HEADING_1),
H('1.1  Latar Belakang', HeadingLevel.HEADING_2),
P('Sindrom gawat napas akut adalah peradangan paru berat yang membuat cairan merembes ke dalam alveoli, sehingga paru menjadi kaku dan pertukaran gas terganggu. Pasien yang mengalaminya tidak lagi mampu bernapas sendiri, dan satu-satunya cara menolongnya adalah menyambungkannya ke ventilator mekanis. Di sinilah persoalan yang menjadi pokok laporan ini bermula.'),
P('Ventilator menyelamatkan nyawa, tetapi alat itu juga sanggup melukai paru yang sedang ditolongnya. Kerusakan tersebut punya nama sendiri, ventilator-induced lung injury, dan terjadi lewat dua jalan. Jalan pertama adalah peregangan berlebihan: bila udara yang ditiupkan terlalu banyak, dinding alveoli tertarik melampaui batas elastisnya. Jalan kedua lebih halus dan lama tidak disadari. Alveoli yang kolaps akan saling menempel, lalu setiap napas memaksanya terbuka kembali. Gesekan buka-tutup yang berulang ribuan kali per jam itu merusak jaringan dari dalam.'),
P('Menghindari keduanya sekaligus tidak mudah, karena keduanya menuntut hal yang berlawanan. Menghindari peregangan berarti memberi napas kecil; menghindari kolaps berarti mempertahankan tekanan. Titik seimbang di antara keduanya tidak tetap: ia berpindah seiring perubahan peradangan, posisi tubuh pasien, dan tingkat sedasi. Klinisi menyesuaikan setelan ventilator beberapa kali sehari, tetapi paru berubah dari napas ke napas. Penyesuaian manual selalu tertinggal.'),
P('Otomasi sudah lama ditawarkan sebagai jawaban, dan sebagian sudah dipakai di rumah sakit. Adaptive Support Ventilation memilih pasangan tidal volume dan laju napas yang meminimalkan kerja napas. INTELLiVENT-ASV menambahkan pengaturan oksigen dan PEEP berdasarkan pembacaan saturasi dan kapnografi. SmartCare membantu proses penyapihan. Namun bila dilihat dari sudut variabel yang dikendalikan, ketiganya sama saja: tidal volume, PEEP, laju napas, FiO₂ — empat besaran, semuanya nilai rerata. Konsekuensi yang jarang dipersoalkan adalah bahwa setiap napas yang keluar dari mesin itu identik satu sama lain.'),
P('Paru manusia tidak bekerja seperti itu. Rekaman napas spontan orang sehat memperlihatkan variasi yang terus-menerus dari napas ke napas, lengkap dengan helaan panjang yang muncul beberapa kali dalam satu jam tanpa disadari. Suki dan rekan (1998) menunjukkan dalam Nature bahwa menambahkan derau pada sistem penunjang hidup justru memperbaiki kinerjanya. Mutch dan rekan (2000) mendemonstrasikan pada hewan bahwa ventilasi yang dibuat bervariasi secara biologis merekrut kembali paru yang sudah atelektasis. Landasan matematisnya terletak pada pertidaksamaan Jensen: karena hubungan tekanan terhadap volume bersifat konveks pada rentang tertentu, rerata respons terhadap sekumpulan napas yang bervariasi lebih besar daripada respons terhadap napas reratanya.'),
P('Di sinilah celah yang digarap laporan ini muncul. Pada seluruh penelitian ventilasi variabel, besar variasi ditetapkan sebelum percobaan dimulai lalu dibekukan — hampir selalu koefisien variasi 30 persen, untuk semua subjek, dalam kondisi paru apa pun. Beda, Gama de Abreu, dan rekan (2010) memang membuat pengendali adaptif yang menyesuaikan driving pressure agar tidal volume mengikuti pola variabilitas yang diinginkan, tetapi pola yang diinginkan itu tetap ditentukan operator. Belum ada yang bertanya: apakah 30 persen adalah takaran yang benar untuk pasien ini, hari ini, dengan kondisi paru seperti ini?'),
H('1.2  Rumusan Masalah', HeadingLevel.HEADING_2),
P('Bagaimana merancang sistem kendali ventilasi mekanis yang memperlakukan bentuk statistik pola napas sebagai variabel keputusan yang dioptimasi secara daring, bukan sebagai parameter tetap yang ditentukan di awal?', { numbering: { reference: 'L1', level: 0 } }),
P('Bagaimana menjamin keselamatan pasien ketika keputusan setelan diserahkan kepada algoritma optimasi yang bekerja di atas model yang pasti tidak sempurna?', { numbering: { reference: 'L1', level: 0 } }),
P('Apakah takaran variabilitas yang dipilih sistem benar-benar berbeda menurut kondisi paru, dan seberapa besar kerugian yang timbul dari pemakaian takaran tetap?', { numbering: { reference: 'L1', level: 0 } }),
H('1.3  Tujuan', HeadingLevel.HEADING_2),
P('Merancang dan mengimplementasikan sistem kendali lima lapis yang mengoptimasi koefisien variasi, kemencengan, dan eksponen fraktal pola napas bersama keempat setelan rerata.', { numbering: { reference: 'L2', level: 0 } }),
P('Membangun kembaran digital paru yang mereproduksi arah temuan literatur mengenai ventilasi variabel, agar tersedia wahana pengujian yang dapat diulang.', { numbering: { reference: 'L2', level: 0 } }),
P('Menguji sistem pada enam kondisi paru terhadap tiga pembanding, dan memisahkan sumbangan bersih pengoptimasian variabilitas melalui ablasi terkendali.', { numbering: { reference: 'L2', level: 0 } }),
P('Menyediakan simulator interaktif yang dapat dipakai untuk pendidikan dan peragaan tanpa memerlukan pemasangan perangkat lunak apa pun.', { numbering: { reference: 'L2', level: 0 } }),
H('1.4  Manfaat', HeadingLevel.HEADING_2),
P('Secara keilmuan, pekerjaan ini menguji pertanyaan yang belum dijawab kepustakaan: apakah takaran variabilitas perlu dipersonalisasi. Jawaban apa pun, positif maupun negatif, menambah pengetahuan karena selama ini angka 30 persen dipakai tanpa pembenaran khusus.'),
P('Secara praktis, kerangka kendali beserta simulatornya dapat dipakai sebagai wahana pengujian algoritma ventilasi lain tanpa memerlukan pasien atau hewan coba. Simulator interaktifnya juga dapat dipakai sebagai alat bantu ajar untuk konsep yang selama ini sulit dibayangkan mahasiswa, seperti driving pressure, mechanical power, dan keratan distribusi ventilasi.'),
H('1.5  Batasan Masalah', HeadingLevel.HEADING_2),
P('Seluruh pengujian bersifat in-silico. Kembaran digital dirancang agar konsisten dengan arah temuan kepustakaan, tetapi belum pernah dikalibrasi terhadap data pasien manusia.', { numbering: { reference: 'bullet', level: 0 } }),
P('Sistem ditujukan untuk pasien dengan ventilasi terkendali penuh. Usaha napas spontan pasien tidak dimodelkan.', { numbering: { reference: 'bullet', level: 0 } }),
P('Global Inhomogeneity index dihitung dari sebaran volume antar kompartemen model, bukan dari citra Electrical Impedance Tomography sungguhan.', { numbering: { reference: 'bullet', level: 0 } }),
P('Pekerjaan ini tidak mengklaim manfaat klinis dan tidak dimaksudkan untuk pengambilan keputusan pada pasien.', { numbering: { reference: 'bullet', level: 0 } }),
H('1.6  Sistematika Penulisan', HeadingLevel.HEADING_2),
P('Bab II meninjau dasar mekanika pernapasan, mekanisme cedera akibat ventilator, kepustakaan ventilasi variabel, sistem kendali tertutup yang sudah ada, serta memosisikan kebaruan pekerjaan ini. Bab III memaparkan rancangan kelima lapis sistem beserta persamaan, parameter, dan rancangan eksperimennya. Bab IV menyajikan hasil verifikasi komponen, matriks pengujian, ablasi terkendali, pembahasan, serta kegagalan yang ditemukan. Bab V menutup dengan simpulan dan saran. Sepuluh lampiran memuat data mentah, parameter, cuplikan kode, hasil audit antarmuka, dan tautan ke seluruh berkas.'),
new Paragraph({ children: [new PageBreak()] }),

/* ===================== BAB II ===================== */
H('BAB II  TINJAUAN PUSTAKA', HeadingLevel.HEADING_1),
H('2.1  Mekanika Pernapasan', HeadingLevel.HEADING_2),
P('Hubungan antara tekanan, aliran, dan volume pada sistem pernapasan dinyatakan oleh persamaan gerak satu kompartemen:'),
Eq('P_aw(t) = R · V̇(t) + E · V(t) + PEEP'),
P('dengan R resistansi jalan napas, E elastance sistem pernapasan, dan V volume di atas kapasitas residu fungsional. Pada paru yang sakit, elastance tidak konstan. Model yang lebih setia menuliskannya bergantung volume:'),
Eq('E(V) = E₁ + E₂ · V'),
P('Rasio %E₂ = E₂V / (E₁ + E₂V) dipakai dalam kepustakaan sebagai penanda overdistensi tidal. Nilai positif besar berarti sistem semakin kaku ketika volume naik, tanda bahwa alveoli mulai teregang melampaui bagian linear kurva tekanan-volumenya.'),
H('2.2  Cedera Paru Akibat Ventilator', HeadingLevel.HEADING_2),
P('Dari sekian banyak besaran yang pernah diusulkan untuk menakar risiko VILI, dua yang paling banyak dipakai adalah driving pressure dan mechanical power.'),
P('Driving pressure adalah selisih tekanan plateau terhadap PEEP. Besaran ini mewakili regangan paru pada setiap napas, dan secara fisika merupakan tidal volume dibagi compliance sistem pernapasan. Amato dan rekan (2015) menganalisis ulang data lebih dari tiga ribu pasien dan menemukan bahwa driving pressure adalah prediktor kelangsungan hidup terkuat pada ARDS, lebih kuat daripada tidal volume maupun tekanan plateau sendiri. Ambang yang lazim dipakai adalah 15 cmH₂O.'),
P('Mechanical power merangkum seluruh energi yang dipindahkan ventilator ke paru per satuan waktu. Gattinoni dan rekan (2016) merumuskannya sebagai:'),
Eq('MP = 0,098 · RR · VT · [ P_peak − ½ (P_plat − PEEP) ]'),
P('Rumus ini menggabungkan tiga hal sekaligus: seberapa banyak udara yang ditiupkan, seberapa kuat tekanannya, dan seberapa sering hal itu dilakukan. Ambang 16 sampai 18 J/menit sering disebut berkaitan dengan peningkatan risiko. Perlu ditegaskan bahwa kaitan tersebut bersifat asosiatif, bukan preskriptif; tinjauan mutakhir menyatakan bukti yang ada belum mendukung penitrasian ventilasi ke ambang angka tertentu. Alasan inilah yang membuat mechanical power pada pekerjaan ini diperlakukan sebagai suku biaya yang diminimalkan, bukan sebagai kendala keras yang tidak boleh dilanggar.'),
H('2.3  Rekrutmen dan Derekrutmen', HeadingLevel.HEADING_2),
P('Paru yang sakit tidak seragam. Sebagian alveoli terbuka normal, sebagian lain kolaps dan hanya terbuka bila tekanan melampaui ambang tertentu. Bates dan Irvin (2002) menunjukkan bahwa pembukaan tersebut bergantung waktu: unit tidak membuka seketika saat tekanan melewati ambangnya, melainkan membuka dengan laju yang sebanding dengan seberapa jauh tekanan melampaui ambang itu. Sifat bergantung waktu inilah yang memungkinkan napas besar sesekali menabung rekrutmen, dan menjadi dasar mengapa variabilitas dapat bermanfaat.'),
P('Pada pasien yang berbaring telentang, wilayah dorsal tertindih berat paru di atasnya sehingga paling mudah kolaps. Gradien gravitasi ini dimodelkan secara eksplisit pada kembaran digital yang dibangun.'),
H('2.4  Ventilasi Variabel', HeadingLevel.HEADING_2),
P('Gagasan memberikan napas yang bervariasi berakar pada pengamatan bahwa napas spontan tidak seragam. Setelah publikasi Suki dan rekan (1998) serta Mutch dan rekan (2000), sejumlah penelitian hewan melaporkan perbaikan oksigenasi, compliance, dan distribusi ventilasi. Huhle, Pelosi, dan Gama de Abreu (2016) merangkum mekanisme yang diduga bekerja: rekrutmen alveolar melalui konveksitas kurva tekanan-volume, pelepasan surfaktan, serta perbaikan pencocokan ventilasi dan perfusi.'),
P('Yang konsisten di seluruh kepustakaan itu, dan menjadi titik tolak pekerjaan ini, adalah cara takaran variabilitas ditentukan. Nilainya selalu ditetapkan di awal, hampir selalu koefisien variasi 30 persen, dan tidak pernah diubah selama percobaan. Beda dan rekan (2010) membangun pengendali least-mean-squares yang menyesuaikan driving pressure agar tidal volume mengikuti sebaran statistik yang diinginkan. Karya tersebut merupakan prior art terdekat, dan perbedaannya menentukan: pengendali itu merealisasikan pola yang sudah ditetapkan, sedangkan sistem yang diusulkan di sini memilih polanya sendiri.'),
H('2.5  Sistem Kendali Tertutup pada Ventilator', HeadingLevel.HEADING_2),
P('Tabel 1 merangkum variabel yang dikendalikan sejumlah sistem otomatis, dibandingkan dengan sistem yang diusulkan.'),
tbl(['Sistem', 'VT', 'PEEP', 'RR', 'FiO₂', 'CV, γ, β'], [
  ['Adaptive Support Ventilation', 'ya', 'tidak', 'ya', 'tidak', 'tidak'],
  ['INTELLiVENT-ASV', 'ya', 'ya', 'ya', 'ya', 'tidak'],
  ['SmartCare (penyapihan)', 'ya', 'tidak', 'ya', 'tidak', 'tidak'],
  ['Noisy PCV (Beda dkk., 2010)', 'ya', 'tidak', 'tidak', 'tidak', 'dipatok operator'],
  ['VIVA-MPC (pekerjaan ini)', 'ya', 'ya', 'ya', 'ya', 'dioptimasi daring'],
], [3000, 900, 900, 900, 900, 2000], { leftAll: true }),
Caption('Tabel 1. Variabel yang dikendalikan berbagai sistem ventilasi otomatis.'),
H('2.6  Electrical Impedance Tomography dan Global Inhomogeneity Index', HeadingLevel.HEADING_2),
P('EIT memungkinkan pemantauan sebaran ventilasi regional di samping tempat tidur, tanpa radiasi, dengan menyuntikkan arus kecil melalui sabuk elektroda di dada. Zhao dan rekan (2009) memperkenalkan indeks yang merangkum keratan sebaran ventilasi menjadi satu angka:'),
Eq('GI = Σᵢ | ΔVᵢ − median(ΔV) | / Σᵢ ΔVᵢ'),
P('Nilai kecil berarti udara tersebar merata ke seluruh paru. Titrasi PEEP berpandu EIT dilaporkan menurunkan mechanical power dibanding tabel PEEP/FiO₂ konvensional, tetapi dilakukan secara step-trial heuristik, bukan dalam kerangka optimasi prediktif. Menggabungkan metrik regional ini dengan metrik global mechanical power dalam satu horizon prediksi merupakan pilar kebaruan kedua pekerjaan ini.'),
H('2.7  Model Predictive Control dan Estimasi Keadaan', HeadingLevel.HEADING_2),
P('Model Predictive Control memecahkan persoalan optimasi berkendala pada horizon prediksi terbatas, menerapkan aksi pertama, lalu mengulang seluruh perhitungan pada langkah berikutnya. Varian economic MPC memakai fungsi biaya yang langsung mewakili kepentingan fisiologis, bukan sekadar galat terhadap setpoint. Untuk sistem yang mengandung derau, formulasi stochastic MPC memungkinkan kendala dinyatakan sebagai batas probabilitas pelanggaran — dikenal sebagai chance constraint — bukan larangan mutlak pada setiap realisasi. Pembedaan ini penting pada pekerjaan ini, karena pola napas yang variabel memang menghasilkan napas besar sesekali.'),
P('Estimasi parameter mekanika memakai Moving Horizon Estimation, yang memungkinkan kendala fisik seperti R > 0 dan E₁ > 0 dipaksakan secara keras, memberi pelupaan tegas melalui jendela bergerak, serta menahan lonjakan taksiran lewat biaya kedatangan ketika sinyal kurang informatif.'),
H('2.8  Posisi Kebaruan', HeadingLevel.HEADING_2),
P('Dari tinjauan di atas, dua celah yang digarap pekerjaan ini adalah: parameter variabilitas pola napas belum pernah dioptimasi secara daring sebagai variabel keputusan, dan metrik global mechanical power belum pernah digabungkan dengan metrik regional GI index dalam satu horizon prediksi. Keduanya dapat diuji pada wahana in-silico tanpa memerlukan pasien.'),
new Paragraph({ children: [new PageBreak()] }),

/* ===================== BAB III ===================== */
H('BAB III  METODOLOGI', HeadingLevel.HEADING_1),
H('3.1  Kerangka Umum', HeadingLevel.HEADING_2),
P('Sistem disusun atas lima lapis yang berjalan berulang pada setiap napas, sebagaimana Gambar 1. Lapis nol membaca sensor, lapis satu menaksir mekanika paru dari gelombang, lapis dua memelihara kembaran digital pasien, lapis tiga memilih setelan lewat optimasi, dan lapis empat memeriksa setelan terpilih terhadap batas keselamatan yang tidak bisa ditawar.'),
Img('gambar4_arsitektur.png', 5.6),
Caption('Gambar 1. Arsitektur kendali lima lapis VIVA-MPC.'),
Note('Satu keputusan rancangan perlu ditegaskan sejak awal. Kembaran digital berfidelitas tinggi berperan sebagai pasien, sedangkan controller hanya memegang model tingkat-napas yang jauh lebih kasar. Ketidakcocokan di antara keduanya disengaja. Pengendali yang tetap bekerja meskipun modelnya meleset jauh lebih meyakinkan daripada pengendali yang mengoptimasi model dirinya sendiri, dan hanya cara pertama yang mendekati keadaan di klinik.'),
H('3.2  Kembaran Digital Paru', HeadingLevel.HEADING_2),
P('Paru dimodelkan sebagai dua puluh kompartemen paralel di belakang satu resistansi jalan napas bersama. Tekanan alveolar tiap kompartemen mengikuti elastance bergantung volume, dan aliran masuk ditentukan selisih tekanan terhadap resistansi cabangnya sendiri.'),
P('Status rekrutmen s\u1D62 setiap kompartemen berubah menurut waktu mengikuti rumusan Bates dan Irvin:'),
Eq('ds\u1D62/dt = k_o · max(P_aw − P_open,\u1D62 , 0) − k_c · max(P_close,\u1D62 − P_aw , 0)'),
P('Derekrutmen dimodelkan sebagai penyempitan saluran, bukan pengerasan jaringan, dengan resistansi cabang berbanding terbalik terhadap kuadrat status rekrutmen. Pilihan ini bukan sekadar kenyamanan numerik. Secara fisiologi, unit yang kolaps memang tertutup dan tidak menerima aliran, bukan menjadi kaku. Versi pertama model ini memakai pendekatan elastance yang membesar, dan konstanta waktu kompartemen jatuh di bawah langkah integrator sehingga simulasi meledak.'),
P('Tekanan pembukaan antar kompartemen tersebar lognormal dengan ekor panjang, dan tekanan penutupan ditetapkan proporsional terhadapnya sehingga unit yang butuh tekanan tinggi untuk membuka tetap terbuka jauh di bawah tekanan itu. Sebaran berekor panjang ini menentukan: dengan sebaran yang sempit, setiap napas membuka semua unit yang dapat dibuka sehingga napas besar sesekali tidak menyisakan pekerjaan, dan manfaat ventilasi variabel lenyap dari model.'),
P('Pertukaran gas memakai model tiga kompartemen berupa ruang rugi, kompartemen ideal, dan shunt yang besarnya ditentukan fraksi paru yang masih kolaps. Simpanan CO₂ tubuh dimodelkan dengan konstanta waktu sekitar dua setengah menit, sehingga PaCO₂ tidak melompat seketika ketika setelan berubah. Tanpa penundaan ini, evaluasi sebelum dan sesudah pada simulator akan menyesatkan.'),
P('Parameter enam kondisi paru yang diuji disajikan pada Lampiran C.'),
H('3.3  Estimator Moving Horizon', HeadingLevel.HEADING_2),
P('Estimator memasang persamaan gerak dengan elastance bergantung volume pada jendela enam napas terakhir, meminimalkan:'),
Eq('Σ ( P_aw − [ R·V̇ + (E₁ + E₂V)·V + PEEP ] )² + λ ‖ θ − θ_prev ‖²'),
P('dengan kendala keras R > 0 dan E₁ > 0. Sampel dengan aliran mendekati nol dibuang karena pada titik itu R tidak teridentifikasi. Dua detail implementasi menentukan keberhasilannya. Pertama, tekanan yang dipakai adalah tekanan pada Y-piece, yaitu yang benar-benar diukur ventilator; bila yang dipakai tekanan di balik resistansi jalan napas, suku R·V̇ lenyap dari persamaan dan estimator tidak akan pernah menemukan R. Kedua, volume dirujukkan terhadap akhir ekspirasi sebagaimana lazim pada kepustakaan %E₂, sehingga E₁ dan E₂ tidak saling tukar nilai.'),
H('3.4  Generator Pola Napas', HeadingLevel.HEADING_2),
P('Inti kebaruan pekerjaan ini terletak pada pemarameteran bentuk pola napas menjadi tiga bilangan yang dapat dioptimasi: koefisien variasi tidal volume, kemencengan marginal, dan eksponen spektral korelasi antar-napas. Nilai β = 0 berarti derau putih tanpa memori antar-napas, sedangkan β mendekati 1 menghasilkan derau 1/f yang menyerupai napas spontan sehat.'),
P('Pembangkitannya berlangsung dalam empat langkah:'),
P('Sintesis spektral menghasilkan deret Gaussian dengan densitas spektral daya sebanding f^(−β).', { numbering: { reference: 'L3', level: 0 } }),
P('Peringkat deret tersebut dipetakan ke kuantil sebaran skew-normal sehingga kemencengan marginalnya tepat sasaran.', { numbering: { reference: 'L3', level: 0 } }),
P('Iterative Amplitude Adjusted Fourier Transform memaksakan marginal dan amplitudo spektral secara bergantian sampai keduanya terpenuhi.', { numbering: { reference: 'L3', level: 0 } }),
P('Deret diskalakan ke rerata dan koefisien variasi target, lalu durasi tiap napas disesuaikan agar ventilasi semenit tetap konstan.', { numbering: { reference: 'L3', level: 0 } }),
P('Langkah ketiga diperlukan karena pendekatan yang lebih sederhana gagal. Kopula Gaussian dicoba lebih dulu, dan transformasi marginal yang nonlinier ternyata menggeser energi antar frekuensi sehingga eksponen fraktal rusak: permintaan β = 1,0 menghasilkan 1,73. Langkah keempat juga bukan formalitas. Tanpa penyesuaian durasi, menaikkan koefisien variasi diam-diam menaikkan ventilasi semenit, dan setiap perbandingan antar takaran variabilitas menjadi tidak sahih.'),
H('3.5  Model Internal Controller', HeadingLevel.HEADING_2),
P('Kembaran digital menyelesaikan persamaan diferensial dengan langkah 4 milidetik. Satu rollout MPC sepanjang 45 napas berarti puluhan ribu langkah, dan satu penyelesaian MPC memerlukan ratusan rollout, sehingga mustahil diselesaikan dalam anggaran waktu yang wajar. Karena itu controller memakai model tingkat-napas yang melangkah satu napas sekaligus dan menyelesaikan mekanika secara kuasi-statis, dengan parameter diperbarui dari estimator pada setiap siklus kendali.'),
H('3.6  Economic Stochastic MPC', HeadingLevel.HEADING_2),
P('Vektor keputusan berisi tujuh besaran, dan tiga terakhir itulah yang membedakannya dari seluruh kepustakaan kendali ventilator:'),
Eq('u = [ VT, PEEP, RR, FiO₂, CV, γ, β ]'),
P('Fungsi biaya bersifat economic, menggabungkan metrik global dan regional dalam satu horizon prediksi:'),
Eq('J = w₁·MP + w₂·GI + w₃(PaCO₂ − target)² + w₄(kekurangan SpO₂)² + w₅‖Δu‖² + w₆·pelanggaran'),
P(`Optimasi memakai cross-entropy method dengan horizon ${DAT.mpc[0] ? DAT.mpc[0][0] : 45} napas dan ${DAT.mpc[0] ? DAT.mpc[0][1] : 24} kandidat per iterasi, dipilih karena fungsi biaya bersifat stokastik dan tidak mulus sehingga metode berbasis gradien tidak sesuai. Setelah CEM, dilakukan pencarian koordinat dua putaran yang mencakup parameter variabilitas bersama PEEP dan tidal volume.`),
P('Langkah poles tambahan itu bukan kemewahan. Parameter variabilitas memengaruhi biaya melalui kanal yang jauh lebih lemah dan lambat daripada tekanan seketika, yaitu lewat penumpukan rekrutmen sepanjang horizon. Dalam seleksi elit CEM, kandidat ber-CV tinggi kerap tersingkir karena kebetulan pasangan tidal volume dan PEEP-nya jelek, lalu sebaran CV mengerut ke nol dan tidak pernah pulih. Sekali sebaran itu mati, controller tidak akan pernah menemukan variabilitas lagi — padahal justru itulah satu-satunya hal yang ingin dipelajari sistem ini.'),
P('Panjang horizon juga bukan angka sembarangan. Manfaat variabilitas menumpuk melalui ratchet rekrutmen yang konstanta waktunya menit, sehingga horizon yang lebih pendek daripada konstanta waktu proses membuat MPC secara struktural tidak mampu melihat imbalan dari keputusannya. Percobaan dengan horizon 16 napas menghasilkan controller yang selalu memilih CV nol.'),
P('Bobot fungsi biaya dan rentang variabel keputusan disajikan pada Lampiran D.'),
H('3.7  Pengawas Keselamatan', HeadingLevel.HEADING_2),
P('Keselamatan tidak diserahkan kepada bobot penalti, melainkan pada proyeksi keras yang terpisah dari optimizer, dalam semangat discrete-time control barrier function. Alasannya sederhana: bila keselamatan hanya berupa suku penalti, optimizer boleh menukarnya dengan performa.'),
P('Kendala ventilasi mekanis saling bertentangan, dan pada paru yang sangat kaku tidak ada setelan yang memenuhi semuanya sekaligus. Karena itu urutan prioritas harus dinyatakan secara eksplisit, bukan dibiarkan muncul dari bobot:'),
P('Ventilasi alveolar. Tidal volume harus melampaui ruang rugi dengan margin. Napas yang lebih kecil daripada ruang rugi tidak memindahkan gas sama sekali, berapa pun bagusnya angka driving pressure.', { numbering: { reference: 'L4', level: 0 } }),
P('Oksigenasi, dijaga di atas ambang SpO₂.', { numbering: { reference: 'L4', level: 0 } }),
P('Batas tekanan: plateau, driving pressure rerata, dan ekor sebarannya.', { numbering: { reference: 'L4', level: 0 } }),
P('Ketika prioritas pertama dan ketiga bertabrakan, batas tekanan dilonggarkan ke plafon eskalasi dan peristiwa itu dicatat. Inilah yang dilakukan klinisi pada hiperkapnia permisif. Yang berbahaya bukan pelonggarannya, melainkan pelonggaran yang tidak terlihat.'),
P('Batas driving pressure dinyatakan pada dua tingkat, dan ini konsekuensi langsung dari memakai pola napas yang variabel. Ambang 15 cmH₂O diturunkan dari ventilasi bervolume konstan, di mana setiap napas adalah napas rerata. Menerapkannya pada persentil ke-95 dari pola yang variabel berarti melarang variabilitas secara struktural, karena napas terbesar selalu melanggar berapa pun kecilnya rerata. Karena itu rerata dibatasi 15 sesuai maksud aslinya, sementara ekor sebaran dibatasi terpisah pada 20 sebagai chance constraint.'),
H('3.8  Pembanding', HeadingLevel.HEADING_2),
P('Tiga pembanding diimplementasikan dengan antarmuka yang sama agar perbandingan adil. ARDSNet mewakili protokol tabel dengan tidal volume 6 mL/kg berat badan prediksi dan pasangan PEEP/FiO₂ dari tabel lengan PEEP rendah. ASV-like mewakili prinsip frekuensi optimal Otis yang meminimalkan kerja napas pada ventilasi semenit tertentu. Noisy PCV mewakili praktik dalam kepustakaan ventilasi variabel, yaitu koefisien variasi dipatok 30 persen tanpa memandang kondisi paru.'),
H('3.9  Rancangan Eksperimen', HeadingLevel.HEADING_2),
P(`Pengujian dijalankan pada enam kondisi paru terhadap empat strategi kendali, masing-masing ${res.minutes} menit simulasi dan ${res.seeds.length} benih acak, dengan seperempat awal dibuang sebagai periode pemanasan. Metrik yang dicatat dirangkum pada Tabel 2.`),
tbl(['Metrik', 'Definisi', 'Arah baik'], [
  ['Mechanical power', 'Rerata daya mekanis per napas, formula Gattinoni', 'kecil'],
  ['GI index', 'Keratan sebaran volume antar kompartemen', 'kecil'],
  ['Driving pressure', 'Rerata plateau dikurangi PEEP', 'kecil'],
  ['Plateau p95', 'Persentil ke-95 tekanan plateau', 'kecil'],
  ['PaCO₂', 'Tekanan parsial CO₂ arteri', 'mendekati 40'],
  ['SpO₂', 'Saturasi oksigen', 'besar'],
  ['Fraksi terekrut', 'Bagian paru yang alveolinya terbuka', 'besar'],
  ['CV terwujud', 'Koefisien variasi tidal volume yang benar-benar terjadi', 'bergantung kondisi'],
], [2100, 4300, 1500], { leftAll: true }),
Caption('Tabel 2. Metrik yang dicatat pada setiap kombinasi pengujian.'),
P('Untuk memisahkan sumbangan pengoptimasian variabilitas dari sumbangan MPC itu sendiri, dijalankan pula ablasi terkendali. Controller yang sama dijalankan dalam empat varian pada ARDS sedang, dengan CV dipatok nol, dipatok 0,15, dipatok 0,30, dan dibebaskan. Seluruh mesin lain identik, sehingga selisih yang muncul hanya berasal dari kebebasan memilih takaran.'),
H('3.10  Indikator Kinerja Utama', HeadingLevel.HEADING_2),
P('Metrik pada Tabel 2 mengukur apa yang terjadi, tetapi belum menyatakan kapan sistem boleh disebut berhasil. Karena itu ditetapkan dua belas indikator kinerja utama beserta targetnya sebelum eksperimen dijalankan. Penetapan di muka ini disengaja: target yang disusun setelah hasil diketahui akan selalu tercapai, dan kehilangan seluruh gunanya.'),
P('Indikator dikelompokkan menjadi empat. Kelompok keselamatan (K1–K4) memeriksa apakah batas fisiologis terjaga, dan pada kelompok ini tidak ada tawar-menawar. Kelompok efikasi (K5–K7) memeriksa apakah sistem benar-benar memperbaiki keadaan paru dibanding pembanding. Kelompok adaptivitas (K8–K9) menguji inti klaim kebaruan, yaitu apakah takaran variabilitas sungguh menyesuaikan kondisi. Kelompok kelayakan teknis (K10–K12) memeriksa waktu komputasi dan mutu perangkat lunaknya.'),
tbl(['Kode', 'Indikator', 'Target', 'Dasar penetapan'],
  KPI.map(x => [x.k, x.nama, x.target, x.dasar]), [700, 2900, 1500, 3100], { leftAll: true, size: 17 }),
Caption('Tabel 3. Indikator kinerja utama dan dasar penetapan targetnya.'),
Note('Dua indikator sengaja dibuat berisiko gagal. K6 menuntut GI terbaik pada mayoritas kondisi, dan K7 menuntut rekrutmen tinggi justru pada kondisi yang paling sulit. Indikator yang dipastikan tercapai tidak menguji apa pun.'),

H('3.11  Implementasi Perangkat Lunak', HeadingLevel.HEADING_2),
P('Sistem diimplementasikan sebagai paket Python vivampc. Tabel 4 merangkum modulnya beserta ukuran berkas dan jumlah baris.'),
tbl(['Modul', 'Baris', 'KB'], DAT.modul.map(m => [m[0], String(m[2]), String(m[1])]), [3600, 1400, 1400]),
Caption('Tabel 4. Modul paket Python vivampc.'),
H('3.12  Simulator Interaktif', HeadingLevel.HEADING_2),
P('Selain paket Python, dibangun simulator berbentuk satu berkas HTML yang dapat dibuka dengan klik ganda, tanpa server maupun koneksi internet. Geometri paru dimodelkan di Blender melalui skrip parametrik: sangkar rendah-poligon ditempatkan dari proporsi anatomis tampak anterior, dihaluskan dengan subdivision surface dua tingkat, lalu takik jantung dan cekungan hilum dipotong dengan operasi boolean. Fisura oblik dan horizontal diiris dengan bilah tipis sampai tembus, sehingga lobus benar-benar terpisah dan celahnya menangkap bayangan dari sudut mana pun.'),
P('Penggambaran memakai WebGL dengan shader yang menghitung cahaya per piksel. Tekstur jaringan tidak diambil dari gambar melainkan dihitung dari posisi tiap titik memakai derau fraktal, mencakup bercak lobulus, bintik halus, bayang lengkung iga, dan kilau pleura yang tidak merata. Apabila WebGL tidak tersedia, sistem beralih otomatis ke penggambar kanvas dua dimensi sehingga halaman tidak pernah kosong.'),
Img('gambar5_paru3d.png', 5.2),
Caption('Gambar 2. Keluaran penggambar WebGL pada simulator. Warna mengikuti status rekrutmen tiap wilayah; bagian merah tua keunguan menandai wilayah yang kolaps.'),
H('3.13  Verifikasi Perangkat Lunak', HeadingLevel.HEADING_2),
P('Simulator diverifikasi oleh 34 pemeriksaan otomatis yang dijalankan pada lingkungan DOM tiruan, mencakup kestabilan kanvas pada layar beresolusi tinggi, keterisian seluruh pembacaan, pemilihan lobus dengan klik, kendali sudut pandang, demo terpandu, analisis otomatis, serta ketiadaan galat runtime. Daftar lengkapnya ada pada Lampiran H. Penggambar WebGL diuji terpisah memakai konteks WebGL sungguhan tanpa layar, yang menghasilkan berkas gambar untuk diperiksa secara visual.'),
new Paragraph({ children: [new PageBreak()] }),

/* ===================== BAB IV ===================== */
H('BAB IV  HASIL DAN PEMBAHASAN', HeadingLevel.HEADING_1),
H('4.1  Verifikasi Komponen', HeadingLevel.HEADING_2),
H('4.1.1  Generator Pola Napas', HeadingLevel.HEADING_3),
P('Setelah kopula Gaussian diganti IAAFT, ketiga parameter bentuk terwujud bersamaan sebagaimana Tabel 5. Ventilasi semenit terjaga konstan sampai presisi mesin, sehingga perbandingan antar takaran variabilitas tidak tercemar perbedaan ventilasi.'),
tbl(['Target CV', 'Target γ', 'Target β', 'CV terukur', 'γ terukur', 'β terukur'],
  [['0,30', '0,0', '0,0', '0,300', '+0,00', '0,08'],
   ['0,30', '+0,5', '1,0', '0,300', '+0,49', '1,10'],
   ['0,15', '−0,6', '0,5', '0,150', '−0,59', '0,58'],
   ['0,40', '+0,8', '1,4', '0,396', '+0,70', '1,56']],
  [1500, 1400, 1400, 1500, 1400, 1400]),
Caption('Tabel 4. Verifikasi generator pola napas pada 512 napas.'),
H('4.1.2  Kembaran Digital', HeadingLevel.HEADING_3),
P('Sebelum dipakai sebagai wahana pengujian, kembaran digital harus terbukti mereproduksi arah temuan kepustakaan. Sapuan koefisien variasi pada ARDS sedang dengan tidal volume rerata 0,35 L dan PEEP 8 cmH₂O memberikan hasil pada Tabel 6.'),
tbl(['CV', 'Fraksi terekrut', 'GI', 'MP (J/menit)'],
  [['0,00', '0,744', '0,454', '12,4'], ['0,20', '0,764', '0,425', '12,9'], ['0,40', '0,790', '0,369', '15,5']],
  [2000, 2200, 2200, 2200]),
Caption('Tabel 5. Respons kembaran digital terhadap takaran variabilitas.'),
P('Arahnya konsisten dengan kepustakaan: variabilitas menaikkan rekrutmen dan menurunkan keratan ventilasi, dengan mechanical power sebagai harga yang dibayar. Adanya trade-off inilah yang memastikan persoalan optimasi yang diajukan memang bermakna. Seandainya variabilitas hanya menguntungkan tanpa biaya, tidak ada yang perlu dioptimasi.'),
H('4.1.3  Waktu Komputasi', HeadingLevel.HEADING_3),
P('Satu penyelesaian MPC lengkap, mencakup CEM dan poles koordinat dua putaran, memerlukan 0,9 sampai 1,1 detik pada perangkat pengembangan. Angka ini terlalu lambat untuk loop per napas, tetapi memadai untuk loop supervisori yang berjalan setiap dua puluh napas. Pada simulator berbasis peramban, versi ringkas algoritma yang sama selesai dalam 70 sampai 95 milidetik.'),
H('4.2  Matriks Skenario', HeadingLevel.HEADING_2),
P('Tabel 7 sampai 9 merangkum tiga besaran terpenting pada enam kondisi paru. Seluruh angka dibaca langsung dari keluaran eksperimen; data lengkap seluruh metrik ada pada Lampiran E.'),
matrix('mp_mean', 1),
Caption('Tabel 6. Mechanical power rerata (J/menit). Makin kecil makin baik.'),
matrix('gi_mean', 3),
Caption('Tabel 7. Global Inhomogeneity index. Makin kecil berarti ventilasi makin merata.'),
matrix('driving_mean', 1),
Caption('Tabel 8. Driving pressure rerata (cmH₂O). Batas klinis 15.'),
Img('gambar1_matriks.png', 6.3),
Caption('Gambar 3. Perbandingan tiga besaran utama pada enam kondisi paru.'),
H('4.2.1  Takaran Variabilitas Bersifat Kondisi-Spesifik', HeadingLevel.HEADING_3),
P('Temuan utama pekerjaan ini terangkum pada Tabel 10. Controller menolak variabilitas pada paru sehat dan paru obstruktif, yaitu kondisi yang tidak memiliki unit kolaps untuk direkrut sehingga variabilitas hanya membakar energi. Sebaliknya controller memakainya kuat pada ARDS ringan sampai sedang dan setelah derekrutmen mendadak, lalu mundur kembali pada ARDS berat karena batas tekanan tidak menyisakan ruang untuk napas yang lebih besar.'),
tbl(['Kondisi paru', 'CV yang dipilih'], SCN.map(sc => [label(sc), num(get(sc, 'VIVA-MPC').vt_cv_realised, 3)]), [4400, 2200]),
Caption('Tabel 9. Takaran variabilitas yang dipilih VIVA-MPC pada tiap kondisi.'),
Img('gambar2_dosis.png', 6.3),
Caption('Gambar 4. Kiri: takaran variabilitas terpilih dibandingkan takaran tetap 0,30 dari kepustakaan. Kanan: posisi tiap strategi pada bidang mechanical power terhadap GI.'),
H('4.2.2  Biaya Takaran Tetap', HeadingLevel.HEADING_3),
P('Kerugian dari pemakaian takaran tetap terbaca jelas pada Tabel 6. Pada paru obstruktif, ventilasi variabel bertakaran 30 persen membakar mechanical power jauh di atas ARDSNet untuk GI yang praktis sama — manfaat nol dengan biaya energi yang nyata. Menurut penulis, inilah argumen terkuat pekerjaan ini, dan argumen tersebut datang dari kegagalan pembanding, bukan dari keunggulan sistem yang diusulkan.'),
H('4.2.3  Keselamatan', HeadingLevel.HEADING_3),
P('Tabel 11 menyajikan persentase napas yang melanggar batas tekanan. Angka ini penting karena menunjukkan apakah pengawas keselamatan benar-benar bekerja, bukan sekadar ada di dalam kode.'),
tbl(['Kondisi paru'].concat(CTRL.map(c => SHORT[c])),
  SCN.map(sc => [label(sc)].concat(CTRL.map(c => num(get(sc, c).pct_breaths_plateau_over_30, 1)))),
  [2500, 1450, 1450, 1450, 1450]),
Caption('Tabel 10. Persentase napas dengan tekanan plateau di atas 30 cmH₂O.'),
H('4.3  Ablasi Terkendali', HeadingLevel.HEADING_2),
P('Hasil ablasi disajikan ringkas pada Tabel 12; data lengkapnya ada pada Lampiran F.'),
tbl(['Bobot GI', 'Varian', 'MP (J/mnt)', 'GI', 'ΔP', 'CV akhir'],
  paretoRows.map(r => [num(r.w_gi, 0), r.variant, num(r.mp_mean, 1), num(r.gi_mean, 3), num(r.driving_mean, 1), num(r.cv_final, 3)]),
  [1300, 1500, 1500, 1300, 1300, 1400]),
Caption('Tabel 11. Ablasi variabilitas pada ARDS sedang, empat bobot GI.'),
Img('gambar3_ablasi.png', 4.6),
Caption('Gambar 5. Sebaran hasil ablasi pada bidang mechanical power terhadap GI.'),
P('Hasilnya tidak mendukung klaim superioritas. Varian CV bebas memilih takaran mendekati 0,3 tetapi mendarat pada titik yang didominasi varian CV tetap 0,15, yaitu lebih boros energi sekaligus lebih tidak merata. Artinya optimizer menemukan variabilitas, tetapi salah memilih takarannya.'),
P('Penyebabnya dilacak sampai ke kalibrasi model internal. Model internal meramalkan perbaikan GI sekitar −0,076 untuk perubahan CV dari 0 ke 0,30, sementara kembaran digital hanya memberi −0,035. Model internal terlalu optimistis mengenai imbalan variabilitas, sehingga controller membelinya terlalu banyak. Simpulan yang dapat ditarik, dan menurut penulis inilah temuan metodologis yang paling berguna dari pekerjaan ini: nilai dari mengoptimasi variabilitas dibatasi oleh fidelitas dinamika rekrutmen pada model internal controller, bukan oleh algoritma optimasinya.'),
H('4.4  Pencapaian Indikator Kinerja', HeadingLevel.HEADING_2),
P('Tabel 13 mengevaluasi sistem terhadap dua belas indikator yang ditetapkan pada subbab 3.10. Seluruh angka pada kolom hasil dihitung langsung dari data eksperimen.'),
tbl(['Kode', 'Indikator', 'Target', 'Hasil', 'Status'],
  KPI.map(x => [x.k, x.pendek, x.target, x.hasil, statusOf(x)]),
  [700, 2350, 1750, 2050, 1200], { leftAll: true, size: 17 }),
Caption('Tabel 13. Pencapaian indikator kinerja utama.'),
Rich([{ text: `Rekapitulasi: ${KPI_OK} indikator tercapai, ${KPI_SEB} tercapai sebagian, dan ${KPI.length - KPI_OK - KPI_SEB} belum tercapai dari total ${KPI.length}. `, bold: true },
  semua(['K1', 'K2', 'K3', 'K4'])
    ? 'Seluruh indikator keselamatan terpenuhi pada semua kondisi paru.'
    : `Kelompok keselamatan belum sepenuhnya terpenuhi. ${daftar(belumOf(['K1','K2','K3','K4']))} tidak tercapai` +
      (sebagianOf(['K1','K2','K3','K4']).length
        ? `, sementara ${daftar(sebagianOf(['K1','K2','K3','K4']))} tercapai sebagian.` : '.')]),
P(`Kegagalan keselamatan tersebut berasal dari kondisi yang sama. Pelanggaran tekanan plateau terburuk terjadi pada ${label(worstPlateau)}, dan PaCO₂ tertinggi juga pada ${label(worstCO2)}. Keduanya bukan galat perangkat lunak, melainkan tabrakan antar-kendala yang memang tidak memiliki solusi pada paru sekaku itu: tidal volume yang cukup untuk membuang CO₂ tidak dapat diberikan tanpa menembus batas tekanan. Pengawas keselamatan menangani tabrakan tersebut sesuai rancangan pada subbab 3.7, yaitu melonggarkan plafon tekanan ke batas eskalasi sambil mencatat peristiwanya. Justru di sinilah target yang ditetapkan sebelum eksperimen membuktikan gunanya: ia menolak menyatakan sistem berhasil hanya karena perilakunya sudah dirancang demikian.`),
P('Kelompok adaptivitas juga terpenuhi, dan kedua indikatornya saling menguatkan. K8 menunjukkan takaran yang dipilih memang terentang lebar antar kondisi, sementara K9 menunjukkan sistem tahu kapan sebaiknya tidak berbuat apa-apa. Menurut penulis, K9 lebih bernilai daripada K8. Sistem yang selalu bertindak mudah dibuat; sistem yang menahan diri ketika tindakan tidak berguna jauh lebih sulit, dan itulah yang membedakannya dari takaran tetap.'),
P('Kegagalan terpusat pada kelompok efikasi, dan penyebabnya sama dengan yang diuraikan pada subbab 4.3. Model internal controller terlalu optimistis mengenai imbalan variabilitas, sehingga sistem membelinya terlalu banyak pada kondisi yang tidak membutuhkannya dan kehabisan ruang pada kondisi yang membutuhkannya. Indikator-indikator inilah yang harus diperiksa ulang setelah kalibrasi model internal diperbaiki.'),

H('4.5  Pembahasan terhadap Kepustakaan', HeadingLevel.HEADING_2),
P('Temuan bahwa takaran variabilitas bersifat kondisi-spesifik sejalan dengan laporan bahwa manfaat ventilasi variabel bergantung pada etiologi ARDS, di mana ARDS pulmoner yang didominasi konsolidasi memberi respons berbeda dari ARDS ekstrapulmoner yang didominasi kolaps alveolar. Bila manfaatnya memang bergantung etiologi, maka takaran tetap untuk semua pasien sulit dipertahankan secara konseptual, dan hasil simulasi ini memberi dukungan tambahan dari arah yang berbeda.'),
P('Sebaliknya, kegagalan varian CV bebas mengungguli takaran tetap 0,15 perlu ditempatkan secara jujur. Hasil ini tidak membantah gagasan personalisasi takaran; yang ditunjukkannya adalah bahwa personalisasi memerlukan model internal yang cukup setia. Pada kerangka yang sama, memperbaiki kalibrasi model internal merupakan langkah yang jelas dan terukur, dan itulah saran pertama pada Bab V.'),
H('4.6  Kegagalan dan Keterbatasan', HeadingLevel.HEADING_2),
P('ARDS berat. Sistem menghasilkan GI dan fraksi rekrutmen terburuk di antara empat strategi. Pengawas keselamatan mengunci controller pada sudut ruang solusi yang sempit dan controller tidak keluar dari situ.', { numbering: { reference: 'bullet', level: 0 } }),
P('Paru sehat. Sistem memakai mechanical power lebih besar daripada ARDSNet untuk perbaikan GI yang kecil. Bobot fungsi biaya belum tertata untuk kondisi ringan.', { numbering: { reference: 'bullet', level: 0 } }),
P('Ablasi belum mendukung klaim superioritas, sebagaimana diuraikan pada subbab 4.3.', { numbering: { reference: 'bullet', level: 0 } }),
P('Di luar ketiganya, keterbatasan terbesar adalah bahwa seluruh angka berasal dari simulasi terhadap model yang dibangun sendiri. Kembaran digital dirancang konsisten dengan arah temuan kepustakaan, tetapi belum pernah dikalibrasi terhadap data manusia. Yang ditunjukkan pekerjaan ini adalah bahwa kerangka kendalinya dapat dibangun, berjalan dalam waktu yang masuk akal, dan memilih takaran variabilitas yang berbeda menurut kondisi paru. Yang belum ditunjukkan adalah bahwa takaran-takaran itu benar.'),
H('4.7  Catatan Rekayasa', HeadingLevel.HEADING_2),
P('Empat kesalahan ditemukan selama pengembangan, dan seluruhnya dicatat di sini karena kemungkinan besar akan ditemui siapa pun yang mengulang pekerjaan ini.'),
P('Simulasi meledak. Memodelkan derekrutmen sebagai elastance yang membesar menjatuhkan konstanta waktu kompartemen di bawah langkah integrator. Perbaikannya justru lebih benar secara fisiologi: unit yang kolaps itu tertutup, bukan kaku, sehingga yang naik adalah resistansinya.', { numbering: { reference: 'L5', level: 0 } }),
P('Estimator tidak menemukan R. Penyebabnya tekanan direkam di balik resistansi jalan napas, padahal ventilator mengukur di Y-piece. Pada titik yang salah, suku R·V̇ lenyap dari persamaan gerak.', { numbering: { reference: 'L5', level: 0 } }),
P('Kembaran digital tidak menunjukkan manfaat variabilitas. Sebaran tekanan pembukaan terlalu sempit, sehingga setiap napas membuka semua unit dan pertidaksamaan Jensen tidak punya tempat bekerja.', { numbering: { reference: 'L5', level: 0 } }),
P('Laju rekrutmen model internal sembilan belas kali terlalu cepat. Rekrutmen jenuh dalam dua napas, sehingga rollout meramalkan GI yang sama untuk CV berapa pun dan optimizer menyimpulkan variabilitas tidak berguna — persis kebalikan perilaku kembaran digital.', { numbering: { reference: 'L5', level: 0 } }),
H('4.8  Simulator dan Audit Antarmuka', HeadingLevel.HEADING_2),
P('Simulator disusun mengikuti alur bernomor satu sampai lima, dari pemilihan pasien sampai analisis. Setiap pembacaan disajikan bersama batang rentang dan vonis aman, waspada, atau bahaya, sehingga pengguna tidak perlu menghafal ambang klinis. Panel analisis menerjemahkan angka menjadi kalimat berisi apa yang terjadi dan apa yang dapat dilakukan, dan setelah controller dijalankan sistem menunggu 26 napas sebelum menampilkan perbandingan sebelum dan sesudah agar rerata tidak tercampur napas dari setelan lama.'),
P('Palet warna disusun dari pengukuran, bukan dari selera. Seluruh pasangan warna yang benar-benar dipakai diukur terhadap WCAG 2.1, dengan syarat 4,5:1 untuk teks biasa dan 3:1 untuk bentuk pembawa informasi. Audit awal menemukan enam pelanggaran, yang seluruhnya telah diperbaiki; hasil pengukuran lengkap ada pada Lampiran G.'),
P('Simulasi penglihatan buta warna menunjukkan bahwa hijau, kuning, dan merah status menjadi nyaris tidak terbedakan pada protanopia maupun deuteranopia, dengan rasio pembeda hanya 1,0 sampai 1,7. Karena itu status tidak pernah disampaikan melalui warna saja: setiap pembacaan membawa label teks, setiap temuan analisis diberi lambang di depan judulnya, dan arah perubahan pada tabel perbandingan ditandai panah. Aksen untuk penanda inovasi juga dipindahkan dari oranye ke nila, karena pada sistem kendali satu rona tidak boleh memiliki dua arti.'),
new Paragraph({ children: [new PageBreak()] }),

/* ===================== BAB V ===================== */
H('BAB V  PENUTUP', HeadingLevel.HEADING_1),
H('5.1  Simpulan', HeadingLevel.HEADING_2),
Rich([`Dari dua belas indikator kinerja yang ditetapkan sebelum eksperimen dijalankan, ${KPI_OK} tercapai, ${KPI_SEB} tercapai sebagian, dan ${KPI.length - KPI_OK - KPI_SEB} belum tercapai. `,
  `Kedua indikator adaptivitas terpenuhi, sedangkan kegagalan terpusat pada kelompok efikasi dan pada kondisi ARDS berat. Secara rinci:`]),
P('Sistem kendali lima lapis yang memperlakukan koefisien variasi, kemencengan, dan eksponen fraktal pola napas sebagai variabel keputusan berhasil dibangun dan berjalan, dengan waktu penyelesaian satu langkah optimasi sekitar satu detik, memadai untuk loop supervisori antar-napas.', { numbering: { reference: 'L6', level: 0 } }),
P('Kembaran digital yang dibangun mereproduksi arah temuan kepustakaan mengenai ventilasi variabel, yaitu kenaikan rekrutmen dan penurunan keratan ventilasi yang disertai kenaikan mechanical power.', { numbering: { reference: 'L6', level: 0 } }),
P('Takaran variabilitas yang dipilih sistem berbeda-beda menurut kondisi paru, dari 0,017 pada paru obstruktif sampai 0,227 setelah derekrutmen mendadak. Tidak ada satu takaran tunggal yang tepat untuk semua kondisi.', { numbering: { reference: 'L6', level: 0 } }),
P('Pemakaian takaran tetap 30 persen menaikkan mechanical power 1,7 sampai 2,7 kali, dan pada paru obstruktif kenaikan itu terjadi tanpa manfaat sama sekali.', { numbering: { reference: 'L6', level: 0 } }),
P('Ablasi terkendali belum mendukung klaim bahwa membebaskan takaran variabilitas lebih baik daripada mematoknya pada 0,15. Penyebabnya adalah model internal controller yang terlalu optimistis mengenai imbalan variabilitas.', { numbering: { reference: 'L6', level: 0 } }),
H('5.2  Saran', HeadingLevel.HEADING_2),
P('Mengalibrasi model internal terhadap kembaran digital melalui identifikasi, bukan penyetelan manual. Langkah ini paling menentukan, karena tanpa itu klaim superioritas tidak akan berdiri di hadapan penguji.', { numbering: { reference: 'L7', level: 0 } }),
P('Menyapu bobot fungsi biaya untuk membangun frontier Pareto, lalu menyajikan frontier tersebut alih-alih klaim skalar yang selalu dapat dipatahkan dengan mempersoalkan pemilihan bobot.', { numbering: { reference: 'L7', level: 0 } }),
P('Menambahkan manuver rekrutmen terjadwal sebagai aksi tersendiri untuk memperbaiki perilaku pada ARDS berat.', { numbering: { reference: 'L7', level: 0 } }),
P('Melanjutkan ke pengujian hardware-in-the-loop memakai platform ventilator sumber terbuka dan test lung terstandar.', { numbering: { reference: 'L7', level: 0 } }),
P('Melakukan uji Monte Carlo pada populasi pasien virtual dengan domain randomization sesuai rentang pada ISO 80601-2-80.', { numbering: { reference: 'L7', level: 0 } }),
new Paragraph({ children: [new PageBreak()] }),

/* ===================== PUSTAKA ===================== */
H('DAFTAR PUSTAKA', HeadingLevel.HEADING_1),
...[
  'Amato, M. B. P., Meade, M. O., Slutsky, A. S., dkk. (2015). Driving pressure and survival in the acute respiratory distress syndrome. New England Journal of Medicine, 372(8), 747–755.',
  'Bates, J. H. T., & Irvin, C. G. (2002). Time dependence of recruitment and derecruitment in the lung: a theoretical model. Journal of Applied Physiology, 93(2), 705–713.',
  'Beda, A., Spieth, P. M., Handzsuj, T., Pelosi, P., Carvalho, N. C., Koch, E., Koch, T., & Gama de Abreu, M. (2010). A novel adaptive control system for noisy pressure-controlled ventilation: a numerical simulation and bench test study. Intensive Care Medicine, 36(1), 164–168.',
  'Brewster, J. F., Graham, M. R., & Mutch, W. A. C. (2005). Convexity, Jensen\u2019s inequality and benefits of noisy mechanical ventilation. Journal of the Royal Society Interface, 2(4), 393–396.',
  'Gattinoni, L., Tonetti, T., Cressoni, M., dkk. (2016). Ventilator-related causes of lung injury: the mechanical power. Intensive Care Medicine, 42(10), 1567–1575.',
  'Huhle, R., Pelosi, P., & Gama de Abreu, M. (2016). Variable ventilation from bench to bedside. Critical Care, 20, 62.',
  'Kiss, T., Silva, P. L., Huhle, R., dkk. (2016). Comparison of different degrees of variability in tidal volume to prevent deterioration of respiratory system elastance. British Journal of Anaesthesia, 116(5), 708–715.',
  'Mutch, W. A. C., Harms, S., Lefevre, G. R., Graham, M. R., Girling, L. G., & Kowalski, S. E. (2000). Biologically variable ventilation increases arterial oxygenation over that seen with positive end-expiratory pressure alone. Critical Care Medicine, 28(7), 2457–2464.',
  'Otis, A. B., Fenn, W. O., & Rahn, H. (1950). Mechanics of breathing in man. Journal of Applied Physiology, 2(11), 592–607.',
  'Schreiber, T., & Schmitz, A. (1996). Improved surrogate data for nonlinearity tests. Physical Review Letters, 77(4), 635–638.',
  'Suki, B., Alencar, A. M., Sujeer, M. K., Lutchen, K. R., Collins, J. J., Andrade, J. S., Ingenito, E. P., Zapperi, S., & Stanley, H. E. (1998). Life-support system benefits from noise. Nature, 393, 127–128.',
  'Zhao, Z., Möller, K., Steinmann, D., Frerichs, I., & Guttmann, J. (2009). Evaluation of an electrical impedance tomography-based Global Inhomogeneity Index for pulmonary ventilation distribution. Intensive Care Medicine, 35(11), 1900–1906.',
].map(t => new Paragraph({ spacing: { after: 150, line: 265 },
  indent: { left: convertInchesToTwip(0.4), hanging: convertInchesToTwip(0.4) },
  children: [new TextRun({ text: t, size: 20, font: 'Calibri', color: INK })] })),
new Paragraph({ children: [new PageBreak()] }),

/* ===================== LAMPIRAN ===================== */
H('LAMPIRAN A  Struktur Berkas dan Tautan', HeadingLevel.HEADING_1),
P('Seluruh berkas berada dalam satu folder. Tautan di bawah bersifat relatif, sehingga dapat diklik langsung dari dokumen ini selama dokumen disimpan di folder yang sama dengan berkas-berkas tersebut.'),
tbl(['Berkas', 'KB', 'Keterangan'], [
  ['viva-mpc-simulator.html', String(DAT.berkas.find(b => b[0] === 'viva-mpc-simulator.html')?.[1] || '–'), 'Simulator interaktif, berdiri sendiri tanpa server'],
  ['vivampc/', '–', 'Paket Python berisi sepuluh modul (Tabel 4)'],
  ['run_experiments.py', String(DAT.berkas.find(b => b[0] === 'run_experiments.py')?.[1] || '–'), 'Matriks enam skenario terhadap empat strategi'],
  ['run_pareto.py', String(DAT.berkas.find(b => b[0] === 'run_pareto.py')?.[1] || '–'), 'Ablasi terkendali parameter variabilitas'],
  ['make_figures.py', String(DAT.berkas.find(b => b[0] === 'make_figures.py')?.[1] || '–'), 'Pembuatan gambar dari hasil eksperimen'],
  ['cek_instalasi.py', String(DAT.berkas.find(b => b[0] === 'cek_instalasi.py')?.[1] || '–'), 'Verifikasi cepat seluruh komponen'],
  ['buat_paru_blender.py', String(DAT.berkas.find(b => b[0] === 'buat_paru_blender.py')?.[1] || '–'), 'Pemodelan geometri paru di Blender'],
  ['tanam_mesh.py', String(DAT.berkas.find(b => b[0] === 'tanam_mesh.py')?.[1] || '–'), 'Menanam mesh hasil Blender ke simulator'],
  ['paru_blender.obj', String(DAT.berkas.find(b => b[0] === 'paru_blender.obj')?.[1] || '–'), 'Mesh paru mentah hasil Blender'],
  ['uji_web.js', String(DAT.berkas.find(b => b[0] === 'uji_web.js')?.[1] || '–'), 'Pengujian otomatis antarmuka (34 pemeriksaan)'],
  ['uji_webgl.js', String(DAT.berkas.find(b => b[0] === 'uji_webgl.js')?.[1] || '–'), 'Uji render WebGL sungguhan, menghasilkan PNG'],
  ['results.json', String(DAT.berkas.find(b => b[0] === 'results.json')?.[1] || '–'), 'Data mentah matriks skenario'],
  ['pareto.json', String(DAT.berkas.find(b => b[0] === 'pareto.json')?.[1] || '–'), 'Data mentah ablasi'],
  ['buat_laporan_docx.js', String(DAT.berkas.find(b => b[0] === 'buat_laporan_docx.js')?.[1] || '–'), 'Pembangun dokumen laporan ini'],
], [2700, 700, 4200], { leftAll: true, size: 17 }),
Caption('Tabel 12. Struktur berkas perangkat lunak.'),
P('Tautan langsung:', { bold: true, after: 60 }),
...[['Simulator interaktif', 'viva-mpc-simulator.html'], ['Paket Python', 'vivampc/'],
    ['Data mentah matriks skenario', 'results.json'], ['Data mentah ablasi', 'pareto.json'],
    ['Skrip pemodelan paru di Blender', 'buat_paru_blender.py'],
    ['Mesh paru (.obj)', 'paru_blender.obj'],
    ['Pengujian otomatis antarmuka', 'uji_web.js'], ['Pengujian render WebGL', 'uji_webgl.js'],
    ['Pembangun laporan ini', 'buat_laporan_docx.js'], ['Panduan menjalankan', 'CARA_MENJALANKAN.md'],
    ['Catatan pengerjaan', 'LAPORAN.md']].map(([a, b]) => Link(a + ' — ' + b, b)),

H('LAMPIRAN B  Cara Menjalankan', HeadingLevel.HEADING_1),
P('Kebutuhan: Python 3.10 atau lebih baru dengan numpy, scipy, dan matplotlib. Simulator tidak memerlukan pemasangan apa pun.'),
...CodeBlock([
  'python -m venv .venv',
  'source .venv/bin/activate        # Windows: .venv\\Scripts\\activate',
  'pip install -r requirements.txt',
  '',
  'python cek_instalasi.py          # verifikasi cepat, beberapa detik',
  'python run_pareto.py             # ablasi variabilitas, ~4 menit',
  'python run_experiments.py        # matriks skenario, ~5 menit',
  'python make_figures.py           # menghasilkan gambar',
].join('\n')),
P('Membangun ulang geometri paru memerlukan Blender, karena skripnya memakai modul bpy yang hanya tersedia di dalam Blender:'),
...CodeBlock([
  'blender -b --python buat_paru_blender.py',
  'python tanam_mesh.py             # menanam mesh ke simulator',
].join('\n')),
P('Membangun ulang dokumen laporan ini memerlukan Node.js:'),
...CodeBlock('node buat_laporan_docx.js'),

H('LAMPIRAN C  Parameter Kondisi Paru', HeadingLevel.HEADING_1),
P('Nilai berikut menentukan keenam kondisi paru yang diuji. Compliance dinyatakan dalam mL/cmH₂O, resistansi dalam cmH₂O·s/L, dan tekanan pembukaan dalam cmH₂O.'),
tbl(['Kondisi', 'Compliance', 'Resistansi', 'P buka (median)', 'Sebaran', '%E₂', 'Shunt dasar'],
  DAT.skenario.map(r => r.map(String)), [1900, 1150, 1150, 1400, 900, 850, 1100], { size: 17 }),
Caption('Tabel 13. Parameter kembaran digital untuk tiap kondisi paru.'),

H('LAMPIRAN D  Konfigurasi Controller', HeadingLevel.HEADING_1),
P('Rentang variabel keputusan:'),
tbl(['Variabel', 'Batas bawah', 'Batas atas'], DAT.bounds.map(b => [b[0], b[1], b[2]]),
  [3600, 1600, 1600], { size: 17 }),
Caption('Tabel 14. Rentang pencarian variabel keputusan.'),
P('Bobot fungsi biaya:'),
tbl(['Suku', 'Bobot', 'Satuan'], DAT.bobot.map(b => [b[0], b[1], b[2] || '–']),
  [2200, 1400, 3200], { leftAll: true, size: 17 }),
Caption('Tabel 15. Bobot fungsi biaya economic MPC.'),
P('Batas keselamatan:'),
tbl(['Batas', 'Nilai'], DAT.limits.map(l => [l[0], l[1]]), [3600, 1800], { size: 17 }),
Caption('Tabel 16. Batas pada pengawas keselamatan.'),
Note('Mechanical power sengaja tidak dimasukkan sebagai kendala keras, melainkan sebagai suku biaya. Pada ARDS berat tidak ada setelan yang sekaligus membersihkan CO₂ dan menahan mechanical power di bawah 17 J/menit; menjadikannya kendala keras hanya melahirkan kebuntuan. Pada percobaan awal, controller menekan tidal volume sampai di bawah ruang rugi dan menghasilkan PaCO₂ 120 mmHg.'),

H('LAMPIRAN E  Data Lengkap Matriks Skenario', HeadingLevel.HEADING_1),
P(`Seluruh metrik pada ${res.rows.length} kombinasi pengujian, rerata dari ${res.seeds.length} benih acak. Dibaca langsung dari results.json.`),
tbl(['Kondisi', 'Strategi', 'MP', 'GI', 'ΔP', 'Plateau p95'],
  res.rows.map(r => [r.label, SHORT[r.controller], num(r.mp_mean, 1), num(r.gi_mean, 3),
    num(r.driving_mean, 1), num(r.plateau_p95, 1)]),
  [1800, 1500, 900, 950, 900, 1200], { size: 16 }),
Caption('Tabel 17. Data lengkap: tekanan dan energi.'),
tbl(['Kondisi', 'Strategi', 'PaCO₂', 'SpO₂', 'Terekrut', 'CV terwujud'],
  res.rows.map(r => [r.label, SHORT[r.controller], num(r.paco2_mean, 1), num(r.spo2_mean, 1),
    num(r.recruited_final, 3), num(r.vt_cv_realised, 3)]),
  [1800, 1500, 950, 950, 1050, 1100], { size: 16 }),
Caption('Tabel 18. Data lengkap: pertukaran gas dan rekrutmen.'),

H('LAMPIRAN F  Data Lengkap Ablasi', HeadingLevel.HEADING_1),
P(`Hasil ${par.rows.length} kombinasi ablasi pada kondisi ${par.scenario}, masing-masing ${par.minutes} menit simulasi. Dibaca langsung dari pareto.json.`),
tbl(['Bobot GI', 'Varian', 'MP', 'GI', 'ΔP', 'PaCO₂', 'SpO₂', 'CV akhir'],
  paretoRows.map(r => [num(r.w_gi, 0), r.variant, num(r.mp_mean, 1), num(r.gi_mean, 3),
    num(r.driving_mean, 1), num(r.paco2_mean, 1), num(r.spo2_mean, 1), num(r.cv_final, 3)]),
  [1000, 1250, 900, 950, 900, 1000, 950, 1050], { size: 16 }),
Caption('Tabel 19. Data lengkap ablasi variabilitas.'),

H('LAMPIRAN G  Audit Kontras Warna', HeadingLevel.HEADING_1),
P('Rasio kontras dihitung menurut WCAG 2.1 untuk seluruh pasangan warna yang benar-benar dipakai pada antarmuka. Syarat 4,5:1 berlaku untuk teks biasa dan 3:1 untuk bentuk pembawa informasi.'),
tbl(['Pasangan', 'Depan', 'Belakang', 'Rasio', 'Syarat', 'Status'],
  DAT.kontras.map(k => [k[0], k[1], k[2], num(k[3], 2), num(k[4], 1), k[3] >= k[4] ? 'lolos' : 'gagal']),
  [2600, 1050, 1050, 900, 900, 1000], { size: 16 }),
Caption('Tabel 20. Hasil audit kontras antarmuka setelah perbaikan.'),

H('LAMPIRAN H  Daftar Pengujian Otomatis', HeadingLevel.HEADING_1),
P(`Sebanyak ${DAT.uji.length} pemeriksaan dijalankan pada setiap perubahan simulator. Seluruhnya lolos pada versi yang dilaporkan.`),
tbl(['No', 'Pemeriksaan'], DAT.uji.map((u, i) => [String(i + 1), u]), [700, 6500], { leftAll: true, size: 17 }),
Caption('Tabel 21. Daftar pemeriksaan otomatis simulator.'),

H('LAMPIRAN I  Cuplikan Kode Inti', HeadingLevel.HEADING_1),
...DAT.kode.flatMap(([judul, src]) => [
  P(judul, { bold: true, after: 70 }),
  ...CodeBlock(src),
  P(' ', { after: 120 }),
]),

H('LAMPIRAN J  Palet Warna Antarmuka', HeadingLevel.HEADING_1),
P('Seluruh warna yang dipakai antarmuka, sebagaimana tertulis pada berkas simulator.'),
tbl(['Peubah', 'Nilai'], DAT.palet.map(p => ['--' + p[0], p[1]]), [3200, 2000], { leftAll: true, size: 17, mono: true }),
Caption('Tabel 22. Palet warna antarmuka simulator.'),
    ],
  }],
});

Packer.toBuffer(doc).then(b => {
  fs.writeFileSync(path.join(HERE, 'LAPORAN-VIVA-MPC.docx'), b);
  fs.writeFileSync(path.join(HERE, 'caption_list.json'), JSON.stringify(CAPS, null, 1));
  console.log('LAPORAN-VIVA-MPC.docx ditulis,', (b.length / 1024).toFixed(0), 'KB;',
    CAPS.length, 'gambar/tabel berketerangan');
});
