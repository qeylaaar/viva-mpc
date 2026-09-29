# Catatan pengerjaan VIVA-MPC

Ini catatan kerja, bukan laporan yang sudah dirapikan. Saya tulis berurutan sesuai
yang terjadi, termasuk bagian yang salah, supaya kalau nanti saya atau orang lain
mengulang pekerjaan ini tidak jatuh ke lubang yang sama.

---

## Kenapa idenya begini

Berangkat dari satu pengamatan waktu baca-baca literatur closed-loop ventilation.
ASV, INTELLiVENT, SmartCare, sampai yang pakai reinforcement learning — semuanya
mengatur empat hal yang sama: tidal volume, PEEP, laju napas, FiO2. Semuanya nilai
rerata.

Di sisi lain ada literatur ventilasi variabel yang sudah dua dekade menunjukkan
napas yang bervariasi itu bagus buat paru. Tapi di situ variabilitasnya selalu
dipatok di depan. CV 30%, distribusi simetris, selesai. Tidak ada yang menanyakan
apakah 30% itu angka yang benar untuk pasien ini, hari ini.

Jadi pertanyaannya: bagaimana kalau bentuk statistik pola napasnya sendiri yang jadi
variabel keputusan?

```
u = [ VT_rerata, PEEP, RR_rerata, FiO2, CV, gamma, beta ]
```

Tiga yang terakhir itu yang belum pernah dioptimasi siapa pun secara online.

---

## Yang akhirnya jadi

Paket Python `vivampc`, delapan modul, jalan semua. Plus simulator web buat demo.

Arsitekturnya lima lapis:

- **Lapis 0** sensing — flow, tekanan jalan napas, EtCO2, SpO2, EIT
- **Lapis 1** estimator — Moving Horizon Estimation buat R, E1, E2
- **Lapis 2** digital twin — 20 kompartemen, rekrutmen bergantung waktu
- **Lapis 3** MPC — economic stochastic, horizon 45 napas
- **Lapis 4** safety supervisor — hierarki keras

Satu keputusan yang saya rasa penting: twin fidelitas tinggi berperan sebagai
pasien, sementara controller cuma pegang model tingkat-napas yang jauh lebih kasar.
Sengaja tidak dibuat cocok. Kalau controller tetap bekerja walau modelnya meleset,
hasilnya jauh lebih meyakinkan ketimbang controller yang mengoptimasi model dirinya
sendiri.

---

## Verifikasi komponen

### Generator pola napas

Awalnya saya pakai kopula Gaussian buat mengatur kemencengan sambil menjaga struktur
korelasi. Gagal. Begitu kemencengannya besar, transformasi marginalnya menggeser
energi antar frekuensi dan eksponen fraktalnya rusak — minta beta = 1,0 dapatnya 1,73.

Solusinya IAAFT (Schreiber & Schmitz, PRL 1996). Algoritmanya memaksakan dua kendala
bergantian sampai dua-duanya terpenuhi: distribusi marginal persis target, amplitudo
spektral persis target. Setelah itu ketiganya kena bersamaan:

| Target CV | Target gamma | Target beta | Dapat CV | Dapat gamma | Dapat beta |
|---|---|---|---|---|---|
| 0,30 | 0,0 | 0,0 | 0,300 | +0,00 | 0,08 |
| 0,30 | +0,5 | 1,0 | 0,300 | +0,49 | 1,10 |
| 0,15 | -0,6 | 0,5 | 0,150 | -0,59 | 0,58 |
| 0,40 | +0,8 | 1,4 | 0,396 | +0,70 | 1,56 |

Ventilasi semenit saya jaga konstan sampai presisi mesin. Ini bukan detail sepele —
tanpa itu, menaikkan CV diam-diam menaikkan ventilasi semenit juga, dan semua
perbandingan antar-CV jadi tidak adil.

### Twin

Yang saya kejar: apakah twin ini mereproduksi arah temuan literatur. Sapuan CV pada
ARDS sedang, VT rerata 0,35 L, PEEP 8:

| CV | Rekrutmen | GI | MP (J/mnt) |
|---|---|---|---|
| 0,00 | 0,744 | 0,454 | 12,4 |
| 0,20 | 0,764 | 0,425 | 12,9 |
| 0,40 | 0,790 | 0,369 | 15,5 |

Variabilitas beli rekrutmen dan homogenitas, bayarnya pakai energi. Trade-off ini
yang harus dinavigasi MPC.

### Kecepatan

Satu solusi MPC 0,9–1,1 detik. Kelamaan buat loop per napas, tapi cukup untuk loop
supervisori tiap 20 napas. Saya anggap wajar.

---

## Empat hal yang bikin macet

Ini bagian yang paling menghabiskan waktu, dan menurut saya paling berguna dicatat.

**1. Simulasinya meledak.** Awalnya derekrutmen saya modelkan sebagai elastance yang
membesar. Kelihatannya masuk akal — paru kolaps jadi kaku. Ternyata konstanta waktu
kompartemennya jatuh sampai di bawah langkah integrator dan semua nilainya jadi NaN.
Baru sadar setelah lama: unit yang kolaps itu bukan kaku, dia *tertutup*. Yang naik
resistansinya, bukan elastance-nya. R sebanding 1/s kuadrat. Setelah diubah, stabil.

**2. Estimator tidak bisa menemukan R sama sekali.** Taksirannya mentok di batas
bawah terus. Penyebabnya konyol: saya merekam tekanan di balik resistansi jalan
napas. Ventilator itu mengukur di Y-piece. Kalau salah titik, suku R dikali laju
aliran hilang dari persamaan gerak dan tidak ada R yang bisa ditemukan. Sekalian
volumenya juga harus dirujuk ke akhir ekspirasi, bukan ke nol absolut, supaya E1 dan
E2 tidak saling tukar nilai.

**3. Twin-nya tidak menunjukkan manfaat variabilitas apa pun.** Ini yang paling bikin
frustrasi karena kalau ini tidak ketemu, seluruh ide saya runtuh. Sebaran tekanan
pembukaannya terlalu sempit — setiap napas membuka semua unit yang bisa dibuka, jadi
napas besar sesekali tidak punya sisa pekerjaan. Pertidaksamaan Jensen tidak punya
tempat bekerja. Baru muncul setelah saya ganti ke lognormal berekor panjang dengan
histeresis proporsional.

**4. Surrogate-nya 19 kali terlalu cepat merekrut.** Gejalanya aneh: controller-nya
konsisten memilih CV = 0 padahal twin jelas-jelas diuntungkan variabilitas. Ternyata
di rollout, rekrutmen jenuh dalam dua napas, jadi GI-nya sama saja untuk CV berapa
pun. Model internal yang terlalu optimistis soal rekrutmen bikin controller buta
terhadap satu-satunya hal yang ingin dia pelajari.

---

## Tiga keputusan yang harus saya bela nanti

**Chance constraint, bukan batas per napas.** Ambang Amato driving pressure kurang
dari 15 itu diturunkan dari ventilasi bervolume konstan, di mana setiap napas *adalah*
napas rerata. Menerapkannya ke persentil ke-95 dari pola variabel sama saja melarang
variabilitas secara struktural — napas terbesar selalu melanggar, sekecil apa pun
reratanya. Jadi saya pisah dua tingkat: rerata dibatasi 15 sesuai maksud aslinya,
ekor sebaran dibatasi 20 sebagai pagar terpisah. Literatur ventilasi variabel sendiri
menjalankan VT rerata 6 mL/kg dengan CV 30%, artinya napas individual di sana sampai
10 mL/kg.

**Mechanical power bukan kendala keras.** Percobaan awal saya jadikan kendala keras di
17 J/menit. Hasilnya controller menekan VT sampai di bawah ruang rugi dan PaCO2 tembus
120 mmHg — pasien virtualnya praktis dicekik. Setelah dipikir, memang tidak ada setting
yang bisa sekaligus bersihkan CO2 dan tahan MP di bawah 17 pada ARDS berat. Literatur
juga menyatakan bukti yang ada belum mendukung penitrasian ke ambang MP tertentu. Jadi
MP saya kembalikan ke perannya yang benar: suku biaya.

**Safety supervisor berhierarki.** Urutannya ventilasi alveolar, oksigenasi, lalu batas
tekanan. Kalau prioritas pertama dan ketiga bertabrakan, batas tekanan dilonggarkan ke
plafon kedua dan kejadiannya dicatat. Ini yang klinisi lakukan waktu hiperkapnia
permisif. Yang bahaya bukan pelonggarannya — yang bahaya kalau pelonggaran itu tidak
kelihatan.

---

## Hasil

### Dosisnya memang beda-beda

Ini inti klaim saya, dan ini yang paling saya senang lihat:

| Kondisi | CV yang dipilih |
|---|---|
| Paru sehat | 0,03 |
| ARDS ringan | 0,22 |
| ARDS sedang | 0,19 |
| ARDS berat | 0,07 |
| Obstruktif | 0,02 |
| Derekrutmen mendadak | 0,23 |

Controller-nya menolak variabilitas di paru sehat dan paru obstruktif. Masuk akal —
di sana tidak ada unit kolaps yang bisa direkrut, jadi variabilitas cuma bakar energi.
Dia pakai kuat di ARDS ringan sampai sedang dan setelah derekrutmen mendadak. Mundur
di ARDS berat karena batas tekanan tidak menyisakan ruang.

Tidak ada satu dosis yang benar untuk semua.

### Dosis tetap 30% itu mahal

| Kondisi | MP Noisy PCV | MP VIVA-MPC | MP ARDSNet |
|---|---|---|---|
| ARDS ringan | 22,4 | 15,5 | 12,2 |
| ARDS sedang | 32,6 | 19,1 | 16,9 |
| ARDS berat | 39,3 | 25,0 | 21,4 |
| Obstruktif | 34,0 | 13,4 | 12,6 |

Mematok CV 30% tanpa lihat kondisi menaikkan mechanical power 1,7 sampai 2,7 kali.
Yang paling telanjang itu kasus obstruktif: 34 J/menit terbakar untuk GI yang persis
sama dengan ARDSNet. Nol manfaat.

Menurut saya ini justru argumen terkuat saya, dan dia datang dari kegagalan pembanding,
bukan dari keunggulan sistem saya.

### Driving pressure

Di ARDS sedang, VIVA-MPC satu-satunya yang di bawah batas 15 (13,9). ARDSNet 21,0,
ASV-like 17,3, Noisy PCV 20,6. GI-nya juga terbaik di empat dari enam skenario.

Bayarannya hiperkapnia permisif — PaCO2 58,8 di ARDS ringan dan 64,9 di sedang,
sementara ARDSNet tahan di 45–46.

---

## Yang gagal

**ARDS berat.** GI 0,762, terburuk dari empat. Rekrutmen 0,53, juga terburuk. PaCO2
69,3. Safety supervisor mengunci controller di sudut sempit dan dia tidak keluar dari
situ. Ini kegagalan beneran dan saya tidak akan sembunyikan di skripsi.

**Paru sehat kena overtreatment.** MP 12,1 lawan ARDSNet 7,5, buat perbaikan GI dari
0,179 ke 0,133. Bayar energi 62% lebih banyak demi homogenitas yang paru sehat tidak
butuh. Bobot fungsi biayanya belum pas untuk kondisi ringan.

**Ablasinya tidak mendukung klaim superioritas.** Ini yang paling penting saya akui.
Mesin identik, cuma variabilitasnya dipatok atau dibebaskan:

| Varian | MP | GI |
|---|---|---|
| CV = 0 | 17,8 | 0,679 |
| CV = 0,15 | 18,6 | 0,330 |
| CV = 0,30 | 22,2 | 0,234 |
| CV bebas | 20,6 | 0,372 |

CV bebas memilih sekitar 0,32 tapi mendarat di titik yang *didominasi* CV = 0,15 —
lebih boros energi sekaligus lebih tidak merata. Jadi optimizer-nya menemukan
variabilitas, tapi salah pilih dosisnya.

Saya sudah lacak penyebabnya. Surrogate meramal perubahan GI sekitar -0,076 untuk CV
0 ke 0,30, sementara twin-nya cuma kasih -0,035. Model internalnya terlalu optimistis
soal imbalan variabilitas, jadi controller-nya beli kebanyakan.

Kesimpulannya: nilai dari mengoptimasi variabilitas dibatasi fidelitas dinamika
rekrutmen di model internal. Ini menurut saya layak jadi satu subbab sendiri, bukan
disembunyikan.

---

## Rencana berikutnya

1. **Kalibrasi surrogate ke twin.** Paling menentukan. Respons GI terhadap CV harus
   dicocokkan lewat identifikasi, bukan ditebak. Tanpa ini klaim superioritas saya
   tidak akan berdiri di sidang.
2. **Sapu bobot buat bikin frontier Pareto**, lalu sajikan itu, bukan klaim skalar.
   Frontier tidak bisa dipatahkan dengan "bobotnya kan Anda pilih sendiri".
3. **Perbaiki ARDS berat.** Kemungkinan butuh recruitment maneuver terjadwal sebagai
   aksi tersendiri, tidak cukup mengandalkan variabilitas.
4. **Hardware-in-the-loop** di PolyVent atau PVP1 dengan test lung QuickLung.
5. **Monte Carlo 200 pasien virtual** dengan domain randomization sesuai ISO
   80601-2-80.

---

## Cara menjalankan

```bash
python3 cek_instalasi.py      # beberapa detik, pastikan semua terpasang
python3 run_pareto.py         # ablasi variabilitas, sekitar 4 menit
python3 run_experiments.py    # matriks enam skenario, sekitar 5 menit
python3 make_figures.py       # tiga gambar
```

Simulator web: buka `viva-mpc-simulator.html` di peramban. Tidak perlu server.

---

## Catatan kejujuran

Semua angka di sini dari simulasi terhadap model yang saya tulis sendiri. Twin-nya
saya rancang supaya konsisten dengan arah temuan literatur, tapi dia bukan pasien dan
belum pernah dikalibrasi ke data manusia.

Yang pekerjaan ini tunjukkan: kerangka kendalinya bisa dibangun, jalan dalam waktu
yang masuk akal, dan memilih dosis variabilitas yang berbeda-beda menurut kondisi paru.

Yang **belum** ditunjukkan: bahwa dosis-dosis itu benar.
