# 📝 Template: Pembuatan Jurnal Ilmiah dari Skripsi menggunakan AI

> Panduan dan template project untuk mengonversi skripsi/tugas akhir menjadi artikel jurnal ilmiah menggunakan bantuan AI. Dapat digunakan untuk berbagai topik penelitian dan berbagai target jurnal.

---

## 📁 Struktur Folder Project

```
📦 [NAMA PROJECT JURNAL]/
│
├── 📁 Skripsi/
│   ├── NAMA_SKRIPSI_ANDA.docx        ← File skripsi asli
│   └── skripsi.md                    ← Hasil konversi (auto-generate)
│
├── 📁 Hasil Model/
│   ├── arsitektur_sistem.png         ← Diagram arsitektur / pipeline sistem
│   ├── [subfolder hasil training]/   ← Hasil evaluasi model (metrik, grafik)
│   └── [subfolder hasil testing]/    ← Hasil pengujian (latensi, benchmark)
│
├── 📁 Dataset/
│   ├── data.csv                      ← Dataset utama (atau format lain)
│   ├── ground_truth.csv              ← Ground truth untuk evaluasi
│   └── [gambar asli]/                ← Foto/citra dataset jika ada
│
├── 📁 Code Convert/
│   ├── convert_docx.py               ← Script konversi DOCX → Markdown
│   └── buat_jurnal.py                ← Script generator jurnal → Word (.docx)
│
├── 📁 Template Jurnal/
│   ├── NAMA_TEMPLATE.docx            ← Template resmi dari publisher
│   └── template.md                   ← Hasil konversi template (auto-generate)
│
├── 📁 Output/
│   ├── jurnal_[TOPIK].md             ← Draft jurnal (Markdown)
│   └── jurnal_[TOPIK].docx           ← Jurnal final (Word, siap submit)
│
└── 📄 README.md                      ← File ini
```

---

## 🚀 Cara Memulai (Setup dari Nol)

### Langkah 1 — Siapkan File

1. Copy file skripsi (`.docx`) ke folder **`Skripsi/`**
2. Copy template jurnal dari website publisher ke folder **`Template Jurnal/`**
3. Copy semua hasil eksperimen model ke folder **`Hasil Model/`**
   - Metrik evaluasi (`.txt`, `.csv`)
   - Grafik dan visualisasi (`.png`, `.jpg`)
   - Diagram arsitektur sistem (`.png`, `.jpg`)
4. Copy dataset atau file evaluasi ke folder **`Dataset/`**

### Langkah 2 — Install Dependensi

Buka terminal di folder project ini dan jalankan:

```bash
pip install python-docx
```

### Langkah 3 — Konversi File ke Markdown

Edit konfigurasi path di `Code Convert/convert_docx.py` lalu jalankan:

```bash
python "Code Convert/convert_docx.py"
```

Hasil konversi `.md` akan muncul di masing-masing folder.

### Langkah 4 — Mulai Sesi dengan AI

Buka AI coding assistant (Antigravity/Gemini/Claude/dll.) dan ikuti alur konsultasi di bawah.

---

## 💬 Alur Konsultasi dengan AI

### 📌 Prompt 1 — Pembuka Sesi

```
Tolong bantu saya untuk membuat artikel/jurnal dari skripsi saya.

Ketentuan penulisan ada di file: Template Jurnal/template.md
Skripsi saya ada di file: Skripsi/skripsi.md
Hasil eksperimen model ada di folder: Hasil Model/
Data evaluasi ada di folder: Dataset/

Tolong analisis semua file dan folder tersebut terlebih dahulu.
Jangan eksekusi apa-apa dulu — saya ingin berkonsultasi terlebih dahulu.
Saya ingin membuat jurnal/artikel yang berfokus pada [ASPEK AI / SISTEM / METODE]
yang saya terapkan.
```

---

### 📌 Prompt 2 — Menjawab Pertanyaan Konsultasi AI

AI akan mengajukan beberapa pertanyaan klarifikasi. Siapkan jawaban:

```
Jawaban konsultasi:

1. FOKUS JURNAL:
   Saya memilih [Opsi/Judul yang diinginkan].
   Contoh: "Pipeline Computer Vision Berbasis [MODEL] untuk [TUGAS] pada [DOMAIN]"

2. DATA MODEL:
   Hasil model ada di folder: Hasil Model/[nama subfolder]/
   File metrik utama: [nama file summary/evaluation]

3. DATA EVALUASI:
   Detail evaluasi ada di: Dataset/[nama file atau folder]

4. DATA LATENSI (jika ada):
   Dihitung dari output di folder: Hasil Model/[folder testing]
   Dilakukan [N] kali pengujian, diambil rata-ratanya.

5. IDENTITAS PENULIS:
   Nama     : [Nama Lengkap]
   Afiliasi : [Nama Prodi], [Nama Jurusan/Fakultas], [Nama Universitas]
   Email    : [email@domain.ac.id]

6. CO-AUTHOR:
   [Ada: nama pembimbing] / [Tidak ada]

7. ASET VISUAL:
   - Foto objek/data asli: ada di Dataset/[folder foto]
   - Diagram arsitektur: ada di Hasil Model/arsitektur_sistem.png

8. DETAIL DATASET:
   - Data asli: [N] sampel
   - Split: [X]% train / [Y]% val / [Z]% test
   - Augmentasi: [ya/tidak, jika ya: teknik dan faktor perbesaran]

9. JURNAL TARGET:
   [Nama jurnal, misal: JTIIK, Telematika, Jurnal RESTI, IEEE Access, dll.]
```

---

### 📌 Prompt 3 — Analisis Folder Spesifik

Jika AI perlu menganalisis folder tertentu:

```
Tolong buka dan analisis folder [NAMA FOLDER] secara lengkap.
Berikan ringkasan semua metrik yang relevan untuk dimasukkan ke dalam jurnal.
```

---

### 📌 Prompt 4 — Minta Planning Detail

```
Berdasarkan semua informasi yang sudah kita diskusikan, sekarang buatkan
2 planning lengkap dan sangat detail:

- Planning A: Output jurnal dalam format file Markdown (.md)
- Planning B: Output jurnal dalam format file Word (.docx)

Tampilkan struktur setiap bagian jurnal beserta sumber datanya.
Jangan eksekusi dulu — tampilkan planning-nya saja untuk saya review.
```

---

### 📌 Prompt 5 — Eksekusi Plan A (Draft Markdown)

```
Planning sudah saya setujui. Silakan mulai eksekusi Plan A:
Buatkan draft jurnal lengkap dalam format Markdown (.md) dan simpan
di folder Output/ dengan nama jurnal_[TOPIK].md
```

---

### 📌 Prompt 6 — Eksekusi Plan B (Word Final)

```
Draft Markdown sudah saya review dan setujui.
Silakan lanjutkan ke Plan B: buatkan script Python menggunakan
template di "Code Convert/buat_jurnal.py" untuk menghasilkan
file Word (.docx) berformat [NAMA JURNAL] yang siap disubmit.

Simpan output ke: Output/jurnal_[TOPIK].docx
```

---

## ✅ Informasi yang Perlu Disiapkan Sebelum Mulai

| Informasi | Wajib? | Keterangan |
|---|---|---|
| File skripsi `.docx` | ✅ Wajib | Sumber konten utama |
| Template jurnal `.docx` | ✅ Wajib | Ketentuan format penulisan |
| File metrik evaluasi model | ✅ Wajib | Data kinerja model (`.txt`/`.csv`) |
| Grafik hasil evaluasi | ✅ Wajib | Gambar untuk jurnal (`.png`/`.jpg`) |
| Diagram arsitektur sistem | ✅ Wajib | Gambar pipeline/arsitektur |
| Foto data asli / dataset | 🟡 Disarankan | Ilustrasi tantangan dataset |
| Data latensi / waktu inferensi | 🟡 Disarankan | Klaim real-time feasibility |
| Identitas penulis & afiliasi | ✅ Wajib | Header jurnal |
| Nama co-author / pembimbing | 🔵 Opsional | Jika ada |

---

## 📊 Template Data Eksperimen

### Format Folder `Hasil Model/` yang Direkomendasikan

```
Hasil Model/
├── arsitektur_sistem.png              ← WAJIB — diagram pipeline/arsitektur
│
├── model_a/                           ← Hasil training/evaluasi Model A
│   ├── summary.txt                    ← Ringkasan metrik utama
│   ├── evaluation_results.txt         ← Detail hasil evaluasi test set
│   ├── confusion_matrix.png           ← Confusion matrix
│   ├── training_curves.png            ← Grafik loss & akurasi training
│   ├── sample_prediction.jpg          ← Contoh hasil prediksi
│   └── error_analysis/
│       └── error_chart.png            ← Grafik distribusi error
│
├── model_b/                           ← (ulangi untuk setiap model)
│   └── ...
│
└── testing_performance/               ← Hasil pengujian latensi
    └── performance_log.txt            ← Log waktu per request API
```

### Format File `summary.txt` yang Direkomendasikan

```
Model Name    : [Nama Model]
Dataset       : [N_train] train / [N_val] val / [N_test] test images
Epochs        : [N]
Batch Size    : [N]
Image Size    : [NxN]

--- METRICS (Test Set) ---
Precision     : 0.XXXX
Recall        : 0.XXXX
mAP@50        : 0.XXXX
mAP@50-95     : 0.XXXX
Inference     : XX.XX ms/image
```

### Format File `rekap_evaluasi.csv` yang Direkomendasikan

```csv
Model,Metrik1,Metrik2,Metrik3,Metrik4
Model A,XX.XX,XX.XX,XX.XX,XX.XX
Model B,XX.XX,XX.XX,XX.XX,XX.XX
Model C,XX.XX,XX.XX,XX.XX,XX.XX
```

---

## ⚙️ Penggunaan Script

### `convert_docx.py` — Konversi Word ke Markdown

```bash
# Edit path di bagian bawah script terlebih dahulu, lalu:
python "Code Convert/convert_docx.py"
```

Konfigurasi di dalam script:
```python
# Konversi skripsi
docx_to_markdown(
    docx_path = r'Skripsi\NAMA_SKRIPSI.docx',    # <-- ganti ini
    md_path   = r'Skripsi\skripsi.md'
)

# Konversi template jurnal
docx_to_markdown(
    docx_path = r'Template Jurnal\TEMPLATE.docx', # <-- ganti ini
    md_path   = r'Template Jurnal\template.md'
)
```

### `buat_jurnal.py` — Generate Jurnal Word

```bash
# Edit KONFIGURASI di bagian atas script, isi konten jurnal,
# lalu jalankan:
python "Code Convert/buat_jurnal.py"
```

Konfigurasi minimal di dalam script:
```python
JUDUL_ID      = "Judul Jurnal Dalam Bahasa Indonesia"
JUDUL_EN      = "Journal Title in English"
PENULIS       = "Nama Lengkap Anda"
UNIVERSITAS   = "Nama Universitas"
PRODI_JURUSAN = "Nama Program Studi / Jurusan"
EMAIL         = "email@anda.ac.id"
OUTPUT_DOCX   = r"Output\jurnal_output.docx"
```

---

## ⚠️ Tips & Catatan Penting

### Menangani File Skripsi yang Sangat Besar

Jika skripsi berisi banyak gambar, file `.md` hasil konversi bisa sangat besar (>20 MB karena Base64). Cara mengatasinya:

1. **Minta AI membaca per bab** — berikan instruksi untuk membaca bagian tertentu saja
2. **Aktifkan pembersihan Base64** di `convert_docx.py`:
   ```python
   # Hapus comment pada baris ini di convert_docx.py:
   content = re.sub(r'data:image/[^;]+;base64,[A-Za-z0-9+/=]{500,}', '[GAMBAR]', content)
   ```
3. **Ekstrak bab penting** ke file terpisah (misal: `bab4_metode.txt`, `bab5_hasil.txt`)

### Konsistensi Nama Model

Pastikan nama model konsisten di seluruh dokumen:
- ✅ Gunakan nama resmi dari paper/dokumentasi model
- ❌ Jangan mencampurkan nama tidak resmi atau singkatan berbeda

### Novelty dari Trade-off Analysis

Jika sistem Anda memilih komponen yang **tidak paling unggul di metrik** (karena alasan praktis), jadikan ini sebagai **kontribusi novelty** dengan mendokumentasikan alasan trade-off:
- Kemandirian sistem (*offline capability*)
- Efisiensi biaya (*zero operational cost*)
- Privasi data (tidak mengirim data ke server eksternal)
- Stabilitas latensi (deterministik vs. bergantung jaringan)
- Keandalan di lapangan

### Mendokumentasikan Dataset Split dengan Benar

Urutkan operasi dengan jelas untuk menghindari ambiguitas:
```
1. Kumpulkan data asli (N total)
2. Split TERLEBIH DAHULU: X% train / Y% val / Z% test
3. Augmentasi HANYA pada bagian train
4. Catat jumlah akhir setiap split
```

### Format Referensi Harvard-Anglia

```
Penulis, A., & Penulis, B. (Tahun). Judul artikel dalam kalimat biasa.
Nama Jurnal, Volume(Issue), Halaman–Halaman.

Penulis, C. (Tahun). Judul buku. Edisi ke-X. Nama Penerbit.
```

---

## 📋 Checklist Sebelum Submit

- [ ] Abstrak sudah dalam 2 bahasa (Indonesia + Inggris), masing-masing 150–250 kata
- [ ] Semua gambar memiliki nomor dan keterangan (*caption*)
- [ ] Semua tabel memiliki nomor dan judul (di atas tabel)
- [ ] Seluruh gambar dan tabel direferensikan dalam teks (misal: "lihat Gambar 1")
- [ ] Daftar pustaka lengkap dan sesuai format jurnal target
- [ ] Semua referensi dalam daftar pustaka benar-benar dikutip di teks
- [ ] Tidak ada referensi yang dikutip di teks tapi tidak ada di daftar pustaka
- [ ] Jumlah halaman sesuai batas maksimum jurnal
- [ ] Nama file dan format file sudah sesuai ketentuan submission

---

## 📦 Dependensi

```bash
pip install python-docx   # Wajib: untuk konversi dan generate Word
pip install mammoth       # Opsional: alternatif konversi DOCX → HTML/MD
```

---

*Template ini dapat digunakan kembali untuk proyek jurnal lain dengan menyesuaikan konten dan konfigurasi.*
