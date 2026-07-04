# IDX Watchlist Report Generator

Generate infografis watchlist HD (PNG) + PDF otomatis dari hasil screener
`.xlsx` / `.xls` / `.csv` / `.pdf` — versi reusable dari yang dibikinin manual
sebelumnya, tinggal jalanin dari CLI atau panggil dari script lain (bot,
cron, GitHub Actions, dsb).

## Install

```bash
pip install -r requirements.txt --break-system-packages

# dependency sistem (WAJIB, di luar pip) -- ini yang render HTML jadi gambar/PDF
sudo apt-get update && sudo apt-get install -y wkhtmltopdf
```

Kalau mau IHSG auto-fetch (opsional): `pip install yfinance`. Kalau nggak
diinstall / nggak ada internet, section konteks pasar cuma di-skip otomatis
(nggak bikin error, nggak fabricate angka).

## Narasi: rule-based (default) atau AI (opsional)

Default: narasi "kenapa masuk prioritas / kenapa waspada" dibuat **rule-based**
dari angka -- gratis, konsisten, jalan tanpa internet.

Opsional: kalau dikasih API key, narasi per saham digenerate LLM:

| Backend | Dari mana ambil key | Model default |
|---|---|---|
| Gemini | https://aistudio.google.com/apikey (gratis) | `gemini-2.5-flash` |
| GitHub Models | GitHub Settings -> Developer settings -> Personal access tokens, centang permission **models: read** (gratis, rate-limited) | `openai/gpt-4o` |

```bash
export GEMINI_API_KEY=xxxx
python idx_watchlist_report.py input.xlsx --ai-backend gemini

export GITHUB_MODELS_TOKEN=ghp_xxxx
python idx_watchlist_report.py input.xlsx --ai-backend github

# --ai-backend auto (DEFAULT): pakai Gemini kalau GEMINI_API_KEY ke-set,
# lanjut GitHub Models kalau Gemini gagal/kosong, lanjut rule-based kalau
# dua-duanya gagal/kosong. Nggak perlu ubah command apapun -- cukup export
# env var yang mana yang mau dipakai.
```

Override model spesifik: `--gemini-model gemini-3-flash-preview` atau
`--github-model openai/gpt-4.1` (atau env `GEMINI_MODEL` / `GITHUB_MODEL`) --
model API ini cepat berubah, jadi sengaja dibikin gampang diganti tanpa edit
script.

**Yang PENTING dipahami:** skor & flag risiko (VOLATILITAS EKSTREM dkk) selalu
dihitung rule-based dari angka, apapun mode narasinya -- AI cuma ganti kalimat
penjelasannya, bukan logic keputusannya. Kalau API gagal/timeout/rate-limit,
otomatis fallback ke rule-based per saham, TIDAK bikin script crash -- aman
dipakai di cron/GitHub Actions tanpa pengawasan. AI diinstruksikan keras untuk
tidak mengarang fakta di luar angka yang dikasih (lihat `AI_SYSTEM_PROMPT` di
script kalau mau tuning gaya bahasanya).

## Jalan di cloud tanpa buka PC (Streamlit Community Cloud, gratis)

Ini alur "buka app dari HP, upload file dari DSI, generate, kirim ke Telegram" --
tanpa perlu laptop/PC nyala sama sekali.

**1. Push ke GitHub.** Repo harus berisi:
```
streamlit_app.py
idx_watchlist_report.py
ticker_names.json
requirements.txt
packages.txt
fonts/
  Poppins-Regular.ttf
  Poppins-Medium.ttf
  Poppins-Bold.ttf
```

**2. Deploy.** Buka https://share.streamlit.io -> New app -> pilih repo & branch
-> main file path: `streamlit_app.py` -> Deploy.

**3. Isi Secrets (opsional tapi direkomendasikan).** Di halaman app -> menu titik
tiga -> Settings -> Secrets, isi:
```toml
GEMINI_API_KEY = "xxxx"
GITHUB_MODELS_TOKEN = "ghp_xxxx"
TELEGRAM_BOT_TOKEN = "123456:xxxx"
TELEGRAM_CHAT_ID = "-100xxxxxxxxxx"
```
Kalau ini diisi, field-field itu otomatis ke-prefill tiap buka app -- nggak perlu
ketik ulang tiap kali. Kalau nggak diisi, tinggal ketik manual di sidebar app
tiap sesi (aman, nggak disimpan ke mana-mana selain sesi browser lu).

**4. Selesai.** URL app-nya bisa dibuka dari HP kapan aja. Alurnya: buka app ->
upload file export dari DSI -> klik Generate -> preview + download muncul ->
isi caption -> klik "Kirim Sekarang" -> masuk ke grup Telegram.

### Kenapa perlu `packages.txt`
`wkhtmltoimage` (yang dipakai buat render infografis) itu bukan library Python,
jadi nggak cukup taruh di `requirements.txt`. Streamlit Cloud punya mekanisme
terpisah: apapun yang ditulis di `packages.txt` bakal di-`apt-get install`
sebelum app-nya jalan. File `packages.txt` di sini isinya:
```
wkhtmltopdf
xvfb
```
`xvfb` ini jaring pengaman: `wkhtmltoimage` dibangun di atas engine yang secara
default mengharapkan ada display (X11), padahal server cloud itu headless
(nggak ada monitor/display). Script sudah otomatis coba mode "offscreen" dulu
(nggak butuh xvfb), dan baru fallback ke `xvfb-run` kalau itu gagal -- jadi
kombinasi ini aman dipasang di `packages.txt` bahkan kalau ternyata nggak
kepake.

### Kenapa ada folder `fonts/`
Server cloud kemungkinan besar nggak punya font Poppins ter-install. Supaya
tampilan infografis tetap konsisten (bukan fallback ke font default yang
kurang match sama desainnya), font di-embed langsung ke HTML sebagai base64 --
makanya file font-nya harus ikut di-push ke repo, bukan cuma di-reference by
name.

### Batasan Streamlit Community Cloud (free tier)
- App "tidur" kalau lama nggak diakses -- buka pertama kali abis lama nganggur
  bisa makan waktu ~30 detik buat "bangun". Normal, tunggu aja.
- RAM terbatas (~1 GB). Render skala 2.0-2.5 harusnya aman; kalau sering
  gagal/timeout pas file-nya banyak saham, turunin ke `scale=1.5` di sidebar.
- Ini BUKAN buat automation terjadwal (jalan sendiri tiap jam sekian tanpa ada
  yang buka app-nya) -- itu di luar cakupan Streamlit, based on alur yang lu
  mau (manual upload -> manual generate -> manual kirim), ini sudah pas.

## Pakai (CLI, lokal/Railway/GitHub Actions)

```bash
# paling simpel -- auto-detect judul & tanggal dari nama file
python idx_watchlist_report.py input.xlsx

# PDF juga sekalian, dengan konteks IHSG manual
python idx_watchlist_report.py input.xlsx --ihsg-close 5876 --ihsg-change 2.28

# dari .csv atau .pdf, top 6, folder output custom
python idx_watchlist_report.py input.csv --top 6 --outdir hasil_senin
python idx_watchlist_report.py input.pdf

# HD ekstra tajam (file lebih besar, render lebih lama)
python idx_watchlist_report.py input.xlsx --scale 3

# skip PDF, PNG doang, lebih cepat
python idx_watchlist_report.py input.xlsx --no-pdf
```

Semua flag: `python idx_watchlist_report.py --help`

## Dipanggil dari script lain (bot Telegram, cron, dsb)

```python
from idx_watchlist_report import generate_report

result = generate_report(
    input_path="screener_hari_ini.xlsx",
    outdir="outputs",
    top_n=5,
    ihsg_close=5876, ihsg_change=2.28,
    scale=2.0,
    ai_backend="auto",             # atau "gemini" / "github" / "none"
    gemini_key="...",              # opsional, bisa juga lewat env GEMINI_API_KEY
    github_token="...",            # opsional, bisa juga lewat env GITHUB_MODELS_TOKEN
)
# result["png"], result["pdf"], result["html"], result["dataframe"], result["ai_backend_used"]
# -> tinggal result["png"] dikirim ke Telegram pakai bot lu yang udah ada
```

## Format file input

Kolom **wajib**: `Code`, `Close`/`Last`, `High`, `Low` (nama kolom fleksibel,
lihat `COLUMN_ALIASES` di script -- case-insensitive, boleh pakai nama
Indonesia juga misalnya `Kode`, `Harga`).

Kolom **opsional** (kalau nggak ada, otomatis fallback + kasih warning di
terminal, bukan error):
- `Prev` -> kalau nggak ada, `%Chg` dihitung dari `Open`
- `Avg`/`VWAP` -> kalau nggak ada, dihitung proxy `(High+Low+Close)/3`
- `Value` -> kalau nggak ada tapi ada `Volume`, dihitung `Volume x Close`
- `Freq` -> dipakai buat info avg lot size, nggak wajib

Untuk `.pdf`: harus PDF berisi tabel data asli (hasil export komputer), bukan
hasil scan/foto. Kalau screenernya bisa export `.xlsx`/`.csv`, pakai itu --
jauh lebih akurat daripada parsing tabel dari PDF.

## Nambah nama emiten / catatan khusus per ticker

Edit `ticker_names.json`. Contoh:

```json
"MMIX": {
  "name": "Multi Medika Internasional",
  "note": "Pernah kena UMA BEI, histori spekulatif -- size kecil."
}
```

Ticker yang belum ada di file ini otomatis muncul sebagai "Saham XXXX" di
infografis (nggak bikin error), dan script kasih tau di terminal ticker mana
aja yang belum punya nama -- tinggal isi kapan aja, otomatis kepake di run
berikutnya.

## Ubah bobot skor / threshold risiko

Semua di bagian atas `idx_watchlist_report.py` (`SCORE_WEIGHTS`,
`RISK_VWAP_THRESHOLD`, dst) -- tinggal edit angkanya, nggak perlu ubah logic
lain.

## Kejujuran soal narasi otomatis

Narasi "kenapa masuk prioritas" / "kenapa waspada" di infografis dibuat
**rule-based dari angka** (close position, VWAP, range, likuiditas) -- BUKAN
hasil riset berita/fundamental kayak histori suspensi BEI, UMA, dsb. Kalau
mau nambahin insight spesifik semacam itu, isi field `note` di
`ticker_names.json` (lihat contoh MMIX di atas) -- itu satu-satunya tempat
buat nyuntik pengetahuan manual ke laporan otomatis ini.

## Kalau `.pdf` tidak ke-detect

Beberapa PDF hasil scan/gambar (bukan teks asli) tidak bisa diparsing
`pdfplumber`. Errornya bakal jelas ("Gagal menemukan tabel di PDF...").
Solusinya: export ulang dari platform screenernya sebagai `.xlsx` atau `.csv`.

## File di folder ini

- `idx_watchlist_report.py` -- script utama (core logic, dipakai CLI maupun Streamlit)
- `streamlit_app.py` -- UI web (upload -> generate -> kirim Telegram), buat deploy di Streamlit Cloud
- `ticker_names.json` -- database nama emiten (extendable)
- `fonts/` -- Poppins TTF, di-embed base64 ke tiap infografis (biar konsisten di semua host)
- `requirements.txt` -- dependency python
- `packages.txt` -- dependency sistem buat Streamlit Cloud (wkhtmltopdf, xvfb)
- `README.md` -- ini
