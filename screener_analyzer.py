"""
╔══════════════════════════════════════════════════════════════════╗
║  SCREENER ANALYZER — analisa daftar saham hasil screener (IDX)    ║
║  Upload PDF/Excel hasil screener → analisa AI + infografis        ║
║                                                                    ║
║  Beda dari Flow Reader: ini baca DAFTAR saham (ranking by value/  ║
║  volume/movers), bukan broker summary. Fokus: momentum, gap,      ║
║  kekuatan close, volatilitas, mana yang worth di-watch.           ║
║                                                                    ║
║  Run:  streamlit run screener_analyzer.py                         ║
║  Dep:  pip install streamlit pandas plotly pillow pdfplumber openpyxl
║  Config: simpan key AI di .streamlit/secrets.toml (sama kayak Flow Reader)
╚══════════════════════════════════════════════════════════════════╝
"""
import os
import io
import re
import requests
import pandas as pd
import streamlit as st

GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")
GITHUB_MODEL = os.environ.get("GITHUB_MODEL", "openai/gpt-4o")
OLLAMA_MODEL = os.environ.get("OLLAMA_MODEL", "llama3.1")
OLLAMA_URL = os.environ.get("OLLAMA_URL", "http://localhost:11434")


# ════════════════════════════════════════════════════════════════════
# CREDENTIALS (baca dari secrets.toml / env, sama pola kayak Flow Reader)
# ════════════════════════════════════════════════════════════════════
def _secret(key, default=""):
    try:
        v = str(st.secrets.get(key, default)).strip()
        return default if v.startswith("PASTE_") else v
    except Exception:
        return default


def get_gemini_key():
    return (st.session_state.get("gemini_key", "").strip()
            or _secret("GEMINI_API_KEY") or os.environ.get("GEMINI_API_KEY", "").strip())


def get_github_token():
    return (st.session_state.get("github_token", "").strip()
            or _secret("GITHUB_TOKEN") or os.environ.get("GITHUB_TOKEN", "").strip())


def get_telegram_config():
    token = (st.session_state.get("tg_token", "").strip()
             or _secret("TELEGRAM_TOKEN") or os.environ.get("TELEGRAM_TOKEN", "").strip())
    chat = (st.session_state.get("tg_chat", "").strip()
            or _secret("TELEGRAM_CHAT_ID") or os.environ.get("TELEGRAM_CHAT_ID", "").strip())
    return token, chat


# ════════════════════════════════════════════════════════════════════
# UTIL — format angka
# ════════════════════════════════════════════════════════════════════
def fmt_value(v):
    """Format Rupiah ke T/B/jt."""
    try:
        v = float(v)
    except (ValueError, TypeError):
        return "-"
    sign = "-" if v < 0 else ""
    av = abs(v)
    if av >= 1e12:
        return f"{sign}{av/1e12:.2f}T"
    if av >= 1e9:
        return f"{sign}{av/1e9:.1f}B"
    if av >= 1e6:
        return f"{sign}{av/1e6:.0f}jt"
    return f"{sign}{av:,.0f}"


def to_num(x):
    """Konversi string angka (termasuk notasi '3.16E+10') ke float."""
    if x is None:
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip().replace(" ", "")
    if s in ("", "-", "—", "n/a", "N/A"):
        return 0.0
    # notasi ilmiah (3,16E+10 atau 3.16E+10)
    s = s.replace(",", ".") if re.match(r"^\d+,\d+E", s, re.I) else s
    # koma ribuan / desimal
    if "," in s and "." in s:
        s = s.replace(",", "")
    elif "," in s:
        after = s.split(",")[-1]
        s = s.replace(",", "") if len(after) == 3 else s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


# ════════════════════════════════════════════════════════════════════
# PARSING — baca PDF / Excel jadi DataFrame
# ════════════════════════════════════════════════════════════════════
# kolom yang dikenali (fleksibel — cocokin nama header apa pun)
COL_ALIASES = {
    "code": ["code", "kode", "ticker", "saham", "symbol", "stock"],
    "prev": ["prev", "previous", "prevclose", "close_prev"],
    "last": ["last", "lastprice"],
    "open": ["open", "buka"],
    "high": ["high", "tertinggi"],
    "low": ["low", "terendah"],
    "close": ["close", "tutup", "penutupan"],
    "volume": ["volume", "vol", "lot"],
    "value": ["value", "val", "nilai", "sortval"],
    "freq": ["freq", "frequency", "frekuensi"],
    "avg": ["avg", "average", "rata", "vwap"],
}


def _match_col(name):
    """Cocokin nama kolom ke kategori standar."""
    n = re.sub(r"[^a-z]", "", str(name).lower())
    for std, aliases in COL_ALIASES.items():
        if n in [re.sub(r"[^a-z]", "", a) for a in aliases]:
            return std
    return None


def normalize_df(df_raw):
    """Ubah df mentah (kolom apa pun) jadi df standar dengan kolom dikenali."""
    if df_raw is None or len(df_raw) == 0:
        return None
    colmap = {}
    for col in df_raw.columns:
        std = _match_col(col)
        if std and std not in colmap.values():
            colmap[col] = std
    if "code" not in colmap.values():
        return None  # gak ada kolom kode saham, gak bisa diproses
    out = df_raw.rename(columns=colmap)
    keep = [c for c in out.columns if c in COL_ALIASES]
    out = out[keep].copy()
    # konversi numerik
    for c in out.columns:
        if c != "code":
            out[c] = out[c].apply(to_num)
    out["code"] = out["code"].astype(str).str.strip().str.upper()
    # buang baris tanpa kode valid
    out = out[out["code"].str.match(r"^[A-Z]{2,6}$", na=False)].reset_index(drop=True)
    return out if len(out) > 0 else None


def _rows_from_text(text):
    """Parse teks 1 halaman jadi (header, list-of-rows). Return (None,None) kalau gagal."""
    lines = [l for l in text.split("\n") if l.strip()]
    if len(lines) < 2:
        return None, None
    # baris pertama yang kebanyakan kata = header
    header = re.split(r"\s+", lines[0].strip())
    rows = []
    for l in lines[1:]:
        parts = re.split(r"\s+", l.strip())
        if len(parts) >= 2:
            rows.append(parts)
    return header, rows


def _merge_split_pages(pages_text):
    """
    Gabung halaman yang kolomnya kepisah (page1=kolom kiri, page2=kolom kanan).
    Disambung horizontal berdasarkan urutan baris. Return DataFrame / None.
    """
    parsed = []
    for txt in pages_text:
        h, r = _rows_from_text(txt)
        if h and r:
            parsed.append((h, r))
    if not parsed:
        return None

    # kalau cuma 1 halaman → langsung
    if len(parsed) == 1:
        h, rows = parsed[0]
        smart = _smart_row_parse(h, rows)
        return smart if smart is not None else _rows_to_df(h, rows)

    # cek apakah halaman2 ini "split column" (jumlah baris data sama, header beda/komplementer)
    row_counts = [len(r) for _, r in parsed]
    # ambil jumlah baris terbanyak yang konsisten
    base_n = min(row_counts)
    all_same = max(row_counts) - min(row_counts) <= 1  # toleransi 1 baris (footer)

    # apakah ada halaman yang punya 'code' dan ada yang punya 'close/value'? → split column
    has_code = any(any(re.match(r"code|kode|ticker", c, re.I) for c in h) for h, _ in parsed)
    has_close = any(any(re.match(r"close|value|volume|last", c, re.I) for c in h) for h, _ in parsed)

    if all_same and has_code and len(parsed) >= 2:
        # SPLIT COLUMN: page kiri (ada code) + page kanan (close/value/dst)
        # parse halaman ber-'code' pakai smart parser (tahan Date+Time nempel)
        left_h, left_rows = None, None
        right_parts = []
        for h, rows in parsed:
            if any(re.match(r"code|kode|ticker", c, re.I) for c in h):
                left_h, left_rows = h, rows
            else:
                right_parts.append((h, rows))
        left_df = _smart_row_parse(left_h, left_rows) if left_h else None
        if left_df is None:
            # fallback: gabung horizontal mentah
            merged_header, merged_rows = [], [[] for _ in range(base_n)]
            for h, rows in parsed:
                merged_header.extend(h)
                for i in range(base_n):
                    merged_rows[i].extend(rows[i] if i < len(rows) else [""]*len(h))
            return _rows_to_df(merged_header, merged_rows)
        # sambung kolom kanan (close, value, dll) by urutan baris
        n = len(left_df)
        for h, rows in right_parts:
            for ci, col in enumerate(h):
                vals = []
                for i in range(n):
                    vals.append(rows[i][ci] if i < len(rows) and ci < len(rows[i]) else "")
                left_df[col] = vals
        return left_df

    # default: tumpuk vertikal (halaman lanjutan dengan kolom sama)
    h0, all_rows = parsed[0][0], []
    for h, rows in parsed:
        all_rows.extend(rows)
    # coba smart parse dulu (kalau ada code), else biasa
    smart = _smart_row_parse(h0, all_rows)
    return smart if smart is not None else _rows_to_df(h0, all_rows)


def _rows_to_df(header, rows):
    """Rakit DataFrame dari header + rows, samakan panjang kolom."""
    if not header or not rows:
        return None
    ncol = len(header)
    fixed = []
    for r in rows:
        if len(r) == ncol:
            fixed.append(r)
        elif len(r) > ncol:
            fixed.append(r[:ncol])
        else:
            fixed.append(list(r) + [""]*(ncol-len(r)))
    # buang kolom header duplikat (kasih suffix)
    seen = {}
    uniq_header = []
    for c in header:
        if c in seen:
            seen[c] += 1
            uniq_header.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            uniq_header.append(c)
    return pd.DataFrame(fixed, columns=uniq_header)


# pola kode saham IDX: 3-4 huruf kapital (kadang diikuti angka jarang)
_TICKER_RE = re.compile(r"^[A-Z]{3,4}$")


def _smart_row_parse(header, rows):
    """
    Parser tahan-banting buat baris yang jumlah tokennya gak konsisten sama header
    (cth: Date+Time nempel). Strategi: temukan posisi kolom KODE dengan nyari token
    yang cocok pola ticker saham, lalu map relatif dari situ.
    """
    # cari index kolom 'code' di header
    code_hidx = None
    for i, c in enumerate(header):
        if re.match(r"code|kode|ticker|symbol", str(c).strip(), re.I):
            code_hidx = i
            break
    if code_hidx is None:
        return None

    # di tiap baris, cari token yang match pola ticker → itu anchor 'code'
    parsed_rows = []
    for r in rows:
        code_pos = None
        for j, tok in enumerate(r):
            if _TICKER_RE.match(str(tok).strip()):
                code_pos = j
                break
        if code_pos is None:
            continue
        # bangun dict relatif: kolom setelah code map ke header setelah code_hidx
        row_dict = {}
        # kolom code & sesudahnya (paling penting: prev,last,close,dst ada di kanan)
        for k, hcol in enumerate(header[code_hidx:]):
            src_idx = code_pos + k
            row_dict[hcol] = r[src_idx] if src_idx < len(r) else ""
        parsed_rows.append(row_dict)

    if not parsed_rows:
        return None
    return pd.DataFrame(parsed_rows)


def parse_pdf(file_bytes):
    """Extract tabel dari PDF screener. Handle tabel normal + split-column antar halaman."""
    import pdfplumber
    with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
        all_tables = []
        pages_text = []
        for page in pdf.pages:
            tbls = page.extract_tables()
            all_tables.extend(tbls)
            pages_text.append(page.extract_text() or "")

    # kalau pdfplumber dapet tabel rapi, pakai itu
    if all_tables:
        rows = []
        for tbl in all_tables:
            if not tbl or len(tbl) < 2:
                continue
            h = tbl[0]
            for row in tbl[1:]:
                rows.append(dict(zip(h, row)))
        if rows:
            df = pd.DataFrame(rows)
            # cek apakah ada kolom code — kalau gak, fallback ke text merge
            if any(re.match(r"code|kode|ticker", str(c), re.I) for c in df.columns):
                return df

    # fallback / split-column: parse dari teks tiap halaman, gabung pintar
    return _merge_split_pages(pages_text)


def _parse_text_table(text):
    """Fallback lama: parse tabel dari teks mentah 1 blok."""
    h, rows = _rows_from_text(text)
    return _rows_to_df(h, rows) if h else None


def parse_excel(file_bytes, filename):
    """
    Baca Excel/CSV jadi DataFrame. ROBUST: file '.xls' dari app trading sering
    bukan Excel asli (kadang HTML table atau TSV). Coba beberapa cara berurutan.
    """
    fn = filename.lower()
    if fn.endswith(".csv"):
        # coba beberapa separator
        for sep in [",", ";", "\t"]:
            try:
                df = pd.read_csv(io.BytesIO(file_bytes), sep=sep)
                if len(df.columns) > 1:
                    return df
            except Exception:
                continue
        return pd.read_csv(io.BytesIO(file_bytes))

    errors = []

    # 1. coba xlsx modern (openpyxl)
    try:
        return pd.read_excel(io.BytesIO(file_bytes), engine="openpyxl")
    except Exception as e:
        errors.append(f"openpyxl: {e}")

    # 2. coba xls lama (xlrd) — kalau library ada
    try:
        return pd.read_excel(io.BytesIO(file_bytes), engine="xlrd")
    except ImportError:
        errors.append("xlrd belum keinstall")
    except Exception as e:
        errors.append(f"xlrd: {e}")

    # 3. coba sebagai HTML table (banyak '.xls' export sebenarnya HTML)
    try:
        tables = pd.read_html(io.BytesIO(file_bytes))
        if tables:
            # ambil tabel dengan kolom terbanyak (biasanya yang utama)
            return max(tables, key=lambda t: t.shape[1])
    except Exception as e:
        errors.append(f"html: {e}")

    # 4. coba sebagai teks TSV/CSV (delimited)
    try:
        text = file_bytes.decode("utf-8", errors="ignore")
        for sep in ["\t", ";", ",", "|"]:
            try:
                df = pd.read_csv(io.StringIO(text), sep=sep)
                if len(df.columns) > 1 and len(df) > 0:
                    return df
            except Exception:
                continue
    except Exception as e:
        errors.append(f"text: {e}")

    # semua gagal
    raise ValueError("Gak bisa baca file. Coba: " + " | ".join(errors[:3]))


def load_screener(file_bytes, filename):
    """Router: deteksi tipe file → parse → normalize. Return df standar / None + pesan."""
    try:
        ext = filename.lower().rsplit(".", 1)[-1]
        if ext == "pdf":
            raw = parse_pdf(file_bytes)
        elif ext in ("xlsx", "xls", "csv"):
            raw = parse_excel(file_bytes, filename)
        else:
            return None, f"Format .{ext} belum didukung. Pakai PDF, Excel, atau CSV."
        if raw is None or len(raw) == 0:
            return None, "Gak nemu data tabel di file. Pastikan filenya berisi tabel screener."
        norm = normalize_df(raw)
        if norm is None:
            return None, ("Tabel kebaca tapi gak nemu kolom kode saham. "
                          "Pastikan ada kolom Code/Kode/Ticker.")
        return norm, f"✓ {len(norm)} saham kebaca dari {filename}"
    except Exception as e:
        return None, f"Gagal baca file: {e}"


# ════════════════════════════════════════════════════════════════════
# ANALISA TEKNIKAL — hitung metrik per saham + insight agregat
# ════════════════════════════════════════════════════════════════════
def analyze_screener(df):
    """Hitung metrik teknikal per saham + ranking + insight."""
    d = df.copy()

    # pastikan kolom inti ada (isi default kalau gak ada)
    for c in ["prev", "last", "open", "high", "low", "close", "volume", "value", "freq", "avg"]:
        if c not in d.columns:
            d[c] = 0.0

    # kalau 'close' kosong tapi 'last' ada, pakai last
    d["close"] = d.apply(lambda r: r["close"] if r["close"] > 0 else r["last"], axis=1)
    ref = d["prev"].where(d["prev"] > 0, d["open"])  # acuan perubahan

    d["chg_pct"] = ((d["close"] - ref) / ref.replace(0, pd.NA) * 100).fillna(0)
    d["range_pct"] = ((d["high"] - d["low"]) / ref.replace(0, pd.NA) * 100).fillna(0)
    d["gap_pct"] = ((d["open"] - ref) / ref.replace(0, pd.NA) * 100).fillna(0)
    rng = (d["high"] - d["low"]).replace(0, pd.NA)
    d["close_pos"] = ((d["close"] - d["low"]) / rng).fillna(0.5).clip(0, 1)

    # klasifikasi sinyal sederhana
    def signal(r):
        if r["chg_pct"] > 3 and r["close_pos"] > 0.7:
            return "🟢 STRONG (naik + tutup kuat)"
        if r["chg_pct"] > 0 and r["close_pos"] > 0.7:
            return "🟢 Bullish (tutup di atas)"
        if r["chg_pct"] < -3 and r["close_pos"] < 0.3:
            return "🔴 WEAK (turun + tutup lemah)"
        if r["chg_pct"] < 0 and r["close_pos"] < 0.3:
            return "🔴 Bearish (tutup di bawah)"
        if r["range_pct"] > 8:
            return "🟡 Volatil (range lebar)"
        return "⚪ Netral"
    d["signal"] = d.apply(signal, axis=1)

    # sortir by value (likuiditas) default
    if d["value"].sum() > 0:
        d = d.sort_values("value", ascending=False).reset_index(drop=True)

    return d


def screener_insight(d):
    """Insight agregat dari screener. Return dict."""
    out = {}
    if len(d) == 0:
        return out
    out["n"] = len(d)
    out["top_gainer"] = (d.loc[d["chg_pct"].idxmax(), "code"], d["chg_pct"].max())
    out["top_loser"] = (d.loc[d["chg_pct"].idxmin(), "code"], d["chg_pct"].min())
    if d["value"].sum() > 0:
        out["most_liquid"] = (d.loc[d["value"].idxmax(), "code"], d["value"].max())
    out["most_volatile"] = (d.loc[d["range_pct"].idxmax(), "code"], d["range_pct"].max())
    out["n_up"] = int((d["chg_pct"] > 0).sum())
    out["n_down"] = int((d["chg_pct"] < 0).sum())
    out["n_flat"] = int((d["chg_pct"] == 0).sum())
    # saham menarik: naik + tutup kuat + likuid
    interesting = d[(d["chg_pct"] > 2) & (d["close_pos"] > 0.6)].nlargest(5, "value")
    out["interesting"] = interesting["code"].tolist()
    # tutup di high
    out["closed_high"] = d[d["close_pos"] > 0.8]["code"].tolist()
    return out


# ════════════════════════════════════════════════════════════════════
# AI BACKENDS (sama pola Flow Reader)
# ════════════════════════════════════════════════════════════════════
def call_gemini(prompt, max_tokens=8192, max_retries=5):
    import time
    key = get_gemini_key()
    if not key:
        return None
    models = [GEMINI_MODEL, "gemini-2.0-flash", "gemini-flash-latest"]
    headers = {"Content-Type": "application/json", "x-goog-api-key": key}
    body = {"contents": [{"role": "user", "parts": [{"text": prompt}]}],
            "generationConfig": {"temperature": 0.7, "maxOutputTokens": max_tokens,
                                 "thinkingConfig": {"thinkingBudget": 0}}}
    RETRY = {429, 500, 502, 503, 504}
    for model in models:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
        for attempt in range(max_retries):
            try:
                r = requests.post(url, headers=headers, json=body, timeout=90)
                if r.status_code == 200:
                    cands = r.json().get("candidates", [])
                    if not cands:
                        return "⚠️ Gemini gak balikin output."
                    parts = cands[0].get("content", {}).get("parts", [])
                    txt = "".join(p.get("text", "") for p in parts).strip()
                    return txt or "⚠️ Gemini balikin kosong."
                if r.status_code in RETRY and attempt < max_retries - 1:
                    time.sleep(min(2 ** attempt, 8)); continue
                if r.status_code in RETRY:
                    break
                return f"⚠️ Gemini error {r.status_code}: {r.text[:200]}"
            except requests.exceptions.Timeout:
                if attempt < max_retries - 1:
                    time.sleep(min(2 ** attempt, 8)); continue
                break
            except Exception as e:
                return f"⚠️ Gemini error: {e}"
    return "⚠️ GEMINI_BUSY"


def call_github_models(prompt, max_tokens=8192, max_retries=4):
    import time
    token = get_github_token()
    if not token:
        return None
    url = "https://models.github.ai/inference/chat/completions"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    body = {"model": GITHUB_MODEL, "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.7, "max_tokens": min(max_tokens, 16384)}
    for attempt in range(max_retries):
        try:
            r = requests.post(url, headers=headers, json=body, timeout=90)
            if r.status_code == 200:
                ch = r.json().get("choices", [])
                return ch[0].get("message", {}).get("content", "").strip() if ch else "⚠️ kosong"
            if r.status_code == 401:
                return "⚠️ GitHub Token salah/expired (perlu permission models:read)."
            if r.status_code == 429:
                if attempt < max_retries - 1:
                    time.sleep(min(2 ** attempt, 8)); continue
                return "⚠️ GitHub Models rate limit (free tier ~50/hari). Coba nanti / pakai Gemini."
            if r.status_code in {500, 502, 503, 504} and attempt < max_retries - 1:
                time.sleep(min(2 ** attempt, 8)); continue
            return f"⚠️ GitHub Models error {r.status_code}"
        except Exception as e:
            return f"⚠️ GitHub Models error: {e}"
    return "⚠️ GITHUB_BUSY"


def call_ollama(prompt, max_tokens=8192):
    try:
        r = requests.post(f"{OLLAMA_URL}/api/generate",
                          json={"model": OLLAMA_MODEL, "prompt": prompt, "stream": False,
                                "options": {"temperature": 0.7, "num_predict": max_tokens}},
                          timeout=180)
        if r.status_code == 200:
            return r.json().get("response", "").strip()
        return f"⚠️ Ollama error {r.status_code}"
    except Exception as e:
        return f"⚠️ Ollama gak konek ({e}). Pastikan Ollama jalan di {OLLAMA_URL}."


def call_ai(prompt, max_tokens=8192):
    backend = st.session_state.get("ai_backend", "Gemini (API)")
    if backend.startswith("Ollama"):
        return call_ollama(prompt, max_tokens)
    if backend.startswith("GitHub"):
        return call_github_models(prompt, max_tokens)
    return call_gemini(prompt, max_tokens)


def ai_ready():
    backend = st.session_state.get("ai_backend", "Gemini (API)")
    if backend.startswith("Ollama"):
        return True
    if backend.startswith("GitHub"):
        return bool(get_github_token())
    return bool(get_gemini_key())


# ════════════════════════════════════════════════════════════════════
# NARASI AI — analisa screener lengkap
# ════════════════════════════════════════════════════════════════════
def build_screener_table_text(d, limit=25):
    """Rakit tabel ringkas buat prompt AI."""
    lines = []
    for _, r in d.head(limit).iterrows():
        lines.append(
            f"{r['code']}: close {r['close']:.0f} ({r['chg_pct']:+.1f}%), "
            f"range {r['range_pct']:.1f}%, gap {r['gap_pct']:+.1f}%, "
            f"close_pos {r['close_pos']:.0%} (0=low 1=high), "
            f"val {fmt_value(r['value'])}, vol {fmt_value(r['volume'])}, "
            f"sinyal: {r['signal']}")
    return "\n".join(lines)


def generate_screener_narrative(d, insight, notes="", title="Screener IDX"):
    table = build_screener_table_text(d)
    notes_block = ""
    if notes and notes.strip():
        notes_block = f'\n⚠️ CATATAN TRADER (pertimbangkan): "{notes.strip()}"\n'

    prompt = f"""Lo analis pasar saham IDX (Indonesia) SENIOR. Gaya tajam, jujur, gak basa-basi, gak pom-pom. Bahasa Indonesia santai gaya 'bro' tapi berbobot.

Gue kasih DATA HASIL SCREENER ({title}) — daftar saham yang lagi rame/bergerak hari ini, diurut by value (likuiditas). Ini BUKAN data broker summary, jadi GAK ADA info bandar/broker. Yang ada: harga OHLC, volume, value, dan metrik teknikal turunan.

═══════════════════════════════
DATA SCREENER ({insight.get('n', 0)} saham):
{table}

RINGKASAN:
- Naik: {insight.get('n_up', 0)} | Turun: {insight.get('n_down', 0)} | Flat: {insight.get('n_flat', 0)}
- Top gainer: {insight.get('top_gainer', ('-', 0))[0]} ({insight.get('top_gainer', ('-', 0))[1]:+.1f}%)
- Top loser: {insight.get('top_loser', ('-', 0))[0]} ({insight.get('top_loser', ('-', 0))[1]:+.1f}%)
- Paling likuid: {insight.get('most_liquid', ('-', 0))[0]}
- Paling volatil: {insight.get('most_volatile', ('-', 0))[0]} ({insight.get('most_volatile', ('-', 0))[1]:.1f}% range)
- Tutup di high (bullish): {', '.join(insight.get('closed_high', [])) or '-'}
{notes_block}
═══════════════════════════════
TUGAS — tulis analisa screener LENGKAP & TAJAM dengan struktur ini:

## 1. GAMBARAN PASAR HARI INI
Baca mood pasar dari data: lebih banyak naik atau turun? Saham likuid lagi bullish atau bearish? Ada tema/sektor yang menonjol? Kasih konteks umum.

## 2. SAHAM YANG MENONJOL (Top Movers)
Bahas 3-5 saham paling menarik. Untuk tiap saham: kenapa menarik (momentum/gap/kekuatan close/volume), karakteristik pergerakannya, dan apa artinya. Fokus yang punya kombinasi bagus (naik + tutup kuat + likuid).

## 3. DETEKSI POLA
Identifikasi pola teknikal yang muncul:
- Gap up/down signifikan
- Saham tutup di high (close_pos tinggi = beli kuat sampai akhir) atau di low (jual sampai akhir)
- Volume/value spike (rame gak biasa)
- Volatilitas ekstrem (range lebar = ada pertarungan bandar vs ritel)

## 4. WATCHLIST BESOK
Dari data ini, saham mana yang worth dipantau besok & kenapa. Bedain:
- Yang momentum-nya kuat (potensi lanjut)
- Yang perlu konfirmasi (jangan FOMO)
- Yang dihindari (sinyal lemah)
Kasih level kunci kalau bisa (support/resistance dari high/low hari ini).

## 5. RISK NOTE
Catatan kehati-hatian. Ingat ini cuma snapshot 1 hari — bukan tren. Saham yang naik tinggi sehari bisa volatil/gorengan. Yang likuid lebih aman daripada yang tipis.

PENTING:
- Ini data SATU HARI (snapshot), bukan tren panjang. Jangan over-claim.
- Saham value/volume kecil = tipis = rawan gorengan, kasih warning.
- Jujur kalau data gak cukup buat kesimpulan kuat.
- Tutup dengan reminder: ini alat bantu screening, keputusan & risk management di tangan trader.
- JANGAN ngarang data broker/bandar — data ini gak punya itu."""

    return call_ai(prompt, max_tokens=8192)


# ════════════════════════════════════════════════════════════════════
# INFOGRAFIS (Pillow murni, HD) + TELEGRAM
# ════════════════════════════════════════════════════════════════════
def _hex(c):
    c = c.lstrip("#")
    if len(c) == 3:
        c = "".join(ch*2 for ch in c)
    return tuple(int(c[i:i+2], 16) for i in (0, 2, 4))


def _load_font(size, bold=False, mono=False):
    from PIL import ImageFont
    size = int(size)
    # 1. COBA FONT BUNDEL LOKAL DULU (folder fonts/ di sebelah script)
    #    ini jamin font selalu ada & bisa diperbesar di komputer mana pun
    try:
        _here = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        _here = os.getcwd()
    _fdir = os.path.join(_here, "fonts")
    if mono:
        local = "DejaVuSansMono-Bold.ttf" if bold else "DejaVuSansMono.ttf"
    elif bold:
        local = "DejaVuSans-Bold.ttf"
    else:
        local = "DejaVuSans.ttf"
    try:
        return ImageFont.truetype(os.path.join(_fdir, local), size)
    except Exception:
        pass

    # 2. fallback ke font sistem (banyak kandidat)
    if mono:
        cands = ["DejaVuSansMono-Bold.ttf" if bold else "DejaVuSansMono.ttf",
                 "LiberationMono-Bold.ttf" if bold else "LiberationMono-Regular.ttf",
                 "consolab.ttf" if bold else "consola.ttf",
                 "Menlo.ttc", "cour.ttf",
                 "DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf",
                 "LiberationSans-Bold.ttf" if bold else "LiberationSans-Regular.ttf"]
    elif bold:
        cands = ["DejaVuSans-Bold.ttf", "LiberationSans-Bold.ttf",
                 "Arial_Bold.ttf", "arialbd.ttf", "Helvetica-Bold.ttf"]
    else:
        cands = ["DejaVuSans.ttf", "LiberationSans-Regular.ttf",
                 "Arial.ttf", "arial.ttf", "Helvetica.ttf"]
    dirs = ["",
            "/usr/share/fonts/truetype/dejavu/",
            "/usr/share/fonts/truetype/liberation/",
            "/usr/share/fonts/TTF/",
            "/Library/Fonts/", "/System/Library/Fonts/",
            "C:/Windows/Fonts/", "C:\\Windows\\Fonts\\"]
    for name in cands:
        for dd in dirs:
            try:
                return ImageFont.truetype(dd + name, size)
            except Exception:
                continue
    # 3. paling akhir: default dengan ukuran (Pillow >=10)
    try:
        return ImageFont.load_default(size=size)
    except TypeError:
        return ImageFont.load_default()


def _clean(s):
    import re as _re
    if not s:
        return ""
    repl = {"—": "-", "–": "-", "−": "-", "'": "'", "'": "'", """: '"', """: '"',
            "…": "...", "•": "-", "≥": ">=", "≤": "<=", "→": "->"}
    for a, b in repl.items():
        s = s.replace(a, b)
    return _re.sub(r'[\U0001F000-\U0001FAFF\U00002600-\U000027BF\uFE0F\u2B00-\u2BFF]', '', s).strip()


def render_screener_infographic(d, insight, narrative, title="Screener IDX", scale=2):
    """Infografis screener HD pakai Pillow. Return PNG bytes."""
    from PIL import Image, ImageDraw
    S = scale
    BG = _hex("#0a0e14"); PANEL = _hex("#11171f"); BORDER = _hex("#2a3441")
    GOLD = _hex("#f0c040"); MUTE = _hex("#aeb9c4"); TXT = _hex("#f2f5f8")
    GREEN = _hex("#16c784"); RED = _hex("#f85149")

    class SD:
        def __init__(s, dr, sc): s._d, s._s = dr, sc
        def _sc(s, v):
            if isinstance(v, (int, float)): return v*s._s
            if isinstance(v, (list, tuple)): return type(v)(s._sc(x) for x in v)
            return v
        def text(s, xy, *a, **k): s._d.text(s._sc(xy), *a, **k)
        def line(s, xy, **k):
            if "width" in k: k["width"] = int(k["width"]*s._s) or 1
            s._d.line(s._sc(xy), **k)
        def rectangle(s, xy, **k): s._d.rectangle(s._sc(xy), **k)
        def rounded_rectangle(s, xy, radius=0, **k):
            if "width" in k: k["width"] = int(k["width"]*s._s) or 1
            s._d.rounded_rectangle(s._sc(xy), radius=s._sc(radius), **k)
        def textlength(s, t, font=None): return s._d.textlength(t, font=font)/s._s

    def F(sz, **kw): return _load_font(int(sz*S), **kw)
    f_brand = F(24, bold=True); f_title = F(52, bold=True); f_h = F(24, bold=True)
    f_big = F(46, bold=True, mono=True); f_body = F(21); f_bodyb = F(21, bold=True)
    f_small = F(17); f_mono = F(21, mono=True); f_monob = F(21, bold=True, mono=True)

    W, H = 1280, 4500
    img = Image.new("RGB", (W*S, H*S), BG)
    d_ = SD(ImageDraw.Draw(img), S)
    M = 32
    y = M

    def panel(x, yy, w, h, **k): d_.rounded_rectangle([x, yy, x+w, yy+h], radius=12,
                                                       fill=k.get("fill", PANEL),
                                                       outline=k.get("border", BORDER), width=1)

    # HEADER
    d_.text((M, y), "SCREENER ANALYZER", font=f_brand, fill=GOLD)
    d_.text((M, y+34), _clean(title), font=f_title, fill=TXT)
    nup, ndn = insight.get("n_up", 0), insight.get("n_down", 0)
    d_.text((W-M, y+10), f"{insight.get('n',0)} SAHAM", font=f_h, fill=MUTE, anchor="ra")
    d_.text((W-M, y+44), f"{nup} naik / {ndn} turun", font=f_body,
            fill=GREEN if nup >= ndn else RED, anchor="ra")
    y += 110
    d_.line([(M, y), (W-M, y)], fill=BORDER, width=2); y += 24

    # RINGKASAN CARDS (top gainer/loser/likuid/volatil)
    cards = [
        ("TOP GAINER", insight.get("top_gainer", ("-", 0)), GREEN, "%"),
        ("TOP LOSER", insight.get("top_loser", ("-", 0)), RED, "%"),
        ("PALING LIKUID", insight.get("most_liquid", ("-", 0)), GOLD, "val"),
        ("VOLATIL", insight.get("most_volatile", ("-", 0)), _hex("#e8b339"), "%"),
    ]
    cw = (W - 2*M - 3*14) // 4
    cx = M
    for label, (code, val), col, kind in cards:
        panel(cx, y, cw, 118)
        d_.text((cx+16, y+14), label, font=f_small, fill=MUTE)
        d_.text((cx+16, y+38), str(code), font=f_big, fill=col)
        sub = f"{val:+.1f}%" if kind == "%" else fmt_value(val)
        d_.text((cx+16, y+92), sub, font=f_small, fill=col)
        cx += cw + 14
    y += 140

    # TABEL SAHAM (top 15)
    rows = d.head(15)
    rowh = 40
    th = 64 + len(rows)*rowh + 16
    panel(M, y, W-2*M, th)
    d_.text((M+22, y+16), "DAFTAR SAHAM (by value)", font=f_h, fill=GOLD)
    hy = y + 58
    cols_x = [M+22, M+200, M+360, M+520, M+700, M+880]
    for hx, ht in zip(cols_x, ["CODE", "CLOSE", "CHG%", "RANGE%", "VALUE", "SINYAL"]):
        d_.text((hx, hy), ht, font=f_small, fill=MUTE)
    ry = hy + 32
    for _, r in rows.iterrows():
        chg_col = GREEN if r["chg_pct"] > 0 else RED if r["chg_pct"] < 0 else MUTE
        d_.text((cols_x[0], ry), str(r["code"]), font=f_monob, fill=TXT)
        d_.text((cols_x[1], ry), f"{r['close']:.0f}", font=f_mono, fill=TXT)
        d_.text((cols_x[2], ry), f"{r['chg_pct']:+.1f}", font=f_mono, fill=chg_col)
        d_.text((cols_x[3], ry), f"{r['range_pct']:.1f}", font=f_mono, fill=MUTE)
        d_.text((cols_x[4], ry), fmt_value(r["value"]), font=f_mono, fill=TXT)
        sig = _clean(r["signal"]).split("(")[0].strip()[:14]
        sig_col = GREEN if "STRONG" in r["signal"] or "Bullish" in r["signal"] else \
            RED if "WEAK" in r["signal"] or "Bearish" in r["signal"] else MUTE
        d_.text((cols_x[5], ry), sig, font=f_small, fill=sig_col)
        ry += rowh
    y += th + 22

    # NARASI AI
    if narrative and not narrative.startswith("⚠️"):
        clean = _clean(narrative.replace("##", "").replace("**", "").replace("*", ""))
        lines = []
        for para in clean.split("\n"):
            para = para.strip()
            if not para:
                lines.append(""); continue
            # wrap
            words, cur = para.split(), ""
            for w in words:
                if d_.textlength((cur+" "+w).strip(), f_body) <= W-2*M-44:
                    cur = (cur+" "+w).strip()
                else:
                    lines.append(cur); cur = w
            if cur: lines.append(cur)
        lh = 27
        nh = 60 + len(lines)*lh + 16
        panel(M, y, W-2*M, nh)
        d_.text((M+22, y+18), "ANALISA AI", font=f_h, fill=GOLD)
        ny = y + 58
        for ln in lines:
            d_.text((M+22, ny), ln, font=f_body, fill=_hex("#c9d1d9")); ny += lh
        y += nh + 18

    # FOOTER
    d_.line([(M, y), (W-M, y)], fill=BORDER, width=1); y += 12
    d_.text((W//2, y), "Dibuat oleh Screener Analyzer - alat bantu screening, BUKAN ajakan beli/jual",
            font=f_small, fill=MUTE, anchor="ma")
    y += 30

    img = img.crop((0, 0, W*S, min(int(y*S), H*S)))
    buf = io.BytesIO(); img.save(buf, "PNG")
    return buf.getvalue()


def png_to_pdf(png_bytes):
    from PIL import Image
    try:
        img = Image.open(io.BytesIO(png_bytes)).convert("RGB")
        buf = io.BytesIO(); img.save(buf, "PDF", resolution=100)
        return buf.getvalue()
    except Exception:
        return None


def _tg_post(url, **kwargs):
    try:
        r = requests.post(url, timeout=60, **kwargs)
        j = r.json()
        if j.get("ok"):
            return True, "ok"
        return False, f"[{j.get('error_code', r.status_code)}] {j.get('description', 'error')}"
    except Exception as e:
        return False, str(e)


def send_telegram_document(file_bytes, filename, caption="", as_photo=False):
    token, chat = get_telegram_config()
    if not token or not chat:
        return False, "Bot token / chat ID belum diisi."
    base = f"https://api.telegram.org/bot{token}"
    cap = _clean(caption)[:1024]
    mime = "application/pdf" if filename.lower().endswith(".pdf") else "image/png"
    field = "photo" if as_photo else "document"
    ep = "sendPhoto" if as_photo else "sendDocument"
    ok, detail = _tg_post(f"{base}/{ep}", data={"chat_id": chat, "caption": cap},
                          files={field: (filename, file_bytes, mime)})
    if not ok:
        d = detail.lower()
        if "chat not found" in d:
            return False, "Chat ID salah / belum chat bot duluan. " + detail
        if "unauthorized" in d:
            return False, "Bot token salah. " + detail
        return False, detail
    return True, "Terkirim ke Telegram ✓"


def telegram_test():
    token, chat = get_telegram_config()
    if not token: return False, "Bot Token belum diisi."
    if not chat: return False, "Chat ID belum diisi."
    base = f"https://api.telegram.org/bot{token}"
    ok, detail = _tg_post(f"{base}/getMe")
    if not ok:
        return False, f"Bot Token salah ({detail})."
    ok, detail = _tg_post(f"{base}/sendMessage",
                          data={"chat_id": chat, "text": "✅ Tes koneksi Screener Analyzer berhasil!"})
    if not ok:
        return False, f"Gagal kirim ({detail}). Cek Chat ID & pastikan udah chat bot duluan."
    return True, "✅ Koneksi OK! Cek HP lo."


# ════════════════════════════════════════════════════════════════════
# UI STREAMLIT
# ════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="Screener Analyzer", page_icon="📡", layout="wide")
st.markdown("""<style>
.metric-box{background:#11161f;border:1px solid #1e2530;border-radius:10px;padding:12px 14px;}
.metric-box .lbl{color:#8b949e;font-size:11px;text-transform:uppercase;}
.metric-box .val{font-size:22px;font-weight:700;font-family:Consolas,monospace;}
</style>""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("# 📡 Screener Analyzer")
    st.caption("Upload hasil screener (PDF/Excel) → analisa AI + infografis")

    st.markdown("---")
    st.markdown("### 🤖 MESIN AI")
    ai_backend = st.selectbox("Pilih AI",
                              ["Gemini (API)", "GitHub Models (GPT-4o)", "Ollama (lokal, gratis)"],
                              key="ai_backend")
    if ai_backend.startswith("Gemini"):
        if _secret("GEMINI_API_KEY") and not st.session_state.get("gemini_key", "").strip():
            st.caption(f"🔐 Key dari secrets.toml · `{GEMINI_MODEL}`")
            with st.expander("Ganti key sementara"):
                st.text_input("Gemini key", type="password", key="gemini_key")
        else:
            st.text_input("🔑 Gemini API key", type="password", key="gemini_key",
                          placeholder="AIzaSy...", help="Gratis dari aistudio.google.com")
            st.caption("✅ Key kebaca" if get_gemini_key() else "⚪ Belum ada key")
    elif ai_backend.startswith("GitHub"):
        if _secret("GITHUB_TOKEN") and not st.session_state.get("github_token", "").strip():
            st.caption(f"🔐 Token dari secrets.toml · `{GITHUB_MODEL}`")
            with st.expander("Ganti token sementara"):
                st.text_input("GitHub token", type="password", key="github_token")
        else:
            st.text_input("🔑 GitHub Token (PAT)", type="password", key="github_token",
                          placeholder="github_pat_...", help="permission models:read")
            st.caption("✅ Token kebaca" if get_github_token() else "⚪ Belum ada token")
    else:
        st.caption(f"🖥️ Ollama lokal · `{OLLAMA_MODEL}`")

    st.markdown("---")
    st.markdown("### 📤 TELEGRAM")
    if _secret("TELEGRAM_TOKEN") and not st.session_state.get("tg_token", "").strip():
        st.caption("🔐 dari secrets.toml")
        with st.expander("Ganti sementara"):
            st.text_input("Bot Token", type="password", key="tg_token")
            st.text_input("Chat ID", key="tg_chat")
    else:
        st.text_input("Bot Token", type="password", key="tg_token", placeholder="123:ABC...")
        st.text_input("Chat ID", key="tg_chat", placeholder="123456789")
    if all(get_telegram_config()):
        if st.button("🔌 Test koneksi", use_container_width=True):
            ok, info = telegram_test()
            (st.success if ok else st.error)(info)

st.markdown("# 📡 SCREENER ANALYZER")
st.caption("Upload hasil screener saham IDX (PDF/Excel/CSV) → analisa teknikal + AI + infografis. "
           "Beda dari Flow Reader: ini baca daftar movers/ranking, bukan broker summary.")

up = st.file_uploader("📁 Upload file screener", type=["pdf", "xlsx", "xls", "csv"],
                      help="Hasil export screener dari aplikasi trading lo (PDF, Excel, atau CSV).")

title = st.text_input("Judul laporan", value="Screener IDX",
                      placeholder="cth: Top Movers 26 Jun 2026")
notes = st.text_area("📝 Catatan / konteks (opsional)",
                     placeholder="cth: hari ini IHSG turun, fokus saham yang lawan arus")

if up is not None:
    file_bytes = up.read()
    df, msg = load_screener(file_bytes, up.name)
    if df is None:
        st.error(f"❌ {msg}")
    else:
        st.success(msg)
        if st.button("🔍 ANALISA SCREENER", type="primary", use_container_width=True):
            analyzed = analyze_screener(df)
            insight = screener_insight(analyzed)
            st.session_state["screener"] = {
                "df": analyzed, "insight": insight, "title": title, "notes": notes,
            }
            st.session_state.pop("scr_narr", None)

# render hasil dari session
S = st.session_state.get("screener")
if S:
    analyzed = S["df"]; insight = S["insight"]; title = S["title"]; notes = S["notes"]

    st.markdown(f"## 📊 {title}")
    # ringkasan cards
    c = st.columns(4)
    tg = insight.get("top_gainer", ("-", 0))
    tl = insight.get("top_loser", ("-", 0))
    ml = insight.get("most_liquid", ("-", 0))
    mv = insight.get("most_volatile", ("-", 0))
    c[0].markdown(f"<div class='metric-box'><div class='lbl'>Top Gainer</div>"
                  f"<div class='val' style='color:#16c784'>{tg[0]} {tg[1]:+.1f}%</div></div>", unsafe_allow_html=True)
    c[1].markdown(f"<div class='metric-box'><div class='lbl'>Top Loser</div>"
                  f"<div class='val' style='color:#f85149'>{tl[0]} {tl[1]:+.1f}%</div></div>", unsafe_allow_html=True)
    c[2].markdown(f"<div class='metric-box'><div class='lbl'>Paling Likuid</div>"
                  f"<div class='val' style='color:#f0c040'>{ml[0]}</div></div>", unsafe_allow_html=True)
    c[3].markdown(f"<div class='metric-box'><div class='lbl'>Naik/Turun</div>"
                  f"<div class='val'>{insight.get('n_up',0)}/{insight.get('n_down',0)}</div></div>", unsafe_allow_html=True)

    # chart: top movers by chg%
    st.markdown("### 📈 Top Movers (% perubahan)")
    import plotly.graph_objects as go
    mv_df = analyzed.reindex(analyzed["chg_pct"].abs().sort_values(ascending=False).index).head(12)
    fig = go.Figure(go.Bar(
        x=mv_df["code"], y=mv_df["chg_pct"],
        marker_color=["#16c784" if v > 0 else "#f85149" for v in mv_df["chg_pct"]],
        text=[f"{v:+.1f}%" for v in mv_df["chg_pct"]], textposition="outside"))
    fig.update_layout(height=300, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                      font=dict(color="#e6edf3"), margin=dict(l=10, r=10, t=10, b=10),
                      yaxis=dict(title="% chg", gridcolor="#1e2530"), xaxis=dict(showgrid=False))
    st.plotly_chart(fig, use_container_width=True)

    # tabel lengkap
    st.markdown("### 📋 Daftar Saham")
    show = analyzed[["code", "close", "chg_pct", "range_pct", "gap_pct", "value", "signal"]].copy()
    show.columns = ["Kode", "Close", "Chg%", "Range%", "Gap%", "Value", "Sinyal"]
    show["Value"] = analyzed["value"].map(fmt_value)
    show["Chg%"] = show["Chg%"].map(lambda v: f"{v:+.1f}")
    show["Range%"] = show["Range%"].map(lambda v: f"{v:.1f}")
    show["Gap%"] = show["Gap%"].map(lambda v: f"{v:+.1f}")
    show["Close"] = show["Close"].map(lambda v: f"{v:,.0f}")
    st.dataframe(show, use_container_width=True, hide_index=True)

    # NARASI AI
    st.markdown("### 🤖 ANALISA AI")
    if not ai_ready():
        st.info("💡 Isi API key AI di sidebar buat narasi. Tabel + chart tetap jalan tanpa ini.")
        narrative = None
    else:
        bname = ("GPT-4o" if ai_backend.startswith("GitHub")
                 else "Ollama" if ai_backend.startswith("Ollama") else "Gemini")
        ck = f"{title}|{len(analyzed)}|{hash(notes)}|{ai_backend}"
        if st.session_state.get("scr_narr_key") == ck and st.session_state.get("scr_narr"):
            narrative = st.session_state["scr_narr"]
        else:
            with st.spinner(f"{bname} lagi analisa screener..."):
                narrative = generate_screener_narrative(analyzed, insight, notes, title)
            st.session_state["scr_narr"] = narrative
            st.session_state["scr_narr_key"] = ck
        if narrative in ("⚠️ GEMINI_BUSY", "⚠️ GITHUB_BUSY"):
            st.warning("⏳ Server AI lagi sibuk. Coba lagi / ganti backend di sidebar.")
            if st.button("🔄 Coba lagi"):
                st.session_state.pop("scr_narr", None); st.rerun()
            narrative = None
        elif narrative and narrative.startswith("⚠️"):
            st.warning(narrative); narrative = None
        elif narrative:
            st.markdown(narrative)
            if st.button("🔄 Regenerate"):
                st.session_state.pop("scr_narr", None); st.rerun()

    # KIRIM TELEGRAM
    st.markdown("---")
    st.markdown("### 📤 KIRIM KE TELEGRAM")
    if not all(get_telegram_config()):
        st.info("💡 Isi Bot Token + Chat ID di sidebar buat ngirim.")
    else:
        fmt = st.radio("Format", ["🖼️ Infografis PNG", "📄 Infografis PDF"], horizontal=True)
        with st.spinner("Render infografis..."):
            try:
                png = render_screener_infographic(analyzed, insight,
                                                  st.session_state.get("scr_narr"), title)
            except Exception as e:
                png = None
                st.error(f"Gagal render: {e}")
        if png:
            st.image(png, caption="Preview infografis", use_container_width=True)
            st.download_button("⬇️ Download PNG", png, file_name=f"{title}_screener.png", mime="image/png")
            if st.button("📤 Kirim ke Telegram", type="primary", use_container_width=True):
                with st.spinner("Kirim..."):
                    is_pdf = fmt.startswith("📄")
                    cap = f"📡 {title} · {insight.get('n',0)} saham · {insight.get('n_up',0)} naik/{insight.get('n_down',0)} turun"
                    if is_pdf:
                        data = png_to_pdf(png)
                        ok, info = send_telegram_document(data, f"{title}_screener.pdf", cap, as_photo=False) if data else (False, "Gagal PDF")
                    else:
                        ok, info = send_telegram_document(png, f"{title}_screener.png", cap, as_photo=False)
                (st.success if ok else st.error)(f"{'✅' if ok else '❌'} {info}")

    st.caption("⚠️ Snapshot 1 hari, bukan tren. Alat bantu screening, BUKAN sinyal beli/jual. "
               "Saham value/volume kecil rawan gorengan. Validasi sendiri + risk management.")
else:
    st.info("👆 Upload file screener (PDF/Excel/CSV) buat mulai analisa.")
