#!/usr/bin/env python3
"""
IDX Watchlist Report Generator
================================
Bikin infografis watchlist otomatis dari hasil screener (TQL/DSI atau
screener sejenis) dalam format .xlsx, .xls, .csv, atau .pdf.

CONTOH PAKAI
    python idx_watchlist_report.py input.xlsx
    python idx_watchlist_report.py input.pdf --top 5
    python idx_watchlist_report.py input.csv --ihsg-close 5876 --ihsg-change 2.28
    python idx_watchlist_report.py input.xlsx --scale 3 --outdir hasil

    # narasi AI (opsional) -- Gemini atau GPT-4 via GitHub Models
    export GEMINI_API_KEY=xxxx
    python idx_watchlist_report.py input.xlsx --ai-backend gemini
    export GITHUB_MODELS_TOKEN=ghp_xxxx   # PAT dengan permission "models: read"
    python idx_watchlist_report.py input.xlsx --ai-backend github

OUTPUT (di folder --outdir, default "outputs/")
    <nama_file>_HD.png   -> infografis resolusi tinggi, siap kirim
    <nama_file>.pdf       -> versi PDF (satu halaman panjang, identik dgn PNG)
    <nama_file>_report.html -> HTML mentahnya (buat debug/edit manual)

DEPENDENSI SISTEM (di luar pip)
    Cuma butuh Pango (untuk WeasyPrint) -- lihat packages.txt kalau deploy di
    Streamlit Community Cloud. Ubuntu/Debian:
        sudo apt-get update && sudo apt-get install -y libpango-1.0-0 libpangoft2-1.0-0 libharfbuzz-subset0
    PyMuPDF (buat konversi PDF->PNG) TIDAK butuh apt sama sekali, murni wheel pip.
    Font Poppins di-embed base64 langsung dari folder fonts/ -- nggak
    bergantung font ter-install di OS host.

DEPENDENSI PYTHON (lihat requirements.txt)
    pandas, openpyxl, xlrd, matplotlib, pillow, pdfplumber, requests,
    weasyprint, pymupdf
    yfinance (OPSIONAL — kalau ada & ada koneksi internet, IHSG auto-fetch)

NARASI: RULE-BASED (default) ATAU AI (opsional)
    Default: narasi "kenapa masuk prioritas / kenapa waspada" dibuat RULE-BASED
    dari angka (close position, VWAP, range, likuiditas) -- konsisten, gratis,
    tidak butuh internet/API key.

    Opsional: kalau GEMINI_API_KEY atau GITHUB_MODELS_TOKEN di-set (atau pakai
    --ai-backend), narasi per saham digenerate oleh LLM (Gemini atau GPT-4o
    lewat GitHub Models -- gratis, cukup GitHub PAT dengan permission
    "models: read", lihat docs.github.com/en/github-models). Flag risiko
    (VOLATILITAS EKSTREM dkk) TETAP dihitung rule-based apapun mode-nya --
    yang berubah cuma narasi kalimatnya. Kalau API gagal/rate-limit/timeout,
    otomatis fallback ke rule-based per-saham, TIDAK bikin script crash.
    AI dilarang mengarang fakta di luar angka yang dikasih (lihat AI_SYSTEM_PROMPT).

CATATAN JUJUR
    Baik versi rule-based maupun AI, ini BUKAN riset berita/fundamental per
    saham (misal: histori UMA/suspensi BEI, capital structure, dsb) -- itu
    cuma bisa dari `note` manual di ticker_names.json (lihat contoh di file
    itu), karena AI di sini cuma dikasih angka hasil scan, bukan akses berita.
"""

from __future__ import annotations

import argparse
import base64
import io
import json
import os
from dataclasses import dataclass
from datetime import datetime, date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

import matplotlib
matplotlib.use("Agg")
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt
from matplotlib.patches import Patch

from PIL import Image
import fitz  # PyMuPDF -- rasterisasi PDF -> PNG, tanpa dependency sistem tambahan
from weasyprint import HTML as WeasyHTML
from weasyprint.text.fonts import FontConfiguration

# ══════════════════════════════════════════════════════════════════════════
# KONFIGURASI — silakan diutak-atik sesuai selera / gaya trading lu
# ══════════════════════════════════════════════════════════════════════════

SCORE_WEIGHTS = {
    "close_pos": 0.35,   # kekuatan close dalam range harian (Wyckoff sign-of-strength)
    "vwap":      0.25,   # posisi close vs VWAP / harga rata-rata
    "momentum":  0.20,   # % kenaikan harian (di-cap biar ga didominasi 1 saham gocap)
    "liquidity": 0.20,   # nilai transaksi, log-scale
}
MOMENTUM_CAP_PCT = 13.0   # clip Chg% ke sini sebelum dinormalisasi 0-100
VWAP_CLIP_PCT    = 3.0    # clip VsVWAP% ke [-x, +x] sebelum dinormalisasi

RISK_VWAP_THRESHOLD  = -5.0   # VsVWAP% di bawah ini -> flag "DISTRIBUSI KUAT"
RISK_CLOSEPOS_THRESH = 0.20   # ClosePos di bawah ini -> flag "CLOSE LEMAH"
RISK_RANGE_THRESHOLD = 15.0   # Range% di atas ini -> flag "VOLATILITAS EKSTREM"

IHSG_GREEN_THRESHOLD = 1.0    # %chg IHSG >= ini -> regime GREEN (aturan override Roy)
IHSG_RED_THRESHOLD   = -1.0   # %chg IHSG <= ini -> regime RED

TOP_N_DEFAULT  = 5
RENDER_SCALE_DEFAULT = 2.0    # >1 = HD asli (rasterisasi PDF di DPI lebih tinggi via PyMuPDF,
                               # BUKAN upscale gambar). 2.0 kira-kira setara "retina".
BASE_WIDTH_PX  = 1080          # lebar desain logis (referensi 1x, dalam CSS px)
PAGE_MAX_HEIGHT_PX = 26000     # plafon tinggi halaman WeasyPrint -- lebih dari cukup untuk
                               # laporan terpanjang sekalipun; kelebihan otomatis di-crop
                               # setelah render (lihat render_report()).

# --- AI narrative (opsional) --------------------------------------------
GEMINI_MODEL_DEFAULT = "gemini-2.5-flash"
GITHUB_MODEL_DEFAULT = "openai/gpt-4o"
GEMINI_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
GITHUB_MODELS_ENDPOINT = "https://models.github.ai/inference/chat/completions"

AI_SYSTEM_PROMPT = (
    "Kamu adalah analis teknikal saham IDX yang paham metodologi Wyckoff dan "
    "bandarmologi (analisis broker flow). Untuk SETIAP saham di data JSON yang "
    "diberikan user, tulis narasi 2-3 kalimat, Bahasa Indonesia gaya trader "
    "profesional (santai tapi kompeten, bukan bahasa formal kaku ala siaran pers). "
    "Interpretasikan angka yang diberikan: close_pos_pct (posisi close dalam "
    "range harian, makin tinggi makin kuat ala Wyckoff sign-of-strength), "
    "vwap_pct (posisi close vs harga rata-rata transaksi), chg_pct (kenaikan "
    "harian), range_pct (lebar range harian), value_miliar_rp (likuiditas), "
    "flag (risiko yang sudah terdeteksi rule-based), kategori (prioritas_watchlist "
    "atau waspada). JANGAN mengarang fakta perusahaan/berita/histori yang tidak "
    "ada di data ini. Balas HANYA JSON valid dengan format persis "
    "{\"KODE_SAHAM\": \"narasi\", ...} untuk semua kode di data -- tanpa markdown, "
    "tanpa code fence, tanpa teks apapun di luar objek JSON tsb."
)

HERE = Path(__file__).resolve().parent
TICKER_DB_PATH = HERE / "ticker_names.json"
FONTS_DIR = HERE / "fonts"

HARI_ID  = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
BULAN_ID = ["", "Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
            "Agustus", "September", "Oktober", "November", "Desember"]

BG, PANEL, GRID = "#0b1220", "#0f1a2e", "#22304a"
TXT, MUT = "#e8edf5", "#8a97ad"
GREEN, GOLD, ORANGE, RED = "#2ee6a6", "#f5b942", "#f57c42", "#ef4a5f"


def fmt_tanggal_id(d: date) -> str:
    return f"{HARI_ID[d.weekday()]}, {d.day} {BULAN_ID[d.month]} {d.year}"


def next_trading_day(d: date) -> date:
    """Hari bursa berikutnya, skip Sabtu/Minggu. TIDAK memperhitungkan libur
    nasional/cuti bersama BEI -- cek manual kalau mepet tanggal merah."""
    nd = d + timedelta(days=1)
    while nd.weekday() >= 5:
        nd += timedelta(days=1)
    return nd


# ══════════════════════════════════════════════════════════════════════════
# 1) LOAD & NORMALISASI DATA (.xlsx/.xls/.csv/.pdf)
# ══════════════════════════════════════════════════════════════════════════

COLUMN_ALIASES = {
    "code":   ["code", "kode", "ticker", "stock", "symbol", "saham"],
    "prev":   ["prev", "previous", "prevclose", "prevclose", "closeprev"],
    "close":  ["close", "last", "harga", "closeprice"],
    "open":   ["open"],
    "high":   ["high"],
    "low":    ["low"],
    "volume": ["volume", "vol"],
    "avg":    ["avg", "vwap", "average", "avgprice"],
    "value":  ["value", "val", "nilai", "turnover"],
    "freq":   ["freq", "frequency", "frek", "frekuensi"],
}
REQUIRED_COLS = ["code", "close", "high", "low"]


def _norm(s: str) -> str:
    return str(s).strip().lower().replace(" ", "").replace("_", "")


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    normed = {_norm(c): c for c in df.columns}
    colmap = {}
    for std_name, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in normed:
                colmap[normed[alias]] = std_name
                break
    out = df.rename(columns=colmap)

    missing = [c for c in REQUIRED_COLS if c not in out.columns]
    if missing:
        raise ValueError(
            f"Kolom wajib tidak ditemukan: {missing}. "
            f"Kolom yang terbaca dari file: {list(df.columns)}"
        )

    if "prev" not in out.columns:
        print("  [warn] kolom Prev tidak ada -> %Chg pakai Open sebagai proxy")
        out["prev"] = out.get("open", out["close"])
    if "open" not in out.columns:
        out["open"] = out["prev"]
    if "avg" not in out.columns:
        print("  [warn] kolom Avg/VWAP tidak ada -> pakai proxy (High+Low+Close)/3")
        out["avg"] = (pd.to_numeric(out["high"], errors="coerce")
                       + pd.to_numeric(out["low"], errors="coerce")
                       + pd.to_numeric(out["close"], errors="coerce")) / 3
    if "volume" not in out.columns:
        out["volume"] = np.nan
    if "value" not in out.columns:
        if out["volume"].notna().any():
            print("  [warn] kolom Value tidak ada -> dihitung dari Volume * Close")
            out["value"] = pd.to_numeric(out["volume"], errors="coerce") * pd.to_numeric(out["close"], errors="coerce")
        else:
            print("  [warn] kolom Value & Volume tidak ada -> likuiditas diasumsikan rata")
            out["value"] = 1.0
    if "freq" not in out.columns:
        out["freq"] = np.nan

    for c in ["prev", "close", "open", "high", "low", "volume", "avg", "value", "freq"]:
        out[c] = pd.to_numeric(
            out[c].astype(str).str.replace(",", "", regex=False) if out[c].dtype == object else out[c],
            errors="coerce",
        )

    out = out.dropna(subset=["close", "high", "low"]).reset_index(drop=True)
    out["code"] = out["code"].astype(str).str.strip().str.upper()
    out = out[out["code"].str.len() > 0].reset_index(drop=True)
    return out


def load_xlsx(path: Path) -> pd.DataFrame:
    return pd.read_excel(path)


def load_csv(path: Path) -> pd.DataFrame:
    for sep in [",", ";", "\t"]:
        try:
            df = pd.read_csv(path, sep=sep, engine="python")
            if df.shape[1] > 1:
                return df
        except Exception:
            continue
    return pd.read_csv(path)


def load_pdf(path: Path) -> pd.DataFrame:
    import pdfplumber
    header, rows = None, []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                if not table:
                    continue
                if header is None:
                    header, body = table[0], table[1:]
                else:
                    body = table[1:] if table[0] == header else table
                rows.extend(body)
    if header is None or not rows:
        raise ValueError(
            "Gagal menemukan tabel di PDF. Pastikan PDF berisi tabel data asli "
            "(bukan hasil scan/foto). Kalau screener-nya bisa export .xlsx/.csv, "
            "pakai itu saja -- jauh lebih akurat."
        )
    return pd.DataFrame(rows, columns=header)


def load_data(path: Path) -> pd.DataFrame:
    ext = path.suffix.lower()
    print(f"[1/6] Membaca file: {path.name} ({ext})")
    loaders = {".xlsx": load_xlsx, ".xls": load_xlsx, ".csv": load_csv, ".pdf": load_pdf}
    if ext not in loaders:
        raise ValueError(f"Format {ext} belum didukung. Pakai .xlsx / .xls / .csv / .pdf")
    raw = loaders[ext](path)
    df = normalize_columns(raw)
    print(f"  -> {len(df)} baris saham terbaca: {', '.join(df['code'].tolist())}")
    return df


# ══════════════════════════════════════════════════════════════════════════
# 2) METRIK & SKOR
# ══════════════════════════════════════════════════════════════════════════

def compute_metrics(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["chg_pct"] = (df["close"] - df["prev"]) / df["prev"].replace(0, np.nan) * 100
    df["range_pct"] = (df["high"] - df["low"]) / df["low"].replace(0, np.nan) * 100
    span = (df["high"] - df["low"]).replace(0, np.nan)
    df["close_pos"] = ((df["close"] - df["low"]) / span).fillna(0.5).clip(0, 1)
    df["vwap_pct"] = (df["close"] - df["avg"]) / df["avg"].replace(0, np.nan) * 100
    df["value_b"] = df["value"] / 1e9
    df["avg_lot"] = df["volume"] / df["freq"]
    df["chg_pct"] = df["chg_pct"].fillna(0)
    df["vwap_pct"] = df["vwap_pct"].fillna(0)
    return df


def compute_score(df: pd.DataFrame, weights: dict = None) -> pd.DataFrame:
    w = weights or SCORE_WEIGHTS
    df = df.copy()
    df["s_close"] = df["close_pos"] * 100
    df["s_vwap"] = ((df["vwap_pct"] + VWAP_CLIP_PCT) / (2 * VWAP_CLIP_PCT)).clip(0, 1) * 100
    df["s_mom"] = (df["chg_pct"].clip(0, MOMENTUM_CAP_PCT) / MOMENTUM_CAP_PCT) * 100

    valueb = df["value_b"].clip(lower=0.001)
    lv = np.log10(valueb)
    lv_range = lv.max() - lv.min()
    df["s_liq"] = ((lv - lv.min()) / lv_range * 100) if lv_range > 0 else pd.Series(50.0, index=df.index)

    df["score"] = (
        df["s_close"] * w["close_pos"]
        + df["s_vwap"] * w["vwap"]
        + df["s_mom"] * w["momentum"]
        + df["s_liq"] * w["liquidity"]
    )

    def _flag(r):
        flags = []
        if r["vwap_pct"] < RISK_VWAP_THRESHOLD:
            flags.append("DISTRIBUSI KUAT")
        if r["close_pos"] < RISK_CLOSEPOS_THRESH:
            flags.append("CLOSE LEMAH")
        if r["range_pct"] > RISK_RANGE_THRESHOLD:
            flags.append("VOLATILITAS EKSTREM")
        return ", ".join(flags) if flags else "-"

    df["flag"] = df.apply(_flag, axis=1)
    df = df.sort_values("score", ascending=False).reset_index(drop=True)
    df["final_rank"] = df.index + 1
    return df


def tier_color(row) -> str:
    if row["flag"] != "-" and row["score"] < 30:
        return RED
    if row["flag"] != "-":
        return ORANGE
    if row["score"] >= 55:
        return GREEN
    if row["score"] >= 40:
        return GOLD
    return MUT


# ══════════════════════════════════════════════════════════════════════════
# 3) TICKER NAME LOOKUP (opsional, extendable)
# ══════════════════════════════════════════════════════════════════════════

def load_ticker_db() -> dict:
    if TICKER_DB_PATH.exists():
        try:
            return json.loads(TICKER_DB_PATH.read_text())
        except Exception as e:
            print(f"  [warn] gagal baca {TICKER_DB_PATH.name}: {e}")
    return {}


def ticker_info(code: str, db: dict) -> dict:
    entry = db.get(code.upper())
    if entry is None:
        return {"name": f"Saham {code.upper()}", "note": None}
    if isinstance(entry, str):
        return {"name": entry, "note": None}
    return {"name": entry.get("name", f"Saham {code.upper()}"), "note": entry.get("note")}


# ══════════════════════════════════════════════════════════════════════════
# 4) KONTEKS PASAR (IHSG) — opsional, manual override atau auto via yfinance
# ══════════════════════════════════════════════════════════════════════════

@dataclass
class MarketContext:
    available: bool
    close: float = None
    change_pct: float = None
    source: str = "unavailable"

    @property
    def regime(self) -> str:
        if not self.available:
            return "UNKNOWN"
        if self.change_pct >= IHSG_GREEN_THRESHOLD:
            return "GREEN"
        if self.change_pct <= IHSG_RED_THRESHOLD:
            return "RED"
        return "NEUTRAL"


def get_market_context(manual_close=None, manual_change=None) -> MarketContext:
    if manual_close is not None and manual_change is not None:
        return MarketContext(True, manual_close, manual_change, "manual")
    try:
        import yfinance as yf
        hist = yf.Ticker("^JKSE").history(period="5d")
        if len(hist) >= 2:
            last, prev = float(hist["Close"].iloc[-1]), float(hist["Close"].iloc[-2])
            return MarketContext(True, last, (last - prev) / prev * 100, "yfinance")
    except Exception as e:
        print(f"  [info] auto-fetch IHSG via yfinance gagal ({e}) -> section konteks pasar dilewati")
    return MarketContext(False)


# ══════════════════════════════════════════════════════════════════════════
# 4b) NARASI AI (opsional) — Gemini atau GPT-4o via GitHub Models
#     Rule-based tetap jadi fallback kalau ini gagal/nggak dikonfigurasi.
# ══════════════════════════════════════════════════════════════════════════

def call_gemini(system_prompt: str, user_prompt: str, api_key: str,
                 model: str = GEMINI_MODEL_DEFAULT, timeout: int = 45) -> str:
    import requests
    url = GEMINI_ENDPOINT.format(model=model)
    body = {
        "systemInstruction": {"parts": [{"text": system_prompt}]},
        "contents": [{"role": "user", "parts": [{"text": user_prompt}]}],
        "generationConfig": {"temperature": 0.4, "maxOutputTokens": 3000},
    }
    resp = requests.post(url, params={"key": api_key}, json=body, timeout=timeout)
    resp.raise_for_status()
    data = resp.json()
    return data["candidates"][0]["content"]["parts"][0]["text"]


def call_github_models(system_prompt: str, user_prompt: str, token: str,
                        model: str = GITHUB_MODEL_DEFAULT, timeout: int = 45) -> str:
    import requests
    body = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "temperature": 0.4,
    }
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "Content-Type": "application/json",
    }
    resp = requests.post(GITHUB_MODELS_ENDPOINT, headers=headers, json=body, timeout=timeout)
    resp.raise_for_status()
    data = resp.json()
    return data["choices"][0]["message"]["content"]


def _parse_ai_json(text: str) -> dict:
    t = text.strip()
    if t.startswith("```"):
        t = t.strip("`")
        if t.lower().startswith("json"):
            t = t[4:]
    t = t.strip()
    return json.loads(t)


def _build_ai_payload(df: pd.DataFrame, top_n: int) -> list:
    top_df = df.head(top_n)
    top_codes = set(top_df["code"])
    risk_df = df[(df["flag"] != "-") & (~df["code"].isin(top_codes))]
    rows = []
    for _, r in pd.concat([top_df, risk_df]).iterrows():
        rows.append({
            "code": r["code"],
            "chg_pct": round(float(r["chg_pct"]), 2),
            "close_pos_pct": round(float(r["close_pos"]) * 100),
            "vwap_pct": round(float(r["vwap_pct"]), 2),
            "range_pct": round(float(r["range_pct"]), 2),
            "value_miliar_rp": round(float(r["value_b"]), 1),
            "score": round(float(r["score"])),
            "flag": r["flag"],
            "kategori": "prioritas_watchlist" if r["code"] in top_codes else "waspada",
        })
    return rows


def get_ai_narratives(df: pd.DataFrame, top_n: int, backend: str = "auto",
                       gemini_key: str = None, gemini_model: str = GEMINI_MODEL_DEFAULT,
                       github_token: str = None, github_model: str = GITHUB_MODEL_DEFAULT):
    """Return (narratives_dict, backend_used_or_None). TIDAK PERNAH raise --
    kegagalan apapun (key kosong, network error, JSON invalid, rate limit)
    bikin fungsi ini return ({}, None) supaya caller fallback ke rule-based.
    'auto' coba Gemini dulu (kalau key ada), lanjut GitHub Models kalau Gemini
    gagal/key kosong (kalau token ada)."""
    if backend == "none":
        return {}, None

    if backend == "auto":
        candidates = [c for c, key in [("gemini", gemini_key), ("github", github_token)] if key]
    elif backend in ("gemini", "github"):
        candidates = [backend]
    else:
        print(f"  [warn] --ai-backend '{backend}' tidak dikenal -> pakai rule-based")
        return {}, None

    if not candidates:
        return {}, None

    payload_rows = _build_ai_payload(df, top_n)
    if not payload_rows:
        return {}, None
    user_prompt = "Data saham (JSON):\n" + json.dumps(payload_rows, ensure_ascii=False)

    for cand in candidates:
        try:
            if cand == "gemini":
                if not gemini_key:
                    print("  [warn] --ai-backend gemini tapi GEMINI_API_KEY kosong, skip")
                    continue
                raw = call_gemini(AI_SYSTEM_PROMPT, user_prompt, gemini_key, gemini_model)
            else:
                if not github_token:
                    print("  [warn] --ai-backend github tapi GITHUB_MODELS_TOKEN kosong, skip")
                    continue
                raw = call_github_models(AI_SYSTEM_PROMPT, user_prompt, github_token, github_model)

            narratives = _parse_ai_json(raw)
            if not isinstance(narratives, dict) or not narratives:
                raise ValueError("respons AI bukan JSON object berisi narasi")
            narratives = {str(k).upper(): str(v) for k, v in narratives.items()}
            print(f"  [info] narasi AI via {cand} berhasil untuk {len(narratives)} saham")
            return narratives, cand
        except Exception as e:
            print(f"  [warn] narasi AI via {cand} gagal ({e}) -> coba fallback berikutnya")
            continue

    print("  [info] semua backend AI gagal/tidak dikonfigurasi -> pakai narasi rule-based")
    return {}, None


# ══════════════════════════════════════════════════════════════════════════
# 5) FILE-NAME METADATA (judul, tanggal data, tanggal trading berikutnya)
# ══════════════════════════════════════════════════════════════════════════

def parse_filename_meta(path: Path, title_override: str = None) -> dict:
    parts = path.stem.split("__")
    slot = screener = source = None
    dt = None
    if len(parts) >= 3:
        slot, screener, source = parts[0].replace("_", " "), parts[1].replace("_", " "), parts[2]
        if len(parts) >= 4:
            try:
                dt = datetime.strptime(parts[3], "%Y%m%d_%H%M%S")
            except ValueError:
                dt = None
    if dt is None:
        dt = datetime.fromtimestamp(path.stat().st_mtime)

    trade_date = dt.date()
    next_day = next_trading_day(trade_date)
    return {
        "title": title_override or (screener or "WATCHLIST REPORT"),
        "eyebrow": " · ".join(x for x in [source, slot, screener] if x) or "SCREENER REPORT",
        "data_date_str": fmt_tanggal_id(trade_date),
        "data_time_str": dt.strftime("%H:%M WIB"),
        "next_day_str": fmt_tanggal_id(next_day),
    }


# ══════════════════════════════════════════════════════════════════════════
# 6) CHARTS (matplotlib, dark theme)
# ══════════════════════════════════════════════════════════════════════════

def _register_fonts():
    search_dirs = [FONTS_DIR, Path("/usr/share/fonts/truetype/google-fonts/"),
                   Path("/usr/share/fonts/truetype/poppins/")]
    for fname in ["Poppins-Regular.ttf", "Poppins-Medium.ttf", "Poppins-Bold.ttf"]:
        for base in search_dirs:
            fp = base / fname
            if fp.exists():
                try:
                    fm.fontManager.addfont(str(fp))
                except Exception:
                    pass
                break
    try:
        plt.rcParams["font.family"] = "Poppins"
    except Exception:
        pass


def chart_quadrant(df: pd.DataFrame, out_path: Path, scale: float = 1.0):
    fig, ax = plt.subplots(figsize=(10, 8.6), dpi=200 * scale)
    fig.patch.set_facecolor(BG)
    ax.set_facecolor(PANEL)

    x, y = df["chg_pct"], df["close_pos"] * 100
    sizes = (df["value_b"].clip(lower=0.01) ** 0.5) * 34 + 90
    colors = df.apply(tier_color, axis=1)

    ax.scatter(x, y, s=sizes, c=colors, alpha=0.92, edgecolors="white", linewidths=0.9, zorder=3)
    for _, r in df.iterrows():
        ax.annotate(r["code"], (r["chg_pct"], r["close_pos"] * 100), textcoords="offset points",
                    xytext=(0, 11), ha="center", fontsize=12.5, fontweight="bold", color=TXT, zorder=4)

    ax.axhline(70, color=GRID, lw=1.1, ls=(0, (5, 4)), zorder=1)
    ax.axvline(float(x.median()), color=GRID, lw=1.1, ls=(0, (5, 4)), zorder=1)
    ax.text(0.985, 0.975, "ZONA KUAT\n(Close tinggi + Momentum solid)", transform=ax.transAxes,
            ha="right", va="top", fontsize=10, color=GREEN, fontweight="bold", linespacing=1.5)
    ax.text(0.015, 0.03, "ZONA DISTRIBUSI\n(Close lemah = waspada)", transform=ax.transAxes,
            ha="left", va="bottom", fontsize=10, color=RED, fontweight="bold", linespacing=1.5)

    ax.set_xlabel("Kenaikan Harga Hari Ini (%)", fontsize=12.5, color=TXT, labelpad=10)
    ax.set_ylabel("Kekuatan Close (posisi close dalam range hari ini, %)", fontsize=12.5, color=TXT, labelpad=10)
    ax.set_title(f"PETA KEKUATAN vs MOMENTUM — {len(df)} Saham", fontsize=16.5, color=TXT, fontweight="bold", pad=18)
    ax.tick_params(colors=MUT, labelsize=11)
    for spine in ax.spines.values():
        spine.set_color(GRID)
    ax.grid(True, color=GRID, alpha=0.35, zorder=0)
    ax.set_ylim(-2, 108)
    ax.set_xlim(float(x.min()) - 1.5, float(x.max()) + 1.5)

    plt.tight_layout()
    plt.savefig(out_path, facecolor=BG, bbox_inches="tight")
    plt.close(fig)


def chart_ranking(df: pd.DataFrame, out_path: Path, scale: float = 1.0):
    n = len(df)
    fig_h = max(6.0, 0.5 * n + 2.2)
    fig, ax = plt.subplots(figsize=(10, fig_h), dpi=200 * scale)
    fig.patch.set_facecolor(BG)
    ax.set_facecolor(PANEL)

    dfr = df.sort_values("score", ascending=True)
    colors = dfr.apply(tier_color, axis=1)
    bars = ax.barh(dfr["code"], dfr["score"], color=colors, height=0.62, zorder=3)
    for bar, score in zip(bars, dfr["score"]):
        ax.text(bar.get_width() + 1.3, bar.get_y() + bar.get_height() / 2, f"{score:.0f}",
                va="center", ha="left", fontsize=11, color=TXT, fontweight="bold")

    ax.set_xlim(0, max(88, dfr["score"].max() + 15))
    ax.set_xlabel("Skor Kekuatan Komposit (0–100)", fontsize=12.5, color=TXT, labelpad=10)
    ax.set_title("RANKING KEKUATAN SETUP — Semua Saham", fontsize=16.5, color=TXT, fontweight="bold", pad=16)
    ax.tick_params(colors=TXT, labelsize=12.5)
    ax.tick_params(axis="x", colors=MUT, labelsize=11)
    for spine in ax.spines.values():
        spine.set_color(GRID)
    ax.grid(True, axis="x", color=GRID, alpha=0.35, zorder=0)

    legend_elems = [
        Patch(facecolor=GREEN, label="Prioritas Kuat"),
        Patch(facecolor=GOLD, label="Layak Dipantau"),
        Patch(facecolor=ORANGE, label="Ada Flag Risiko"),
        Patch(facecolor=RED, label="Waspada/Hindari"),
    ]
    ax.legend(handles=legend_elems, loc="lower right", fontsize=10.5, facecolor=PANEL,
              edgecolor=GRID, labelcolor=TXT, framealpha=0.9)

    plt.tight_layout()
    plt.savefig(out_path, facecolor=BG, bbox_inches="tight")
    plt.close(fig)


def make_charts(df: pd.DataFrame, workdir: Path, scale: float = 1.0) -> dict:
    _register_fonts()
    quad = workdir / "_chart_quadrant.png"
    rank = workdir / "_chart_ranking.png"
    chart_quadrant(df, quad, scale=scale)
    chart_ranking(df, rank, scale=scale)
    return {"quadrant": quad.name, "ranking": rank.name}


# ══════════════════════════════════════════════════════════════════════════
# 7) HTML BUILDER
# ══════════════════════════════════════════════════════════════════════════

CSS = """
  * { margin:0; padding:0; box-sizing:border-box; font-family:'Poppins', sans-serif; }
  body { width:1080px; background: linear-gradient(180deg, #070c16 0%, #0b1220 6%, #0b1220 94%, #070c16 100%); color:#e8edf5; }
  .wrap { padding: 56px 56px 40px 56px; }
  .eyebrow { display:inline-block; letter-spacing:2.5px; font-size:14px; font-weight:600; color:#2ee6a6;
    background:rgba(46,230,166,0.10); border:1px solid rgba(46,230,166,0.35); padding:7px 16px; border-radius:20px; margin-bottom:22px; }
  h1.title { font-size:58px; font-weight:800; line-height:1.05; letter-spacing:-1px;
    color:#f5f8fd; margin-bottom:10px; }
  .subtitle { font-size:21px; color:#9fb0c9; font-weight:500; margin-bottom:2px;}
  .subtitle b { color:#e8edf5; }
  .sourceline { font-size:14.5px; color:#5f6d84; margin-top:8px; font-weight:500;}
  .ctx { margin-top:32px; border-radius:22px; padding:26px 30px;
    background: linear-gradient(135deg, rgba(46,230,166,0.10), rgba(74,157,245,0.07)); border:1px solid rgba(255,255,255,0.08); }
  .ctx-top { display:flex; align-items:center; justify-content:space-between; margin-bottom:14px;}
  .ctx-label { font-size:15px; font-weight:700; color:#8a97ad; letter-spacing:1.5px;}
  .regime-pill { font-size:14.5px; font-weight:700; color:#0b1220; padding:6px 16px; border-radius:20px; }
  .ctx-headline { font-size:26px; font-weight:700; color:#f2f6fb; margin-bottom:10px; }
  .ctx-headline .up { color:#2ee6a6; } .ctx-headline .down { color:#ef4a5f; }
  .ctx-detail { font-size:15.5px; color:#aab6c9; line-height:1.65; }
  .ctx-detail b { color:#dbe4f0; }
  .sec-title { display:flex; align-items:baseline; gap:14px; margin:44px 0 20px 0; }
  .sec-title .num { font-size:15px; font-weight:800; color:#0b1220; background:#2ee6a6; width:34px; height:34px; border-radius:10px; display:flex; align-items:center; justify-content:center; }
  .sec-title h2 { font-size:28px; font-weight:800; color:#f2f6fb; }
  .sec-title .desc { font-size:14.5px; color:#6b7891; font-weight:500; margin-left:6px; }
  .pick-card { border-radius:22px; padding:26px 28px; margin-bottom:16px;
    background: linear-gradient(135deg, #101c33 0%, #0e1728 100%); border:1px solid rgba(255,255,255,0.07); position:relative; }
  .pick-card.rank1 { border:1px solid rgba(46,230,166,0.45); }
  .pick-top { display:flex; align-items:center; justify-content:space-between; }
  .pick-left { display:flex; align-items:center; gap:18px; }
  .rank-badge { width:52px; height:52px; border-radius:16px; display:flex; align-items:center; justify-content:center;
    font-size:24px; font-weight:800; color:#0b1220; background:#2ee6a6; flex-shrink:0; }
  .rank-badge.gold { background:#f5b942; }
  .ticker-block .code { font-size:30px; font-weight:800; color:#fff; letter-spacing:0.5px;}
  .ticker-block .name { font-size:14.5px; color:#7c8aa1; font-weight:500; margin-top:2px;}
  .pick-right { text-align:right; }
  .price { font-size:30px; font-weight:800; color:#fff; }
  .chg { font-size:18px; font-weight:700; color:#2ee6a6; }
  .chg.neg { color:#ef4a5f; }
  .pick-bars { display:flex; gap:10px; margin-top:20px; flex-wrap:wrap; }
  .stat-chip { background:rgba(255,255,255,0.04); border:1px solid rgba(255,255,255,0.07); border-radius:13px; padding:11px 16px; flex:1; min-width:150px; }
  .stat-chip .lbl { font-size:12px; color:#6b7891; font-weight:600; letter-spacing:0.5px; text-transform:uppercase; margin-bottom:4px;}
  .stat-chip .val { font-size:18px; font-weight:700; color:#e8edf5; }
  .stat-chip .val.pos { color:#2ee6a6; } .stat-chip .val.neg { color:#ef4a5f; }
  .range-viz { margin-top:18px; }
  .range-viz .rlabel { display:flex; justify-content:space-between; font-size:12px; color:#5f6d84; margin-bottom:6px; font-weight:600;}
  .range-track { position:relative; height:10px; border-radius:6px; background:rgba(255,255,255,0.08); }
  .range-fill { position:absolute; top:0; bottom:0; border-radius:6px; background: linear-gradient(90deg,#f5b942,#2ee6a6); }
  .range-marker { position:absolute; top:-5px; width:3px; height:20px; background:#fff; border-radius:2px; }
  .pick-note { margin-top:16px; font-size:14.5px; color:#aab6c9; line-height:1.55; border-top:1px solid rgba(255,255,255,0.06); padding-top:14px;}
  .pick-note b { color:#dbe4f0; }
  .warn-tag { display:inline-block; font-size:12.5px; font-weight:700; color:#f5b942; background:rgba(245,185,66,0.12); border:1px solid rgba(245,185,66,0.35); padding:3px 11px; border-radius:12px; margin-top:10px; }
  .chart-block { border-radius:22px; margin-bottom:8px; background:#0b1220; border:1px solid rgba(255,255,255,0.07); padding:12px;}
  .chart-block img { width:100%; display:block; border-radius:14px; }
  table.wl { width:100%; border-collapse:collapse; border-radius:18px; overflow:hidden; }
  table.wl thead th { background:rgba(255,255,255,0.05); color:#8a97ad; font-size:12.5px; font-weight:700; text-transform:uppercase; letter-spacing:0.6px; padding:14px 14px; text-align:left; }
  table.wl tbody td { padding:15px 14px; font-size:16px; border-top:1px solid rgba(255,255,255,0.06); color:#dbe4f0;}
  table.wl tbody tr:nth-child(odd) { background:rgba(255,255,255,0.02); }
  table.wl td.code { font-weight:700; color:#fff; font-size:17px;}
  table.wl td.pos { color:#2ee6a6; font-weight:700; } table.wl td.neg { color:#ef4a5f; font-weight:700; }
  table.wl td.num { text-align:right; font-variant-numeric: tabular-nums;}
  .score-pill { display:inline-block; padding:4px 12px; border-radius:10px; font-weight:700; font-size:14.5px; }
  .risk-card { border-radius:22px; padding:26px 28px; margin-bottom:16px;
    background: linear-gradient(135deg, rgba(239,74,95,0.09), rgba(239,74,95,0.03)); border:1px solid rgba(239,74,95,0.30); }
  .risk-top { display:flex; align-items:center; gap:16px; margin-bottom:14px;}
  .risk-icon { width:46px; height:46px; border-radius:14px; background:rgba(239,74,95,0.18); display:flex; align-items:center; justify-content:center; font-size:22px;}
  .risk-code { font-size:24px; font-weight:800; color:#fff; }
  .risk-sub { font-size:13.5px; color:#f0a3ad; font-weight:600; letter-spacing:0.5px; text-transform:uppercase;}
  .risk-body { font-size:15px; color:#d8bfc4; line-height:1.65; }
  .risk-body b { color:#fff; }
  .method { border-radius:18px; padding:22px 26px; background:rgba(255,255,255,0.03); border:1px solid rgba(255,255,255,0.06); font-size:14px; color:#8a97ad; line-height:1.7; margin-top:8px;}
  .method b { color:#c3cee0; }
  .footer { margin-top:36px; padding-top:24px; border-top:1px solid rgba(255,255,255,0.08); display:flex; justify-content:space-between; align-items:center; }
  .footer .l { font-size:13.5px; color:#5f6d84; line-height:1.6; }
  .footer .brand { font-size:15px; font-weight:800; color:#8a97ad; letter-spacing:1px;}
"""


def _fmt_price(v: float) -> str:
    return f"{v:,.0f}".replace(",", ".")


def _fmt_rp_m(v_b: float) -> str:
    s = f"{v_b:,.1f}"
    s = s.replace(",", "#").replace(".", ",").replace("#", ".")
    return f"Rp {s} M"


def _close_pos_note(close_pos: float) -> str:
    if close_pos >= 0.9:
        return "closing nyaris di titik tertinggi hari ini — sinyal sign-of-strength yang kuat"
    if close_pos >= 0.7:
        return "closing solid di area atas range harian, buyer masih pegang kendali sampai akhir sesi"
    if close_pos >= 0.5:
        return "closing di area tengah range, tekanan beli-jual relatif seimbang"
    return "closing lemah, mendekati area bawah range hari ini"


def _vwap_note(vwap_pct: float) -> str:
    if vwap_pct >= 1.0:
        return f"ditutup <b>{vwap_pct:+.2f}%</b> di atas VWAP — ada late buying yang cukup meyakinkan"
    if vwap_pct >= 0:
        return f"ditutup tipis di atas VWAP ({vwap_pct:+.2f}%)"
    return f"ditutup <b>{vwap_pct:+.2f}%</b> di bawah VWAP — indikasi ada tekanan jual menjelang close"


FLAG_EXPLAIN = {
    "DISTRIBUSI KUAT": "harga ditutup jauh di bawah rata-rata transaksi harian (VWAP) — indikasi ada pihak yang melepas barang saat harga tinggi",
    "CLOSE LEMAH": "closing dekat titik terendah hari ini meski sempat naik lebih tinggi — waspada pola failed breakout",
    "VOLATILITAS EKSTREM": "range harian sangat lebar — pergerakan liar, risiko whipsaw tinggi kalau entry di harga yang salah",
}


def build_pick_note(r: pd.Series, note_extra: str = None) -> str:
    parts = [
        f"Close di <b>{r['close_pos']*100:.0f}% dari range</b> harian — {_close_pos_note(r['close_pos'])}. "
        f"Saham ini {_vwap_note(r['vwap_pct'])}. Kenaikan hari ini <b>{r['chg_pct']:+.2f}%</b> "
        f"dengan nilai transaksi <b>{_fmt_rp_m(r['value_b'])}</b>."
    ]
    if note_extra:
        parts.append(f"<b>Catatan:</b> {note_extra}")
    return " ".join(parts)


def build_risk_note(r: pd.Series, note_extra: str = None) -> str:
    flags = [f.strip() for f in r["flag"].split(",")] if r["flag"] != "-" else []
    explain = " ".join(f"{FLAG_EXPLAIN.get(f, f)}." for f in flags)
    base = (
        f"Naik <b>{r['chg_pct']:+.2f}%</b> di atas kertas, tapi close cuma di "
        f"<b>{r['close_pos']*100:.0f}% dari range</b> dan <b>{r['vwap_pct']:+.2f}%</b> vs VWAP. {explain}"
    )
    if note_extra:
        base += f" <b>Catatan:</b> {note_extra}"
    return base


def build_range_viz(r: pd.Series) -> str:
    close_pos_pct = r["close_pos"] * 100
    return f"""
    <div class="range-viz">
      <div class="rlabel"><span>Low {_fmt_price(r['low'])}</span><span>High {_fmt_price(r['high'])}</span></div>
      <div class="range-track">
        <div class="range-fill" style="left:0%; width:100%;"></div>
        <div class="range-marker" style="left:calc({close_pos_pct:.1f}% - 2px);"></div>
      </div>
    </div>"""


def build_pick_card(rank: int, r: pd.Series, ticker_db: dict,
                     ai_narratives: dict = None, ai_label: str = None) -> str:
    info = ticker_info(r["code"], ticker_db)
    rankcls = "rank1" if rank == 1 else ""
    badgecls = "" if rank == 1 else "gold"
    chg_cls = "pos" if r["chg_pct"] >= 0 else "neg"
    vwap_cls = "pos" if r["vwap_pct"] >= 0 else "neg"
    warn_html = ""
    if r["flag"] != "-":
        warn_html = f'<div class="warn-tag">⚠️ {r["flag"]} — cek size &amp; stop loss lebih ketat</div>'

    ai_text = (ai_narratives or {}).get(r["code"])
    if ai_text:
        note_label = f'Baca cepat <span style="color:#6b7891;font-weight:600;">(AI · {ai_label})</span>'
        note_body = ai_text
    else:
        note_label = "Baca cepat"
        note_body = build_pick_note(r, info.get("note"))

    return f"""
  <div class="pick-card {rankcls}">
    <div class="pick-top">
      <div class="pick-left">
        <div class="rank-badge {badgecls}">#{rank}</div>
        <div class="ticker-block">
          <div class="code">{r['code']}</div>
          <div class="name">{info['name']}</div>
        </div>
      </div>
      <div class="pick-right">
        <div class="price">{_fmt_price(r['close'])}</div>
        <div class="chg {chg_cls}">{r['chg_pct']:+.2f}%</div>
      </div>
    </div>
    <div class="pick-bars">
      <div class="stat-chip"><div class="lbl">Close Position</div><div class="val pos">{r['close_pos']*100:.0f}% dari range</div></div>
      <div class="stat-chip"><div class="lbl">vs VWAP</div><div class="val {vwap_cls}">{r['vwap_pct']:+.2f}%</div></div>
      <div class="stat-chip"><div class="lbl">Nilai Transaksi</div><div class="val">{_fmt_rp_m(r['value_b'])}</div></div>
      <div class="stat-chip"><div class="lbl">Skor Komposit</div><div class="val pos">{r['score']:.0f} / 100</div></div>
    </div>
    {build_range_viz(r)}
    <div class="pick-note"><b>{note_label}:</b> {note_body}</div>
    {warn_html}
  </div>"""


def build_table_row(r: pd.Series) -> str:
    chg_cls = "pos" if r["chg_pct"] >= 0 else "neg"
    vwap_cls = "pos" if r["vwap_pct"] >= 0 else "neg"
    score = r["score"]
    if score >= 55:
        pillbg, pillfg = "#1c3d33", "#2ee6a6"
    elif score >= 40:
        pillbg, pillfg = "#3d3419", "#f5b942"
    else:
        pillbg, pillfg = "#3d2419", "#f57c42"
    return f"""<tr>
      <td class="code">{r['code']}</td>
      <td class="num">{_fmt_price(r['close'])}</td>
      <td class="num {chg_cls}">{r['chg_pct']:+.2f}%</td>
      <td class="num">{r['close_pos']*100:.0f}%</td>
      <td class="num {vwap_cls}">{r['vwap_pct']:+.2f}%</td>
      <td class="num">{_fmt_rp_m(r['value_b'])}</td>
      <td class="num"><span class="score-pill" style="background:{pillbg};color:{pillfg};">{score:.0f}</span></td>
    </tr>"""


def build_risk_card(r: pd.Series, ticker_db: dict, ai_narratives: dict = None, ai_label: str = None) -> str:
    info = ticker_info(r["code"], ticker_db)
    ai_text = (ai_narratives or {}).get(r["code"])
    if ai_text:
        body = ai_text
        badge = f'<span style="color:#6b7891;font-weight:600;font-size:12px;">AI · {ai_label}</span>'
    else:
        body = build_risk_note(r, info.get("note"))
        badge = ""
    return f"""
  <div class="risk-card">
    <div class="risk-top">
      <div class="risk-icon">⚠️</div>
      <div>
        <div class="risk-code">{r['code']} <span style="color:#f0a3ad;font-weight:700;font-size:16px;">{info['name']}</span></div>
        <div class="risk-sub">Skor {r['score']:.0f}/100 · {r['flag']} {badge}</div>
      </div>
    </div>
    <div class="risk-body">{body}</div>
  </div>"""


def build_market_ctx_html(mkt: MarketContext) -> str:
    if not mkt.available:
        return ""
    regime_bg = {"GREEN": GREEN, "RED": RED, "NEUTRAL": GOLD}[mkt.regime]
    up_cls = "up" if mkt.change_pct >= 0 else "down"
    sign = "+" if mkt.change_pct >= 0 else ""
    return f"""
  <div class="ctx">
    <div class="ctx-top">
      <div class="ctx-label">KONTEKS PASAR — IHSG</div>
      <div class="regime-pill" style="background:{regime_bg};">REGIME: {mkt.regime}</div>
    </div>
    <div class="ctx-headline">IHSG <span class="{up_cls}">{sign}{mkt.change_pct:.2f}% ke {mkt.close:,.0f}</span></div>
    <div class="ctx-detail">
      Aturan override: perubahan harian &ge; {IHSG_GREEN_THRESHOLD:.0f}% otomatis regime <b>GREEN</b>,
      &le; {IHSG_RED_THRESHOLD:.0f}% otomatis <b>RED</b>. Data: {mkt.source}. Selalu cek ulang sentimen
      global &amp; flow asing sebelum entry Senin.
    </div>
  </div>"""


def _font_b64(filename: str) -> str:
    fp = FONTS_DIR / filename
    if not fp.exists():
        return None
    return base64.b64encode(fp.read_bytes()).decode("ascii")


def build_font_face_css() -> str:
    """@font-face dengan font di-embed base64 langsung -- supaya tampilan
    identik di mesin manapun (nggak bergantung font ter-install di OS host)."""
    faces = []
    for fname, weight in [("Poppins-Regular.ttf", 400), ("Poppins-Medium.ttf", 500), ("Poppins-Bold.ttf", 700)]:
        b64 = _font_b64(fname)
        if b64:
            faces.append(
                f"@font-face{{font-family:'Poppins';font-weight:{weight};font-style:normal;"
                f"src:url(data:font/truetype;charset=utf-8;base64,{b64}) format('truetype');}}"
            )
    if not faces:
        # fonts/ nggak ke-ship -- fallback ke font sistem kalau ada, else sans-serif default
        return "@font-face{font-family:'Poppins';src:local('Poppins');}"
    return "\n".join(faces)


def build_page_css(width_px: int = BASE_WIDTH_PX) -> str:
    """@page terpisah dari CSS utama (yang isinya string biasa penuh kurung
    kurawal literal) supaya gampang di-generate dinamis tanpa perlu bikin
    seluruh CSS jadi f-string. Tinggi dibikin sangat generous (lihat
    PAGE_MAX_HEIGHT_PX) karena WeasyPrint itu paged-media renderer -- kalau
    kontennya lebih tinggi dari @page, dia bakal mecah jadi banyak halaman
    kayak dokumen cetak. Kelebihan tinggi di-crop otomatis pas render."""
    return f"@page {{ size: {width_px}px {PAGE_MAX_HEIGHT_PX}px; margin: 0; }}"


def build_html(df: pd.DataFrame, ticker_db: dict, mkt: MarketContext, meta: dict,
                top_n: int, chart_files: dict,
                ai_narratives: dict = None, ai_label: str = None) -> str:
    top_df = df.head(top_n)
    top_codes = set(top_df["code"])
    risk_df = df[(df["flag"] != "-") & (~df["code"].isin(top_codes))].sort_values("score")
    risk_codes = set(risk_df["code"])
    sec_df = df[(~df["code"].isin(top_codes)) & (~df["code"].isin(risk_codes))]

    top_cards = "\n".join(
        build_pick_card(i, r, ticker_db, ai_narratives, ai_label)
        for i, (_, r) in enumerate(top_df.iterrows(), start=1)
    )
    table_rows = "\n".join(build_table_row(r) for _, r in sec_df.iterrows())
    risk_cards = "\n".join(
        build_risk_card(r, ticker_db, ai_narratives, ai_label) for _, r in risk_df.iterrows()
    )
    ctx_html = build_market_ctx_html(mkt)
    ai_badge = f'<div class="eyebrow" style="margin-left:10px;">🤖 NARASI AI · {ai_label}</div>' if ai_label else ""

    risk_section = ""
    if len(risk_df) > 0:
        risk_section = f"""
  <div class="sec-title"><div class="num">5</div><h2>Waspada / Hindari</h2>
    <div class="desc">Sinyal distribusi di penutupan — risiko lanjutan turun</div></div>
  {risk_cards}"""

    sec_table = ""
    if len(sec_df) > 0:
        sec_table = f"""
  <div class="sec-title"><div class="num">3</div><h2>Watchlist Sekunder</h2>
    <div class="desc">Layak dipantau, entry lebih selektif / tunggu konfirmasi</div></div>
  <table class="wl">
    <thead><tr><th>Kode</th><th>Close</th><th>%Chg</th><th>Close Pos</th><th>vs VWAP</th><th>Value</th><th>Skor</th></tr></thead>
    <tbody>{table_rows}</tbody>
  </table>"""

    font_css = build_font_face_css()
    page_css = build_page_css(BASE_WIDTH_PX)
    return f"""<!DOCTYPE html>
<html lang="id"><head><meta charset="UTF-8"><style>{font_css}
{page_css}
{CSS}</style></head>
<body><div class="wrap">

  <div class="eyebrow">{meta['eyebrow']}</div>{ai_badge}
  <h1 class="title">{meta['title'].upper()}</h1>
  <div class="subtitle">Watchlist untuk <b>{meta['next_day_str']}</b> · Data screening {meta['data_date_str']}, {meta['data_time_str']}</div>
  <div class="sourceline">{len(df)} saham dianalisis — diurutkan berdasarkan skor kekuatan close, bukan sekadar % kenaikan</div>

  {ctx_html}

  <div class="sec-title"><div class="num">1</div><h2>Prioritas Watchlist</h2>
    <div class="desc">Skor tertinggi dari kombinasi kekuatan close, VWAP, momentum &amp; likuiditas</div></div>
  {top_cards}

  <div class="sec-title"><div class="num">2</div><h2>Peta Kekuatan vs Momentum</h2>
    <div class="desc">Kuadran kanan-atas = setup paling sehat untuk continuation</div></div>
  <div class="chart-block"><img src="{chart_files['quadrant']}"></div>

  {sec_table}

  <div class="sec-title"><div class="num">4</div><h2>Ranking Kekuatan Setup</h2>
    <div class="desc">Semua saham, satu skor komposit</div></div>
  <div class="chart-block"><img src="{chart_files['ranking']}"></div>

  {risk_section}

  <div class="method">
    <b>Metodologi skor (0–100):</b> Kekuatan Close dalam range harian {SCORE_WEIGHTS['close_pos']*100:.0f}% ·
    Posisi Close vs VWAP {SCORE_WEIGHTS['vwap']*100:.0f}% · Momentum kenaikan harian {SCORE_WEIGHTS['momentum']*100:.0f}% ·
    Likuiditas nilai transaksi {SCORE_WEIGHTS['liquidity']*100:.0f}%. Skor & flag risiko selalu dihitung rule-based
    (deterministik, sama tiap run). Narasi per saham {"digenerate AI (<b>" + ai_label + "</b>) dari angka di atas -- AI diinstruksikan untuk tidak mengarang fakta di luar data ini" if ai_label else "dibuat rule-based dari angka di atas, bukan riset berita/fundamental"}.
    Ini adalah pembacaan objektif atas data historis,
    <b>bukan sinyal beli/jual dan bukan jaminan pergerakan hari berikutnya</b>.
  </div>

  <div class="footer">
    <div class="l">Dibuat otomatis · {datetime.now().strftime('%d %b %Y, %H:%M')}<br>Bukan rekomendasi investasi — keputusan trading sepenuhnya tanggung jawab pengguna.</div>
    <div class="brand">MESIN PRESISI</div>
  </div>

</div></body></html>"""


# ══════════════════════════════════════════════════════════════════════════
# 8) RENDER: HTML -> PDF (WeasyPrint) -> crop ke tinggi konten -> PNG HD (PyMuPDF)
# ══════════════════════════════════════════════════════════════════════════
#
# Kenapa bukan wkhtmltopdf/wkhtmltoimage lagi: proyek itu MATI sejak 2020 dan
# beberapa distro (termasuk base image Streamlit Community Cloud per pertengahan
# 2026) sudah BUANG paketnya dari apt sama sekali -- bukan cuma butuh xvfb, tapi
# betul-betul "Package has no installation candidate". WeasyPrint aktif
# di-develop, dependency sistemnya jauh lebih ringan (cuma Pango), dan
# PyMuPDF adalah wheel pip murni (nggak butuh apt sama sekali) buat konversi
# PDF -> PNG. Kombinasi ini juga menghilangkan seluruh masalah "butuh X11
# display" yang selalu menghantui wkhtmltopdf di server headless.
#
# WeasyPrint itu paged-media renderer (didesain buat cetak/PDF berhalaman),
# bukan continuous-viewport renderer kayak browser. Makanya CSS @page dikasih
# tinggi yang sangat generous (PAGE_MAX_HEIGHT_PX) supaya seluruh konten laporan
# muat di SATU halaman (nggak kepotong-potong kayak dokumen bersambung), lalu
# kelebihan ruang kosong di bawah di-crop otomatis berdasarkan tinggi konten
# yang sebenarnya -- baik untuk PDF (mediabox-nya dipangkas) maupun PNG.

_BG_RGB = (0x07, 0x0c, 0x16)  # harus sama dengan warna body paling luar di CSS


def _detect_content_height_pt(page: "fitz.Page", sample_width_px: int = 200) -> float:
    """Render halaman (yang tingginya masih PAGE_MAX_HEIGHT_PX, penuh ruang
    kosong) di resolusi RENDAH buat deteksi baris terakhir yang bukan cuma
    warna background -- cepat & ringan memory karena cuma dipakai untuk
    menentukan titik potong, bukan output final. Return dalam POINTS (satuan
    native PDF/fitz) supaya langsung nyambung ke set_mediabox() tanpa
    konversi unit lagi -- zoom di sini dihitung relatif terhadap page.rect
    yang satuannya point, bukan pixel."""
    zoom = sample_width_px / page.rect.width
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat)
    im = Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB")
    w, h = im.size
    px = im.load()
    last_content_row = 0
    tol = 10
    for y in range(h - 1, -1, -1):
        row_is_bg = True
        for x in range(0, w, max(1, w // 20)):
            r, g, b = px[x, y]
            if abs(r - _BG_RGB[0]) > tol or abs(g - _BG_RGB[1]) > tol or abs(b - _BG_RGB[2]) > tol:
                row_is_bg = False
                break
        if not row_is_bg:
            last_content_row = y
            break
    content_height_pt = (last_content_row / zoom) + 45  # padding bawah ~60px (~45pt)
    return min(content_height_pt, page.rect.height)


def render_report(html_str: str, workdir: Path, pdf_path: Path, png_path: Path,
                   width_px: int = BASE_WIDTH_PX, scale: float = RENDER_SCALE_DEFAULT):
    """HTML -> pdf_path (PDF asli, mediabox dipangkas ke tinggi konten) dan
    png_path (raster HD dari PDF yang sama, jadi dua-duanya dijamin identik)."""
    font_config = FontConfiguration()
    weasy_doc = WeasyHTML(string=html_str, base_url=str(workdir)).render(font_config=font_config)
    page = weasy_doc.pages[0]

    pdf_bytes = weasy_doc.write_pdf(font_config=font_config)
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    fitz_page = doc[0]

    content_height_pt = _detect_content_height_pt(fitz_page)
    page_width_pt = fitz_page.rect.width
    full_height_pt = fitz_page.rect.height
    # Koordinat mediabox PDF itu bottom-up: konten yang tampak "di atas" secara
    # visual sebenarnya berada di y mendekati full_height_pt, bukan mendekati 0.
    # Makanya potongan yang dipertahankan adalah ujung ATAS (dekat full_height_pt),
    # bukan dari y=0.
    fitz_page.set_mediabox(fitz.Rect(0, full_height_pt - content_height_pt, page_width_pt, full_height_pt))

    doc.save(pdf_path, garbage=3, deflate=True)

    final_zoom = (96 / 72) * scale
    pix = fitz_page.get_pixmap(matrix=fitz.Matrix(final_zoom, final_zoom))
    pix.save(png_path)
    doc.close()


def optimize_png(png_path: Path):
    """Buang alpha channel & re-save supaya file size masuk akal buat dikirim."""
    im = Image.open(png_path).convert("RGB")
    im.save(png_path, optimize=True)


# ══════════════════════════════════════════════════════════════════════════
# 9) MAIN / CLI
# ══════════════════════════════════════════════════════════════════════════

def generate_report(input_path: str, outdir: str = "outputs", top_n: int = TOP_N_DEFAULT,
                     title: str = None, ihsg_close: float = None, ihsg_change: float = None,
                     scale: float = RENDER_SCALE_DEFAULT, make_pdf: bool = True,
                     ai_backend: str = "auto", gemini_key: str = None,
                     gemini_model: str = GEMINI_MODEL_DEFAULT, github_token: str = None,
                     github_model: str = GITHUB_MODEL_DEFAULT) -> dict:
    """Fungsi utama -- bisa dipanggil langsung dari script lain (bot, cron, dsb),
    tidak harus lewat CLI. Return dict path-path output.

    ai_backend: "auto" (pakai kalau gemini_key/github_token ke-isi, else rule-based),
    "gemini", "github", atau "none" (paksa rule-based)."""
    in_path = Path(input_path)
    out_dir = Path(outdir)
    out_dir.mkdir(parents=True, exist_ok=True)

    raw_df = load_data(in_path)
    df = compute_metrics(raw_df)
    df = compute_score(df)
    print(f"[2/7] Skor dihitung untuk {len(df)} saham (top score: {df.iloc[0]['code']} = {df.iloc[0]['score']:.0f})")

    ticker_db = load_ticker_db()
    unknown = [c for c in df["code"] if c.upper() not in ticker_db]
    if unknown:
        preview = ", ".join(unknown[:10]) + ("..." if len(unknown) > 10 else "")
        print(f"  [info] {len(unknown)} ticker belum ada di {TICKER_DB_PATH.name}: {preview}")

    mkt = get_market_context(ihsg_close, ihsg_change)
    print(f"[3/7] Konteks IHSG: {'OK (' + mkt.source + ', regime ' + mkt.regime + ')' if mkt.available else 'tidak tersedia (lewati section ini)'}")

    print(f"[4/7] Narasi ({ai_backend})...")
    ai_narratives, ai_used = get_ai_narratives(
        df, top_n, backend=ai_backend,
        gemini_key=gemini_key, gemini_model=gemini_model,
        github_token=github_token, github_model=github_model,
    )
    ai_label = None
    if ai_used == "gemini":
        ai_label = f"Gemini {gemini_model}"
    elif ai_used == "github":
        ai_label = f"GitHub Models {github_model}"

    meta = parse_filename_meta(in_path, title)

    print(f"[5/7] Membuat chart (quadrant + ranking) di skala {scale}x...")
    chart_files = make_charts(df, out_dir, scale=scale)

    print("[6/7] Merangkai HTML...")
    html = build_html(df, ticker_db, mkt, meta, top_n, chart_files,
                       ai_narratives=ai_narratives, ai_label=ai_label)
    html_path = out_dir / f"{in_path.stem}_report.html"
    html_path.write_text(html, encoding="utf-8")

    png_path = out_dir / f"{in_path.stem}_HD.png"
    pdf_path = out_dir / f"{in_path.stem}.pdf" if make_pdf else out_dir / f"_tmp_{in_path.stem}.pdf"
    print(f"[7/7] Render PDF (WeasyPrint) -> crop -> PNG HD skala {scale}x (PyMuPDF)...")
    render_report(html, out_dir, pdf_path, png_path, width_px=BASE_WIDTH_PX, scale=scale)
    optimize_png(png_path)

    result = {"html": html_path, "png": png_path, "pdf": pdf_path if make_pdf else None,
              "dataframe": df, "ai_backend_used": ai_used}
    if not make_pdf and pdf_path.exists():
        pdf_path.unlink()

    print("\nSelesai:")
    print(f"  PNG (HD) : {png_path}  ({png_path.stat().st_size/1024:.0f} KB)")
    if result["pdf"]:
        print(f"  PDF      : {result['pdf']}  ({result['pdf'].stat().st_size/1024:.0f} KB)")
    print(f"  HTML     : {html_path}")
    print(f"  Narasi   : {'AI (' + ai_label + ')' if ai_label else 'rule-based'}")
    return result


def main():
    ap = argparse.ArgumentParser(
        description="Generate infografis watchlist HD (PNG + PDF) dari screener .xlsx/.xls/.csv/.pdf")
    ap.add_argument("input", type=str, help="Path ke file screener (.xlsx/.xls/.csv/.pdf)")
    ap.add_argument("--top", type=int, default=TOP_N_DEFAULT, help=f"Jumlah saham di 'Prioritas Watchlist' (default {TOP_N_DEFAULT})")
    ap.add_argument("--title", type=str, default=None, help="Judul screener (override auto-detect dari nama file)")
    ap.add_argument("--outdir", type=str, default="outputs", help="Folder output (default 'outputs')")
    ap.add_argument("--ihsg-close", type=float, default=None, help="Override manual: harga close IHSG")
    ap.add_argument("--ihsg-change", type=float, default=None, help="Override manual: %%chg IHSG")
    ap.add_argument("--scale", type=float, default=RENDER_SCALE_DEFAULT,
                     help="Faktor resolusi HD asli, bukan upscale (2.0=tajam/rekomendasi, 1.0=standar 1080px, 3.0=ekstra tajam/file lebih besar & lebih lama render)")
    ap.add_argument("--no-pdf", action="store_true", help="Skip generate PDF, cuma PNG")

    ap.add_argument("--ai-backend", choices=["auto", "gemini", "github", "none"], default="auto",
                     help="Sumber narasi per saham: auto (pakai kalau ada API key/token di env, default), "
                          "gemini, github (GPT-4o via GitHub Models), none (paksa rule-based)")
    ap.add_argument("--gemini-key", type=str, default=os.environ.get("GEMINI_API_KEY"),
                     help="API key Gemini (default: env GEMINI_API_KEY)")
    ap.add_argument("--gemini-model", type=str, default=os.environ.get("GEMINI_MODEL", GEMINI_MODEL_DEFAULT),
                     help=f"Model Gemini (default {GEMINI_MODEL_DEFAULT}, atau env GEMINI_MODEL)")
    ap.add_argument("--github-token", type=str,
                     default=os.environ.get("GITHUB_MODELS_TOKEN") or os.environ.get("GITHUB_TOKEN"),
                     help="GitHub PAT dengan permission 'models: read' (default: env GITHUB_MODELS_TOKEN atau GITHUB_TOKEN)")
    ap.add_argument("--github-model", type=str, default=os.environ.get("GITHUB_MODEL", GITHUB_MODEL_DEFAULT),
                     help=f"Model ID di GitHub Models (default {GITHUB_MODEL_DEFAULT}, atau env GITHUB_MODEL)")
    args = ap.parse_args()

    try:
        generate_report(
            input_path=args.input, outdir=args.outdir, top_n=args.top, title=args.title,
            ihsg_close=args.ihsg_close, ihsg_change=args.ihsg_change, scale=args.scale,
            make_pdf=not args.no_pdf, ai_backend=args.ai_backend,
            gemini_key=args.gemini_key, gemini_model=args.gemini_model,
            github_token=args.github_token, github_model=args.github_model,
        )
    except (ValueError, RuntimeError) as e:
        print(f"\n[GAGAL] {e}")
        raise SystemExit(1)
    except FileNotFoundError:
        print(f"\n[GAGAL] File tidak ditemukan: {args.input}")
        raise SystemExit(1)


if __name__ == "__main__":
    main()
