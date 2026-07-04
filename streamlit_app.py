"""
IDX Watchlist Report Generator -- Streamlit App
=================================================
UI web buat idx_watchlist_report.py: upload file screener dari browser (HP/PC
mana aja, PC lu sendiri nggak perlu nyala), generate infografis HD + PDF,
lalu kirim langsung ke grup Telegram.

DEPLOY KE STREAMLIT COMMUNITY CLOUD (gratis)
1. Push folder ini (streamlit_app.py, idx_watchlist_report.py, ticker_names.json,
   fonts/, requirements.txt, packages.txt) ke sebuah repo GitHub.
2. Buka share.streamlit.io -> New app -> pilih repo ini -> main file: streamlit_app.py
3. Di menu app -> Settings -> Secrets, isi (opsional tapi direkomendasikan):

    GEMINI_API_KEY = "xxxx"
    GITHUB_MODELS_TOKEN = "ghp_xxxx"
    TELEGRAM_BOT_TOKEN = "123456:xxxx"
    TELEGRAM_CHAT_ID = "-100xxxxxxxxxx"

   Kalau diisi di sini, field-field itu otomatis ke-prefill di sidebar tiap
   buka app -- nggak perlu ketik ulang tiap kali.
4. Deploy. Selesai -- akses dari HP kapan aja, PC nggak perlu nyala.

PENTING: packages.txt WAJIB ada isinya "wkhtmltopdf" dan "xvfb" (lihat file
packages.txt di folder ini) -- tanpa itu, generate PNG/PDF bakal gagal karena
wkhtmltoimage butuh binary itu ter-install di server.
"""

import os
import tempfile
from pathlib import Path

import requests
import streamlit as st

from idx_watchlist_report import (
    GEMINI_MODEL_DEFAULT,
    GITHUB_MODEL_DEFAULT,
    RENDER_SCALE_DEFAULT,
    TOP_N_DEFAULT,
    generate_report,
)

st.set_page_config(page_title="IDX Watchlist Generator", page_icon="📈", layout="centered")


# ──────────────────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────────────────

def secret_or_env(key: str, default: str = "") -> str:
    try:
        if key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return os.environ.get(key, default)


def send_telegram_document(token: str, chat_id: str, file_bytes: bytes, filename: str,
                            caption: str = "", mime: str = "application/octet-stream"):
    url = f"https://api.telegram.org/bot{token}/sendDocument"
    files = {"document": (filename, file_bytes, mime)}
    data = {"chat_id": chat_id, "caption": caption[:1024]}
    r = requests.post(url, data=data, files=files, timeout=90)
    try:
        payload = r.json()
    except Exception:
        payload = {"ok": False, "description": r.text[:300]}
    return payload.get("ok", False), payload.get("description", "OK")


def send_telegram_photo(token: str, chat_id: str, image_bytes: bytes, filename: str, caption: str = ""):
    url = f"https://api.telegram.org/bot{token}/sendPhoto"
    files = {"photo": (filename, image_bytes, "image/png")}
    data = {"chat_id": chat_id, "caption": caption[:1024]}
    r = requests.post(url, data=data, files=files, timeout=90)
    try:
        payload = r.json()
    except Exception:
        payload = {"ok": False, "description": r.text[:300]}
    return payload.get("ok", False), payload.get("description", "OK")


# ──────────────────────────────────────────────────────────────────────────
# Sidebar -- pengaturan
# ──────────────────────────────────────────────────────────────────────────

st.title("📈 IDX Watchlist Generator")
st.caption("Upload hasil screener DSI/TQL → infografis HD + PDF → kirim ke Telegram. Semua jalan di cloud, PC nggak perlu nyala.")

with st.sidebar:
    st.header("⚙️ Pengaturan")
    top_n = st.slider("Jumlah saham di 'Prioritas Watchlist'", 3, 10, TOP_N_DEFAULT)
    scale = st.select_slider("Kualitas render (HD)", options=[1.0, 1.5, 2.0, 2.5, 3.0], value=RENDER_SCALE_DEFAULT,
                              help="2.0 direkomendasikan. Naikin ke 2.5-3.0 kalau mau ekstra tajam (render lebih lama & file lebih besar).")

    st.subheader("IHSG (opsional)")
    st.caption("Kosongin kalau mau auto-fetch (butuh yfinance+internet), atau isi manual.")
    ihsg_close_in = st.text_input("Close IHSG", value="")
    ihsg_change_in = st.text_input("%Chg IHSG", value="")
    ihsg_close = float(ihsg_close_in) if ihsg_close_in.strip() else None
    ihsg_change = float(ihsg_change_in) if ihsg_change_in.strip() else None

    st.subheader("🤖 Narasi AI (opsional)")
    ai_backend = st.selectbox("Sumber narasi", ["auto", "gemini", "github", "none"], index=0,
                               help="auto: pakai Gemini kalau key ada, fallback GitHub Models, fallback rule-based.")
    gemini_key = st.text_input("Gemini API Key", value=secret_or_env("GEMINI_API_KEY"), type="password")
    gemini_model = st.text_input("Model Gemini", value=secret_or_env("GEMINI_MODEL", GEMINI_MODEL_DEFAULT))
    github_token = st.text_input("GitHub Models Token", value=secret_or_env("GITHUB_MODELS_TOKEN"), type="password")
    github_model = st.text_input("Model GitHub Models", value=secret_or_env("GITHUB_MODEL", GITHUB_MODEL_DEFAULT))

    st.subheader("📤 Telegram")
    tg_token = st.text_input("Bot Token", value=secret_or_env("TELEGRAM_BOT_TOKEN"), type="password")
    tg_chat_id = st.text_input("Chat ID / Group ID", value=secret_or_env("TELEGRAM_CHAT_ID"))
    with st.expander("Cara dapetin Bot Token & Chat ID"):
        st.markdown(
            "1. Chat **@BotFather** di Telegram → `/newbot` → ikuti langkahnya → salin token.\n"
            "2. Add bot itu ke grup diskusi lu.\n"
            "3. Kirim 1 pesan apa aja di grup itu.\n"
            "4. Buka `https://api.telegram.org/bot<TOKEN>/getUpdates` di browser "
            "(ganti `<TOKEN>` dengan token bot lu).\n"
            "5. Cari `\"chat\":{\"id\": -100xxxxxxxxxx` di hasilnya -- itu Chat ID grupnya "
            "(biasanya dimulai dengan minus)."
        )

# ──────────────────────────────────────────────────────────────────────────
# Main -- upload & generate
# ──────────────────────────────────────────────────────────────────────────

uploaded = st.file_uploader("Upload file screener (.xlsx / .xls / .csv / .pdf)",
                             type=["xlsx", "xls", "csv", "pdf"])

if uploaded is not None:
    st.caption(f"File: **{uploaded.name}** ({uploaded.size/1024:.0f} KB)")

generate_clicked = st.button("🚀 Generate Infografis", type="primary", disabled=uploaded is None)

if generate_clicked and uploaded is not None:
    with st.spinner("Memproses screener... (parsing → skor → chart → render HD)"):
        with tempfile.TemporaryDirectory() as tmpdir:
            in_path = Path(tmpdir) / uploaded.name
            in_path.write_bytes(uploaded.getvalue())
            try:
                result = generate_report(
                    input_path=str(in_path), outdir=tmpdir, top_n=top_n,
                    ihsg_close=ihsg_close, ihsg_change=ihsg_change, scale=scale,
                    ai_backend=ai_backend, gemini_key=gemini_key or None, gemini_model=gemini_model,
                    github_token=github_token or None, github_model=github_model,
                )
            except Exception as e:
                st.error(f"Gagal generate: {e}")
                st.stop()

            png_bytes = result["png"].read_bytes()
            pdf_bytes = result["pdf"].read_bytes() if result["pdf"] else None
            st.session_state["png_bytes"] = png_bytes
            st.session_state["pdf_bytes"] = pdf_bytes
            st.session_state["png_name"] = result["png"].name
            st.session_state["pdf_name"] = result["pdf"].name if result["pdf"] else None
            st.session_state["ai_used"] = result["ai_backend_used"]

    narasi_label = st.session_state["ai_used"] or "rule-based"
    st.success(f"Selesai! Narasi: **{narasi_label}** · Ukuran PNG: {len(png_bytes)/1024:.0f} KB")

# ──────────────────────────────────────────────────────────────────────────
# Hasil -- preview, download, kirim Telegram
# ──────────────────────────────────────────────────────────────────────────

if st.session_state.get("png_bytes"):
    st.divider()
    st.subheader("Hasil")
    st.image(st.session_state["png_bytes"], use_container_width=True)

    col1, col2 = st.columns(2)
    with col1:
        st.download_button("⬇️ Download PNG (HD)", st.session_state["png_bytes"],
                            file_name=st.session_state["png_name"], mime="image/png", use_container_width=True)
    with col2:
        if st.session_state.get("pdf_bytes"):
            st.download_button("⬇️ Download PDF", st.session_state["pdf_bytes"],
                                file_name=st.session_state["pdf_name"], mime="application/pdf",
                                use_container_width=True)

    st.divider()
    st.subheader("📤 Kirim ke Telegram")
    caption = st.text_input("Caption", value="Watchlist Senin 🚀")
    also_photo = st.checkbox(
        "Kirim juga sebagai Photo (preview inline di chat, tapi dikompres Telegram)",
        value=False,
        help="Defaultnya PNG dikirim sebagai Document supaya kualitas HD-nya tetap utuh -- "
             "Telegram mengompres foto yang dikirim lewat mode Photo.",
    )

    if st.button("📤 Kirim Sekarang", type="primary"):
        if not tg_token or not tg_chat_id:
            st.error("Bot Token / Chat ID belum diisi di sidebar.")
        else:
            with st.spinner("Mengirim ke Telegram..."):
                ok1, msg1 = send_telegram_document(
                    tg_token, tg_chat_id, st.session_state["png_bytes"],
                    st.session_state["png_name"], caption=caption, mime="image/png",
                )
                results = [("PNG (Document/HD)", ok1, msg1)]

                if also_photo:
                    ok2, msg2 = send_telegram_photo(
                        tg_token, tg_chat_id, st.session_state["png_bytes"],
                        st.session_state["png_name"], caption=caption,
                    )
                    results.append(("PNG (Photo/preview)", ok2, msg2))

                if st.session_state.get("pdf_bytes"):
                    ok3, msg3 = send_telegram_document(
                        tg_token, tg_chat_id, st.session_state["pdf_bytes"],
                        st.session_state["pdf_name"], caption=caption, mime="application/pdf",
                    )
                    results.append(("PDF", ok3, msg3))

            for label, ok, msg in results:
                if ok:
                    st.success(f"{label}: terkirim ✅")
                else:
                    st.error(f"{label}: gagal -- {msg}")
