import streamlit as st
import pandas as pd
import io
import re
from youtube_transcript_api import YouTubeTranscriptApi, TranscriptsDisabled, NoTranscriptFound
from deep_translator import GoogleTranslator
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="YouTube Transcript Extractor",
    page_icon="🎬",
    layout="wide",
)

# ── Helpers ─────────────────────────────────────────────────────────────────────

def extract_video_id(url: str) -> str | None:
    """Extract YouTube video ID from various URL formats."""
    patterns = [
        r"(?:v=|\/)([0-9A-Za-z_-]{11})(?:[&?\/]|$)",
        r"youtu\.be\/([0-9A-Za-z_-]{11})",
        r"embed\/([0-9A-Za-z_-]{11})",
        r"shorts\/([0-9A-Za-z_-]{11})",
    ]
    for pattern in patterns:
        match = re.search(pattern, url)
        if match:
            return match.group(1)
    return None


LANG_CODE_MAP = {
    "None (keep original)": None,
    # Major national languages
    "Burmese (Myanmar) → English": "my",
    "Filipino / Tagalog → English": "tl",
    "Indonesian → English": "id",
    "Khmer (Cambodian) → English": "km",
    "Lao → English": "lo",
    "Malay → English": "ms",
    "Thai → English": "th",
    "Vietnamese → English": "vi",
    # South Asian languages with SEA presence
    "Bengali → English": "bn",
    # Regional / additional languages with YouTube presence
    "Cebuano → English": "ceb",
    "Javanese → English": "jw",
    "Sundanese → English": "su",
    "Tetum (Timor-Leste) → English": "tet",
}


def fetch_transcript(video_id: str, source_lang_code: str | None) -> dict:
    """
    Fetch transcript for a video.
    Returns a dict with keys: video_id, title_url, transcript, language, translated, error
    """
    result = {
        "video_id": video_id,
        "url": f"https://www.youtube.com/watch?v={video_id}",
        "transcript": "",
        "language_detected": "",
        "translated": False,
        "error": "",
    }

    try:
        transcript_list = YouTubeTranscriptApi.list_transcripts(video_id)

        # Prefer manually created over auto-generated; prefer source lang if specified
        transcript_obj = None
        if source_lang_code:
            try:
                transcript_obj = transcript_list.find_transcript([source_lang_code])
            except Exception:
                pass
        if transcript_obj is None:
            try:
                transcript_obj = transcript_list.find_manually_created_transcript(
                    [source_lang_code] if source_lang_code else []
                )
            except Exception:
                pass
        if transcript_obj is None:
            transcript_obj = transcript_list.find_generated_transcript(
                [source_lang_code] if source_lang_code else ["en", "id", "ms"]
            )

        result["language_detected"] = transcript_obj.language_code
        raw = transcript_obj.fetch()
        full_text = " ".join(segment["text"] for segment in raw)

        # Translate if requested
        if source_lang_code and transcript_obj.language_code == source_lang_code:
            translator = GoogleTranslator(source=source_lang_code, target="en")
            # GoogleTranslator has a char limit; chunk if necessary
            chunks = _chunk_text(full_text, 4500)
            translated_chunks = [translator.translate(c) for c in chunks]
            result["transcript"] = " ".join(translated_chunks)
            result["translated"] = True
        else:
            result["transcript"] = full_text

    except TranscriptsDisabled:
        result["error"] = "Transcripts are disabled for this video."
    except NoTranscriptFound:
        result["error"] = "No transcript found for the selected language."
    except Exception as e:
        result["error"] = str(e)

    return result


def _chunk_text(text: str, max_chars: int) -> list[str]:
    """Split text into chunks of max_chars without breaking words."""
    words = text.split()
    chunks, current = [], []
    length = 0
    for word in words:
        if length + len(word) + 1 > max_chars:
            chunks.append(" ".join(current))
            current, length = [word], len(word)
        else:
            current.append(word)
            length += len(word) + 1
    if current:
        chunks.append(" ".join(current))
    return chunks


def build_excel(rows: list[dict]) -> bytes:
    """Build a formatted Excel workbook and return as bytes."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Transcripts"

    # ── Styles ──
    header_fill = PatternFill("solid", start_color="1F4E79", end_color="1F4E79")
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    alt_fill = PatternFill("solid", start_color="D6E4F0", end_color="D6E4F0")
    white_fill = PatternFill("solid", start_color="FFFFFF", end_color="FFFFFF")

    cell_font = Font(name="Arial", size=10)
    cell_align = Alignment(vertical="top", wrap_text=True)

    thin_border = Border(
        left=Side(style="thin", color="BFBFBF"),
        right=Side(style="thin", color="BFBFBF"),
        top=Side(style="thin", color="BFBFBF"),
        bottom=Side(style="thin", color="BFBFBF"),
    )

    # ── Headers ──
    headers = ["#", "YouTube URL", "Video ID", "Language Detected", "Translated to EN?", "Status", "Transcript"]
    col_widths = [5, 45, 15, 18, 16, 22, 100]

    for col_idx, (header, width) in enumerate(zip(headers, col_widths), start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A2"

    # ── Data rows ──
    for row_num, row in enumerate(rows, start=2):
        fill = alt_fill if row_num % 2 == 0 else white_fill
        status = "Error" if row["error"] else "Success"
        transcript_or_error = row["error"] if row["error"] else row["transcript"]

        values = [
            row_num - 1,
            row["url"],
            row["video_id"],
            row["language_detected"],
            "Yes" if row["translated"] else "No",
            status,
            transcript_or_error,
        ]

        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row=row_num, column=col_idx, value=value)
            cell.font = cell_font
            cell.fill = fill
            cell.alignment = cell_align
            cell.border = thin_border

        # Row height — transcript cells can be tall; cap at 400pts
        ws.row_dimensions[row_num].height = min(400, max(20, len(str(transcript_or_error)) // 15))

    # ── Auto-filter ──
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ── UI ───────────────────────────────────────────────────────────────────────────

st.title("🎬 YouTube Transcript Extractor")
st.caption("Paste YouTube URLs → download a formatted Excel file with full transcripts.")

with st.sidebar:
    st.header("⚙️ Options")
    translation_option = st.selectbox(
        "Translate to English from…",
        options=list(LANG_CODE_MAP.keys()),
        help="Select a Southeast Asian source language to translate transcripts into English.",
    )
    st.markdown("---")
    st.markdown(
        "**Supported URL formats**\n"
        "- `https://www.youtube.com/watch?v=...`\n"
        "- `https://youtu.be/...`\n"
        "- `https://www.youtube.com/shorts/...`\n"
        "- `https://www.youtube.com/embed/...`"
    )
    st.markdown("---")
    st.markdown(
        "**Notes**\n"
        "- Transcripts must be enabled on the video.\n"
        "- Translation covers all major SEA languages via Google Translate (free tier).\n"
        "- Very long videos may take a moment to process."
    )

st.subheader("📋 Enter YouTube URLs")

tab_paste, tab_csv = st.tabs(["✏️ Paste URLs", "📂 Upload CSV"])

raw_urls = []

with tab_paste:
    url_input = st.text_area(
        label="One URL per line",
        placeholder="https://www.youtube.com/watch?v=dQw4w9WgXcQ\nhttps://youtu.be/abc123xyz",
        height=200,
    )
    if url_input.strip():
        raw_urls = [u.strip() for u in url_input.strip().splitlines() if u.strip()]

with tab_csv:
    st.markdown(
        "Upload a **Filmot CSV or Excel export**, or any file with a `url` / `URL` column. "
        "Filmot exports are automatically detected and the YouTube URL is reconstructed from the `id` column."
    )
    uploaded_file = st.file_uploader("Choose a CSV or Excel file", type=["csv", "xlsx", "xls"])
    if uploaded_file:
        try:
            fname = uploaded_file.name.lower()
            if fname.endswith((".xlsx", ".xls")):
                df_upload = pd.read_excel(uploaded_file)
            else:
                df_upload = pd.read_csv(uploaded_file)

            cols_lower = [c.strip().lower() for c in df_upload.columns]

            # Filmot format: has an 'id' column containing bare video IDs
            if "id" in cols_lower and "url" not in cols_lower:
                id_col = df_upload.columns[cols_lower.index("id")]
                raw_urls = [
                    f"https://www.youtube.com/watch?v={vid}"
                    for vid in df_upload[id_col].dropna().astype(str).str.strip()
                    if vid
                ]
                st.success(
                    f"Detected **Filmot format** — reconstructed **{len(raw_urls)} YouTube URLs** "
                    f"from the `{id_col}` column."
                )
            # Generic format: look for a url column
            else:
                url_col_idx = next(
                    (i for i, c in enumerate(cols_lower) if c == "url"), 0
                )
                url_col = df_upload.columns[url_col_idx]
                raw_urls = df_upload[url_col].dropna().astype(str).str.strip().tolist()
                st.success(f"Loaded **{len(raw_urls)} URLs** from column `{url_col}`.")

            with st.expander("Preview reconstructed URLs"):
                st.dataframe(
                    pd.DataFrame({"YouTube URL": raw_urls[:20]}),
                    use_container_width=True,
                )
        except Exception as e:
            st.error(f"Could not read file: {e}")

run = st.button(
    f"🚀 Extract Transcripts{f' ({len(raw_urls)} URLs)' if raw_urls else ''}",
    type="primary",
    use_container_width=True,
    disabled=not raw_urls,
)

if run:
    if not raw_urls:
        st.warning("Please enter at least one YouTube URL.")
        st.stop()

    source_lang_code = LANG_CODE_MAP[translation_option]
    results = []

    progress = st.progress(0, text="Starting…")
    status_box = st.empty()

    for i, url in enumerate(raw_urls):
        status_box.info(f"Processing ({i+1}/{len(raw_urls)}): `{url}`")
        video_id = extract_video_id(url)
        if not video_id:
            results.append({
                "video_id": "INVALID",
                "url": url,
                "transcript": "",
                "language_detected": "",
                "translated": False,
                "error": "Could not parse a valid YouTube video ID from this URL.",
            })
        else:
            results.append(fetch_transcript(video_id, source_lang_code))
        progress.progress((i + 1) / len(raw_urls), text=f"{i+1}/{len(raw_urls)} processed")

    status_box.empty()
    progress.empty()

    # Persist to session state so the download button survives reruns
    st.session_state["results"] = results
    st.session_state["excel_bytes"] = build_excel(results)

# ── Results panel (shown whenever session_state has data) ──
if "results" in st.session_state:
    results = st.session_state["results"]

    # ── Summary ──
    successes = sum(1 for r in results if not r["error"])
    failures = len(results) - successes

    col1, col2, col3 = st.columns(3)
    col1.metric("Total URLs", len(results))
    col2.metric("✅ Succeeded", successes)
    col3.metric("❌ Failed", failures)

    if failures:
        with st.expander("⚠️ Failed URLs"):
            for r in results:
                if r["error"]:
                    st.error(f"**{r['url']}** — {r['error']}")

    # ── Preview ──
    with st.expander("👀 Preview transcripts"):
        for r in results:
            if not r["error"]:
                label = f"🎬 {r['url']} | lang: `{r['language_detected']}`" + (
                    " | translated ✓" if r["translated"] else ""
                )
                with st.expander(label):
                    st.write(r["transcript"][:3000] + ("…" if len(r["transcript"]) > 3000 else ""))

    # ── Download — persists across reruns ──
    st.download_button(
        label="📥 Download Excel",
        data=st.session_state["excel_bytes"],
        file_name="youtube_transcripts.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )
