# 🎬 YouTube Transcript Extractor

A Streamlit app that takes a list of YouTube video URLs, fetches their full transcripts, and downloads them as a formatted Excel file — ready for narrative analysis, sentiment analysis, or entity extraction in downstream tools.

---

## What It's For

This tool is designed to work as the **second step** in a YouTube intelligence pipeline that begins on **[Filmot.com](https://filmot.com)**.

### The Full Workflow

```
Filmot.com  →  CSV of YouTube URLs  →  This app  →  Excel with full transcripts  →  LLM / NLP analysis
```

**Step 1 — Search on Filmot**

[Filmot](https://filmot.com) is a search engine that indexes YouTube transcripts. You can search for a keyword or phrase and retrieve a list of videos whose transcripts contain that term. Before exporting:
- Filter by **language** so all your results are in the same language (e.g., all Indonesian, all Thai)
- Filter by date range, channel, or other criteria as needed
- Export results as a **CSV file**

Filmot's CSV includes video URLs, titles, channel names, and other metadata.

**Step 2 — Extract full transcripts (this app)**

Filmot only gives you a snippet around your search term. This app fetches the **full transcript** for each video. Upload your Filmot CSV, optionally select a translation language, and download an Excel file with complete transcripts.

**Step 3 — Downstream analysis**

Feed the Excel file into your analysis tool of choice:
- **LLMs** (ChatGPT, Claude, Gemini) for narrative summarization or thematic analysis
- **Sentiment analysis** tools or Python libraries (VADER, TextBlob, Transformers)
- **Entity extraction** (spaCy, AWS Comprehend, Azure Text Analytics) for people, organizations, and locations
- **Network analysis** tools for relationship mapping between entities

---

## Features

- **Paste URLs or upload a CSV** — one URL per line, or upload a Filmot-exported CSV directly
- **Full transcript extraction** — retrieves complete caption data, not just snippets
- **Translation to English** — optionally translate transcripts from Southeast Asian and South Asian languages (see supported languages below)
- **Formatted Excel output** — one row per video, with URL, video ID, detected language, translation status, and full transcript text
- **Error handling** — failed videos (disabled captions, private videos, bad URLs) are flagged clearly rather than silently dropped

---

## Supported Translation Languages

The app can translate transcripts from any of the following languages into English using Google Translate:

| Language | Region |
|---|---|
| Indonesian | Southeast Asia |
| Malay | Southeast Asia |
| Filipino / Tagalog | Southeast Asia |
| Vietnamese | Southeast Asia |
| Thai | Southeast Asia |
| Burmese (Myanmar) | Southeast Asia |
| Khmer (Cambodian) | Southeast Asia |
| Lao | Southeast Asia |
| Cebuano | Southeast Asia |
| Javanese | Southeast Asia |
| Sundanese | Southeast Asia |
| Tetum (Timor-Leste) | Southeast Asia |
| Bengali | South Asia |

> **Note:** Translation is only triggered when the detected transcript language matches your selected source language. Translation uses Google Translate's free tier via `deep-translator` — no API key required.

---

## Excel Output Format

The downloaded file contains one row per video with the following columns:

| Column | Description |
|---|---|
| # | Row number |
| YouTube URL | Full video URL |
| Video ID | The `v=` ID extracted from the URL |
| Language Detected | Language code of the transcript as found on YouTube |
| Translated to EN? | Yes/No — whether Google Translate was applied |
| Status | Success or Error |
| Transcript | Full transcript text (or error message if failed) |

---

## Installation & Local Setup

**Requirements:** Python 3.10+

```bash
git clone https://github.com/YOUR_USERNAME/youtube-transcript-extractor.git
cd youtube-transcript-extractor

pip install -r requirements.txt

streamlit run app.py
```

---

## Deploying on Streamlit Cloud

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in
3. Click **New app** and select your repo
4. Set the main file path to `app.py`
5. Click **Deploy**

Streamlit Cloud will install dependencies from `requirements.txt` automatically. No environment variables or API keys are needed.

---

## Dependencies

```
streamlit
youtube-transcript-api
deep-translator
openpyxl
pandas
```

---

## Known Limitations

- **Transcripts must exist on YouTube.** If a video's uploader has disabled captions, or if YouTube hasn't auto-generated them yet, the video will return an error.
- **Private and age-restricted videos** cannot be accessed.
- **Translation quality** depends on Google Translate. For low-resource languages (Lao, Tetum, Javanese), quality may be inconsistent.
- **Very long videos** with large transcripts may take a few seconds longer to process due to chunking for translation.

---

## Use Case Context

This tool was built for **OSINT and intelligence analysis workflows** where YouTube is a primary source — particularly for monitoring information environments in Southeast Asia and the broader Indo-Pacific. Common applications include:

- Tracking narratives in foreign-language media
- Monitoring influence operations across YouTube channels
- Building corpora for sentiment or framing analysis around specific keywords or events
- Extracting named entities (people, organizations, locations) from video content at scale
