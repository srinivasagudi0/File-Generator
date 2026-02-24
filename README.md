# File Generator

File Generator is a Python app that can create, append, read, and delete content across multiple file types (text, docs, spreadsheets, images, charts, PDF, PPTX, audio/video file stubs).  
It includes:

- CLI workflow (`main.py`)
- Streamlit UI workflow (`app_ui.py`)
- AI-assisted content generation and summarization (HackClub API only, configured via `.env`)

## Supported File Types

- `txt`, `docx`, `xlsx`, `csv`, `pdf`, `pptx`
- `markdown`, `html`, `code`
- `image`, `chart`, `audio`, `video`

## Requirements

- Python 3.10+
- Optional system binary for OCR: `tesseract` (used with `pytesseract`)

## Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
```

Set your HackClub API config in `.env` before running:

```env
FILEGEN_AI_PROVIDER=hackclub  # optional; defaults to hackclub
HACKCLUB_API_KEY=your-hackclub-api-key
HACKCLUB_BASE_URL=https://your-hackclub-base-url/v1
```

## Run

CLI:

```bash
python main.py
```

Streamlit UI:

```bash
streamlit run app_ui.py
```

## Security Notes

- Do not commit `.env` or real API keys.
- This repository is configured to ignore generated files and local runtime artifacts.

## Publish to GitHub

Run these commands from this project folder:

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin <your-github-repo-url>
git push -u origin main
```
