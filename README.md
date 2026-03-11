# File Generator

File Generator is a Python app that can create, append, read, and delete content across multiple file types (text, docs, spreadsheets, images, charts, PDF, PPTX, audio/video file stubs).  
It includes:

- CLI workflow (`main.py`)
- Streamlit UI workflow (`app_ui.py`)
- AI-assisted content generation and summarization (OpenAI by default, HackClub optional, configured via `.env`)

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

Set your API config in `.env` before running.

OpenAI (default / easiest setup):

```env
FILEGEN_AI_PROVIDER=openai
OPENAI_API_KEY=your-openai-api-key
```

Optional HackClub configuration:

```env
FILEGEN_AI_PROVIDER=hackclub
HACKCLUB_API_KEY=your-hackclub-api-key
HACKCLUB_BASE_URL=https://ai.hackclub.com/proxy/v1
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

## Local Accounts in Streamlit

The Streamlit UI now requires a local account before you can generate, preview, read, or download files.

- Accounts are stored in a local SQLite database under `.filegen_data/`
- Passwords are stored as salted PBKDF2 hashes
- Generated files are written into per-user storage folders under `.filegen_data/users/<user_id>/files/`
- The UI only shows history entries that belong to the signed-in user
- Read, append, preview, download, and delete actions are limited to the current user's files

The CLI workflow in `main.py` remains a direct local workflow and does not enforce the Streamlit account system.

## Security Notes

- Do not commit `.env` or real API keys.
- This repository is configured to ignore generated files and local runtime artifacts, including `.filegen_data/`.

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
