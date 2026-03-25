# 📁 File Generator

File Generator is a Python app that helps you **create and manage different types of files easily**.
It works through both CLI and a Streamlit UI.

---

## 🧠 What It Does

- Create files (text, docs, spreadsheets, etc.)
- Read and edit existing files
- Delete files when not needed
- Use AI to generate or summarize content (optional)

---

## ⚙️ Features

- 🖥️ CLI workflow (`main.py`)
- 🌐 Streamlit UI (`app_ui.py`)
- 🤖 AI support (OpenAI default, HackClub optional)
- 📂 Multi-file support (docs, images, charts, etc.)

---

## 📦 Supported File Types

- Text & docs → `txt`, `docx`, `pdf`, `markdown`
- Data → `xlsx`, `csv`
- Code → `html`, `code`
- Media → `image`, `chart`, `audio`, `video`

---

## 🧪 Requirements

- Python 3.10+
- Optional: `tesseract` (for OCR features)

---

## ⚡ Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
```

---

## 🔐 Configuration

Default (OpenAI):

```env
FILEGEN_AI_PROVIDER=openai
OPENAI_API_KEY=your-api-key
```

Optional (HackClub):

```env
FILEGEN_AI_PROVIDER=hackclub
HACKCLUB_API_KEY=your-key
HACKCLUB_BASE_URL=https://ai.hackclub.com/proxy/v1
```

---

## ▶️ Run

CLI:

```bash
python main.py
```

UI:

```bash
streamlit run app_ui.py
```

---

## 👤 User System (Streamlit)

- Requires login before using features
- Data stored locally using SQLite
- Passwords are securely hashed
- Each user has their own file storage
- You can only access your own files

---

## 🔒 Security

- Do NOT upload `.env` or API keys
- Generated files are ignored by Git
- Local data is stored in `.filegen_data/`

---

## 🏁 Final

Simple tool.

Does the job.

More features coming.

