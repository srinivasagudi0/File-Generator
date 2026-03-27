# File Generator

I built this because I got tired of making the same kinds of files over and over frequent either it is for school or clubs.

Sometimes I just want a quick `.txt` or `.md` file. Sometimes I need a `.docx`, `.xlsx`, `.pdf`, chart image, or a rough slide deck. This project is a small utility that handles that from either the terminal or a Streamlit UI. This is a small but very helpful project.

It is not meant to be a big platform. It is a file tool. You tell it what you want, pick the format, and it does the repetitive part.

## What it does

- Create files
- Read files
- Append to supported files
- Delete files
- Use AI for some text, image, and file-processing flows if you configure a provider such as OpenAI or HackCLUb

## Supported file types

The current code supports these file types:

- `txt`
- `docx`
- `xlsx`
- `csv`
- `pdf`
- `pptx`
- `markdown` / `.md`
- `html`
- `code`
- `image`
- `chart`
- `audio`
- `video`

Two important notes here:

1. `chart` output is generated as an image using `matplotlib`.
2. `audio` and `video` are not generated from scratch. Right now the tool copies them from a source path you provide.

## How to run it

### CLI

```bash
python main.py
```

The CLI walks through the action, file type, content, and output name.

### Streamlit UI

```bash
streamlit run app_ui.py
```

The UI does the same basic job, just with forms instead of terminal prompts.

## Setup

Use Python `3.10+`.

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## AI setup

AI is optional for basic local file operations. It matters when you want AI-assisted content generation, image work, OCR fallback, or file summarization flows. I would personally recommend using this with AI since the main idea is to ease the process.

There is an example config in `.env.example`.

### OpenAI

```env
FILEGEN_AI_PROVIDER=openai
OPENAI_API_KEY=your_openai_api_key_here
```

### Hack Club

```env
FILEGEN_AI_PROVIDER=hackclub
HACKCLUB_API_KEY=your_hackclub_api_key_here
HACKCLUB_BASE_URL=https://ai.hackclub.com/proxy/v1
```

If you want model overrides, check `.env.example` and uncomment what you need.

## A few honest notes

- OCR support depends on `pytesseract`, and Tesseract needs to be installed on the machine. It took about 3 minutes so not that time consuming and is one time.
- `.docx`, `.xlsx`, `.pdf`, `.pptx`, and charts depend on their respective Python packages from `requirements.txt`.
- Some read/process flows use AI only for specific file types, not everything.
- Cloud-style paths such as `dropbox:report.docx` or `s3:folder/data.csv` are supported in the codebase, with `rclone` used when available.
- Audio and video write mode expects a source path in the content, like `path:/full/path/to/file.mp3`.

## Main files

- `main.py`: CLI flow
- `app_ui.py`: Streamlit UI
- `file_generator.py`: file routing and generation logic
- `intel.py`: AI, OCR, and image-related helpers

## Why this exists

Because doing file strcuture or the file (sometimes) by hand is boring.

That is really it. I wanted something that could take a prompt, turn it into a usable file, and save a few minutes every time. Please feel free to test it out.
