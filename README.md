# Geo PDF Pipeline 📄

Extract all text from PDFs — handwritten and typed — using Claude Vision. Upload a PDF, see structured results instantly, export to Excel.

![Python](https://img.shields.io/badge/Python-3.10+-blue) ![Flask](https://img.shields.io/badge/Flask-3.1-green) ![Claude](https://img.shields.io/badge/Claude-Vision-orange)

## Quick Start

```bash
git clone https://github.com/qtyree1328/geo-pdf-pipeline.git
cd geo-pdf-pipeline
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

# Add your Anthropic API key
cp .env.example .env
# Edit .env → set ANTHROPIC_API_KEY

python app.py
# Open http://localhost:3014
```

### Prerequisites
- Python 3.10+
- [Poppler](https://poppler.freedesktop.org/) for PDF rendering (`brew install poppler` on macOS)
- Anthropic API key

## How It Works

1. **Upload** a PDF through the web interface
2. **Extract** — each page is rendered at 300 DPI and sent to Claude Vision
3. **Review** — see structured fields, tables, and raw text with handwriting highlighted
4. **Export** — download everything as a formatted Excel spreadsheet

### What Gets Extracted

- **Typed text** — headers, paragraphs, form labels, field values
- **Handwritten text** — notes, signatures, filled-in forms (highlighted in the UI)
- **Tables** — detected and parsed into rows/columns
- **Confidence levels** — high/medium/low for each extracted section

### Excel Output

Three sheets:
- **Extracted Fields** — all sections with type, label, content, handwritten flag, confidence
- **Tables** — all detected tables with headers and rows
- **Raw Text** — full page text in reading order

## Tech Stack

- **Backend:** Flask + Anthropic SDK
- **OCR:** Claude Vision (claude-sonnet-4-20250514)
- **PDF Processing:** pdf2image + Poppler
- **Export:** openpyxl (formatted Excel)
- **Frontend:** Vanilla HTML/CSS/JS

## Configuration

Edit `.env`:
```
ANTHROPIC_API_KEY=sk-ant-xxxxx
PORT=3014
```

## License

MIT
