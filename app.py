"""
Geo PDF Pipeline — Extract text from PDFs (handwritten + typed) via Claude Vision.
Upload a PDF → see extracted text → export to spreadsheet.
"""

import os
import json
import time
import base64
import uuid
from pathlib import Path
from datetime import datetime

from flask import Flask, request, jsonify, send_from_directory, send_file
from dotenv import load_dotenv
from pdf2image import convert_from_path
import anthropic
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

load_dotenv()

app = Flask(__name__, static_folder="static")

UPLOAD_DIR = Path("uploads")
RESULTS_DIR = Path("results")
UPLOAD_DIR.mkdir(exist_ok=True)
RESULTS_DIR.mkdir(exist_ok=True)

# In-memory store for processed documents
documents = {}


def pdf_to_images(pdf_path: str, dpi: int = 300) -> list[str]:
    """Convert PDF pages to base64-encoded PNG images."""
    images = convert_from_path(pdf_path, dpi=dpi)
    encoded = []
    for img in images:
        from io import BytesIO
        buf = BytesIO()
        img.save(buf, format="PNG")
        encoded.append(base64.standard_b64encode(buf.getvalue()).decode("utf-8"))
    return encoded


def extract_with_claude(image_b64: str, page_num: int, total_pages: int) -> dict:
    """Send a page image to Claude Vision and extract all text."""
    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))

    response = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=4096,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/png",
                            "data": image_b64,
                        },
                    },
                    {
                        "type": "text",
                        "text": """Extract ALL text from this document page. Include both typed/printed text AND handwritten text.

Return a JSON object with this structure:
{
  "page_number": <int>,
  "sections": [
    {
      "type": "header" | "field" | "table" | "paragraph" | "handwritten_note" | "label" | "other",
      "label": "<field label if applicable, e.g. 'Project Name', 'Date'>",
      "content": "<the extracted text>",
      "is_handwritten": <boolean>,
      "confidence": "high" | "medium" | "low"
    }
  ],
  "tables": [
    {
      "title": "<table title if visible>",
      "headers": ["col1", "col2", ...],
      "rows": [["val1", "val2", ...], ...]
    }
  ],
  "raw_text": "<all text on the page in reading order, preserving layout as much as possible>"
}

Be thorough. Capture every piece of text including stamps, signatures, marginal notes, form field labels AND their filled-in values. For handwritten text, do your best to interpret it and flag confidence level.""",
                    },
                ],
            }
        ],
    )

    text = response.content[0].text

    # Try to parse as JSON, fall back to wrapping raw text
    try:
        # Find JSON in the response
        start = text.find("{")
        end = text.rfind("}") + 1
        if start >= 0 and end > start:
            result = json.loads(text[start:end])
            result["page_number"] = page_num
            return result
    except json.JSONDecodeError:
        pass

    return {
        "page_number": page_num,
        "sections": [{"type": "paragraph", "content": text, "is_handwritten": False, "confidence": "medium"}],
        "tables": [],
        "raw_text": text,
    }


def results_to_excel(doc_data: dict, output_path: str):
    """Export extracted data to a formatted Excel spreadsheet."""
    wb = Workbook()

    # --- Sheet 1: Extracted Fields ---
    ws_fields = wb.active
    ws_fields.title = "Extracted Fields"

    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="2B5B84", end_color="2B5B84", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell_font = Font(name="Calibri", size=10)
    cell_align = Alignment(vertical="top", wrap_text=True)
    hw_fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    headers = ["Page", "Type", "Label", "Content", "Handwritten", "Confidence"]
    for col, h in enumerate(headers, 1):
        cell = ws_fields.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

    row = 2
    for page in doc_data.get("pages", []):
        page_num = page.get("page_number", "?")
        for section in page.get("sections", []):
            ws_fields.cell(row=row, column=1, value=page_num).border = thin_border
            ws_fields.cell(row=row, column=2, value=section.get("type", "")).border = thin_border
            ws_fields.cell(row=row, column=3, value=section.get("label", "")).border = thin_border
            content_cell = ws_fields.cell(row=row, column=4, value=section.get("content", ""))
            content_cell.border = thin_border
            content_cell.alignment = cell_align

            is_hw = section.get("is_handwritten", False)
            hw_cell = ws_fields.cell(row=row, column=5, value="Yes" if is_hw else "No")
            hw_cell.border = thin_border
            if is_hw:
                for c in range(1, 7):
                    ws_fields.cell(row=row, column=c).fill = hw_fill

            ws_fields.cell(row=row, column=6, value=section.get("confidence", "")).border = thin_border

            for c in range(1, 7):
                ws_fields.cell(row=row, column=c).font = cell_font

            row += 1

    ws_fields.column_dimensions["A"].width = 8
    ws_fields.column_dimensions["B"].width = 16
    ws_fields.column_dimensions["C"].width = 22
    ws_fields.column_dimensions["D"].width = 60
    ws_fields.column_dimensions["E"].width = 14
    ws_fields.column_dimensions["F"].width = 12

    # --- Sheet 2: Tables ---
    ws_tables = wb.create_sheet("Tables")
    t_row = 1
    for page in doc_data.get("pages", []):
        for table in page.get("tables", []):
            title = table.get("title", "Table")
            title_cell = ws_tables.cell(row=t_row, column=1, value=f"Page {page.get('page_number', '?')}: {title}")
            title_cell.font = Font(name="Calibri", size=12, bold=True)
            t_row += 1

            table_headers = table.get("headers", [])
            for col, h in enumerate(table_headers, 1):
                cell = ws_tables.cell(row=t_row, column=col, value=h)
                cell.font = header_font
                cell.fill = header_fill
                cell.border = thin_border
            t_row += 1

            for data_row in table.get("rows", []):
                for col, val in enumerate(data_row, 1):
                    cell = ws_tables.cell(row=t_row, column=col, value=val)
                    cell.font = cell_font
                    cell.border = thin_border
                t_row += 1

            t_row += 1

    # --- Sheet 3: Raw Text ---
    ws_raw = wb.create_sheet("Raw Text")
    ws_raw.cell(row=1, column=1, value="Page").font = Font(bold=True)
    ws_raw.cell(row=1, column=2, value="Raw Text").font = Font(bold=True)
    ws_raw.column_dimensions["A"].width = 8
    ws_raw.column_dimensions["B"].width = 100

    r_row = 2
    for page in doc_data.get("pages", []):
        ws_raw.cell(row=r_row, column=1, value=page.get("page_number", "?"))
        raw_cell = ws_raw.cell(row=r_row, column=2, value=page.get("raw_text", ""))
        raw_cell.alignment = Alignment(wrap_text=True, vertical="top")
        r_row += 1

    wb.save(output_path)


# ─── Routes ───

@app.route("/")
def index():
    return send_from_directory("static", "index.html")


@app.route("/api/upload", methods=["POST"])
def upload_pdf():
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    file = request.files["file"]
    if not file.filename.lower().endswith(".pdf"):
        return jsonify({"error": "Only PDF files are supported"}), 400

    doc_id = str(uuid.uuid4())[:8]
    filename = f"{doc_id}_{file.filename}"
    filepath = UPLOAD_DIR / filename
    file.save(filepath)

    documents[doc_id] = {
        "id": doc_id,
        "filename": file.filename,
        "filepath": str(filepath),
        "status": "uploaded",
        "uploaded_at": datetime.now().isoformat(),
        "pages": [],
        "page_count": 0,
        "processing_time": 0,
    }

    return jsonify({"id": doc_id, "filename": file.filename, "status": "uploaded"})


@app.route("/api/process/<doc_id>", methods=["POST"])
def process_pdf(doc_id):
    if doc_id not in documents:
        return jsonify({"error": "Document not found"}), 404

    doc = documents[doc_id]
    if doc["status"] == "processing":
        return jsonify({"error": "Already processing"}), 409

    doc["status"] = "processing"
    start_time = time.time()

    try:
        # Convert PDF to images
        images = pdf_to_images(doc["filepath"])
        doc["page_count"] = len(images)

        # Extract text from each page
        pages = []
        for i, img_b64 in enumerate(images):
            page_result = extract_with_claude(img_b64, i + 1, len(images))
            pages.append(page_result)

        doc["pages"] = pages
        doc["status"] = "complete"
        doc["processing_time"] = round(time.time() - start_time, 1)

        # Save JSON results
        result_path = RESULTS_DIR / f"{doc_id}.json"
        with open(result_path, "w") as f:
            json.dump(doc, f, indent=2)

        return jsonify({
            "id": doc_id,
            "status": "complete",
            "page_count": len(pages),
            "processing_time": doc["processing_time"],
            "pages": pages,
        })

    except Exception as e:
        doc["status"] = "error"
        doc["error"] = str(e)
        return jsonify({"error": str(e)}), 500


@app.route("/api/export/<doc_id>")
def export_excel(doc_id):
    if doc_id not in documents:
        return jsonify({"error": "Document not found"}), 404

    doc = documents[doc_id]
    if doc["status"] != "complete":
        return jsonify({"error": "Document not yet processed"}), 400

    excel_path = RESULTS_DIR / f"{doc_id}.xlsx"
    results_to_excel(doc, str(excel_path))

    return send_file(
        excel_path,
        as_attachment=True,
        download_name=f"{Path(doc['filename']).stem}_extracted.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/documents")
def list_documents():
    docs = [
        {
            "id": d["id"],
            "filename": d["filename"],
            "status": d["status"],
            "uploaded_at": d["uploaded_at"],
            "page_count": d.get("page_count", 0),
            "processing_time": d.get("processing_time", 0),
        }
        for d in documents.values()
    ]
    return jsonify(docs)


@app.route("/api/documents/<doc_id>")
def get_document(doc_id):
    if doc_id not in documents:
        return jsonify({"error": "Document not found"}), 404
    return jsonify(documents[doc_id])


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 3011))
    print(f"\n  📄 Geo PDF Pipeline running at http://localhost:{port}\n")
    app.run(host="0.0.0.0", port=port, debug=True)
