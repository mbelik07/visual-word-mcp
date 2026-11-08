import base64
import io
import json
import os
import re
import uuid
import threading
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Literal, Optional

import httpx
from docx import Document
from docx.shared import Mm
from mcp.server.fastmcp import FastMCP
from PIL import Image, ImageDraw, ImageFont
from pydantic import BaseModel, Field

# ---------------------------------------------------------------------------
# Word Document Editor for Mac mini 2024 - Apple silicon
# Workspace shape: document_id + pages[] + image_url = thumbnail workspace
# All user facing labels use Document / Page with Australian spelling.
# ---------------------------------------------------------------------------

mcp = FastMCP(
    "visual-word-editor",
    instructions="""When an <edit_request> block is present with object_type="document", the document_id is ALREADY PROVIDED in <metadata><id>. Use it directly - do NOT call find_word_document. Only use find_word_document when no edit context exists and the user references a document by name or title.

For create_word_document:
- Provide a `pages` array; each page: title, content, type, notes.
- Available page types: cover_page, content_page, worksheet, infographic, comparison_table, case_study, checklist_summary, closing_page.
- Available styles: modern, minimal, corporate, academic, creative, vibrant, monochrome.
- Each page is rendered as a high-resolution A4 PORTRAIT page image for Word.
- After create or edit, ALWAYS call export_word_document(format="docx") so a Word file is compiled into the exports folder.
- Use Australian English spelling in all rendered text.

For edit_word_document:
- Operations: regenerate_page, add_page, remove_page, reorder_pages, update_page_content, update_style.
- regenerate_page: re-renders the page image. Provide page_number and optionally title, content, page_type, additional_instructions.
- update_page_content: ONLY for teacher notes. Any visual change MUST use regenerate_page.
- add_page: provide title, content, page_type, position (0-indexed).
- remove_page: page_number (1-indexed). reorder_pages: order as 0-indexed positions."""
)

# ---------------------------------------------------------------------------
# Storage + local image server so image_url is viewable in the workspace
# ---------------------------------------------------------------------------

BASE_DIR = Path(os.environ.get("VISUAL_WORD_MCP_DIR", str(Path.home() / "Documents" / "VisualWordMCP")))
RECORDS_DIR = BASE_DIR / "records"
IMAGES_DIR = BASE_DIR / "images"
EXPORTS_DIR = BASE_DIR / "exports"
for directory in (RECORDS_DIR, IMAGES_DIR, EXPORTS_DIR):
    directory.mkdir(parents=True, exist_ok=True)

STATIC_PORT = int(os.environ.get("VISUAL_WORD_STATIC_PORT", "8765"))
IMAGE_BASE = os.environ.get("VISUAL_WORD_IMAGE_BASE_URL", f"http://127.0.0.1:{STATIC_PORT}")

A4_WIDTH_PX = 1240
A4_HEIGHT_PX = 1754


def _start_static_server() -> None:
    from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler

    class _QuietHandler(SimpleHTTPRequestHandler):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, directory=str(BASE_DIR), **kwargs)

        def log_message(self, *args, **kwargs):
            pass

    try:
        server = ThreadingHTTPServer(("127.0.0.1", STATIC_PORT), _QuietHandler)
        threading.Thread(target=server.serve_forever, daemon=True).start()
    except OSError:
        pass


_start_static_server()

# ---------------------------------------------------------------------------
# Schemas - Document / Page only
# ---------------------------------------------------------------------------

PageType = Literal[
    "cover_page",
    "content_page",
    "worksheet",
    "infographic",
    "comparison_table",
    "case_study",
    "checklist_summary",
    "closing_page"
]

DocumentStyle = Literal["modern", "minimal", "corporate", "academic", "creative", "vibrant", "monochrome"]

EditOperationType = Literal[
    "regenerate_page",
    "add_page",
    "remove_page",
    "reorder_pages",
    "update_page_content",
    "update_style"
]


class PageDefinition(BaseModel):
    title: str = Field(..., description="Page heading")
    content: str = Field(..., description="Text, dot points, diagram descriptions or worksheet questions to render visually")
    type: PageType = Field(..., description="Visual layout for this Word page")
    notes: Optional[str] = Field(default="", description="Teacher notes, not rendered on the page")


class EditOperation(BaseModel):
    type: EditOperationType = Field(..., description="Edit operation to perform")
    page_number: Optional[int] = Field(default=None, description="1-indexed page number")
    position: Optional[int] = Field(default=None, description="For add_page: 0-indexed insert position")
    order: Optional[List[int]] = Field(default=None, description="For reorder_pages: new order as 0-indexed positions")
    title: Optional[str] = Field(default=None)
    content: Optional[str] = Field(default=None)
    notes: Optional[str] = Field(default=None)
    page_type: Optional[PageType] = Field(default=None)
    style: Optional[DocumentStyle] = Field(default=None)
    colour_scheme: Optional[str] = Field(default=None)
    include_page_numbers: Optional[bool] = Field(default=None)
    additional_instructions: Optional[str] = Field(default=None, description="Extra visual instructions when regenerating a page image")
    annotations: Optional[Any] = Field(default=None, description="Drawn annotations from the workspace editor")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _slugify(text: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-")
    return slug or "visual-document"


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _save_record(record: Dict[str, Any]) -> None:
    (RECORDS_DIR / f"{record['document_id']}.json").write_text(json.dumps(record, indent=2), encoding="utf-8")


def _load_record(document_id: str) -> Dict[str, Any]:
    path = RECORDS_DIR / f"{document_id}.json"
    if not path.exists():
        raise ValueError(f"Document '{document_id}' not found.")
    return json.loads(path.read_text(encoding="utf-8"))


def _build_page_prompt(doc_title, page_title, page_content, layout, style, colour_scheme, page_number, include_numbers, additional_instructions=None) -> str:
    parts = [
        f"Design a vertical A4 portrait Word document page ({layout} layout) in a '{style}' editorial style.",
        "Use Australian English spelling for all rendered text.",
        f"Document Context: {doc_title}",
        f"Page Heading: {page_title}",
        f"Page Content and Visual Elements to Include:\n{page_content}",
        "Ensure high legibility, clean typography, structured visual hierarchy and a print-ready layout on a clean background.",
    ]
    if colour_scheme:
        parts.append(f"Colour Palette: {colour_scheme}")
    if include_numbers:
        parts.append(f"Include subtle page number '{page_number}' in the bottom footer.")
    if additional_instructions:
        parts.append(f"Specific Visual Modifications: {additional_instructions}")
    return "\n\n".join(parts)


def _render_fallback_page_image(page_title, page_content, layout, style, page_number, output_path) -> None:
    img = Image.new("RGB", (A4_WIDTH_PX, A4_HEIGHT_PX), color=(248, 250, 252))
    draw = ImageDraw.Draw(img)
    draw.rectangle([(0, 0), (A4_WIDTH_PX, 220)], fill=(30, 58, 138))
    draw.rectangle([(0, 220), (A4_WIDTH_PX, 232)], fill=(249, 115, 22))
    try:
        title_font = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial Bold.ttf", 48)
        sub_font = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial.ttf", 28)
        body_font = ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial.ttf", 30)
    except Exception:
        title_font = sub_font = body_font = ImageFont.load_default()
    draw.text((70, 65), page_title[:45], fill=(255, 255, 255), font=title_font)
    draw.text((70, 145), f"{layout.upper()} | {style.upper()}", fill=(219, 234, 254), font=sub_font)
    draw.rounded_rectangle([(60, 290), (A4_WIDTH_PX - 60, A4_HEIGHT_PX - 120)], radius=20, fill=(255, 255, 255), outline=(203, 213, 225), width=3)
    y_text = 340
    for paragraph in page_content.split("\n"):
        line = ""
        for word in paragraph.split(" "):
            test_line = f"{line} {word}".strip()
            if len(test_line) > 58:
                draw.text((100, y_text), line, fill=(30, 41, 59), font=body_font)
                y_text += 46
                line = word
            else:
                line = test_line
        if line:
            draw.text((100, y_text), line, fill=(30, 41, 59), font=body_font)
            y_text += 52
        if y_text > A4_HEIGHT_PX - 180:
            break
    draw.text((A4_WIDTH_PX // 2 - 40, A4_HEIGHT_PX - 75), f"Page {page_number}", fill=(100, 116, 139), font=sub_font)
    img.save(output_path, format="PNG")


def _generate_page_image(doc_title, page_title, page_content, layout, style, colour_scheme, page_number, include_numbers=True, additional_instructions=None) -> Dict[str, str]:
    slug = _slugify(page_title)[:35]
    short_id = uuid.uuid4().hex[:8]
    output_path = IMAGES_DIR / f"page_{page_number}_{slug}_{short_id}.png"
    prompt = _build_page_prompt(doc_title, page_title, page_content, layout, style, colour_scheme, page_number, include_numbers, additional_instructions)
    api_key = os.environ.get("OPENAI_API_KEY")
    if api_key:
        try:
            from openai import OpenAI
            client = OpenAI(api_key=api_key)
            result = client.images.generate(model="gpt-image-1", prompt=prompt, size="1024x1536", quality="high")
            image_bytes = base64.b64decode(result.data[0].b64_json)
            with Image.open(io.BytesIO(image_bytes)) as im:
                im.convert("RGB").resize((A4_WIDTH_PX, A4_HEIGHT_PX), Image.Resampling.LANCZOS).save(output_path, format="PNG")
        except Exception:
            _render_fallback_page_image(page_title, page_content, layout, style, page_number, output_path)
    else:
        _render_fallback_page_image(page_title, page_content, layout, style, page_number, output_path)
    return {
        "image_url": f"{IMAGE_BASE}/images/{output_path.name}",
        "local_path": str(output_path),
    }


# ---------------------------------------------------------------------------
# MCP Tools - Document / Page only
# ---------------------------------------------------------------------------

@mcp.tool()
def create_word_document(
    title: str,
    pages: List[PageDefinition],
    colour_scheme: Optional[str] = None,
    include_page_numbers: bool = True,
    style: DocumentStyle = "modern",
    topic: Optional[str] = None,
    auto_export_docx: bool = True,
) -> Dict[str, Any]:
    """Creates a new visual Word document; each page is a full-page A4 portrait image."""
    document_id = str(uuid.uuid4())
    context_topic = topic or title
    rendered = []
    for idx, page in enumerate(pages):
        imgs = _generate_page_image(context_topic, page.title, page.content, page.type, style, colour_scheme, idx + 1, include_page_numbers)
        rendered.append({
            "id": str(uuid.uuid4()),
            "order": idx,
            "title": page.title,
            "content": page.content,
            "notes": page.notes or "",
            "type": page.type,
            "image_url": imgs["image_url"],
            "local_path": imgs["local_path"],
            "image_asset_id": None,
            "annotations": None,
        })
    record = {
        "document_id": document_id,
        "title": title,
        "slug": _slugify(title),
        "revision": 1,
        "updated_at": _now_iso(),
        "versions": [{"revision": 1, "label": "Initial creation", "timestamp": _now_iso()}],
        "data": {
            "colour_scheme": colour_scheme or "",
            "include_page_numbers": include_page_numbers,
            "style": style,
            "pages": rendered,
        },
    }
    _save_record(record)
    result: Dict[str, Any] = {
        "document_id": document_id,
        "title": title,
        "revision": 1,
        "page_count": len(rendered),
        "status": "completed",
        "data": record["data"],
    }
    if auto_export_docx:
        result["docx_path"] = export_word_document(document_id, format="docx")["output_path"]
    return result


@mcp.tool()
def edit_word_document(
    document_id: str,
    operations: List[EditOperation],
    label: str,
    auto_export_docx: bool = True,
) -> Dict[str, Any]:
    """Applies batch edits to a visual Word document and recompiles the .docx."""
    record = _load_record(document_id)
    data = record["data"]
    pages: List[Dict[str, Any]] = data["pages"]
    for op in operations:
        if op.type == "update_style":
            if op.style is not None:
                data["style"] = op.style
            if op.colour_scheme is not None:
                data["colour_scheme"] = op.colour_scheme
            if op.include_page_numbers is not None:
                data["include_page_numbers"] = op.include_page_numbers
        elif op.type == "update_page_content":
            if not op.page_number or not (1 <= op.page_number <= len(pages)):
                raise ValueError(f"Invalid page_number: {op.page_number}")
            if op.notes is not None:
                pages[op.page_number - 1]["notes"] = op.notes
        elif op.type == "regenerate_page":
            if not op.page_number or not (1 <= op.page_number <= len(pages)):
                raise ValueError(f"Invalid page_number: {op.page_number}")
            target = pages[op.page_number - 1]
            if op.title is not None:
                target["title"] = op.title
            if op.content is not None:
                target["content"] = op.content
            if op.notes is not None:
                target["notes"] = op.notes
            if op.page_type is not None:
                target["type"] = op.page_type
            if op.annotations is not None:
                target["annotations"] = op.annotations
            imgs = _generate_page_image(record["title"], target["title"], target["content"], target["type"], data["style"], data.get("colour_scheme"), op.page_number, data.get("include_page_numbers", True), op.additional_instructions)
            target["image_url"] = imgs["image_url"]
            target["local_path"] = imgs["local_path"]
        elif op.type == "add_page":
            insert_pos = op.position if op.position is not None and 0 <= op.position <= len(pages) else len(pages)
            s_title = op.title or "New Page"
            s_content = op.content or ""
            s_type = op.page_type or "content_page"
            imgs = _generate_page_image(record["title"], s_title, s_content, s_type, data["style"], data.get("colour_scheme"), insert_pos + 1, data.get("include_page_numbers", True), op.additional_instructions)
            pages.insert(insert_pos, {
                "id": str(uuid.uuid4()),
                "order": insert_pos,
                "title": s_title,
                "content": s_content,
                "notes": op.notes or "",
                "type": s_type,
                "image_url": imgs["image_url"],
                "local_path": imgs["local_path"],
                "image_asset_id": None,
                "annotations": None,
            })
        elif op.type == "remove_page":
            if not op.page_number or not (1 <= op.page_number <= len(pages)):
                raise ValueError(f"Invalid page_number: {op.page_number}")
            pages.pop(op.page_number - 1)
        elif op.type == "reorder_pages":
            if op.order is None or sorted(op.order) != list(range(len(pages))):
                raise ValueError("Invalid reorder positions.")
            pages[:] = [pages[i] for i in op.order]
    for idx, s in enumerate(pages):
        s["order"] = idx
    record["revision"] += 1
    record["updated_at"] = _now_iso()
    record.setdefault("versions", []).append({"revision": record["revision"], "label": label, "timestamp": _now_iso()})
    _save_record(record)
    response: Dict[str, Any] = {
        "document_id": document_id,
        "title": record["title"],
        "revision": record["revision"],
        "page_count": len(pages),
        "label": label,
        "status": "completed",
        "data": data,
    }
    if auto_export_docx:
        response["docx_path"] = export_word_document(document_id, format="docx")["output_path"]
    return response


@mcp.tool()
def get_word_document(document_id: str, page_number: Optional[int] = None) -> Dict[str, Any]:
    """Returns full page data for reasoning before edits."""
    record = _load_record(document_id)
    data = dict(record["data"])
    if page_number is not None:
        if not (1 <= page_number <= len(data["pages"])):
            raise ValueError(f"page_number {page_number} is out of range.")
        data["pages"] = [data["pages"][page_number - 1]]
    return {
        "document_id": record["document_id"],
        "title": record["title"],
        "revision": record["revision"],
        "status": "completed",
        "data": data,
    }


@mcp.tool()
def open_word_document(document_id: str) -> Dict[str, Any]:
    """Loads lightweight metadata and page summaries for the workspace thumbnail list."""
    record = _load_record(document_id)
    pages = record["data"]["pages"]
    return {
        "document_id": record["document_id"],
        "title": record["title"],
        "revision": record["revision"],
        "page_count": len(pages),
        "style": record["data"]["style"],
        "pages": [
            {
                "id": s["id"],
                "page_number": idx + 1,
                "title": s["title"],
                "type": s["type"],
                "image_url": s["image_url"],
                "has_notes": bool(s.get("notes")),
            }
            for idx, s in enumerate(pages)
        ],
        "status": "completed",
    }


@mcp.tool()
def find_word_document(query: Optional[str] = "") -> Dict[str, Any]:
    """Searches visual Word documents by title or partial query."""
    q = (query or "").strip().lower()
    results = []
    for file_path in RECORDS_DIR.glob("*.json"):
        try:
            rec = json.loads(file_path.read_text(encoding="utf-8"))
            if not q or q in rec["title"].lower() or q in rec.get("slug", ""):
                results.append({
                    "id": rec["document_id"],
                    "title": rec["title"],
                    "slug": rec.get("slug", _slugify(rec["title"])),
                    "revision": rec["revision"],
                    "page_count": len(rec["data"]["pages"]),
                    "updated_at": rec.get("updated_at", ""),
                })
        except Exception:
            continue
    results.sort(key=lambda x: x["updated_at"], reverse=True)
    return {"count": len(results), "documents": results[:20], "status": "completed"}


@mcp.tool()
def export_word_document(
    document_id: str,
    format: Literal["docx", "pdf"] = "docx",
    filename: Optional[str] = None,
) -> Dict[str, Any]:
    """Compiles all A4 page images into a full-bleed Word (.docx) or PDF file."""
    record = _load_record(document_id)
    pages = record["data"]["pages"]
    if not pages:
        raise ValueError("Document has no pages to export.")
    base_name = filename or record.get("slug") or _slugify(record["title"])
    output_path = EXPORTS_DIR / f"{base_name}.{format}"
    if format == "docx":
        doc = Document()
        section = doc.sections[0]
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.top_margin = Mm(0)
        section.bottom_margin = Mm(0)
        section.left_margin = Mm(0)
        section.right_margin = Mm(0)
        section.header_distance = Mm(0)
        section.footer_distance = Mm(0)
        for page in pages:
            paragraph = doc.add_paragraph()
            paragraph.paragraph_format.space_before = Mm(0)
            paragraph.paragraph_format.space_after = Mm(0)
            paragraph.paragraph_format.line_spacing = 1.0
            paragraph.add_run().add_picture(page["local_path"], width=Mm(210), height=Mm(297))
        doc.save(str(output_path))
    else:
        pil_images = [Image.open(s["local_path"]).convert("RGB") for s in pages]
        pil_images[0].save(str(output_path), "PDF", resolution=150.0, save_all=True, append_images=pil_images[1:])
    return {
        "document_id": document_id,
        "filename": f"{base_name}.{format}",
        "format": format,
        "page_count": len(pages),
        "output_path": str(output_path),
        "status": "completed",
    }


def main() -> None:
    transport = os.environ.get("MCP_TRANSPORT", "stdio")
    if transport == "streamable-http":
        mcp.settings.host = os.environ.get("MCP_HOST", "0.0.0.0")
        mcp.settings.port = int(os.environ.get("MCP_PORT", "8766"))
    mcp.run(transport=transport)


if __name__ == "__main__":
    main()
