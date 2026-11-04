import base64
import io
import json
import os
import re
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Literal, Optional

import httpx
from docx import Document
from docx.shared import Mm
from mcp.server.fastmcp import FastMCP
from PIL import Image, ImageDraw, ImageFont
from pydantic import BaseModel, Field

# Initialise the MCP Server
mcp = FastMCP(
    "visual-word-editor",
    instructions="""Use this MCP to create, edit, and export visual Microsoft Word (.docx) documents where every page is composed as a full-page A4 portrait graphic/image.

For create_visual_document:
- Provide a `pages` array with each page containing: title, content, type (cover_page|content_page|worksheet|infographic|comparison_table|case_study|checklist_summary).
- Available styles: modern, minimal, corporate, academic, creative, vibrant, monochrome.
- Each page is generated as a high-resolution portrait image fitted to A4.
- Call export_visual_document to compile all page images into a ready-to-use .docx Word document or .pdf file.

For edit_visual_document:
- Use operations to modify the document: regenerate_page, add_page, remove_page, reorder_pages, update_page_notes, update_style.
- Any visual change (changing layout, text on the graphic, colours, or diagrams) MUST use regenerate_page."""
)

# Storage Paths
BASE_DIR = Path(os.environ.get("VISUAL_WORD_MCP_DIR", str(Path.home() / "Documents" / "VisualWordMCP")))
RECORDS_DIR = BASE_DIR / "records"
IMAGES_DIR = BASE_DIR / "images"
EXPORTS_DIR = BASE_DIR / "exports"

for directory in (RECORDS_DIR, IMAGES_DIR, EXPORTS_DIR):
    directory.mkdir(parents=True, exist_ok=True)

# A4 @ 150 DPI pixel dimensions
A4_WIDTH_PX = 1240
A4_HEIGHT_PX = 1754

# Type Definitions
PageType = Literal[
    "cover_page",
    "content_page",
    "worksheet",
    "infographic",
    "comparison_table",
    "case_study",
    "checklist_summary",
]

DocumentStyle = Literal[
    "modern",
    "minimal",
    "corporate",
    "academic",
    "creative",
    "vibrant",
    "monochrome",
]

EditOperationType = Literal[
    "regenerate_page",
    "add_page",
    "remove_page",
    "reorder_pages",
    "update_page_notes",
    "update_style",
]


class PageDefinition(BaseModel):
    title: str = Field(..., description="Heading or title for this page")
    content: str = Field(
        ...,
        description="Detailed text, bullet points, diagram descriptions, or worksheet questions to render visually on the page",
    )
    type: PageType = Field(..., description="Visual page layout type")
    notes: Optional[str] = Field(
        default="", description="Internal author/teacher notes (not rendered on page)"
    )


class EditOperation(BaseModel):
    type: EditOperationType = Field(..., description="Edit operation to perform")
    page_number: Optional[int] = Field(
        default=None, description="1-indexed page number"
    )
    position: Optional[int] = Field(
        default=None, description="For add_page: 0-indexed insertion position"
    )
    order: Optional[List[int]] = Field(
        default=None, description="For reorder_pages: new order as 0-indexed positions"
    )
    title: Optional[str] = Field(default=None)
    content: Optional[str] = Field(default=None)
    notes: Optional[str] = Field(default=None)
    page_type: Optional[PageType] = Field(default=None)
    style: Optional[DocumentStyle] = Field(default=None)
    color_scheme: Optional[str] = Field(default=None)
    include_page_numbers: Optional[bool] = Field(default=None)
    additional_instructions: Optional[str] = Field(
        default=None,
        description="Extra visual instructions when regenerating a page image",
    )


# Helper Functions
def _slugify(text: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-")
    return slug or "visual-document"


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _save_record(record: Dict[str, Any]) -> None:
    path = RECORDS_DIR / f"{record['document_id']}.json"
    path.write_text(json.dumps(record, indent=2), encoding="utf-8")


def _load_record(document_id: str) -> Dict[str, Any]:
    path = RECORDS_DIR / f"{document_id}.json"
    if not path.exists():
        raise ValueError(f"Document '{document_id}' not found.")
    return json.loads(path.read_text(encoding="utf-8"))


def _build_page_prompt(
    doc_title: str,
    page_title: str,
    page_content: str,
    page_type: str,
    style: str,
    color_scheme: Optional[str],
    page_number: int,
    include_page_numbers: bool,
    additional_instructions: Optional[str] = None,
) -> str:
    parts = [
        f"Design a vertical A4 portrait document page ({page_type} layout) in a '{style}' editorial style.",
        "Use Australian English spelling for all rendered text.",
        f"Document Context: {doc_title}",
        f"Page Heading: {page_title}",
        f"Page Content & Visual Elements to Include:\n{page_content}",
        "Ensure high legibility, clean typography, structured visual hierarchy, and print-ready layout on a clean background.",
    ]
    if color_scheme:
        parts.append(f"Colour Palette: {color_scheme}")
    if include_page_numbers:
        parts.append(f"Include subtle page number '{page_number}' at the bottom footer.")
    if additional_instructions:
        parts.append(f"Specific Visual Modifications: {additional_instructions}")
    return "\n\n".join(parts)


def _render_fallback_page_image(
    page_title: str,
    page_content: str,
    page_type: str,
    style: str,
    page_number: int,
    output_path: Path,
) -> None:
    """Local fallback renderer using Pillow for offline testing."""
    img = Image.new("RGB", (A4_WIDTH_PX, A4_HEIGHT_PX), color=(248, 250, 252))
    draw = ImageDraw.Draw(img)

    # Top header banner
    draw.rectangle([(0, 0), (A4_WIDTH_PX, 220)], fill=(30, 58, 138))
    draw.rectangle([(0, 220), (A4_WIDTH_PX, 232)], fill=(249, 115, 22))

    # Load macOS system font if available
    try:
        title_font = ImageFont.truetype(
            "/System/Library/Fonts/Supplemental/Arial Bold.ttf", 48
        )
        sub_font = ImageFont.truetype(
            "/System/Library/Fonts/Supplemental/Arial.ttf", 28
        )
        body_font = ImageFont.truetype(
            "/System/Library/Fonts/Supplemental/Arial.ttf", 30
        )
    except Exception:
        title_font = ImageFont.load_default()
        sub_font = ImageFont.load_default()
        body_font = ImageFont.load_default()

    draw.text((70, 65), page_title[:45], fill=(255, 255, 255), font=title_font)
    draw.text(
        (70, 145),
        f"Layout: {page_type.upper()}  |  Style: {style.upper()}",
        fill=(219, 234, 254),
        font=sub_font,
    )

    # Content card box
    draw.rounded_rectangle(
        [(60, 290), (A4_WIDTH_PX - 60, A4_HEIGHT_PX - 120)],
        radius=20,
        fill=(255, 255, 255),
        outline=(203, 213, 225),
        width=3,
    )

    # Wrap and draw body text
    y_text = 340
    for paragraph in page_content.split("\n"):
        words = paragraph.split(" ")
        line = ""
        for word in words:
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

    # Footer
    draw.text(
        (A4_WIDTH_PX // 2 - 40, A4_HEIGHT_PX - 75),
        f"Page {page_number}",
        fill=(100, 116, 139),
        font=sub_font,
    )
    img.save(output_path, format="PNG")


def _generate_page_image(
    doc_title: str,
    page_title: str,
    page_content: str,
    page_type: str,
    style: str,
    color_scheme: Optional[str],
    page_number: int,
    include_page_numbers: bool = True,
    additional_instructions: Optional[str] = None,
) -> str:
    """Generates a portrait page image and saves it locally."""
    slug = _slugify(page_title)[:35]
    short_id = uuid.uuid4().hex[:8]
    output_path = IMAGES_DIR / f"page_{page_number}_{slug}_{short_id}.png"

    prompt = _build_page_prompt(
        doc_title=doc_title,
        page_title=page_title,
        page_content=page_content,
        page_type=page_type,
        style=style,
        color_scheme=color_scheme,
        page_number=page_number,
        include_page_numbers=include_page_numbers,
        additional_instructions=additional_instructions,
    )

    api_key = os.environ.get("OPENAI_API_KEY")
    if api_key:
        try:
            from openai import OpenAI

            client = OpenAI(api_key=api_key)
            result = client.images.generate(
                model="gpt-image-1",
                prompt=prompt,
                size="1024x1536",
                quality="high",
            )
            image_bytes = base64.b64decode(result.data[0].b64_json)
            with Image.open(io.BytesIO(image_bytes)) as im:
                a4_im = im.convert("RGB").resize(
                    (A4_WIDTH_PX, A4_HEIGHT_PX), Image.Resampling.LANCZOS
                )
                a4_im.save(output_path, format="PNG")
            return str(output_path)
        except Exception:
            pass

    _render_fallback_page_image(
        page_title=page_title,
        page_content=page_content,
        page_type=page_type,
        style=style,
        page_number=page_number,
        output_path=output_path,
    )
    return str(output_path)


def _load_image_file(image_path_or_url: str) -> Path:
    """Ensures we have a local PNG path on disk."""
    if image_path_or_url.startswith(("http://", "https://")):
        temp_path = IMAGES_DIR / f"downloaded_{uuid.uuid4().hex[:8]}.png"
        with httpx.Client(timeout=30.0) as client:
            resp = client.get(image_path_or_url)
            resp.raise_for_status()
            temp_path.write_bytes(resp.content)
        return temp_path
    return Path(image_path_or_url)


# MCP Tools
@mcp.tool()
def create_visual_document(
    title: str,
    pages: List[PageDefinition],
    color_scheme: Optional[str] = None,
    include_page_numbers: bool = True,
    style: DocumentStyle = "modern",
    topic: Optional[str] = None,
    auto_export_docx: bool = True,
) -> Dict[str, Any]:
    """Creates a new visual Word document where each page is generated as a full-page portrait image."""
    document_id = str(uuid.uuid4())
    context_topic = topic or title

    rendered_pages = []
    for idx, page in enumerate(pages):
        image_path = _generate_page_image(
            doc_title=context_topic,
            page_title=page.title,
            page_content=page.content,
            page_type=page.type,
            style=style,
            color_scheme=color_scheme,
            page_number=idx + 1,
            include_page_numbers=include_page_numbers,
        )
        rendered_pages.append(
            {
                "id": str(uuid.uuid4()),
                "order": idx,
                "title": page.title,
                "content": page.content,
                "notes": page.notes or "",
                "type": page.type,
                "image_path": image_path,
            }
        )

    record = {
        "document_id": document_id,
        "title": title,
        "slug": _slugify(title),
        "revision": 1,
        "updated_at": _now_iso(),
        "versions": [
            {"revision": 1, "label": "Initial creation", "timestamp": _now_iso()}
        ],
        "data": {
            "page_size": "A4_PORTRAIT",
            "color_scheme": color_scheme or "",
            "include_page_numbers": include_page_numbers,
            "style": style,
            "pages": rendered_pages,
        },
    }
    _save_record(record)

    result: Dict[str, Any] = {
        "document_id": document_id,
        "title": title,
        "revision": 1,
        "page_count": len(rendered_pages),
        "status": "completed",
        "data": record["data"],
    }

    if auto_export_docx:
        export_info = export_visual_document(document_id=document_id, format="docx")
        result["docx_path"] = export_info["output_path"]

    return result


@mcp.tool()
def edit_visual_document(
    document_id: str,
    operations: List[EditOperation],
    label: str,
    auto_export_docx: bool = True,
) -> Dict[str, Any]:
    """Applies edits to a visual Word document."""
    record = _load_record(document_id)
    data = record["data"]
    pages: List[Dict[str, Any]] = data["pages"]

    for op in operations:
        if op.type == "update_style":
            if op.style is not None:
                data["style"] = op.style
            if op.color_scheme is not None:
                data["color_scheme"] = op.color_scheme
            if op.include_page_numbers is not None:
                data["include_page_numbers"] = op.include_page_numbers

        elif op.type == "update_page_notes":
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

            target["image_path"] = _generate_page_image(
                doc_title=record["title"],
                page_title=target["title"],
                page_content=target["content"],
                page_type=target["type"],
                style=data["style"],
                color_scheme=data.get("color_scheme"),
                page_number=op.page_number,
                include_page_numbers=data.get("include_page_numbers", True),
                additional_instructions=op.additional_instructions,
            )

        elif op.type == "add_page":
            insert_pos = (
                op.position
                if op.position is not None and 0 <= op.position <= len(pages)
                else len(pages)
            )
            p_title = op.title or "New Page"
            p_content = op.content or ""
            p_type = op.page_type or "content_page"
            image_path = _generate_page_image(
                doc_title=record["title"],
                page_title=p_title,
                page_content=p_content,
                page_type=p_type,
                style=data["style"],
                color_scheme=data.get("color_scheme"),
                page_number=insert_pos + 1,
                include_page_numbers=data.get("include_page_numbers", True),
                additional_instructions=op.additional_instructions,
            )
            pages.insert(
                insert_pos,
                {
                    "id": str(uuid.uuid4()),
                    "order": insert_pos,
                    "title": p_title,
                    "content": p_content,
                    "notes": op.notes or "",
                    "type": p_type,
                    "image_path": image_path,
                },
            )

        elif op.type == "remove_page":
            if not op.page_number or not (1 <= op.page_number <= len(pages)):
                raise ValueError(f"Invalid page_number: {op.page_number}")
            pages.pop(op.page_number - 1)

        elif op.type == "reorder_pages":
            if op.order is None or sorted(op.order) != list(range(len(pages))):
                raise ValueError("Invalid reorder positions.")
            pages[:] = [pages[i] for i in op.order]

    for idx, p in enumerate(pages):
        p["order"] = idx

    record["revision"] += 1
    record["updated_at"] = _now_iso()
    record.setdefault("versions", []).append(
        {"revision": record["revision"], "label": label, "timestamp": _now_iso()}
    )
    _save_record(record)

    response: Dict[str, Any] = {
        "document_id": document_id,
        "title": record["title"],
        "revision": record["revision"],
        "page_count": len(pages),
        "label": label,
        "status": "completed",
    }

    if auto_export_docx:
        export_info = export_visual_document(document_id=document_id, format="docx")
        response["docx_path"] = export_info["output_path"]

    return response


@mcp.tool()
def get_visual_document(
    document_id: str,
    page_number: Optional[int] = None,
) -> Dict[str, Any]:
    """Returns the full page data of a visual Word document for AI reasoning."""
    record = _load_record(document_id)
    data = dict(record["data"])

    if page_number is not None:
        pages = data["pages"]
        if not (1 <= page_number <= len(pages)):
            raise ValueError(f"page_number {page_number} is out of range.")
        data["pages"] = [pages[page_number - 1]]

    return {
        "document_id": record["document_id"],
        "title": record["title"],
        "revision": record["revision"],
        "status": "completed",
        "data": data,
    }


@mcp.tool()
def open_visual_document(document_id: str) -> Dict[str, Any]:
    """Loads lightweight visual document metadata and page summaries."""
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
                "id": p["id"],
                "page_number": idx + 1,
                "title": p["title"],
                "type": p["type"],
                "image_path": p["image_path"],
                "has_notes": bool(p.get("notes")),
            }
            for idx, p in enumerate(pages)
        ],
        "status": "completed",
    }


@mcp.tool()
def find_visual_document(query: Optional[str] = "") -> Dict[str, Any]:
    """Searches for visual Word documents by title or partial query match."""
    q = (query or "").strip().lower()
    results = []

    for file_path in RECORDS_DIR.glob("*.json"):
        try:
            rec = json.loads(file_path.read_text(encoding="utf-8"))
            if not q or q in rec["title"].lower() or q in rec.get("slug", ""):
                results.append(
                    {
                        "id": rec["document_id"],
                        "title": rec["title"],
                        "slug": rec.get("slug", _slugify(rec["title"])),
                        "revision": rec["revision"],
                        "page_count": len(rec["data"]["pages"]),
                        "updated_at": rec.get("updated_at", ""),
                    }
                )
        except Exception:
            continue

    results.sort(key=lambda x: x["updated_at"], reverse=True)
    return {
        "count": len(results),
        "documents": results[:20],
        "status": "completed",
    }


@mcp.tool()
def export_visual_document(
    document_id: str,
    format: Literal["docx", "pdf"] = "docx",
    filename: Optional[str] = None,
) -> Dict[str, Any]:
    """Compiles all generated page images into a full-bleed A4 Word (.docx) or PDF (.pdf) file."""
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
            img_file = _load_image_file(page["image_path"])
            paragraph = doc.add_paragraph()
            paragraph.paragraph_format.space_before = Mm(0)
            paragraph.paragraph_format.space_after = Mm(0)
            paragraph.paragraph_format.line_spacing = 1.0

            run = paragraph.add_run()
            run.add_picture(str(img_file), width=Mm(210), height=Mm(297))

        doc.save(str(output_path))

    elif format == "pdf":
        pil_images = [
            Image.open(_load_image_file(page["image_path"])).convert("RGB")
            for page in pages
        ]
        pil_images[0].save(
            str(output_path),
            "PDF",
            resolution=150.0,
            save_all=True,
            append_images=pil_images[1:],
        )

    return {
        "document_id": document_id,
        "filename": f"{base_name}.{format}",
        "format": format,
        "page_count": len(pages),
        "output_path": str(output_path),
        "status": "completed",
    }


def main() -> None:
    """Entry point for uvx / GitHub execution."""
    mcp.run()


if __name__ == "__main__":
    main()
