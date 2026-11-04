# Visual Word Editor MCP

An MCP (Model Context Protocol) server that composes full-page A4 portrait images and compiles them into Microsoft Word (.docx) and PDF documents.

## Overview

Instead of generating Word documents with text and styling, this MCP generates **complete A4 portrait page graphics** (using AI image generation or local rendering) and embeds them seamlessly into Word documents with zero margins—creating print-ready visual documents that look exactly as designed.

### Key Features

- **Visual Page Generation**: Each page is rendered as a high-resolution A4 portrait image
- **Multiple Page Layouts**: cover_page, content_page, worksheet, infographic, comparison_table, case_study, checklist_summary
- **Flexible Styling**: modern, minimal, corporate, academic, creative, vibrant, monochrome
- **Full-Bleed A4 Embedding**: Images fill complete A4 pages with zero margins in Word
- **.docx & .pdf Export**: Compile page images into ready-to-print documents
- **Batch Edit Operations**: Regenerate, add, remove, or reorder pages atomically
- **Revision History**: Track document versions and changes
- **OpenAI Integration**: Uses gpt-image-1 (1024x1536 portrait) when API key provided
- **Fallback Rendering**: Local macOS Pillow renderer for offline testing
- **Australian English**: All generated text uses Australian spelling conventions

## Installation

### Option 1: Run Locally via `uv`

```bash
# Clone the repository
git clone https://github.com/mbelik07/visual-word-mcp.git
cd visual-word-mcp

# Install dependencies (requires uv)
uv sync

# Run the server
uv run visual-word-mcp
```

### Option 2: Run via `uvx` Directly from GitHub

No local installation needed:

```bash
uvx --from git+https://github.com/mbelik07/visual-word-mcp visual-word-mcp
```

## Configuration

### Environment Variables

Set these variables before running the server:

| Variable | Description | Default |
|----------|-------------|----------|
| `OPENAI_API_KEY` | OpenAI API key for gpt-image-1 image generation | (none—uses fallback) |
| `VISUAL_WORD_MCP_DIR` | Directory for storing document records, images, and exports | `~/Documents/VisualWordMCP` |

### Claude Desktop Setup

Add this to `~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "visual-word-editor": {
      "command": "uvx",
      "args": [
        "--from",
        "git+https://github.com/mbelik07/visual-word-mcp",
        "visual-word-mcp"
      ],
      "env": {
        "OPENAI_API_KEY": "sk-your-openai-api-key-here"
      }
    }
  }
}
```

## Tools

### `create_visual_document`

Creates a new visual Word document with AI-generated page images.

**Parameters:**
- `title` (string, required): Document title
- `pages` (array, required): Array of page definitions
  - `title`: Page heading
  - `content`: Page text/content to render visually
  - `type`: Page layout type (cover_page, content_page, worksheet, infographic, comparison_table, case_study, checklist_summary)
  - `notes`: Internal speaker/author notes (not rendered)
- `style` (string): Visual style (modern, minimal, corporate, academic, creative, vibrant, monochrome) — default: modern
- `color_scheme` (string, optional): Colour palette description
- `include_page_numbers` (boolean): Add page numbers to footer — default: true
- `topic` (string, optional): Context/topic for image generation
- `auto_export_docx` (boolean): Automatically export .docx after creation — default: true

**Returns:**
- `document_id`: Unique document identifier
- `title`: Document title
- `revision`: Current revision number
- `page_count`: Total number of pages
- `docx_path`: Path to exported .docx file (if auto_export_docx=true)

### `edit_visual_document`

Apply batch edit operations to an existing document.

**Parameters:**
- `document_id` (string, required): Document ID to edit
- `operations` (array, required): Array of edit operations
  - `type`: Operation type (regenerate_page, add_page, remove_page, reorder_pages, update_page_notes, update_style)
  - `page_number`: 1-indexed page number (for page operations)
  - `title`, `content`, `notes`, `page_type`: Updated values
  - `style`, `color_scheme`, `include_page_numbers`: Global style updates
  - `additional_instructions`: Extra visual instructions for regeneration
- `label` (string, required): Description of changes (for revision history)
- `auto_export_docx` (boolean): Automatically export .docx after edits — default: true

**Returns:**
- Updated document metadata and .docx path

### `get_visual_document`

Retrieve full page data for AI reasoning before edits.

**Parameters:**
- `document_id` (string, required): Document ID
- `page_number` (integer, optional): Get single page (1-indexed)

**Returns:**
- Complete document data including all page content and images

### `open_visual_document`

Load lightweight metadata and page summaries.

**Parameters:**
- `document_id` (string, required): Document ID

**Returns:**
- Document title, revision, page count, style
- Lightweight page summaries (no full content)

### `find_visual_document`

Search for documents by title or query.

**Parameters:**
- `query` (string, optional): Search term

**Returns:**
- List of matching documents (up to 20)
- Includes ID, title, slug, revision, page count, update timestamp

### `export_visual_document`

Export compiled page images into .docx or .pdf.

**Parameters:**
- `document_id` (string, required): Document ID
- `format` (string): Export format — docx or pdf — default: docx
- `filename` (string, optional): Custom output filename

**Returns:**
- `output_path`: Full path to exported file
- `format`: Export format used
- `page_count`: Number of pages exported

## Workflow Example

```
1. create_visual_document("Health Worksheet", pages=[...], style="academic")
   → Returns: document_id, auto-exports .docx

2. User reviews the .docx in Word/Preview

3. User asks: "Change page 2 to a comparison_table layout"
   → Call: edit_visual_document(document_id, operations=[
       {type: "regenerate_page", page_number: 2, page_type: "comparison_table"}
     ], label: "Updated page 2 layout")
   → Auto-exports updated .docx

4. User asks: "Add a checklist summary page at the end"
   → Call: edit_visual_document(document_id, operations=[
       {type: "add_page", title: "Summary", content: "...", page_type: "checklist_summary"}
     ], label: "Added summary page")
   → Auto-exports updated .docx

5. export_visual_document(document_id, format="pdf")
   → Returns path to final PDF for printing/sharing
```

## Data Storage

Documents, page images, and exports are stored in subdirectories under `VISUAL_WORD_MCP_DIR`:

```
~/Documents/VisualWordMCP/
├── records/           # JSON metadata for each document
├── images/            # Generated page images (PNG)
└── exports/           # Compiled .docx and .pdf files
```

## Image Generation

### With OpenAI API (Recommended)

When `OPENAI_API_KEY` is set, uses `gpt-image-1` model with `1024x1536` portrait aspect ratio for high-quality, print-ready page graphics.

### Fallback: Local macOS Renderer

If OpenAI API key is not set, the server falls back to a local **Pillow-based renderer** that generates test A4 pages using macOS system fonts. This allows offline development and testing without spending API credits.

## Requirements

- Python 3.10+
- mcp >= 1.2.0
- pydantic >= 2.0.0
- python-docx >= 1.1.0
- pillow >= 10.0.0
- httpx >= 0.27.0
- openai >= 1.50.0 (optional, for image generation)

## Development

To develop and test locally:

```bash
# Clone and setup
git clone https://github.com/mbelik07/visual-word-mcp.git
cd visual-word-mcp
uv sync

# Run in development mode with inspector
uv run mcp dev src/visual_word_mcp/server.py

# Or run directly
uv run visual-word-mcp
```

## License

MIT License — feel free to use, modify, and distribute.

## Author

Created for Australian educators and content creators who need visual, publication-ready Word documents.