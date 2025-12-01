# Document Commenter

A Python utility to programmatically add comments and annotations to Microsoft Word (`.docx`) and PDF documents based on a JSON specification.

## Features

- **DOCX Support**: Adds comments to paragraphs matching specific text patterns.
  - Supports exact matching and Regular Expressions.
  - Options for case sensitivity and whole-word matching.
  - Control over which occurrence to annotate (first, all, or specific index).
- **PDF Support**: Adds highlight annotations with comments.
  - specific text matching (literal search).
  - Coordinate-based positioning (page number and bounding box).
- **CLI Interface**: Simple command-line tool for easy integration.

## Requirements

- Python 3.6+
- `python-docx`
- `pymupdf` (fitz)

## Installation

1. Clone this repository.
2. Install the required Python packages:

```bash
pip install python-docx pymupdf
```

## Usage

Run the script from the command line, providing the document path and the annotations JSON file:

```bash
python comment.py <document_path> <json_path> [-o <output_path>]
```

### Arguments

- `document_path`: Path to the input `.docx` or `.pdf` file.
- `json_path`: Path to the JSON file containing annotation specifications.
- `-o`, `--output` (Optional): Path for the output file. If not provided, it defaults to `<filename>-annotated.<ext>`.

### Example

```bash
python comment.py test-example.docx annotations-example.json -o my-annotated-doc.docx
```

## JSON Configuration Format

The annotations file should be a JSON array of objects. Each object defines a target and the comment to apply.

### Structure

```json
[
  {
    "target": {
      "mode": "text",             // "text" (default) or "position" (PDF only)
      "text": "target text",      // The text to find
      "match_type": "exact",      // "exact" or "regex" (DOCX only)
      "case_sensitive": false,    // Boolean
      "whole_word": true,         // Boolean (DOCX only)
      "occurrence": "first"       // "first", "all", or an integer (1-based)
    },
    "comment": {
      "text": "Comment content",
      "author": "Reviewer Name"
    }
  }
]
```

### PDF Position Targeting

For PDFs, you can also target by specific coordinates:

```json
{
  "target": {
    "mode": "position",
    "pdf": {
      "page": 1,           // 1-based page number
      "bbox": [100, 100, 200, 150] // [x1, y1, x2, y2]
    }
  },
  "comment": {
    "text": "Area comment",
    "author": "System"
  }
}
```
