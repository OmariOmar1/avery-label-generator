# Avery 5160 Label Generator

A Python script that generates Avery 5160-compatible 30-up label documents (`.docx`) with sequential numbering. No external dependencies required — uses only Python standard library.

## Features

- Generates `.docx` files compatible with **Avery 5160** (30-up) label sheets
- Configurable starting number, page count, and label prefix
- Exact dimensions matching Avery 5160 specifications
- No third-party dependencies (uses `zipfile` and `xml` from stdlib)
- Bold, centered, 24pt text on each label

## Label Specifications

| Property | Value |
|---|---|
| Template | Avery 5160 (30-up) |
| Page Size | 8.5" x 11" (US Letter) |
| Labels per Page | 30 (3 columns x 10 rows) |
| Label Size | 2.625" x 1.0" |
| Column Gap | 0.125" |
| Top/Bottom Margins | 0.5" |
| Left/Right Margins | 0.1875" |
| Font | 24pt Bold, Centered |

## Requirements

- Python 3.6+

## Usage

```bash
python generate_labels.py --pages <NUM_PAGES> --start <START_NUMBER> [--prefix PREFIX] [--output FILENAME]
```

### Arguments

| Argument | Required | Default | Description |
|---|---|---|---|
| `--pages` | Yes | — | Number of pages (30 labels per page) |
| `--start` | Yes | — | Starting label number |
| `--prefix` | No | `BG` | Text prefix before the number |
| `--output` | No | Auto-generated | Output `.docx` filename |

### Examples

```bash
# Generate 10 pages of labels starting from 1 (BG 01 - BG 300)
python generate_labels.py --pages 10 --start 1

# Generate 30 pages starting from 800 (BG 800 - BG 1699)
python generate_labels.py --pages 30 --start 800

# Custom prefix and output filename
python generate_labels.py --pages 5 --start 301 --prefix "ITEM" --output inventory_labels.docx

# Single page of labels
python generate_labels.py --pages 1 --start 1
```

### Output

```
Generated: BG_01-300_30up_labels.docx
  Pages: 10
  Labels: BG 01 - BG 300
  Total labels: 300
  Layout: Avery 5160 (30-up, 3x10)
  Page: 8.5" x 11" Letter
  Margins: T/B 0.5", L/R 0.1875"
  Label size: 2.625" x 1.0"
```

## How It Works

The script builds a `.docx` file (which is a ZIP archive of XML files) from scratch using the Office Open XML (OOXML) standard. Each page is a table with:
- **3 label columns** (2.625" wide) for the label content
- **2 spacer columns** (0.125" wide) for gaps between labels
- **10 rows** at 1.0" height each

Labels are numbered sequentially left-to-right, top-to-bottom.

## License

MIT
