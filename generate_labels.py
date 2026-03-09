"""
Generate Avery 5160-compatible 30-up label documents (.docx)
Matches the exact layout of BG_01-300_30up_labels_v3.

Usage:
    python generate_labels.py --pages 10 --start 1 --prefix "BG" --output labels.docx
"""

import argparse
import zipfile
import os
import sys
from io import BytesIO


def twips(inches):
    return int(inches * 1440)


def build_document_xml(num_pages, start_num, prefix):
    """Build the word/document.xml content."""

    # Avery 5160 specs (matching the reference document exactly)
    PAGE_W = 12240        # 8.5 inches
    PAGE_H = 15840        # 11 inches
    MARGIN_TOP = 720      # 0.5 inches
    MARGIN_BOTTOM = 720   # 0.5 inches
    MARGIN_LEFT = 270     # 0.1875 inches
    MARGIN_RIGHT = 270    # 0.1875 inches
    LABEL_W = 3780        # 2.625 inches
    SPACER_W = 180        # 0.125 inches
    ROW_H = 1438          # ~1.0 inches (2 twips less to fit trailing paragraph)
    COLS = 3
    ROWS = 10
    LABELS_PER_PAGE = COLS * ROWS  # 30
    FONT_SIZE_HALF_PT = 48  # 24pt

    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    ns_r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    ns_mc = 'http://schemas.openxmlformats.org/markup-compatibility/2006'

    lines = []
    lines.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    lines.append(f'<w:document xmlns:w="{ns}" xmlns:r="{ns_r}">')
    lines.append('  <w:body>')

    current_num = start_num

    for page in range(num_pages):
        # Table (each table fills the page, so no explicit page breaks needed)
        lines.append('    <w:tbl>')

        # Table properties
        lines.append('      <w:tblPr>')
        lines.append(f'        <w:tblW w:type="dxa" w:w="11700"/>')
        lines.append('        <w:jc w:val="left"/>')
        lines.append('        <w:tblLayout w:type="fixed"/>')
        lines.append('        <w:tblCellMar>')
        lines.append('          <w:top w:w="0" w:type="dxa"/>')
        lines.append('          <w:start w:w="0" w:type="dxa"/>')
        lines.append('          <w:bottom w:w="0" w:type="dxa"/>')
        lines.append('          <w:end w:w="0" w:type="dxa"/>')
        lines.append('        </w:tblCellMar>')
        lines.append('        <w:tblBorders>')
        lines.append('          <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>')
        lines.append('          <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>')
        lines.append('          <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>')
        lines.append('          <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>')
        lines.append('          <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>')
        lines.append('          <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>')
        lines.append('        </w:tblBorders>')
        lines.append('      </w:tblPr>')

        # Table grid (5 equal columns as in the original template)
        GRID_COL_W = 2340  # 11700 / 5
        lines.append('      <w:tblGrid>')
        lines.append(f'        <w:gridCol w:w="{GRID_COL_W}"/>')
        lines.append(f'        <w:gridCol w:w="{GRID_COL_W}"/>')
        lines.append(f'        <w:gridCol w:w="{GRID_COL_W}"/>')
        lines.append(f'        <w:gridCol w:w="{GRID_COL_W}"/>')
        lines.append(f'        <w:gridCol w:w="{GRID_COL_W}"/>')
        lines.append('      </w:tblGrid>')

        # Rows
        for row in range(ROWS):
            lines.append('      <w:tr>')
            lines.append('        <w:trPr>')
            lines.append(f'          <w:trHeight w:val="{ROW_H}" w:hRule="exact"/>')
            lines.append('        </w:trPr>')

            for col in range(5):  # 5 actual columns (3 labels + 2 spacers)
                is_spacer = (col == 1 or col == 3)
                cell_w = SPACER_W if is_spacer else LABEL_W

                lines.append('        <w:tc>')
                lines.append('          <w:tcPr>')
                lines.append(f'            <w:tcW w:w="{cell_w}" w:type="dxa"/>')
                lines.append('            <w:vAlign w:val="center"/>')
                lines.append('          </w:tcPr>')

                if is_spacer:
                    # Empty spacer cell
                    lines.append('          <w:p>')
                    lines.append('            <w:pPr>')
                    lines.append('              <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>')
                    lines.append('            </w:pPr>')
                    lines.append('          </w:p>')
                else:
                    # Label cell
                    # Format number with leading zeros based on max number
                    max_num = start_num + (num_pages * LABELS_PER_PAGE) - 1
                    if max_num < 100:
                        num_str = f"{current_num:02d}"
                    elif max_num < 1000:
                        num_str = f"{current_num:02d}" if current_num < 100 else str(current_num)
                    else:
                        num_str = f"{current_num:02d}" if current_num < 100 else str(current_num)

                    label_text = f"{prefix} {num_str}"

                    lines.append('          <w:p>')
                    lines.append('            <w:pPr>')
                    lines.append('              <w:jc w:val="center"/>')
                    lines.append('              <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>')
                    lines.append('            </w:pPr>')
                    lines.append('            <w:r>')
                    lines.append('              <w:rPr>')
                    lines.append(f'                <w:sz w:val="{FONT_SIZE_HALF_PT}"/>')
                    lines.append(f'                <w:szCs w:val="{FONT_SIZE_HALF_PT}"/>')
                    lines.append('                <w:b/>')
                    lines.append('                <w:bCs/>')
                    lines.append('              </w:rPr>')
                    lines.append(f'              <w:t>{label_text}</w:t>')
                    lines.append('            </w:r>')
                    lines.append('          </w:p>')

                lines.append('        </w:tc>')
                if not is_spacer:
                    current_num += 1

            lines.append('      </w:tr>')

        lines.append('    </w:tbl>')

    # Final paragraph with section properties embedded to prevent blank trailing page
    lines.append('    <w:p>')
    lines.append('      <w:pPr>')
    lines.append('        <w:spacing w:before="0" w:after="0" w:line="0" w:lineRule="exact"/>')
    lines.append('        <w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>')
    lines.append('        <w:sectPr>')
    lines.append(f'          <w:pgSz w:w="{PAGE_W}" w:h="{PAGE_H}"/>')
    lines.append(f'          <w:pgMar w:top="{MARGIN_TOP}" w:right="{MARGIN_RIGHT}" '
                 f'w:bottom="{MARGIN_BOTTOM}" w:left="{MARGIN_LEFT}" '
                 f'w:header="720" w:footer="720" w:gutter="0"/>')
    lines.append('          <w:cols w:space="720"/>')
    lines.append('        </w:sectPr>')
    lines.append('      </w:pPr>')
    lines.append('    </w:p>')

    lines.append('  </w:body>')
    lines.append('</w:document>')

    return '\n'.join(lines)


def build_content_types():
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''


def build_rels():
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''


def build_document_rels():
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''


def build_styles():
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{ns}">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:sz w:val="24"/>
        <w:szCs w:val="24"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
</w:styles>'''


def build_settings():
    ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="{ns}">
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
</w:settings>'''


def generate_labels(num_pages, start_num, prefix, output_path):
    """Generate the .docx file."""
    buf = BytesIO()

    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', build_content_types())
        zf.writestr('_rels/.rels', build_rels())
        zf.writestr('word/_rels/document.xml.rels', build_document_rels())
        zf.writestr('word/document.xml', build_document_xml(num_pages, start_num, prefix))
        zf.writestr('word/styles.xml', build_styles())
        zf.writestr('word/settings.xml', build_settings())

    with open(output_path, 'wb') as f:
        f.write(buf.getvalue())

    labels_per_page = 30
    total = num_pages * labels_per_page
    end_num = start_num + total - 1
    print(f"Generated: {output_path}")
    print(f"  Pages: {num_pages}")
    print(f"  Labels: {prefix} {start_num:02d} - {prefix} {end_num}")
    print(f"  Total labels: {total}")
    print(f"  Layout: Avery 5160 (30-up, 3x10)")
    print(f"  Page: 8.5\" x 11\" Letter")
    print(f"  Margins: T/B 0.5\", L/R 0.1875\"")
    print(f"  Label size: 2.625\" x 1.0\"")


def main():
    parser = argparse.ArgumentParser(
        description='Generate Avery 5160-compatible 30-up label documents (.docx)',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  python generate_labels.py --pages 10 --start 1
  python generate_labels.py --pages 5 --start 301 --prefix "BG" --output labels_301-450.docx
  python generate_labels.py --pages 1 --start 1 --prefix "ITEM"
        '''
    )
    parser.add_argument('--pages', type=int, required=True,
                        help='Number of pages (30 labels per page)')
    parser.add_argument('--start', type=int, required=True,
                        help='Starting label number')
    parser.add_argument('--prefix', type=str, default='BG',
                        help='Label prefix (default: BG)')
    parser.add_argument('--output', type=str, default=None,
                        help='Output file path (default: auto-generated name)')

    args = parser.parse_args()

    if args.pages < 1:
        print("Error: --pages must be at least 1")
        sys.exit(1)
    if args.start < 0:
        print("Error: --start must be non-negative")
        sys.exit(1)

    if args.output is None:
        end_num = args.start + (args.pages * 30) - 1
        args.output = f"{args.prefix}_{args.start:02d}-{end_num}_30up_labels.docx"

    generate_labels(args.pages, args.start, args.prefix, args.output)


if __name__ == '__main__':
    main()
