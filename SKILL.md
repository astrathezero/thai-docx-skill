---
name: thai-docx
description: "Create, format, edit, repair, and manipulate professional Thai Word documents (.docx) following Thai government standards (ระเบียบงานสารบรรณ), Thai typography rules (TH Sarabun New 16pt, Complex Script OpenXML, thaiDistribute justification), and Anthropic docx engineering best practices. Use whenever creating or modifying Thai documents, fixing Thai line breaks or spacing bugs (distribute letter-stretching vs huge justify gaps), formatting Thai tables, numbered/hanging lists, converting Thai PDF/reports to DOCX, or editing document.xml, merging runs, and managing tracked changes or comments."
---

# Thai DOCX Creation, Editing, and Typography Engineering

A `.docx` file is a ZIP archive containing XML parts. Thai typography in Microsoft Word requires specific OpenXML Complex Script (CS) and language declarations that differ from standard Western documents.

Choose your approach by task:

| Task | Approach |
|---|---|
| **Create a new Thai document** | Use `scripts/thai_docx.py` via Python (`/usr/bin/python3` or active venv) |
| **Fix broken spacing / justification** | Set `<w:jc w:val="thaiDistribute"/>` and Complex Script tagging |
| **Edit an existing document** | `unzip` → `python scripts/merge_runs.py unpacked/` → edit `word/document.xml` → `zip` |
| **Read / Inspect text** | `pandoc -t markdown file.docx` or `python-docx` |
| **Tracked changes & comments** | Use `scripts/comment.py` and `scripts/accept_changes.py` |

---

## 1. Quick Start (Python with `thai_docx.py`)

A battle-tested helper module is located at `scripts/thai_docx.py` (relative to this skill). It abstracts all OpenXML Complex Script and Thai Distributed complexities.

```python
import sys, os
# Add skill scripts directory to sys.path
SKILL_SCRIPTS = os.path.expanduser("~/.gemini/config/skills/thai-docx/scripts")
if SKILL_SCRIPTS not in sys.path:
    sys.path.append(SKILL_SCRIPTS)

import thai_docx
from docx.shared import Inches, Pt, RGBColor

# 1. Initialize Document with standard 1-inch margins and TH Sarabun New
doc = thai_docx.create_thai_document(
    margin_top_in=1.0, margin_bottom_in=1.0,
    margin_left_in=1.0, margin_right_in=1.0,
    font_name="TH Sarabun New", body_size_pt=16
)

# 2. Add Headings (Left-aligned, clean, keep_with_next=True)
thai_docx.add_heading_1(doc, "1. บทสรุปสำหรับผู้บริหาร (Executive Summary)")
thai_docx.add_heading_2(doc, "1.1 ข้อมูลภาพรวมการดำเนินงาน")

# 3. Add Body Paragraphs (Thai Distributed, first-line indent ~1 cm / 0.4")
thai_docx.add_body_paragraph(
    doc,
    "อุตสาหกรรมศูนย์ข้อมูล (Data Center) เป็นโครงสร้างพื้นฐานสำคัญที่สุดของเศรษฐกิจดิจิทัล...",
    indent=True
)

# 4. Add Numbered Items with Hanging Indent (Left: 0.45", Hanging: 0.25", Thai Distributed)
thai_docx.add_numbered_item(
    doc,
    num_str="1)",
    bold_title="วิกฤตคำขอจองไฟฟ้าเกินจริง",
    text="ปัจจุบันมียอดคำขอใช้ไฟฟ้าสะสมสูงถึงประมาณ 35 กิกะวัตต์..."
)

# 5. Add and Style Table
table = doc.add_table(rows=3, cols=3)
thai_docx.style_table(
    table,
    col_widths=[1.5, 2.5, 2.5],
    header_bg="1B365D",
    alt_bg="F4F6F9"
)

doc.save("output_thai.docx")
```

---

## 2. Thai Typography & The OpenXML Architecture

Thai script flows continuously without word spaces (*Scriptio continua*). In Microsoft Word (especially on macOS), standard Western formatting breaks Thai layout.

### The Two Justification Pitfalls

1. **The Word-Gap Explosion (`เว้นว่างไว้เยอะเกิน`)**:
   - Occurs when standard Western Justify (`<w:jc w:val="both"/>` / `WD_ALIGN_PARAGRAPH.JUSTIFY`) is applied to Thai text.
   - Word only expands the ASCII spaces between words, resulting in enormous white gaps across the line.
2. **The Letter-Stretching Defect (`อ ุ ต ส า ห ก ร ร ม...`)**:
   - Occurs when Thai Distributed (`<w:jc w:val="thaiDistribute"/>`) is applied to a line where Microsoft Word (especially on macOS) cannot find a dictionary break in a long compound word or technical term.
   - Word wraps the entire compound word prematurely to the next line.
   - Having no recognized spaces or break points to expand on the remaining short line, Word's `thaiDistribute` engine forces letter-spacing tracking across all individual Thai characters and tone marks.

### The Solution: True Thai Justify (`thaiDistribute`) with Complex Script & ZWSP Word Segmentation

To achieve perfect Thai newspaper/book justification without letter-stretching:

1. **Inject Unicode Zero-Width Space (`\u200b` / ZWSP)**:
   - Thai is written without spaces between words (*Scriptio continua*).
   - Injecting `\u200b` (Zero-Width Space) at Thai word boundaries via `thai_docx.add_zwsp(text)` gives Microsoft Word invisible, discretionary line-break opportunities.
   - When a line approaches the right margin, Word breaks cleanly at the nearest syllable/word boundary instead of wrapping whole compound words. This completely eliminates letter-stretching!
   ```python
   import thai_docx

   # Add ZWSP to a string
   text_with_zwsp = thai_docx.add_zwsp("การบริหารจัดการข้อมูลและการเปลี่ยนผ่านสู่รัฐบาลดิจิทัล")

   # Or inject ZWSP across an entire existing document (.docx)
   thai_docx.inject_zwsp_to_document("input.docx", "output_zwsp.docx")
   ```

2. **Set `thaiDistribute` on Paragraphs**:
   - Use **`WD_ALIGN_PARAGRAPH.THAI_JUSTIFY` (`<w:jc w:val="thaiDistribute"/>`)** on all body paragraphs and list items.

3. **Declare Complex Script (CS) Properties on Every Run**:
   - Every text run (`<w:r>`) must declare **Complex Script (CS)** properties:
   ```xml
   <w:rPr>
     <!-- 1. Map all font slots (TH Sarabun New) -->
     <w:rFonts w:ascii="TH Sarabun New" w:hAnsi="TH Sarabun New" w:cs="TH Sarabun New" w:eastAsia="TH Sarabun New"/>
     <!-- 2. Dual font sizes: Western w:sz and Complex Script w:szCs (in half-points; 32 = 16pt) -->
     <w:sz w:val="32"/>
     <w:szCs w:val="32"/>
     <!-- 3. Complex Script flag -->
     <w:cs/>
     <!-- 4. Complex Script bold/italic if applicable -->
     <w:bCs/>
     <!-- 5. Thai Language tagging -->
     <w:lang w:val="th-TH" w:eastAsia="th-TH" w:bidi="th-TH"/>
   </w:rPr>
   ```
4. **Paragraph Properties (`<w:pPr>`)**:
   ```xml
   <w:pPr>
     <w:jc w:val="thaiDistribute"/>
     <w:rPr>
       <w:rFonts w:cs="TH Sarabun New"/>
       <w:szCs w:val="32"/>
       <w:lang w:val="th-TH" w:eastAsia="th-TH" w:bidi="th-TH"/>
     </w:rPr>
   </w:pPr>
   ```

---

## 3. Thai Government & Publishing Standards

When formatting formal Thai documents (หนังสือราชการ / รายงานผลการศึกษา):

| Element | Standard Value | Notes |
|---|---|---|
| **Font Family** | `TH Sarabun New` (preferred) or `TH Sarabun PSK` | Standard Thai government 13 fonts |
| **Body Size** | 16 pt (`w:sz="32"`, `w:szCs="32"`) | Regular weight |
| **Heading 1** | 18 pt (`w:sz="36"`, `w:szCs="36"`) | Bold, Deep Navy `#1B365D` or Black |
| **Heading 2** | 16 pt (`w:sz="32"`, `w:szCs="32"`) | Bold, Medium Blue `#006699` or Black |
| **Line Spacing** | `1.1` (Reports) / `1.0` (Memos & Tables) | `TH Sarabun New`'s internal font metrics already handle diacritics natively. Use `1.1` for executive reports/summaries (tight, readable), and `1.0` (Single) for government memos (ระเบียบงานสารบรรณ) and table cells. Avoid 1.15–1.25 as it creates loose, fragmented layouts. |
| **Paragraph After** | 3 to 4 pt | Clean separation without extra blank lines |
| **First-Line Indent**| 0.4 to 0.5 inches (~1.0 to 1.25 cm) | Standard paragraph indentation |
| **Hanging Indent** | Left: `0.45"`, First-line: `-0.25"` | Clean alignment for `1)`, `2)`, `1.1)` lists |
| **Numerals** | Arabic `0–9` or Thai `๐–๙` | Keep consistent. Government prefers Thai digits or Arabic depending on agency directive |
| **Punctuation** | **Never use colons (`:`)** in Thai headings | In Thai, do not write `ได้แก่:` or `หัวข้อ:`. Write `ได้แก่` and `หัวข้อ` |

### Eliminating Red Squiggly Underlines (Proofing Errors)

Microsoft Word frequently marks Thai technical terms, proper nouns, and English abbreviations (PUE, Data Center, GDCC) with red squiggly underlines. To eliminate them completely on both macOS and Windows:

1. **Tag Runs with `<w:noProof/>`**:
   Add `<w:noProof/>` to `<w:rPr>` for all runs and paragraph properties.
2. **Strip Proof Error Tags**:
   Remove all `<w:proofErr w:type="spellStart"/>` and `<w:proofErr w:type="spellEnd"/>` tags from `document.xml`.
3. **Configure `word/settings.xml`**:
   Insert `<w:hideSpellingErrors/>` and `<w:hideGrammaticalErrors/>` directly before `<w:proofState>` in `word/settings.xml`.
4. **Automated Helper**:
   ```python
   thai_docx.suppress_proofing_in_docx("document.docx")
   ```

### OpenXML ECMA-376 Schema Ordering

To avoid corrupt document errors when modifying table and cell properties:
- **`w:tblPr` child order**: `tblStyle` → `tblW` → `jc` → `tblBorders` → `shd` → `tblCellMar` → `tblLook`.
- **`w:tcPr` child order**: `tcW` → `shd` (must include `w:val="clear"`) → `tcMar` → `vAlign`.
- **`w:hint` attribute**: Must only be `default` or `eastAsia`. Never set `w:hint="cs"` as it violates ECMA-376 schema. Complex Script is declared via `<w:cs/>` and `w:cs="TH Sarabun New"`.

### Text Normalization Helpers

In `thai_docx`:
```python
thai_docx.thai_to_arabic("พ.ศ. ๒๕๖๙ จำนวน ๔๒ โครงการ")
# -> "พ.ศ. 2567 จำนวน 42 โครงการ"

thai_docx.arabic_to_thai("ปี 2026")
# -> "ปี ๒๐๒๖"

thai_docx.remove_colons("วิกฤตคำขอจองไฟฟ้าเกินจริง: ได้แก่:")
# -> "วิกฤตคำขอจองไฟฟ้าเกินจริง ได้แก่"
```

---

## 4. Editing Existing Thai Documents (XML Surgery)

When updating existing `.docx` documents to fix Thai formatting or edit content:

```bash
unzip -q document.docx -d unpacked/
find unpacked -type l -delete   # strip untrusted symlinks

# 1. Merge fragmented runs so Thai text is contiguous and findable
python scripts/merge_runs.py unpacked/

# 2. Edit unpacked/word/document.xml in place (e.g., regex replacement or Python script)
# To fix justification across all body paragraphs:
# Replace: <w:jc w:val="both"/> -> <w:jc w:val="thaiDistribute"/>

# 3. Repack the document
(cd unpacked && rm -f ../repacked.docx && zip -XrD ../repacked.docx .)

# 4. Validate output
python scripts/office/validate.py repacked.docx --original document.docx
```

### Adding Comments and Tracked Changes

- **Add Comment**:
  ```bash
  python scripts/comment.py unpacked/ "ปรับแก้การจัดแนวเป็น thaiDistribute"
  ```
  Anchor with the printed `<w:commentRangeStart>` / `<w:commentRangeEnd>` in `document.xml`.
- **Accept Tracked Changes**:
  ```bash
  python scripts/accept_changes.py in.docx out_clean.docx
  ```

---

## 5. Verification & Visual Quality Inspection

Always verify the output visually to confirm Thai line-breaking and justification:

```bash
# 1. Verify XML properties with Python
/usr/bin/python3 -c "
import docx
doc = docx.Document('output.docx')
for i, p in enumerate(doc.paragraphs[:10]):
    print(f'P{i}: jc={p.alignment}, text={p.text[:50]}')
"

# 2. Render to PDF via LibreOffice headless (if available)
python scripts/office/soffice.py --headless --convert-to pdf output.docx
pdftoppm -jpeg -r 150 output.pdf page
# Then read/inspect page-*.jpg using read_file
```
