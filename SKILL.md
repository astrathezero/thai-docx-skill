---
name: thai-docx
description: "Create, format, edit, repair, and manipulate professional Thai Word documents (.docx) following Thai government standards (ระเบียบงานสารบรรณ), Thai typography rules (TH Sarabun New 16pt, Complex Script OpenXML, thaiDistribute justification). Use whenever creating or modifying Thai documents, fixing Thai line breaks or spacing bugs (distribute letter-stretching vs huge justify gaps), formatting Thai tables, numbered/hanging lists, converting Thai PDF/reports to DOCX, or normalizing numerals and punctuation."
---

# Thai DOCX Creation, Typography Engineering, and Document Repair

A `.docx` file is a ZIP archive containing XML parts. Thai typography in Microsoft Word requires specific OpenXML Complex Script (CS) and language declarations that differ from standard Western documents.

Choose your approach by task:

| Task | Approach |
|---|---|
| **Create a new Thai document** | Use `scripts/thai_docx.py` via Python (`create_thai_document`, `add_body_paragraph`, `add_numbered_item`) |
| **Fix broken spacing / letter stretching** | Use `thai_docx.apply_thai_distribute_to_document("in.docx", "out.docx")` |
| **Inspect document alignments** | Use `thai_docx.inspect_document_alignments("in.docx")` |
| **Table formatting & styling** | Use `thai_docx.style_table(table, col_widths=...)` |
| **Numerals / punctuation normalization** | Use `thai_to_arabic()`, `arabic_to_thai()`, `remove_colons()` |

---

## 1. Quick Start (Python with `thai_docx.py`)

A helper module is located at `scripts/thai_docx.py`. It handles all OpenXML Complex Script and Thai Distributed alignment details:

```python
import sys, os
sys.path.append(os.path.expanduser("~/.claude/skills/thai-docx/scripts"))
import thai_docx

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
table = doc.add_table(rows=3, cols=2)
table.cell(0, 0).text = "หัวข้อ"
table.cell(0, 1).text = "รายละเอียด"
table.cell(1, 0).text = "ข้อมูล 1"
table.cell(1, 1).text = "รายละเอียด 1"
thai_docx.style_table(table, col_widths=[2.0, 4.5], header_bg="1B365D", alt_bg="F4F6F9")

doc.save("output_thai.docx")
```

---

## 2. Thai Typography & The OpenXML Architecture

Thai script flows continuously without word spaces (*Scriptio continua*). In Microsoft Word (especially on macOS), standard Western formatting breaks Thai layout.

### The Two Justification Pitfalls

1. **The Word-Gap Explosion (`เว้นว่างไว้เยอะเกิน`)**:
   - Occurs when standard Western Justify (`<w:jc w:val="both"/>` / `WD_ALIGN_PARAGRAPH.JUSTIFY`) is applied to Thai text.
   - Word only expands ASCII spaces between words, resulting in huge white gaps across the line.
2. **The Letter-Stretching Defect (`อ ุ ต ส า ห ก ร ร ม...`)**:
   - Occurs when Thai Distributed (`<w:jc w:val="thaiDistribute"/>`) is applied without proper Complex Script (CS) tags, or when a line wraps prematurely because Word cannot find a dictionary break in a long compound word.
   - Word treats the short line as needing full distribution and forces letter-spacing tracking across all individual Thai characters.

### The Solution: True Thai Justify (`thaiDistribute`) with Complex Script

To achieve clean Thai justification:
- Use **`WD_ALIGN_PARAGRAPH.THAI_JUSTIFY` (`<w:jc w:val="thaiDistribute"/>`)** on all body paragraphs and list items.
- Every text run (`<w:r>`) must declare **Complex Script (CS)** properties:
  ```xml
  <w:rPr>
    <!-- 1. Map all font slots with hint="cs" -->
    <w:rFonts w:ascii="TH Sarabun New" w:hAnsi="TH Sarabun New" w:cs="TH Sarabun New" w:eastAsia="TH Sarabun New" w:hint="cs"/>
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
- Every paragraph (`<w:pPr>`) must declare paragraph-level language and font hint:
  ```xml
  <w:pPr>
    <w:jc w:val="thaiDistribute"/>
    <w:rPr>
      <w:rFonts w:hint="cs"/>
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
| **Line Spacing** | 1.15 to 1.25 | Prevents vowel collision (`่`, `้`, `๊`, `๋`, `์`, `ุ`, `ู`) |
| **Paragraph After** | 3 to 4 pt | Clean separation without extra blank lines |
| **First-Line Indent**| 0.4 to 0.5 inches (~1.0 to 1.25 cm) | Standard paragraph indentation |
| **Hanging Indent** | Left: `0.45"`, First-line: `-0.25"` | Clean alignment for `1)`, `2)`, `1.1)` lists |
| **Numerals** | Arabic `0–9` or Thai `๐–๙` | Keep consistent. Government prefers Thai digits or Arabic depending on agency directive |
| **Punctuation** | **Never use colons (`:`)** in Thai headings | In Thai, do not write `ได้แก่:` or `หัวข้อ:`. Write `ได้แก่` and `หัวข้อ` |

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

## 4. Repairing Existing Thai Documents

To batch-fix an existing `.docx` document that suffers from bad word breaks, giant gaps, or stretched characters:

```python
import thai_docx

# Automatically updates all body paragraphs and tables to Thai Distributed
# with full Complex Script properties
thai_docx.apply_thai_distribute_to_document("broken_report.docx", "fixed_report.docx")
```
