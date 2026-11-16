# Thai DOCX Skill for Anthropic Claude

A specialized, production-ready skill for **Anthropic Claude** (Claude Code, Claude Desktop, and Agent SDK) to create, format, edit, and repair professional Thai Word documents (`.docx`).

This skill bridges the gap between **Anthropic's official DOCX engineering** and **Thai OpenXML Typography standards** (solving the infamous letter-stretching and word-gap bugs on macOS and Windows).

---

## 🌟 Key Features

1. **True Thai Justify (`thaiDistribute`)**:
   - Eliminates the **Word-Gap Explosion** (`เว้นว่างไว้เยอะเกิน`) caused by standard Western Justify (`both`).
   - Eliminates the **Letter-Stretching Defect** (`อ ุ ต ส า ห ก ร ร ม...`) by correctly injecting OpenXML Complex Script (CS) and Language definitions.
2. **Thai Government & Standard Typography (มาตรฐานงานสารบรรณ)**:
   - Fonts: `TH Sarabun New` (16pt body, 16–18pt bold headings).
   - Line spacing: `1.1` for executive reports/summaries; `1.0` (Single) for government memos (ระเบียบงานสารบรรณ) and tables (eliminates loose character gaps while TH Sarabun New natively handles diacritics).
   - Paragraph indents: 0.4 inches (~1 cm).
   - Hanging indents for numbered lists: Left `0.45"`, First-line `-0.25"`.
   - Normalization: Thai ↔ Arabic numeral conversion and colon (`:`) removal from headings/lists.
3. **Professional Table Styling**:
   - Explicit column widths, custom borders, cell padding (`tcMar`), alternate row zebra striping, Single line spacing (1.0), and `cantSplit`/`tblHeader`.
4. **Zero Red Squiggly Underlines**:
   - Completely eliminates spell/grammar check red squiggles across macOS and Windows Word using OpenXML `noProof` and `word/settings.xml` error hiding.
5. **Full Anthropic DOCX Tooling**:
   - XML surgery: Unzip, coalesce runs with `merge_runs.py`, edit `word/document.xml`, re-zip, and schema validate.
   - Tracked changes acceptance (`accept_changes.py`).
   - Multi-threaded comment injection (`comment.py`).

---

## 📦 Installation for Anthropic Claude

### Option A: Install Globally in Claude Code (Recommended)

Run the included install script:
```bash
./install.sh
```

Or manually copy / symlink to your Claude skills directory:
```bash
# Global installation for Claude Code
mkdir -p ~/.claude/skills
cp -R thai-docx ~/.claude/skills/

# Or symlink
ln -s "$(pwd)/thai-docx" ~/.claude/skills/thai-docx
```

### Option B: Install in a Specific Project Repository

For repository-scoped skills in Claude Code:
```bash
mkdir -p .claude/skills
cp -R thai-docx .claude/skills/
```

---

## 🚀 How to Use with Claude Code

Once installed, Claude automatically detects the skill based on your prompts. You can trigger it naturally:

- *"สร้างเอกสารสรุปรายงานเป็น Word ภาษาไทย จัดหน้าตามระเบียบงานสารบรรณ ใช้ font TH Sarabun New"*
- *"แก้ปัญหาเอกสาร Word ภาษาไทยที่ตัวอักษรยืดห่าง หรือช่องว่างระหว่างคำกว้างเกินไป"*
- *"จัดหน้าตารางและ bullet ใน docx นี้ให้เป็นเลขอารบิกและจัดแบบ thaiDistribute"*
- *"Use the thai-docx skill to generate an executive summary Word document"*

---

## 💻 Python Quick Start (`scripts/thai_docx.py`)

You can also use the bundled Python module directly in scripts:

```python
import sys, os
# Add skill scripts to sys.path
sys.path.append(os.path.expanduser("~/.claude/skills/thai-docx/scripts"))
import thai_docx

# 1. Create document with standard 1-inch margins and TH Sarabun New
doc = thai_docx.create_thai_document(font_name="TH Sarabun New", body_size_pt=16)

# 2. Add Title and Headings
thai_docx.add_heading_1(doc, "1. บทสรุปสำหรับผู้บริหาร (Executive Summary)")
thai_docx.add_heading_2(doc, "1.1 ข้อมูลภาพรวม")

# 3. Add Body Paragraph (Thai Distributed justification + 1 cm indent)
thai_docx.add_body_paragraph(
    doc,
    "อุตสาหกรรมศูนย์ข้อมูล (Data Center) เป็นโครงสร้างพื้นฐานสำคัญที่สุดของเศรษฐกิจดิจิทัล...",
    indent=True
)

# 4. Add Numbered Item with Hanging Indent
thai_docx.add_numbered_item(
    doc,
    num_str="1)",
    bold_title="วิกฤตคำขอจองไฟฟ้าเกินจริง",
    text="ปัจจุบันมียอดคำขอใช้ไฟฟ้าสะสมสูงถึงประมาณ 35 กิกะวัตต์..."
)

# 5. Add Professional Table
table = doc.add_table(rows=3, cols=2)
table.cell(0, 0).text = "หัวข้อ"
table.cell(0, 1).text = "รายละเอียด"
table.cell(1, 0).text = "ด้านพลังงาน"
table.cell(1, 1).text = "ความต้องการใช้ไฟฟ้าสูงขึ้นอย่างมีนัยสำคัญ"
thai_docx.style_table(table, col_widths=[2.0, 4.5], header_bg="1B365D", alt_bg="F4F6F9")

doc.save("รายงานสรุป.docx")
```

---

## 🛠️ CLI Utilities Included

| Tool | Command | Description |
|---|---|---|
| **Merge Runs** | `python scripts/merge_runs.py unpacked/` | Coalesces fragmented XML text runs so Thai words can be searched/replaced. |
| **Comments** | `python scripts/comment.py unpacked/ "คอมเมนต์"` | Injects threaded comments into Word XML. |
| **Accept Changes** | `python scripts/accept_changes.py in.docx out.docx` | Cleans document and accepts all tracked revisions. |
| **Validate** | `python scripts/office/validate.py out.docx` | Validates against OpenXML ECMA-376 schemas. |

---

## 📋 Requirements

- Python 3.9+ with `python-docx` (`pip install python-docx lxml`)
- (Optional) LibreOffice (`soffice`) and `pdftoppm` for headless PDF export and visual verification.

---

## 📄 License

MIT License. Feel free to use and distribute in any personal or enterprise workflow.
