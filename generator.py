"""
AGM Generator — DOCX Generator
Recreates the Singapore AGM Directors' Resolution document format.
Uses python-docx; faithfully mirrors the formatting of the sample template.
"""

from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from config import ORDINALS
import math
import re


def _add_run(para, text: str, bold=False, size_pt: float = None, underline=False, italic=False):
    run = para.add_run(text)
    run.bold = bold
    run.underline = underline
    run.italic = italic
    if size_pt:
        run.font.size = Pt(size_pt)
    return run


def _set_para_spacing(para, before_pt=0, after_pt=0, line_pt=None):
    pf = para.paragraph_format
    pf.space_before = Pt(before_pt)
    pf.space_after = Pt(after_pt)
    if line_pt:
        pf.line_spacing = Pt(line_pt)


def _add_double_border(para):
    """Add a top+bottom double-border to a paragraph (like the DIRECTORS' RESOLUTION banner)."""
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    for side in ("top", "bottom"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "double")
        el.set(qn("w:sz"), "6")
        el.set(qn("w:space"), "5")
        el.set(qn("w:color"), "auto")
        pBdr.append(el)
    pPr.append(pBdr)


class AGMGenerator:

    # ------------------------------------------------------------------
    # Public
    # ------------------------------------------------------------------

    def generate_to_bytes(self, data: dict) -> tuple[bool, bytes, str]:
        """Generate an AGM document and return its bytes (for zip download)."""
        try:
            import io
            doc = self._build_document(data)
            buf = io.BytesIO()
            doc.save(buf)
            return True, buf.getvalue(), ""
        except Exception as exc:
            import traceback
            return False, b"", f"Error: {exc}\n{traceback.format_exc()}"

    def generate(self, data: dict, output_path: str) -> tuple[bool, str]:
        """
        Generate an AGM document.
        data keys: company_name, reg_no, address, financial_year_end,
                   agm_number (int or str), agm_date, directors (list[str])
        Returns (success: bool, message: str).
        """
        try:
            doc = self._build_document(data)
            doc.save(output_path)
            
            # Generate PDF alongside the DOCX
            from pathlib import Path
            pdf_path = str(Path(output_path).with_suffix(".pdf"))
            pdf_ok = self.convert_docx_to_pdf(output_path, pdf_path)
            
            if pdf_ok:
                return True, f"Saved: {output_path} (and PDF)"
            else:
                return True, f"Saved: {output_path} (PDF conversion failed)"
        except Exception as exc:
            import traceback
            return False, f"Error: {exc}\n{traceback.format_exc()}"

    def convert_docx_to_pdf(self, docx_path: str, pdf_path: str) -> bool:
        import sys
        import subprocess
        from pathlib import Path
        
        docx_path_abs = str(Path(docx_path).resolve())
        pdf_path_abs = str(Path(pdf_path).resolve())
        
        if sys.platform == "darwin":
            # macOS: Use AppleScript to control Microsoft Word
            applescript = f'''
            tell application "Microsoft Word"
                set original_setting to screen updating
                set screen updating to false
                try
                    open "{docx_path_abs}"
                    set theDoc to active document
                    save as theDoc file format format pdf file name "{pdf_path_abs}"
                    close theDoc saving no
                on error errText number errNum
                    set screen updating to original_setting
                    error errText number errNum
                end try
                set screen updating to original_setting
            end tell
            '''
            try:
                # Run AppleScript via subprocess
                subprocess.run(["osascript", "-e", applescript], capture_output=True, text=True, check=True)
                return True
            except Exception as e:
                print(f"macOS PDF conversion error: {e}")
                return False
                
        elif sys.platform == "win32":
            # Windows: Use comtypes or win32com to control Word
            try:
                import win32com.client
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False
                doc = word.Documents.Open(docx_path_abs)
                # wdFormatPDF is 17
                doc.SaveAs(pdf_path_abs, FileFormat=17)
                doc.Close()
                word.Quit()
                return True
            except Exception as e1:
                try:
                    import comtypes.client
                    word = comtypes.client.CreateObject("Word.Application")
                    word.Visible = False
                    doc = word.Documents.Open(docx_path_abs)
                    doc.SaveAs(pdf_path_abs, FileFormat=17)
                    doc.Close()
                    word.Quit()
                    return True
                except Exception as e2:
                    print(f"Windows PDF conversion error: {e1} / {e2}")
                    return False
        else:
            # Linux/other: Try libreoffice if installed
            try:
                subprocess.run([
                    "libreoffice", "--headless", "--convert-to", "pdf", 
                    "--outdir", str(Path(pdf_path).parent), docx_path_abs
                ], check=True, capture_output=True)
                return True
            except Exception as e:
                print(f"Headless PDF conversion error: {e}")
                return False

    # ------------------------------------------------------------------
    # Document construction
    # ------------------------------------------------------------------

    def _build_document(self, data: dict) -> Document:
        doc = Document()

        # ---- Page setup (A4, 1.5 cm left/right, 2 cm top/bottom) ----
        for section in doc.sections:
            section.page_width  = Cm(21)
            section.page_height = Cm(29.7)
            section.left_margin   = Cm(3.17)
            section.right_margin  = Cm(3.17)
            section.top_margin    = Cm(2.54)
            section.bottom_margin = Cm(2.54)

        # ---- Prepare data ----
        company_name = data.get("company_name", "").strip().upper()
        reg_no       = data.get("reg_no", "").strip()
        address      = data.get("address", "").strip()
        year_end     = data.get("financial_year_end", "").strip()
        directors    = [d.strip() for d in data.get("directors", []) if d.strip()]
        agm_date     = data.get("agm_date", "").strip()

        try:
            agm_num = int(re.sub(r"\D", "", str(data.get("agm_number", "1"))) or 1)
        except ValueError:
            agm_num = 1
        ordinal_cap   = ORDINALS.get(agm_num, f"{agm_num}th")
        ordinal_upper = ordinal_cap.upper()

        # ---- Company name ----
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_para_spacing(p, after_pt=2)
        _add_run(p, company_name, bold=True, size_pt=14)

        # ---- Dotted separator ----
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_para_spacing(p, after_pt=0)
        _add_run(p, "." * 75, size_pt=10)

        # ---- Registration number ----
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_para_spacing(p, after_pt=0)
        _add_run(p, "Company Regn No.  ", size_pt=9)
        _add_run(p, reg_no, size_pt=9)

        # ---- Incorporated line ----
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_para_spacing(p, after_pt=4)
        _add_run(p, "(Incorporated in the Republic of Singapore)", size_pt=10)

        # ---- Blank line ----
        p = doc.add_paragraph()
        _set_para_spacing(p, after_pt=0)

        # ---- DIRECTORS' RESOLUTION banner ----
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_para_spacing(p, before_pt=4, after_pt=4)
        _add_double_border(p)
        _add_run(p, "DIRECTORS\u2019 RESOLUTION", bold=True, size_pt=30)

        # ---- Authority paragraph ----
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_para_spacing(p, before_pt=6, after_pt=0)
        _add_run(p, "In writing pursuant to the authority given by Article 93", size_pt=9.5)

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _set_para_spacing(p, after_pt=8)
        _add_run(p, "of the Company\u2019s Articles of Association, hereby RESOLVED:-", size_pt=9.5)

        # ---- Section: AUDITED ACCOUNTS ----
        self._section_heading(doc, "AUDITED ACCOUNTS/FINANCIAL STATEMENTS")
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _set_para_spacing(p, after_pt=6)
        _add_run(
            p,
            f"That the audited accounts/financial statements of the Company for the year ended "
            f"{year_end} having been examined by the Directors be and are hereby approved "
            f"and authorised for issue.",
            size_pt=9.5,
        )

        # ---- Section: DIRECTORS' STATEMENT ----
        self._section_heading(doc, "DIRECTORS\u2019 STATEMENT & AUDITORS\u2019 REPORT")
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _set_para_spacing(p, after_pt=6)
        _add_run(
            p,
            f"That the Directors\u2019 Statement and Report of Auditors of the Company for the "
            f"year ended {year_end} be and are hereby adopted and approved and that any two "
            f"Directors be authorised to sign on behalf of the Board the Directors\u2019 Statement.",
            size_pt=9.5,
        )

        # ---- Section: AGM notice ----
        self._section_heading(doc, f"{ordinal_upper} ANNUAL GENERAL MEETING")
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _set_para_spacing(p, after_pt=6)
        _add_run(
            p,
            f"That the {ordinal_cap} Annual General Meeting of members of the Company will be "
            f"held at {address} on  ",
            size_pt=9.5,
        )

        # ---- Section: AGENDA ----
        self._section_heading(doc, f"AGENDA OF {ordinal_upper} ANNUAL GENERAL MEETING")
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _set_para_spacing(p, after_pt=2)
        _add_run(
            p,
            f"That the Agenda of the {ordinal_cap} Annual General Meeting of the members of "
            f"the Company be respectively as follows :-",
            size_pt=9.5,
        )

        agenda_items = [
            f"To receive and approve the Company\u2019s Audited Accounts/financial statements "
            f"together with the Directors\u2019 Statement and Report of Auditors for the year ended {year_end}.",
            "To approve Auditors\u2019 Remuneration",
            "To elect Directors",
            "To appoint Auditors",
            "To transact any other business, if any.",
        ]
        letters = "abcde"
        for i, item in enumerate(agenda_items):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent  = Cm(1)
            p.paragraph_format.first_line_indent = Cm(-1)
            _set_para_spacing(p, after_pt=2)
            _add_run(p, f"({letters[i]})\t", size_pt=9.5)
            _add_run(p, item, size_pt=9.5)

        # ---- Directors signature section ----
        p = doc.add_paragraph()
        _set_para_spacing(p, before_pt=10, after_pt=2)
        _add_run(p, "DIRECTORS  ", bold=True, size_pt=9.5)

        self._add_director_signatures(doc, directors)

        # ---- Dated line ----
        p = doc.add_paragraph()
        _set_para_spacing(p, before_pt=10, after_pt=0)
        _add_run(p, "Dated this                   day of", size_pt=9.5)

        return doc

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _section_heading(self, doc: Document, text: str):
        p = doc.add_paragraph()
        _set_para_spacing(p, before_pt=8, after_pt=2)
        _add_run(p, text, bold=True, size_pt=9.5)

    def _add_director_signatures(self, doc: Document, directors: list):
        """
        Render director signature blocks in pairs (2 per row).
        Each row:  __underline__  [tab]  __underline__
                   Name           [tab]  Name
        """
        LINE = "_" * 26
        WIDE_LINE = "_" * 28

        pairs = []
        for i in range(0, len(directors), 2):
            pairs.append(directors[i: i + 2])

        for pair in pairs:
            # Signature lines row
            p = doc.add_paragraph()
            _set_para_spacing(p, before_pt=6, after_pt=0)
            if len(pair) >= 1:
                _add_run(p, LINE, size_pt=9.5)
            if len(pair) >= 2:
                _add_run(p, "\t\t\t\t", size_pt=9.5)
                _add_run(p, WIDE_LINE, size_pt=9.5)

            # Names row
            p = doc.add_paragraph()
            _set_para_spacing(p, after_pt=4)
            if len(pair) >= 1:
                _add_run(p, pair[0], size_pt=9.5)
            if len(pair) >= 2:
                _add_run(p, "\t\t\t\t\t", size_pt=9.5)
                _add_run(p, pair[1], size_pt=9.5)

            # Spacer
            p = doc.add_paragraph()
            _set_para_spacing(p, after_pt=0)
