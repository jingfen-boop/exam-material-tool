
import io
from pathlib import Path
import os
import re
import json
import zipfile
from dataclasses import dataclass, asdict
from typing import List, Dict, Optional, Tuple

import streamlit as st
import fitz  # PyMuPDF
from PIL import Image
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

APP_VERSION = "Web v3.0 詳解工作台＋範本套版"

# -----------------------------
# Models
# -----------------------------
@dataclass
class Question:
    source_no: int
    page_no: int
    text: str = ""
    options: Dict[str, str] = None
    answer: str = ""
    pass_rate: Optional[float] = None
    category: str = ""
    explanation: str = ""
    teaching_focus: str = ""
    teaching: str = ""
    note_strategy: str = ""
    workbench_reviewed: bool = False
    group_id: str = ""
    group_intro: str = ""
    material: str = ""
    crop_png: bytes = b""
    visual_mode: bool = False
    reviewed: bool = False
    layout_style: str = "一般直列"
    include_image: bool = True
    image_pngs: list = None
    body_crop_png: bytes = b""
    render_mode: str = "自動"
    selected: bool = True

    def __post_init__(self):
        if self.options is None:
            self.options = {}
        if self.image_pngs is None:
            self.image_pngs = []

# -----------------------------
# PDF parsing
# -----------------------------
def pdf_text(pdf_bytes: bytes) -> str:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    return "\n".join(page.get_text("text") for page in doc)

def _row_words(words, y_center, tol=5.0):
    return [w for w in words if abs(((w[1]+w[3])/2)-y_center) <= tol]

def _header_center(words, label):
    hits = [w for w in words if w[4].strip() == label]
    if not hits:
        return None
    return sum((w[0]+w[2])/2 for w in hits) / len(hits)

def _chinese_column_bounds(words):
    hx = _header_center(words, "國文")
    if hx is None:
        return None
    other_centers = []
    for label in ["英語", "數學", "社會", "自然"]:
        for w in words:
            if w[4].strip() == label:
                other_centers.append((w[0]+w[2])/2)
    right_candidates = [x for x in other_centers if x > hx]
    right = min(right_candidates) if right_candidates else hx + 100
    left_anchor = 95.0
    xmin = (left_anchor + hx) / 2
    xmax = (hx + right) / 2
    return hx, xmin, xmax

def parse_answers(answer_pdf: bytes) -> Dict[int, str]:
    doc = fitz.open(stream=answer_pdf, filetype="pdf")
    out = {}
    for page in doc:
        words = page.get_text("words")
        bounds = _chinese_column_bounds(words)
        if not bounds:
            continue
        hx, xmin, xmax = bounds
        for w in words:
            token = w[4].strip()
            if not re.fullmatch(r"\d{1,3}", token):
                continue
            qno = int(token)
            yc = (w[1]+w[3])/2
            row = _row_words(words, yc, tol=4.8)
            candidates = []
            for rw in row:
                cx = (rw[0]+rw[2])/2
                if xmin <= cx <= xmax and re.fullmatch(r"[ABCD]", rw[4].strip()):
                    candidates.append(rw)
            if candidates:
                chosen = min(candidates, key=lambda rw: abs(((rw[0]+rw[2])/2)-hx))
                out[qno] = chosen[4].strip()
    return out

def parse_pass_rates(rate_pdf: bytes) -> Dict[int, float]:
    doc = fitz.open(stream=rate_pdf, filetype="pdf")
    out = {}
    for page in doc:
        words = page.get_text("words")
        bounds = _chinese_column_bounds(words)
        if not bounds:
            continue
        hx, xmin, xmax = bounds
        for w in words:
            token = w[4].strip()
            if not re.fullmatch(r"\d{1,3}", token):
                continue
            qno = int(token)
            yc = (w[1]+w[3])/2
            row = _row_words(words, yc, tol=5.0)
            candidates = []
            for rw in row:
                t = rw[4].strip()
                cx = (rw[0]+rw[2])/2
                if xmin <= cx <= xmax and re.fullmatch(r"(?:0\.\d+|1(?:\.0+)?)", t):
                    candidates.append(rw)
            if candidates:
                chosen = min(candidates, key=lambda rw: abs(((rw[0]+rw[2])/2)-hx))
                out[qno] = float(chosen[4])
    return out

def _line_text(line):
    return "".join(span.get("text", "") for span in line.get("spans", []))

def _all_page_lines(page):
    """Return text lines in visual reading order."""
    rows = []
    d = page.get_text("dict")
    for block in d.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            txt = _line_text(line).strip()
            if not txt:
                continue
            bbox = line.get("bbox", [0,0,0,0])
            rows.append((bbox[1], bbox[0], txt, bbox))
    return sorted(rows, key=lambda x: (x[0], x[1]))

def _sequential_question_starts(doc, expected_count):
    """
    v1.3:
    Supports both official PDF patterns:
      1.
      10. 題幹……
    Scan in visual order and accept only the next expected question number.
    """
    expected = 1
    starts = []
    for pi, page in enumerate(doc, start=1):
        if pi == 1:
            continue
        for _, _, txt, bbox in _all_page_lines(page):
            m = re.match(r"^(\d{1,2})\.\s*(.*)$", txt)
            if not m:
                continue
            qno = int(m.group(1))
            if qno == expected:
                starts.append({
                    "qno": qno,
                    "page": pi,
                    "bbox": bbox,
                    "inline_after_number": m.group(2).strip()
                })
                expected += 1
                if expected > expected_count:
                    return starts
    return starts

def _region_text_and_options(page, y0, y1, qno):
    rows = []
    for y, x, t, bbox in _all_page_lines(page):
        cy = (bbox[1]+bbox[3])/2
        if y0 <= cy < y1:
            if t in {"請翻頁繼續作答", "請不要翻到次頁！"}:
                continue
            if re.fullmatch(r"\d{1,2}", t) and bbox[1] > page.rect.height - 100:
                continue
            rows.append((y, x, t))

    stem_lines, options = [], {}
    current_opt = None
    removed_q_prefix = False

    for _, _, t in rows:
        if not removed_q_prefix:
            mq = re.match(rf"^{qno}\.\s*(.*)$", t)
            if mq:
                rest = mq.group(1).strip()
                removed_q_prefix = True
                if rest:
                    stem_lines.append(rest)
                continue

        mo = re.match(r"^\s*[\(（]\s*([ABCD])\s*[\)）]\s*(.*)", t)
        if mo:
            current_opt = mo.group(1)
            options[current_opt] = mo.group(2).strip()
        elif current_opt:
            if t.startswith("【") or t.startswith("《") or t.startswith(""):
                current_opt = None
                stem_lines.append(t)
            else:
                options[current_opt] = (options[current_opt] + " " + t).strip()
        else:
            stem_lines.append(t)

    return "\n".join(stem_lines).strip(), options

def extract_questions(question_pdf: bytes, expected_count: int = 42) -> Tuple[List[Question], Dict[int, bytes]]:
    """Web v1.3 question parser."""
    doc = fitz.open(stream=question_pdf, filetype="pdf")
    page_images = {}
    for pi, page in enumerate(doc, start=1):
        pix = page.get_pixmap(matrix=fitz.Matrix(1.25,1.25), alpha=False)
        page_images[pi] = pix.tobytes("png")

    full_text = "\n".join(page.get_text("text") for page in doc)
    groups = []
    for m in re.finditer(r"回答\s*(\d{1,2})\s*[～~\-至]\s*(\d{1,2})\s*題", full_text):
        a, b = int(m.group(1)), int(m.group(2))
        if a <= b:
            groups.append((a,b))

    starts = _sequential_question_starts(doc, expected_count)
    questions = []

    for idx, s in enumerate(starts):
        qno, pi = s["qno"], s["page"]
        page = doc[pi-1]
        y0 = max(0, s["bbox"][1]-6)

        y1 = page.rect.height - 35
        for later in starts[idx+1:]:
            if later["page"] == pi:
                y1 = max(y0+35, later["bbox"][1]-6)
                break
            if later["page"] > pi:
                break

        stem, options = _region_text_and_options(page, y0, y1, qno)

        crop_rect = fitz.Rect(0, y0, page.rect.width, min(page.rect.height, y1))
        render_scale = 1.5
        cpix = page.get_pixmap(matrix=fitz.Matrix(render_scale,render_scale), clip=crop_rect, alpha=False)
        crop = cpix.tobytes("png")

        # Create another version with ONLY the original question number erased.
        # This allows Word to use an editable answer bracket + new question number
        # while preserving the rest of the official visual layout as an image.
        body_crop = crop
        try:
            qim = Image.open(io.BytesIO(crop)).convert("RGB")
            from PIL import ImageDraw
            draw = ImageDraw.Draw(qim)
            nb = s["bbox"]
            rx0 = max(0, int((nb[0] - crop_rect.x0 - 4) * render_scale))
            ry0 = max(0, int((nb[1] - crop_rect.y0 - 4) * render_scale))
            rx1 = min(qim.width, int((nb[2] - crop_rect.x0 + 8) * render_scale))
            ry1 = min(qim.height, int((nb[3] - crop_rect.y0 + 5) * render_scale))
            draw.rectangle([rx0, ry0, rx1, ry1], fill="white")
            bout = io.BytesIO()
            qim.save(bout, format="PNG")
            body_crop = bout.getvalue()
        except Exception:
            body_crop = crop

        visual = False
        image_pngs = []
        d = page.get_text("dict")
        for block in d.get("blocks", []):
            if block.get("type") != 1:
                continue
            br = fitz.Rect(block.get("bbox"))
            if not br.intersects(crop_rect):
                continue
            # Ignore tiny decorative fragments.
            if br.width < 25 or br.height < 20 or br.width * br.height < 1200:
                continue
            visual = True
            data = block.get("image")
            if data:
                image_pngs.append(data)

        q = Question(
            source_no=qno,
            page_no=pi,
            text=stem,
            options=options,
            crop_png=crop,
            body_crop_png=body_crop,
            visual_mode=visual,
            image_pngs=image_pngs
        )
        for a,b in groups:
            if a <= qno <= b:
                q.group_id = f"{a}-{b}"
                break
        questions.append(q)

    return questions, page_images

def _effective_render_mode(q: Question) -> str:
    if q.render_mode and q.render_mode != "自動":
        return q.render_mode
    filled = len([v for v in q.options.values() if (v or "").strip()])
    if filled < 4 and q.crop_png:
        return "整題圖像"
    if q.visual_mode and q.image_pngs:
        return "圖文混合"
    return "可編輯文字"

def _question_structure_status(q: Question) -> str:
    """QA indicator that respects each question's output mode."""
    mode = _effective_render_mode(q)
    issues = []
    if mode == "整題圖像":
        if not q.body_crop_png and not q.crop_png:
            issues.append("缺題圖")
    else:
        if not q.text.strip():
            issues.append("缺題幹")
        filled = len([v for v in q.options.values() if (v or "").strip()])
        if filled < 4:
            issues.append(f"選項{filled}/4")
    if not q.answer:
        issues.append("缺答案")
    if q.pass_rate is None:
        issues.append("缺通過率")
    return "已確認" if q.reviewed else ("待校對：" + "、".join(issues) if issues else "待確認")

def _clean_option_text(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "").strip())

def _apply_structure_edit(q: Question, material: str, stem: str,
                          opt_a: str, opt_b: str, opt_c: str, opt_d: str,
                          group_id: str, visual_mode: bool, reviewed: bool):
    q.material = material.strip()
    q.text = stem.strip()
    q.options = {
        "A": _clean_option_text(opt_a),
        "B": _clean_option_text(opt_b),
        "C": _clean_option_text(opt_c),
        "D": _clean_option_text(opt_d),
    }
    q.group_id = group_id.strip()
    q.visual_mode = bool(visual_mode)
    q.reviewed = bool(reviewed)

def merge_metadata(questions, answers, rates):
    for q in questions:
        q.answer = answers.get(q.source_no, q.answer)
        q.pass_rate = rates.get(q.source_no, q.pass_rate)
    return questions

# -----------------------------
# Word output
# -----------------------------
def set_cell_shading(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)

def set_eastasia(run, font="Microsoft JhengHei"):
    run.font.name = font
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font)

def add_meta_runs(p, year, source_no, pass_rate, category):
    r = p.add_run(f"{year} 會考-{source_no}")
    set_eastasia(r)
    r.font.size = Pt(10)
    r.font.highlight_color = None
    # Grey via shading on run
    rPr = r._element.get_or_add_rPr()
    shd = OxmlElement("w:shd"); shd.set(qn("w:fill"), "D9D9D9"); rPr.append(shd)

    rate = "" if pass_rate is None else f"{pass_rate:.2f}"
    r = p.add_run(rate)
    set_eastasia(r)
    r.font.size = Pt(10)
    rPr = r._element.get_or_add_rPr()
    shd = OxmlElement("w:shd"); shd.set(qn("w:fill"), "00E5FF"); rPr.append(shd)

    if category:
        r = p.add_run(category)
        set_eastasia(r)
        r.font.size = Pt(10)
        rPr = r._element.get_or_add_rPr()
        shd = OxmlElement("w:shd"); shd.set(qn("w:fill"), "FFF200"); rPr.append(shd)

def setup_doc(doc: Document):
    sec = doc.sections[0]
    sec.page_width = Cm(21)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(1.5)
    sec.bottom_margin = Cm(1.5)
    sec.left_margin = Cm(1.6)
    sec.right_margin = Cm(1.6)
    normal = doc.styles["Normal"]
    normal.font.name = "Microsoft JhengHei"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft JhengHei")
    normal.font.size = Pt(11)

def add_header(doc, title):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title)
    set_eastasia(r)
    r.bold = True
    r.font.size = Pt(18)

    table = doc.add_table(rows=3, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.RIGHT
    table.style = "Table Grid"
    labels = ["姓名", "目標題數", "答對題數"]
    for i, lab in enumerate(labels):
        table.cell(i,0).text = lab
        table.cell(i,0).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        table.cell(i,1).text = ""
    doc.add_paragraph()

def add_question(doc, q: Question, display_no: int, year: int, teacher=False, use_crop=False):
    p = doc.add_paragraph()
    r = p.add_run(f"（{' ' + q.answer + ' ' if teacher and q.answer else '   '}）{display_no}. ")
    set_eastasia(r)
    if teacher and q.answer:
        r.font.color.rgb = RGBColor(255,0,0)

    if use_crop and q.crop_png:
        p2 = doc.add_paragraph()
        p2.add_run().add_picture(io.BytesIO(q.crop_png), width=Cm(16.6))
    else:
        if q.material.strip():
            pmaterial = doc.add_paragraph()
            rm = pmaterial.add_run(q.material)
            set_eastasia(rm)
        r = p.add_run(q.text)
        set_eastasia(r)
        for key in ["A","B","C","D"]:
            if key in q.options:
                po = doc.add_paragraph()
                ro = po.add_run(f"({key}){q.options[key]}")
                set_eastasia(ro)

    pm = doc.add_paragraph()
    add_meta_runs(pm, year, q.source_no, q.pass_rate, q.category)

    if teacher:
        box = doc.add_table(rows=1, cols=1)
        box.style = "Table Grid"
        cell = box.cell(0,0)
        p = cell.paragraphs[0]
        r = p.add_run("解析：\n")
        set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
        r2 = p.add_run(q.explanation or "（待補）")
        set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)

        if q.teaching_focus.strip():
            p = cell.add_paragraph()
            r = p.add_run("【教學重點】：\n")
            set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
            r2 = p.add_run(q.teaching_focus)
            set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)

        p = cell.add_paragraph()
        r = p.add_run("【教學步驟】：\n")
        set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
        r2 = p.add_run(q.teaching or "（待補）")
        set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)

        if q.note_strategy.strip():
            p = cell.add_paragraph()
            r = p.add_run("【筆記策略】：\n")
            set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
            r2 = p.add_run(q.note_strategy)
            set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)

    doc.add_paragraph()

def setup_exam_style_doc(doc: Document):
    """A4 exam-like page with restrained margins and no worksheet metadata."""
    sec = doc.sections[0]
    sec.page_width = Cm(21)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(1.35)
    sec.bottom_margin = Cm(1.35)
    sec.left_margin = Cm(1.45)
    sec.right_margin = Cm(1.45)
    normal = doc.styles["Normal"]
    normal.font.name = "Microsoft JhengHei"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft JhengHei")
    normal.font.size = Pt(11)

def add_exam_style_header(doc, title):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(title)
    set_eastasia(r)
    r.bold = True
    r.font.size = Pt(16)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_after = Pt(7)
    r = p.add_run("姓名：____________________")
    set_eastasia(r)
    r.font.size = Pt(10)

def add_source_crop_question(doc, q: Question, display_no: int, teacher=False,
                             show_new_number=False, show_source_meta=False):
    """
    Preserve the original PDF question block as an image.
    This is the safest way to retain dialogue balloons, tables, figures,
    vertical text and other official-exam visual layout.
    """
    if show_new_number:
        pno = doc.add_paragraph()
        pno.paragraph_format.space_before = Pt(2)
        pno.paragraph_format.space_after = Pt(2)
        r = pno.add_run(f"{display_no}.")
        set_eastasia(r)
        r.bold = True
        r.font.size = Pt(11)

    if q.crop_png:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1.5)
        p.add_run().add_picture(io.BytesIO(q.crop_png), width=Cm(17.7))
    else:
        # Fallback only when a crop is unavailable.
        add_question(doc, q, display_no, 0, teacher=False, use_crop=False)

    if show_source_meta:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1.5)
        r = p.add_run(f"原題：{q.source_no}｜通過率：{q.pass_rate if q.pass_rate is not None else '—'}")
        set_eastasia(r)
        r.font.size = Pt(8)

    if teacher:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        r = p.add_run(f"答案：{q.answer or '—'}")
        set_eastasia(r)
        r.bold = True
        r.font.color.rgb = RGBColor(255,0,0)

        box = doc.add_table(rows=1, cols=1)
        box.style = "Table Grid"
        cell = box.cell(0,0)
        p = cell.paragraphs[0]
        r = p.add_run("解析：\n")
        set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
        r2 = p.add_run(q.explanation or "（待補）")
        set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)

        if q.teaching_focus.strip():
            p = cell.add_paragraph()
            r = p.add_run("【教學重點】：\n")
            set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
            r2 = p.add_run(q.teaching_focus)
            set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)

        p = cell.add_paragraph()
        r = p.add_run("【教學步驟】：\n")
        set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
        r2 = p.add_run(q.teaching or "（待補）")
        set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)

        if q.note_strategy.strip():
            p = cell.add_paragraph()
            r = p.add_run("【筆記策略】：\n")
            set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
            r2 = p.add_run(q.note_strategy)
            set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)

    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(2)

def _add_editable_options(doc, q: Question, two_columns=False):
    keys = ["A", "B", "C", "D"]
    if two_columns:
        table = doc.add_table(rows=2, cols=2)
        table.autofit = False
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for idx, key in enumerate(keys):
            cell = table.cell(idx // 2, idx % 2)
            p = cell.paragraphs[0]
            r = p.add_run(f"({key}) {q.options.get(key, '')}")
            set_eastasia(r)
            r.font.size = Pt(10.5)
        # Remove table borders for exam-like appearance.
        tblPr = table._tbl.tblPr
        borders = OxmlElement("w:tblBorders")
        for edge in ("top","left","bottom","right","insideH","insideV"):
            el = OxmlElement(f"w:{edge}")
            el.set(qn("w:val"), "nil")
            borders.append(el)
        tblPr.append(borders)
    else:
        for key in keys:
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.7)
            p.paragraph_format.space_after = Pt(1)
            r = p.add_run(f"({key}) {q.options.get(key, '')}")
            set_eastasia(r)
            r.font.size = Pt(10.5)

def _add_first_image_to_cell(cell, q: Question, width_cm=6.3):
    if not q.include_image or not q.image_pngs:
        return
    try:
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(io.BytesIO(q.image_pngs[0]), width=Cm(width_cm))
    except Exception:
        pass



def trim_full_image_left_gutter(data: bytes, ratio: float = 0.075) -> bytes:
    """Trim the blank gutter left after the original source question number is erased."""
    try:
        im = Image.open(io.BytesIO(data)).convert("RGB")
        w, h = im.size
        cut = max(0, min(int(w * ratio), int(w * 0.12)))
        if cut <= 0:
            return data
        im = im.crop((cut, 0, w, h))
        out = io.BytesIO()
        im.save(out, format="PNG")
        return out.getvalue()
    except Exception:
        return data

def add_full_image_exam_question(doc, q: Question, display_no: int, teacher=False):
    """
    Full-image mode for complex source questions.
    Editable: answer bracket + new question number.
    Visual: original question body as one image, with the old number gutter removed.
    """
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0)
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    answer_text = q.answer if teacher and q.answer else "　"
    r = p.add_run(f"（{answer_text}）{display_no}.")
    set_eastasia(r)
    r.font.size = Pt(10.5)
    if teacher and q.answer:
        r.font.color.rgb = RGBColor(255,0,0)

    data = q.body_crop_png or q.crop_png
    if data:
        data = trim_full_image_left_gutter(data)
        pimg = doc.add_paragraph()
        pimg.paragraph_format.left_indent = Cm(0.35)
        pimg.paragraph_format.right_indent = Cm(0)
        pimg.paragraph_format.space_before = Pt(0)
        pimg.paragraph_format.space_after = Pt(2)
        try:
            # Preserve aspect ratio and keep within the usable A4 text width.
            im = Image.open(io.BytesIO(data))
            w, h = im.size
            max_w = 16.0
            # Compact natural-flow width; preserve aspect ratio and avoid forced page gaps.
            target_w = min(max_w, max(12.8, max_w if w >= 900 else 14.8))
            pimg.add_run().add_picture(io.BytesIO(data), width=Cm(target_w))
        except Exception:
            pass

    if teacher:
        p = doc.add_paragraph()
        r = p.add_run("解析：\n")
        set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
        r2 = p.add_run(q.explanation or "（待補）")
        set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)

        if q.teaching_focus.strip():
            p = doc.add_paragraph()
            r = p.add_run("【教學重點】：\n")
            set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
            r2 = p.add_run(q.teaching_focus)
            set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)

        p = doc.add_paragraph()
        r = p.add_run("【教學步驟】：\n")
        set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
        r2 = p.add_run(q.teaching or "（待補）")
        set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)



def _clean_word_text(value) -> str:
    if value is None:
        return ""
    s = str(value).replace("\u00a0", " ").replace("\u3000", " ")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n[ \t]*\n+", "\n", s)
    return s.strip()

def _remove_empty_body_paragraphs(doc):
    body = doc._element.body
    for p in list(body.findall(qn("w:p"))):
        has_text = any((t.text or "").strip() for t in p.findall(".//" + qn("w:t")))
        has_drawing = bool(p.findall(".//" + qn("w:drawing")))
        has_pagebreak = any(br.get(qn("w:type")) == "page"
                            for br in p.findall(".//" + qn("w:br")))
        if not has_text and not has_drawing and not has_pagebreak:
            body.remove(p)


def _question_is_short_for_keep(q: Question) -> bool:
    """Only compact ordinary text questions are kept on one page.
    Long material questions and image-heavy questions remain naturally pageable.
    """
    if _effective_render_mode(q) == "整題圖像":
        return False
    if q.include_image and q.image_pngs:
        return False

    material = _clean_word_text(q.material)
    stem = _clean_word_text(q.text)
    options = [_clean_word_text(q.options.get(k, "")) for k in ("A", "B", "C", "D")]

    # Conservative threshold: intended for questions like the observed Q5.
    total_chars = len(material) + len(stem) + sum(len(x) for x in options)
    nonempty_opts = sum(bool(x) for x in options)
    return (not material) and nonempty_opts >= 3 and total_chars <= 260

def _wrap_body_elements_in_keep_table(doc, start_index: int):
    """Wrap newly-added question body elements in a borderless 1-cell table.
    Word will keep the row together when it fits on one page. If it does not fit,
    Word may move the entire short question to the next page.
    """
    body = doc._element.body
    children = list(body)
    new_children = children[start_index:]
    # Exclude sectPr from moving.
    new_children = [el for el in new_children if el.tag != qn("w:sectPr")]
    if not new_children:
        return

    tbl = OxmlElement("w:tbl")
    tblPr = OxmlElement("w:tblPr")
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:w"), "0")
    tblW.set(qn("w:type"), "auto")
    tblPr.append(tblW)

    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement("w:" + edge)
        e.set(qn("w:val"), "nil")
        borders.append(e)
    tblPr.append(borders)

    cellMar = OxmlElement("w:tblCellMar")
    for edge in ("top", "left", "bottom", "right"):
        m = OxmlElement("w:" + edge)
        m.set(qn("w:w"), "0")
        m.set(qn("w:type"), "dxa")
        cellMar.append(m)
    tblPr.append(cellMar)
    tbl.append(tblPr)

    tr = OxmlElement("w:tr")
    trPr = OxmlElement("w:trPr")
    cantSplit = OxmlElement("w:cantSplit")
    trPr.append(cantSplit)
    tr.append(trPr)

    tc = OxmlElement("w:tc")
    tcPr = OxmlElement("w:tcPr")
    tcW = OxmlElement("w:tcW")
    tcW.set(qn("w:w"), "0")
    tcW.set(qn("w:type"), "auto")
    tcPr.append(tcW)
    tc.append(tcPr)

    # Move generated paragraphs/images into the cell.
    for el in new_children:
        body.remove(el)
        tc.append(el)

    # Word requires a final paragraph in a cell.
    if not list(tc) or list(tc)[-1].tag != qn("w:p"):
        tc.append(OxmlElement("w:p"))

    tr.append(tc)
    tbl.append(tr)

    # Insert before sectPr if present.
    sectPr = body.find(qn("w:sectPr"))
    if sectPr is not None:
        body.insert(list(body).index(sectPr), tbl)
    else:
        body.append(tbl)

def add_editable_exam_question(doc, q: Question, display_no: int, teacher=False):
    material_text = _clean_word_text(q.material)
    stem_text = _clean_word_text(q.text)
    clean_options = {k: _clean_word_text(q.options.get(k, "")) for k in ("A", "B", "C", "D")}

    """
    Editable-first output:
    - all extracted/manually corrected text is real Word text;
    - non-text visual material is inserted as image files;
    - layout approximates the original exam with a small set of reusable templates.
    """
    style = q.layout_style or "一般直列"

    if material_text:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(q.material)
        set_eastasia(r)
        r.font.size = Pt(10.5)

    if style == "圖片在右" and q.include_image and q.image_pngs:
        table = doc.add_table(rows=1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        table.columns[0].width = Cm(10.7)
        table.columns[1].width = Cm(6.4)

        c0, c1 = table.cell(0,0), table.cell(0,1)
        p = c0.paragraphs[0]
        r = p.add_run(f"（{' ' + q.answer + ' ' if teacher and q.answer else '   '}）{display_no}. {q.text}")
        set_eastasia(r)
        r.font.size = Pt(10.5)
        _add_first_image_to_cell(c1, q, width_cm=5.9)

        tblPr = table._tbl.tblPr
        borders = OxmlElement("w:tblBorders")
        for edge in ("top","left","bottom","right","insideH","insideV"):
            el = OxmlElement(f"w:{edge}")
            el.set(qn("w:val"), "nil")
            borders.append(el)
        tblPr.append(borders)
    else:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(f"（{' ' + q.answer + ' ' if teacher and q.answer else '   '}）{display_no}. {q.text}")
        set_eastasia(r)
        r.font.size = Pt(10.5)

        if style == "圖片在上" and q.include_image and q.image_pngs:
            ip = doc.add_paragraph()
            ip.alignment = WD_ALIGN_PARAGRAPH.CENTER
            try:
                ip.add_run().add_picture(io.BytesIO(q.image_pngs[0]), width=Cm(14.8))
            except Exception:
                pass

    # If there are additional images, preserve them below the stem.
    if q.include_image and len(q.image_pngs) > 1:
        for data in q.image_pngs[1:]:
            ip = doc.add_paragraph()
            ip.alignment = WD_ALIGN_PARAGRAPH.CENTER
            try:
                ip.add_run().add_picture(io.BytesIO(data), width=Cm(13.5))
            except Exception:
                pass

    _add_editable_options(doc, q, two_columns=(style == "選項兩欄"))

    if teacher:
        p = doc.add_paragraph()
        r = p.add_run("解析：\n")
        set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
        r2 = p.add_run(q.explanation or "（待補）")
        set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)

        if q.teaching_focus.strip():
            p = doc.add_paragraph()
            r = p.add_run("【教學重點】：\n")
            set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
            r2 = p.add_run(q.teaching_focus)
            set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)

        p = doc.add_paragraph()
        r = p.add_run("【教學步驟】：\n")
        set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
        r2 = p.add_run(q.teaching or "（待補）")
        set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)

        if q.note_strategy.strip():
            p = doc.add_paragraph()
            r = p.add_run("【筆記策略】：\n")
            set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
            r2 = p.add_run(q.note_strategy)
            set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)

    # Keep only normal paragraph spacing; do not insert an empty spacer paragraph.


def _zh_num(n: int) -> str:
    digits = "零一二三四五六七八九"
    if n < 10:
        return digits[n]
    if n < 20:
        return "十" + (digits[n % 10] if n % 10 else "")
    if n < 100:
        return digits[n // 10] + "十" + (digits[n % 10] if n % 10 else "")
    return str(n)

def _template_path(kind: str, teacher: bool) -> Path:
    base = Path(__file__).resolve().parent
    mapping = {
        ("八成以上", False): "template_80_student.docx",
        ("八成以上", True): "template_80_teacher.docx",
        ("六成至七成", False): "template_60_70_student.docx",
        ("六成至七成", True): "template_60_70_teacher.docx",
    }
    return base / mapping[(kind, teacher)]

def _load_clean_template(kind: str, teacher: bool, year: int, count: int, booklet_no: str = ""):
    """Load the user's real sample Word and retain its page/style/header area through 壹、單題."""
    path = _template_path(kind, teacher)
    if not path.exists():
        return None
    doc = Document(str(path))

    # Update the real sample title but retain its formatting.
    if doc.paragraphs:
        p0 = doc.paragraphs[0]
        old = p0.text
        label = "通過率達八成以上" if kind == "八成以上" else "通過率達六成至七成"
        booklet_label = (booklet_no or "").strip()
        if booklet_label:
            # User controls the booklet identifier exactly as entered.
            new_title = f"{year}會考國文題本-{label}-題本（{booklet_label}）"
        else:
            # Fallback only when the user leaves it blank.
            new_title = f"{year}會考國文題本-{label}-題本（{_zh_num(count)}）"
        if p0.runs:
            p0.runs[0].text = new_title
            for r in p0.runs[1:]:
                r.text = ""
        else:
            p0.text = new_title

    # Keep everything through the actual sample's "壹、單題", remove old sample questions.
    body = doc._element.body
    children = list(body)
    keep_until = None
    for i, child in enumerate(children):
        if child.tag.endswith("}p"):
            txt = "".join(child.itertext())
            if "壹、單題" in txt:
                keep_until = i
                break
    if keep_until is not None:
        for child in children[keep_until + 1:]:
            if child.tag.endswith("}sectPr"):
                continue
            body.remove(child)
    return doc


def _estimated_question_lines(q: Question) -> float:
    """Estimate vertical space without using Word keep-with-next flags.
    This avoids the black square pagination marks and large artificial gaps.
    """
    mode = _effective_render_mode(q)
    if mode == "整題圖像":
        data = q.body_crop_png or q.crop_png
        if data:
            try:
                im = Image.open(io.BytesIO(data))
                w, h = im.size
                # Approximate the displayed image height at 14.8–16.0 cm wide.
                aspect = h / max(w, 1)
                return max(12.0, min(28.0, 25.0 * aspect + 3.0))
            except Exception:
                return 18.0
        return 15.0

    chars_per_line = 34.0
    total = 2.0
    if q.material.strip():
        total += max(1.0, len(q.material.strip()) / chars_per_line)
    total += max(1.0, len((q.text or "").strip()) / chars_per_line)
    for k in ("A", "B", "C", "D"):
        opt = (q.options.get(k, "") or "").strip()
        total += max(1.0, len(opt) / chars_per_line)

    if q.include_image and q.image_pngs:
        total += 10.0
    return min(total + 2.0, 32.0)

def _add_stable_page_break_if_needed(doc, q: Question, used_lines: float, first_page: bool):
    """Return (new_used_lines, did_break).
    Uses explicit page breaks only between questions. It never sets keep_with_next,
    keep_together, bullets, or list styles.
    """
    capacity = 43.0 if first_page else 50.0
    need = _estimated_question_lines(q)

    # If the whole question is reasonably sized but won't fit, start it on next page.
    # Large questions are allowed to flow naturally in Word.
    if used_lines > 8.0 and need <= 31.0 and used_lines + need > capacity:
        doc.add_page_break()
        return need, True
    return used_lines + need, False

def make_editable_exam_layout_docx(questions: List[Question], year: int, title_suffix: str,
                                   teacher=False, template_kind="自訂簡版", booklet_no=""):
    selected = [q for q in questions if q.selected]

    doc = None
    if template_kind in ("八成以上", "六成至七成"):
        doc = _load_clean_template(template_kind, teacher, year, len(selected), booklet_no=booklet_no)

    if doc is None:
        doc = Document()
        setup_exam_style_doc(doc)
        title = (title_suffix or "").strip() or f"{year}年國中教育會考 國文科"
        add_exam_style_header(doc, title)
        p = doc.add_paragraph()
        r = p.add_run("壹、單題")
        set_eastasia(r)
        r.bold = True

    # v2.7 compact natural pagination:
    # Do not estimate or reserve an entire question's height.
    # Word is allowed to flow naturally, preventing large blank areas.
    for i, q in enumerate(selected, start=1):
        mode = _effective_render_mode(q)
        if mode == "整題圖像":
            add_full_image_exam_question(doc, q, i, teacher=teacher)
        else:
            keep_short = _question_is_short_for_keep(q)
            body = doc._element.body
            sectPr = body.find(qn("w:sectPr"))
            start_index = list(body).index(sectPr) if sectPr is not None else len(list(body))
            add_editable_exam_question(doc, q, i, teacher=teacher)
            if keep_short:
                _wrap_body_elements_in_keep_table(doc, start_index)

    out = io.BytesIO()
    _remove_empty_body_paragraphs(doc)
    doc.save(out)
    return out.getvalue()


def make_exam_layout_docx(questions: List[Question], year: int, title_suffix: str,
                          teacher=False, show_new_number=False,
                          show_source_meta=False) -> bytes:
    """
    Original-layout mode: every selected item uses its source PDF crop.
    This intentionally prioritizes visual fidelity over text editability.
    """
    doc = Document()
    setup_exam_style_doc(doc)
    title = f"{year}會考國文題本-{title_suffix}"
    add_exam_style_header(doc, title)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(5)
    r = p.add_run("壹、單題")
    set_eastasia(r)
    r.bold = True

    selected = [q for q in questions if q.selected]
    for i, q in enumerate(selected, start=1):
        add_source_crop_question(
            doc, q, i, teacher=teacher,
            show_new_number=show_new_number,
            show_source_meta=show_source_meta
        )

    out = io.BytesIO()
    _remove_empty_body_paragraphs(doc)
    doc.save(out)
    return out.getvalue()

def make_docx(questions: List[Question], year: int, title_suffix: str, teacher=False, preserve_visual=True) -> bytes:
    doc = Document()
    setup_doc(doc)
    title = f"{year}會考國文題本-{title_suffix}"
    add_header(doc, title)

    doc.add_paragraph("壹、單題").runs[0].bold = True

    selected = [q for q in questions if q.selected]
    for i, q in enumerate(selected, start=1):
        # visual-mode questions retain source crop by default
        use_crop = preserve_visual and q.visual_mode
        add_question(doc, q, i, year, teacher=teacher, use_crop=use_crop)

    out = io.BytesIO()
    _remove_empty_body_paragraphs(doc)
    doc.save(out)
    return out.getvalue()


# -----------------------------
# Manual explanation / teaching
# -----------------------------
# v1.8 deliberately contains no external AI/API calls.
# Explanations and teaching steps are edited manually in the web interface.


def parse_question_spec(spec: str, available_numbers):
    """
    Parse inputs like:
      3
      3,8,10
      1-10
      1-5,8,10,15-20
    Returns a sorted unique list limited to available question numbers.
    """
    available = set(available_numbers)
    result = set()
    spec = (spec or "").replace("，", ",").replace("～", "-").replace("~", "-").strip()
    if not spec:
        return []
    for part in [p.strip() for p in spec.split(",") if p.strip()]:
        if "-" in part:
            pieces = [x.strip() for x in part.split("-", 1)]
            if len(pieces) != 2 or not pieces[0].isdigit() or not pieces[1].isdigit():
                raise ValueError(f"無法辨識：{part}")
            a, b = int(pieces[0]), int(pieces[1])
            if a > b:
                a, b = b, a
            for n in range(a, b + 1):
                if n in available:
                    result.add(n)
        else:
            if not part.isdigit():
                raise ValueError(f"無法辨識：{part}")
            n = int(part)
            if n in available:
                result.add(n)
    return sorted(result)


# -----------------------------
# Explanation workbench reference library
# -----------------------------
def _load_reference_library():
    path = Path(__file__).resolve().parent / "reference_library_v30.json"
    if not path.exists():
        return {"publisher": {}, "strategy": {}, "history_raw": {}}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {"publisher": {}, "strategy": {}, "history_raw": {}}

def _publisher_analysis_only(block: str) -> str:
    """Extract the explanation portion when headings are available.
    Otherwise return the reference block intact so the user can judge it.
    """
    block = (block or "").strip()
    for marker in ("試題解析：", "詳解：", "詳解 ", "解析："):
        if marker in block:
            tail = block.split(marker, 1)[1].strip()
            if tail:
                return tail
    return block

def _history_examples_for_category(refdb, category: str, limit=4):
    """Show source-grounded historical excerpts around the same ability label."""
    if not category:
        return []
    out = []
    for source, raw in refdb.get("history_raw", {}).items():
        pos = raw.find(category)
        if pos < 0 and category == "表層文意理解":
            pos = raw.find("表層文意")
        if pos < 0:
            continue
        # Prefer a nearby teaching-step section after the ability label.
        teach = raw.find("【教學步驟】", pos)
        if teach >= 0 and teach - pos < 4500:
            start = max(0, pos - 350)
            end = min(len(raw), teach + 1600)
        else:
            start = max(0, pos - 350)
            end = min(len(raw), pos + 1800)
        excerpt = raw[start:end].strip()
        out.append((source, excerpt))
        if len(out) >= limit:
            break
    return out

def _draft_explanation_from_publishers(refdb, qno: int) -> str:
    """Free/no-API initial draft: use the most detailed available publisher analysis
    as an editable base. This is deliberately NOT presented as an AI synthesis.
    """
    candidates = []
    for pub in ("翰林", "康軒", "南一"):
        block = refdb.get("publisher", {}).get(pub, {}).get(str(qno), "")
        ana = _publisher_analysis_only(block)
        if ana:
            candidates.append((len(ana), pub, ana))
    if not candidates:
        return ""
    _, pub, ana = max(candidates)
    return ana.strip()

def _strategy_for_category(refdb, category: str):
    return refdb.get("strategy", {}).get(category) or refdb.get("strategy", {}).get("其他", {})

# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(page_title="會考教材產製工具", page_icon="📘", layout="wide")
st.title("📘 會考教材產製工具")
st.caption(f"{APP_VERSION}｜題本＋答案＋通過率 → 組題 → 詳解／教學 → Word")

if "questions" not in st.session_state:
    st.session_state.questions = []
if "page_images" not in st.session_state:
    st.session_state.page_images = {}
if "year" not in st.session_state:
    st.session_state.year = 115

with st.sidebar:
    st.header("設定")
    st.session_state.year = st.number_input("年度（民國）", min_value=100, max_value=150, value=int(st.session_state.year), step=1)
    st.subheader("免費版")
    st.caption("本版不使用任何外部 AI API，不需要 API Key，也不會產生 API 費用。")

tab1, tab2, tab3, tab4 = st.tabs(["① 建立題庫", "② 篩選組題", "③ 詳解工作台", "④ 產生 Word"])

with tab1:
    st.subheader("上傳三份來源")
    c1,c2,c3 = st.columns(3)
    with c1:
        qfile = st.file_uploader("原始會考題本 PDF", type=["pdf"], key="qfile")
    with c2:
        afile = st.file_uploader("官方答案 PDF", type=["pdf"], key="afile")
    with c3:
        rfile = st.file_uploader("各題通過率 PDF", type=["pdf"], key="rfile")

    if st.button("建立／更新題庫", type="primary", disabled=not (qfile and afile and rfile)):
        with st.spinner("解析 PDF、題號、答案與通過率…"):
            qbytes = qfile.getvalue()
            answers = parse_answers(afile.getvalue())
            rates = parse_pass_rates(rfile.getvalue())
            expected_count = len(answers)
            questions, page_images = extract_questions(qbytes, expected_count=expected_count)
            questions = merge_metadata(questions, answers, rates)
            st.session_state.questions = questions
            st.session_state.page_images = page_images
        expected = len(answers)
        found_nos = {q.source_no for q in questions}
        expected_nos = set(range(1, expected+1))
        missing = sorted(expected_nos - found_nos)
        if len(questions) == expected and len(rates) == expected and not missing and expected:
            st.success(f"完成：辨識 {len(questions)} 題；國文答案 {len(answers)} 題；國文通過率 {len(rates)} 題。1～{expected} 題完整。")
        else:
            miss_text = "、".join(map(str,missing)) if missing else "無"
            st.warning(f"解析完成但需校對：辨識 {len(questions)} 題；國文答案 {len(answers)} 題；國文通過率 {len(rates)} 題；缺題：{miss_text}。")

    if st.session_state.questions:
        data = []
        for q in st.session_state.questions:
            data.append({
                "原題號": q.source_no,
                "頁": q.page_no,
                "答案": q.answer,
                "通過率": q.pass_rate,
                "題組": q.group_id,
                "輸出模式": _effective_render_mode(q),
                "選項": f"{len([v for v in q.options.values() if (v or '').strip()])}/4",
                "有圖/複雜版面": "是" if q.visual_mode else "",
                "校對": _question_structure_status(q),
                "題幹預覽": q.text[:55].replace("\n"," ")
            })
        st.dataframe(data, use_container_width=True, hide_index=True)

        st.divider()
        st.markdown("### 題目結構校對")
        st.caption("免費版流程：左側看原 PDF 裁圖；右側人工修正共用材料、題幹、A–D、題組與版型。")
        st.caption("左側看原 PDF 裁圖；右側可直接修正共用材料、題幹、A–D、題組與圖片模式。修正後按「儲存本題校對」。")

        review_choices = [q.source_no for q in st.session_state.questions]
        review_no = st.selectbox("選擇要校對的原題號", review_choices, key="review_qno")
        rq = next(x for x in st.session_state.questions if x.source_no == review_no)

        rc1, rc2 = st.columns([1.05, 1])
        with rc1:
            st.markdown(f"#### 原 PDF｜第 {rq.source_no} 題")
            st.caption(f"原頁：{rq.page_no}｜答案：{rq.answer or '—'}｜通過率：{rq.pass_rate if rq.pass_rate is not None else '—'}")
            if rq.crop_png:
                st.image(rq.crop_png, caption="原 PDF 題目區塊", use_container_width=True)
            else:
                st.info("本題目前沒有裁圖。")

        with rc2:
            material_edit = st.text_area(
                "閱讀／共用材料（沒有可留白）",
                value=rq.material,
                height=130,
                key=f"review_material_{review_no}"
            )
            stem_edit = st.text_area(
                "題幹",
                value=rq.text,
                height=120,
                key=f"review_stem_{review_no}"
            )
            oa = st.text_area("A", value=rq.options.get("A",""), height=70, key=f"review_A_{review_no}")
            ob = st.text_area("B", value=rq.options.get("B",""), height=70, key=f"review_B_{review_no}")
            oc = st.text_area("C", value=rq.options.get("C",""), height=70, key=f"review_C_{review_no}")
            od = st.text_area("D", value=rq.options.get("D",""), height=70, key=f"review_D_{review_no}")
            group_edit = st.text_input(
                "題組 ID（例如 21-23；非題組留白）",
                value=rq.group_id,
                key=f"review_group_{review_no}"
            )
            render_choices = ["自動", "可編輯文字", "圖文混合", "整題圖像"]
            current_render = rq.render_mode if rq.render_mode in render_choices else "自動"
            render_edit = st.selectbox(
                "本題輸出模式",
                render_choices,
                index=render_choices.index(current_render),
                key=f"review_render_{review_no}",
                help="整題圖像：題號與答案括弧是可編輯文字，題目本體使用原 PDF 圖片；最適合第3題這類特殊排版。"
            )
            if _effective_render_mode(rq) == "整題圖像":
                st.info("本題目前採「整題圖像」：不需要補 A–D。Word 會使用可編輯的（　）與新題號，題目本體保留原 PDF 排版。")

            layout_choices = ["一般直列", "圖片在右", "圖片在上", "選項兩欄"]
            current_layout = rq.layout_style if rq.layout_style in layout_choices else "一般直列"
            layout_edit = st.selectbox(
                "可編輯 Word 版型",
                layout_choices,
                index=layout_choices.index(current_layout),
                key=f"review_layout_{review_no}",
                help="一般直列適合純文字題；圖片在右適合題幹左、圖右；圖片在上適合大型圖表；選項兩欄適合短選項。"
            )
            include_image_edit = st.checkbox(
                "可編輯版輸出獨立圖片",
                value=rq.include_image,
                key=f"review_include_image_{review_no}",
                help="若原圖本身含有大量文字且你已手動把文字重建到題幹/選項，可取消，以避免文字重複。"
            )
            visual_edit = st.checkbox(
                "保留原 PDF 裁圖作為備用／原版型輸出",
                value=rq.visual_mode,
                key=f"review_visual_{review_no}"
            )
            reviewed_edit = st.checkbox(
                "本題內容已人工確認",
                value=rq.reviewed,
                key=f"review_done_{review_no}"
            )

            if st.button("💾 儲存本題校對", type="primary", key=f"save_review_{review_no}"):
                _apply_structure_edit(
                    rq, material_edit, stem_edit, oa, ob, oc, od,
                    group_edit, visual_edit, reviewed_edit
                )
                rq.render_mode = render_edit
                rq.layout_style = layout_edit
                rq.include_image = include_image_edit
                st.success(f"第 {rq.source_no} 題已儲存。")
                st.rerun()

        reviewed_count = sum(1 for q in st.session_state.questions if q.reviewed)
        st.progress(reviewed_count / len(st.session_state.questions))
        st.caption(f"人工校對進度：{reviewed_count}/{len(st.session_state.questions)} 題。")

with tab2:
    if not st.session_state.questions:
        st.info("請先在「建立題庫」上傳三份來源並建立題庫。")
    else:
        st.subheader("依通過率篩選")
        preset = st.radio("快速條件", ["八成以上", "六成至七成", "自訂"], horizontal=True)
        if preset == "八成以上":
            lo, hi = 0.80, 1.00
        elif preset == "六成至七成":
            lo, hi = 0.60, 0.79
        else:
            cc1, cc2 = st.columns(2)
            lo = cc1.number_input("最低通過率", 0.0, 1.0, 0.60, 0.01)
            hi = cc2.number_input("最高通過率", 0.0, 1.0, 0.79, 0.01)

        keep_groups = st.checkbox("題組不可拆開（選中其中一題就保留同題組）", value=True)

        if st.button("套用通過率篩選"):
            qs = st.session_state.questions
            selected_nos = {q.source_no for q in qs if q.pass_rate is not None and lo <= q.pass_rate <= hi}
            if keep_groups:
                group_ids = {q.group_id for q in qs if q.source_no in selected_nos and q.group_id}
                for q in qs:
                    if q.group_id and q.group_id in group_ids:
                        selected_nos.add(q.source_no)
            for q in qs:
                q.selected = q.source_no in selected_nos
                st.session_state[f"sel_{q.source_no}"] = q.selected
            st.success(f"目前選取 {len(selected_nos)} 題。")
            st.rerun()

        st.divider()
        st.markdown("### 快速指定題號")
        st.caption("可以直接輸入：3、3,8,10、1-10、1-5,8,10,15-20。這會直接取代目前所有勾選。")

        quick_spec = st.text_input(
            "指定題號",
            placeholder="例如：3 或 3,8,10 或 1-10",
            key="quick_question_spec"
        )

        qc1, qc2, qc3 = st.columns(3)
        with qc1:
            if st.button("只選指定題號", use_container_width=True):
                try:
                    available = [q.source_no for q in st.session_state.questions]
                    chosen = set(parse_question_spec(quick_spec, available))
                    if not chosen:
                        st.warning("沒有選到任何題目，請檢查輸入。")
                    else:
                        if keep_groups:
                            group_ids = {q.group_id for q in st.session_state.questions if q.source_no in chosen and q.group_id}
                            for q in st.session_state.questions:
                                if q.group_id and q.group_id in group_ids:
                                    chosen.add(q.source_no)
                        for q in st.session_state.questions:
                            q.selected = q.source_no in chosen
                            st.session_state[f"sel_{q.source_no}"] = q.selected
                        st.success("已只選：" + "、".join(map(str, sorted(chosen))))
                        st.rerun()
                except Exception as e:
                    st.error(str(e))

        with qc2:
            if st.button("全部取消", use_container_width=True):
                for q in st.session_state.questions:
                    q.selected = False
                    st.session_state[f"sel_{q.source_no}"] = False
                st.rerun()

        with qc3:
            if st.button("全部選取", use_container_width=True):
                for q in st.session_state.questions:
                    q.selected = True
                    st.session_state[f"sel_{q.source_no}"] = True
                st.rerun()

        selected_now = [q.source_no for q in st.session_state.questions if q.selected]
        st.info(
            f"目前共選取 {len(selected_now)} 題"
            + (f"：{'、'.join(map(str, selected_now))}" if len(selected_now) <= 20 else "")
        )

        st.markdown("### 手動微調")
        st.caption("快速選題或通過率篩選後，可在這裡做最後加減。")
        for q in st.session_state.questions:
            rate = "—" if q.pass_rate is None else f"{q.pass_rate:.2f}"
            key = f"sel_{q.source_no}"
            if key not in st.session_state:
                st.session_state[key] = q.selected
            new_val = st.checkbox(
                f"原第 {q.source_no} 題｜通過率 {rate}｜答案 {q.answer or '—'}｜{q.category or '未分類'}",
                key=key
            )
            q.selected = new_val

with tab3:
    if not st.session_state.questions:
        st.info("請先建立題庫。")
    else:
        refdb = _load_reference_library()
        selected_nums = [q.source_no for q in st.session_state.questions if q.selected]
        choices = selected_nums or [q.source_no for q in st.session_state.questions]

        reviewed_count = sum(1 for q in st.session_state.questions if q.workbench_reviewed)
        st.caption(f"教師詳解人工確認進度：{reviewed_count}/{len(st.session_state.questions)} 題")

        qno = st.selectbox("選擇題目", choices, key="workbench_qno")
        q = next(x for x in st.session_state.questions if x.source_no == qno)

        left, right = st.columns([0.92, 1.08])

        with left:
            st.markdown(f"### 原第 {q.source_no} 題")
            if q.material.strip():
                st.markdown("**閱讀／共用材料**")
                st.write(q.material)
            st.markdown("**題幹**")
            st.write(q.text)
            for k,v in q.options.items():
                st.write(f"({k}) {v}")
            st.caption(
                f"官方答案：{q.answer or '—'}｜通過率："
                f"{q.pass_rate if q.pass_rate is not None else '—'}｜原頁：{q.page_no}"
            )
            if q.crop_png:
                st.image(q.crop_png, caption="原 PDF 題目區塊", use_container_width=True)

        with right:
            st.markdown("### 一、三家出版社詳解")
            pub_cols = st.columns(3)
            for col, pub in zip(pub_cols, ["翰林", "康軒", "南一"]):
                block = refdb.get("publisher", {}).get(pub, {}).get(str(q.source_no), "")
                with col:
                    st.markdown(f"**{pub}**")
                    if block:
                        st.text_area(
                            f"{pub}參考內容",
                            value=block,
                            height=290,
                            key=f"ref_{pub}_{qno}",
                            label_visibility="collapsed"
                        )
                    else:
                        st.warning("此題目前未解析到出版社文字，請回原檔確認。")

            st.markdown("### 二、歷年本團隊參考與能力類型")
            ability_options = ["", "字詞辨識", "表層文意理解", "文意統整", "推論理解", "分析評鑑", "其他"]
            current = q.category if q.category in ability_options else "其他"
            q.category = st.selectbox(
                "能力類型",
                ability_options,
                index=ability_options.index(current),
                key=f"cat_{qno}"
            )

            strategy = _strategy_for_category(refdb, q.category)
            if q.category:
                with st.expander("查看歷年同能力類型的教學框架", expanded=False):
                    st.markdown("**整理出的共通教學重點**")
                    st.write(strategy.get("教學重點", ""))
                    st.markdown("**整理出的共通教學步驟**")
                    st.write(strategy.get("教學步驟", ""))
                    examples = _history_examples_for_category(refdb, q.category)
                    if examples:
                        st.markdown("**112～114 年原教師版摘錄**")
                        for source, excerpt in examples:
                            st.markdown(f"**{source}**")
                            st.text_area(
                                f"{source}_{qno}",
                                value=excerpt,
                                height=180,
                                key=f"hist_{source}_{qno}",
                                label_visibility="collapsed"
                            )

            st.markdown("### 三、本次建議稿（可直接修改）")

            # Initialize widget state from question object.
            for k, value in {
                f"exp_{qno}": q.explanation,
                f"focus_{qno}": q.teaching_focus,
                f"teach_{qno}": q.teaching,
                f"note_{qno}": q.note_strategy,
            }.items():
                if k not in st.session_state:
                    st.session_state[k] = value

            b1, b2 = st.columns(2)
            with b1:
                if st.button("帶入出版社解析初稿", use_container_width=True, key=f"draft_exp_{qno}"):
                    draft = _draft_explanation_from_publishers(refdb, q.source_no)
                    if draft:
                        st.session_state[f"exp_{qno}"] = draft
                        q.explanation = draft
                        st.rerun()
                    else:
                        st.warning("目前沒有可帶入的出版社解析。")
            with b2:
                if st.button("帶入本團隊教學框架", use_container_width=True, key=f"draft_teach_{qno}"):
                    strategy = _strategy_for_category(refdb, q.category)
                    st.session_state[f"focus_{qno}"] = strategy.get("教學重點", "")
                    st.session_state[f"teach_{qno}"] = strategy.get("教學步驟", "")
                    st.session_state[f"note_{qno}"] = strategy.get("筆記策略", "")
                    q.teaching_focus = st.session_state[f"focus_{qno}"]
                    q.teaching = st.session_state[f"teach_{qno}"]
                    q.note_strategy = st.session_state[f"note_{qno}"]
                    st.rerun()

            st.caption(
                "以上兩個按鈕均為免費的規則式帶入，不使用 AI/API。"
                "「出版社解析初稿」只是方便複製修改，不代表已完成三家綜合判斷。"
            )

            q.explanation = st.text_area(
                "建議詳解",
                height=260,
                key=f"exp_{qno}"
            )
            q.teaching_focus = st.text_area(
                "教學重點",
                height=90,
                key=f"focus_{qno}"
            )
            q.teaching = st.text_area(
                "建議教學步驟",
                height=230,
                key=f"teach_{qno}"
            )
            q.note_strategy = st.text_area(
                "筆記策略（選填）",
                height=110,
                key=f"note_{qno}"
            )

            q.workbench_reviewed = st.checkbox(
                "本題詳解與教學步驟已人工確認",
                value=q.workbench_reviewed,
                key=f"wb_reviewed_{qno}"
            )
            q.visual_mode = st.checkbox(
                "輸出時保留原題裁圖（適合圖片／表格／複雜版面）",
                value=q.visual_mode,
                key=f"vis_{qno}"
            )

            st.info(
                "建議工作方式：先比較三家出版社 → 查看歷年同能力類型 → "
                "把你要的內容貼到『建議詳解／教學步驟』修改 → 人工確認。"
            )

with tab4:
    if not st.session_state.questions:
        st.info("請先建立題庫。")
    else:
        st.markdown("### 輸出題目")
        st.caption("若只想臨時輸出幾題，可直接在這裡輸入題號，不必回②重選。例如：3、3,8,10、1-10。")
        export_spec = st.text_input("本次輸出題號（留白＝使用②目前選取）", key="export_question_spec")

        selected = [q for q in st.session_state.questions if q.selected]
        export_questions = selected

        if export_spec.strip():
            try:
                available = [q.source_no for q in st.session_state.questions]
                chosen_export = set(parse_question_spec(export_spec, available))
                export_questions = [q for q in st.session_state.questions if q.source_no in chosen_export]
            except Exception as e:
                st.error(str(e))
                export_questions = []

        st.write(f"本次輸出：**{len(export_questions)} 題**"
                 + (f"（{'、'.join(str(q.source_no) for q in export_questions)}）" if len(export_questions) <= 20 else ""))

        # Temporarily map selection state for output generators without permanently changing step②.
        original_selected_states = {q.source_no: q.selected for q in st.session_state.questions}
        if export_spec.strip():
            export_nos = {q.source_no for q in export_questions}
            for q in st.session_state.questions:
                q.selected = q.source_no in export_nos

        selected = [q for q in st.session_state.questions if q.selected]
        unreviewed = [q.source_no for q in selected if not q.reviewed]
        if unreviewed:
            st.warning("目前選取題目中仍有未人工確認的題目：" + "、".join(map(str, unreviewed)) + "。原版型輸出仍可保留原 PDF 題目外觀。")
        else:
            st.success("目前選取題目皆已完成結構校對。")

        output_mode = st.radio(
            "Word 題本版型",
            ["可編輯原會考風格（推薦）", "原 PDF 圖像版", "一般可編輯文字版"],
            horizontal=True,
            help="可編輯原會考風格：文字可編輯、圖片獨立插入並套用近似版型。原 PDF 圖像版：最像原卷但題目文字不可編輯。一般可編輯文字版：版面較簡化。"
        )
        suffix = st.text_input("學生題本標題", value=f"{int(st.session_state.year)}年國中教育會考 國文科", help="這裡會原樣成為 Word 頁首主標題，可直接改成你的正式範本標題。")
        template_kind = st.radio(
            "Word 母版",
            ["八成以上", "六成至七成", "自訂簡版"],
            horizontal=True,
            help="前兩項直接套用你先前提供的114年正式成品 Word 版型；自訂簡版才使用上方自行輸入的標題。"
        )
        if template_kind != "自訂簡版":
            st.info(f"目前使用「{template_kind}」實際成品 Word 作為母版：保留原本頁面設定、標題格式、姓名欄與「壹、單題」區塊，再插入本次選題。")
        booklet_no = st.text_input(
            "題本編號",
            value="1",
            help="這是標題最後的題本編號。你可以輸入 1、2、A、甲、01 等；程式會原樣放入「題本（　）」中。"
        )
        st.caption(f"目前標題題本編號會顯示為：題本（{booklet_no.strip() or '自動'}）")


        if output_mode == "可編輯原會考風格（推薦）":
            st.info("此模式會依每題設定自動混合輸出：一般題用可編輯文字；圖文題用文字＋獨立圖片；特殊版面題可用「整題圖像」；（　）與新的組題題號仍是可編輯文字，題圖會改為正文寬度置於題號下方，不再塞在右欄。")
            incomplete = [q.source_no for q in selected if _effective_render_mode(q) != "整題圖像" and len([v for v in q.options.values() if (v or "").strip()]) < 4]
            if incomplete:
                st.warning("以下非「整題圖像」題目尚未有完整 A–D：" + "、".join(map(str, incomplete)) + "。可補齊文字，或到①題目結構校對改成「整題圖像」。")
            student_bytes = make_editable_exam_layout_docx(
                st.session_state.questions,
                int(st.session_state.year),
                suffix,
                teacher=False,
                template_kind=template_kind,
                booklet_no=booklet_no
            )
            teacher_bytes = make_editable_exam_layout_docx(
                st.session_state.questions,
                int(st.session_state.year),
                suffix+"(詳解_教學法)",
                teacher=True,
                template_kind=template_kind,
                booklet_no=booklet_no
            )

        elif output_mode == "原 PDF 圖像版":
            st.info("此模式以原 PDF 題目區塊作為圖檔放入 Word，因此視覺最接近原會考題本；圖內文字本身不可直接編輯。詳解與教學文字仍是可編輯 Word 文字。")
            show_new_number = st.checkbox(
                "在每題上方另外顯示新的組題序號（1、2、3…）",
                value=False,
                help="原裁圖可能仍含原始題號。若希望完全維持原卷外觀，建議不要勾選；若重組後需要新序號，可勾選。"
            )
            show_source_meta = st.checkbox(
                "教師版顯示原題號／通過率來源資訊",
                value=False
            )

            student_bytes = make_exam_layout_docx(
                st.session_state.questions,
                int(st.session_state.year),
                suffix,
                teacher=False,
                show_new_number=show_new_number,
                show_source_meta=False
            )
            teacher_bytes = make_exam_layout_docx(
                st.session_state.questions,
                int(st.session_state.year),
                suffix+"(詳解_教學法)",
                teacher=True,
                show_new_number=show_new_number,
                show_source_meta=show_source_meta
            )
        else:
            preserve_visual = st.checkbox("圖片／圖表／複雜題保留原 PDF 裁圖", value=True)
            st.caption("一般題目以可編輯文字輸出；圖片、圖表與複雜版面可保留原 PDF 裁圖。")
            student_bytes = make_docx(
                st.session_state.questions,
                int(st.session_state.year),
                suffix,
                teacher=False,
                preserve_visual=preserve_visual
            )
            teacher_bytes = make_docx(
                st.session_state.questions,
                int(st.session_state.year),
                suffix+"(詳解_教學法)",
                teacher=True,
                preserve_visual=preserve_visual
            )

        if export_spec.strip():
            for q in st.session_state.questions:
                q.selected = original_selected_states.get(q.source_no, q.selected)

        c1,c2 = st.columns(2)
        with c1:
            st.download_button(
                "⬇️ 下載學生題本 Word",
                data=student_bytes,
                file_name=f"{st.session_state.year}年會考國文_{suffix}_題本{booklet_no.strip() or '自動'}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True
            )
        with c2:
            st.download_button(
                "⬇️ 下載詳解／教學法 Word",
                data=teacher_bytes,
                file_name=f"{st.session_state.year}年會考國文_{suffix}_題本{booklet_no.strip() or '自動'}(詳解_教學法).docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

        st.divider()
        st.markdown("#### 儲存工作進度")
        project = []
        for q in st.session_state.questions:
            d = asdict(q)
            d["crop_png"] = ""
            d["image_pngs"] = []
            d["body_crop_png"] = ""
            project.append(d)
        project_json = json.dumps({
            "version": APP_VERSION,
            "year": st.session_state.year,
            "questions": project
        }, ensure_ascii=False, indent=2)
        st.download_button(
            "下載題庫/詳解 JSON",
            data=project_json.encode("utf-8"),
            file_name=f"{st.session_state.year}_國文題庫.json",
            mime="application/json"
        )
