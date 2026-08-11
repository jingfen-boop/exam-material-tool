
import io
from pathlib import Path
import os
import re
import difflib
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
from pptx import Presentation

APP_VERSION = "Web v4.4 同行選項自動拆分＋考題總覽暨校對"

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
    synthesis_notes: str = ""
    teaching_focus: str = ""
    teaching: str = ""
    note_strategy: str = ""
    workbench_reviewed: bool = False
    group_id: str = ""
    group_intro: str = ""
    group_crop_pngs: list = None
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
        if self.group_crop_pngs is None:
            self.group_crop_pngs = []

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

def _option_markers(text: str):
    """Return A-D option-marker matches in visual text order.

    Supports:
      (A)  （A）  (Ａ)  （Ａ）
    including multiple options printed on the SAME PDF line.
    """
    pat = re.compile(r"[\(（]\s*([ABCDＡＢＣＤ])\s*[\)）]")
    return list(pat.finditer(text or ""))


def _normalize_option_letter(letter: str) -> str:
    table = {
        "Ａ": "A", "Ｂ": "B", "Ｃ": "C", "Ｄ": "D",
        "A": "A", "B": "B", "C": "C", "D": "D",
    }
    return table.get(letter, letter)


def _split_inline_option_line(text: str):
    """Split an option line into [(letter, value), ...].

    Example:
      "(A)①   (B)②   (C)③   (D)④"
    becomes:
      [("A","①"), ("B","②"), ("C","③"), ("D","④")]

    Returns (prefix, pairs). `prefix` is any text before the first marker.
    """
    text = (text or "").strip()
    marks = _option_markers(text)
    if not marks:
        return text, []

    prefix = text[:marks[0].start()].strip()
    pairs = []

    for i, m in enumerate(marks):
        letter = _normalize_option_letter(m.group(1))
        end = marks[i + 1].start() if i + 1 < len(marks) else len(text)
        value = text[m.end():end].strip()
        pairs.append((letter, value))

    return prefix, pairs


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

    for _, _, raw_t in rows:
        t = (raw_t or "").strip()

        # Remove official source question number once.
        if not removed_q_prefix:
            mq = re.match(rf"^{qno}\.\s*(.*)$", t)
            if mq:
                t = mq.group(1).strip()
                removed_q_prefix = True
                if not t:
                    continue

        # v4.4: handle one OR MANY A-D option markers on the same PDF text line.
        prefix, inline_pairs = _split_inline_option_line(t)
        if inline_pairs:
            # Any text before the first option marker belongs to the stem unless
            # we were already inside a previous option continuation.
            if prefix:
                if current_opt:
                    options[current_opt] = (
                        options.get(current_opt, "") + " " + prefix
                    ).strip()
                else:
                    stem_lines.append(prefix)

            for letter, value in inline_pairs:
                current_opt = letter
                options[letter] = value.strip()
            continue

        # No option marker on this line: it is either an option continuation
        # or ordinary stem text.
        if current_opt:
            if t.startswith("【") or t.startswith("《") or t.startswith(""):
                current_opt = None
                stem_lines.append(t)
            else:
                options[current_opt] = (
                    options.get(current_opt, "") + " " + t
                ).strip()
        else:
            stem_lines.append(t)

    # Always return all four keys when any option was found. This makes
    # structure review clearer and avoids "A contains everything, B-D blank".
    if options:
        options = {
            key: _clean_option_text(options.get(key, ""))
            for key in ["A", "B", "C", "D"]
        }

    return "\n".join(stem_lines).strip(), options


def _clean_group_intro_text(text: str) -> str:
    """Clean page furniture while preserving the actual common passage."""
    lines = []
    for raw in (text or "").splitlines():
        s = raw.strip()
        if not s:
            continue
        if re.fullmatch(r"\d+", s):
            continue
        if any(x in s for x in (
            "請翻頁繼續作答", "尚有試題", "第頁，共頁",
            "國中教育會考", "國文科"
        )):
            continue
        lines.append(s)
    return "\n".join(lines).strip()


def _group_headings_with_positions(doc, expected_count=42):
    """Find '請閱讀…並回答X～Y題' headings with page/y coordinates."""
    found = []
    pat = re.compile(r"回答\s*(\d{1,2})\s*[～~\-－至]\s*(\d{1,2})\s*題")
    for pi, page in enumerate(doc, start=1):
        for y, x, txt, bbox in _all_page_lines(page):
            m = pat.search(txt)
            if not m:
                continue
            a, b = int(m.group(1)), int(m.group(2))
            if 1 <= a <= b <= expected_count:
                found.append({
                    "a": a, "b": b, "page": pi,
                    "y0": bbox[1], "y1": bbox[3], "heading": txt
                })
    return found


def _extract_group_shared_material(doc, heading, starts_by_q):
    """Extract the full common material from the group heading to the first subquestion.

    Returns:
      text: editable extracted text
      crops: original PDF image slices so tables/images/poems can also be reviewed
    """
    a = heading["a"]
    first = starts_by_q.get(a)
    if not first:
        return "", []

    start_page = heading["page"]
    end_page = first["page"]
    text_parts = []
    crops = []

    for pi in range(start_page, end_page + 1):
        page = doc[pi - 1]

        if pi == start_page:
            y0 = max(0, heading["y0"] - 4)
        else:
            y0 = 42

        if pi == end_page:
            y1 = max(y0 + 20, first["bbox"][1] - 6)
        else:
            y1 = page.rect.height - 38

        if y1 <= y0 + 10:
            continue

        # Text in visual reading order.
        segment_lines = []
        for y, x, txt, bbox in _all_page_lines(page):
            cy = (bbox[1] + bbox[3]) / 2
            if y0 <= cy <= y1:
                segment_lines.append(txt)
        segment = _clean_group_intro_text("\n".join(segment_lines))
        if segment:
            text_parts.append(segment)

        # Original visual crop: preserves tables, diagrams, unusual typography.
        try:
            rect = fitz.Rect(0, y0, page.rect.width, y1)
            pix = page.get_pixmap(
                matrix=fitz.Matrix(1.35, 1.35),
                clip=rect,
                alpha=False
            )
            crops.append(pix.tobytes("png"))
        except Exception:
            pass

    merged = "\n".join(text_parts).strip()
    return merged, crops


def extract_questions(question_pdf: bytes, expected_count: int = 42) -> Tuple[List[Question], Dict[int, bytes]]:
    """Web v1.3 question parser."""
    doc = fitz.open(stream=question_pdf, filetype="pdf")
    page_images = {}
    for pi, page in enumerate(doc, start=1):
        pix = page.get_pixmap(matrix=fitz.Matrix(1.25,1.25), alpha=False)
        page_images[pi] = pix.tobytes("png")

    starts = _sequential_question_starts(doc, expected_count)
    starts_by_q = {s["qno"]: s for s in starts}

    # Detect each actual reading group and extract its common material BEFORE
    # splitting individual child questions.
    group_headings = _group_headings_with_positions(doc, expected_count)
    groups = []
    group_material = {}
    for h in group_headings:
        a, b = h["a"], h["b"]
        groups.append((a, b))
        intro, intro_crops = _extract_group_shared_material(doc, h, starts_by_q)
        group_material[(a, b)] = {
            "intro": intro,
            "crops": intro_crops
        }

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
                gm = group_material.get((a, b), {})
                q.group_intro = gm.get("intro", "")
                q.group_crop_pngs = list(gm.get("crops", []))
                # material is also populated so existing editing/output paths
                # can use the shared passage without requiring re-entry.
                if q.group_intro:
                    q.material = q.group_intro
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
        r = p.add_run(_clean_word_text(category))
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
    r = p.add_run(f"（{' ' + _clean_word_text(q.answer) + ' ' if teacher and q.answer else '   '}）{display_no}. ")
    set_eastasia(r)
    if teacher and q.answer:
        r.font.color.rgb = RGBColor(255,0,0)

    if use_crop and q.crop_png:
        p2 = doc.add_paragraph()
        p2.add_run().add_picture(io.BytesIO(q.crop_png), width=Cm(16.6))
    else:
        if q.material.strip():
            pmaterial = doc.add_paragraph()
            rm = pmaterial.add_run(_clean_word_text(q.material))
            set_eastasia(rm)
        r = p.add_run(_clean_word_text(q.text))
        set_eastasia(r)
        for key in ["A","B","C","D"]:
            if key in q.options:
                po = doc.add_paragraph()
                ro = po.add_run(f"({key}){_clean_word_text(q.options[key])}")
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
        r2 = p.add_run(_clean_word_text(q.explanation or "（待補）"))
        set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)

        if q.teaching_focus.strip():
            p = cell.add_paragraph()
            r = p.add_run("【教學重點】：\n")
            set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
            r2 = p.add_run(_clean_word_text(q.teaching_focus))
            set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)

        p = cell.add_paragraph()
        r = p.add_run("【教學步驟】：\n")
        set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
        r2 = p.add_run(_clean_word_text(q.teaching or "（待補）"))
        set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)

        if q.note_strategy.strip():
            p = cell.add_paragraph()
            r = p.add_run("【筆記策略】：\n")
            set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
            r2 = p.add_run(_clean_word_text(q.note_strategy))
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
        r = p.add_run(f"答案：{_clean_word_text(q.answer or '—')}")
        set_eastasia(r)
        r.bold = True
        r.font.color.rgb = RGBColor(255,0,0)

        box = doc.add_table(rows=1, cols=1)
        box.style = "Table Grid"
        cell = box.cell(0,0)
        p = cell.paragraphs[0]
        r = p.add_run("解析：\n")
        set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
        r2 = p.add_run(_clean_word_text(q.explanation or "（待補）"))
        set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)

        if q.teaching_focus.strip():
            p = cell.add_paragraph()
            r = p.add_run("【教學重點】：\n")
            set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
            r2 = p.add_run(_clean_word_text(q.teaching_focus))
            set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)

        p = cell.add_paragraph()
        r = p.add_run("【教學步驟】：\n")
        set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
        r2 = p.add_run(_clean_word_text(q.teaching or "（待補）"))
        set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)

        if q.note_strategy.strip():
            p = cell.add_paragraph()
            r = p.add_run("【筆記策略】：\n")
            set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
            r2 = p.add_run(_clean_word_text(q.note_strategy))
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
            r = p.add_run(f"({key}) {_clean_word_text(q.options.get(key, ''))}")
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
            r = p.add_run(f"({key}) {_clean_word_text(q.options.get(key, ''))}")
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
        r2 = p.add_run(_clean_word_text(q.explanation or "（待補）"))
        set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)

        if q.teaching_focus.strip():
            p = doc.add_paragraph()
            r = p.add_run("【教學重點】：\n")
            set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
            r2 = p.add_run(_clean_word_text(q.teaching_focus))
            set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)

        p = doc.add_paragraph()
        r = p.add_run("【教學步驟】：\n")
        set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
        r2 = p.add_run(_clean_word_text(q.teaching or "（待補）"))
        set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)



def _clean_word_text(value) -> str:
    """Return text safe for python-docx / XML 1.0.

    PDF/PPT/DOCX extraction may contain invisible control characters such as
    form-feed, vertical-tab, NUL, surrogate code points or Unicode noncharacters.
    Streamlit can display them, but lxml cannot write them into a .docx XML node.
    """
    if value is None:
        return ""

    s = str(value).replace("\u00a0", " ").replace("\u3000", " ")

    cleaned = []
    for ch in s:
        cp = ord(ch)

        # XML 1.0 valid characters:
        # TAB, LF, CR, U+0020–D7FF, U+E000–FFFD, U+10000–10FFFF.
        valid = (
            cp in (0x09, 0x0A, 0x0D)
            or 0x20 <= cp <= 0xD7FF
            or 0xE000 <= cp <= 0xFFFD
            or 0x10000 <= cp <= 0x10FFFF
        )

        # Explicitly exclude Unicode noncharacters that can still cause trouble
        # in XML consumers.
        noncharacter = (
            0xFDD0 <= cp <= 0xFDEF
            or (cp & 0xFFFF) in (0xFFFE, 0xFFFF)
        )

        if valid and not noncharacter:
            cleaned.append(ch)
        elif ch in ("\x0b", "\x0c"):
            # Preserve intended visual separation rather than silently gluing text.
            cleaned.append("\n")
        else:
            # Drop illegal/invisible controls.
            pass

    s = "".join(cleaned)
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
        r = p.add_run(material_text)
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
        r = p.add_run(f"（{' ' + _clean_word_text(q.answer) + ' ' if teacher and q.answer else '   '}）{display_no}. {stem_text}")
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
        r = p.add_run(f"（{' ' + _clean_word_text(q.answer) + ' ' if teacher and q.answer else '   '}）{display_no}. {stem_text}")
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
        r2 = p.add_run(_clean_word_text(q.explanation or "（待補）"))
        set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)

        if q.teaching_focus.strip():
            p = doc.add_paragraph()
            r = p.add_run("【教學重點】：\n")
            set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
            r2 = p.add_run(_clean_word_text(q.teaching_focus))
            set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)

        p = doc.add_paragraph()
        r = p.add_run("【教學步驟】：\n")
        set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
        r2 = p.add_run(_clean_word_text(q.teaching or "（待補）"))
        set_eastasia(r2); r2.font.color.rgb = RGBColor(255,0,0)

        if q.note_strategy.strip():
            p = doc.add_paragraph()
            r = p.add_run("【筆記策略】：\n")
            set_eastasia(r); r.bold = True; r.font.color.rgb = RGBColor(255,0,0)
            r2 = p.add_run(_clean_word_text(q.note_strategy))
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
# Annual reference package / explanation workbench
# -----------------------------
ABILITY_OPTIONS = ["", "字詞辨識", "表層文意理解", "文意統整", "推論理解", "分析評鑑", "其他"]

DEFAULT_STRATEGY_LIBRARY = {
    "字詞辨識": {
        "教學重點": "學生能辨識字音、字形、詞義、標點或語文知識，並能依語境正確判斷與運用。",
        "教學步驟": (
            "1. 引導學生讀題，圈出題幹中的作答關鍵詞，確認本題要判斷的是字音、字形、詞義、標點或語文知識。\n"
            "2. 教師帶領學生逐項找出需要判斷的字詞／語文知識點，先說明其基本定義或用法，再回到完整句子中判讀。\n"
            "3. 逐項比對選項：請學生說明每一選項正確或錯誤的具體原因；若為一字多義，需將字詞代回語境判斷，不只背單一字義。\n"
            "4. 將容易混淆的字詞、讀音或概念進行對照整理，必要時搭配造句、同義／反義比較或錯字訂正。\n"
            "5. 請學生口頭說明正確答案及排除其他選項的依據，最後整理可遷移到相似題型的判斷原則。"
        ),
        "筆記策略": "依本題建立「字詞／定義或讀音／本題語境／易混淆點／例句」對照表。"
    },
    "表層文意理解": {
        "教學重點": "學生能從文本中定位明確訊息，劃記關鍵字句，並以原文證據檢核各選項。",
        "教學步驟": (
            "1. 引導學生讀題，圈出「最符合、最恰當、依本文」等限制詞，確認題目要求。\n"
            "2. 依題目中的人物、事件、時間、因果或關鍵詞回到文本定位，劃記直接相關的句子。\n"
            "3. 將題幹與文本切分成可比對的資訊單位，逐項對照 A～D：標記為「原文支持、原文相反、原文未提及」。\n"
            "4. 排除與文本不符、偷換概念、擴大或縮小範圍的選項；要求學生指出原文證據而非只說『感覺比較像』。\n"
            "5. 讓學生個別或分組說明答案與證據，若出現不同答案，再回到原文比對關鍵句。"
        ),
        "筆記策略": "使用「選項／文本證據／判斷理由」三欄表。"
    },
    "文意統整": {
        "教學重點": "學生能整合段落或多則資料的重點，找出反覆或互相呼應的概念，形成整體理解。",
        "教學步驟": (
            "1. 先確認題目要找的是主旨、核心概念、共同點、關係或整體觀點，而非單一細節。\n"
            "2. 依標點、段落或資料一／資料二切分文本，為每一部分寫下一句重點。\n"
            "3. 圈畫不同段落中反覆出現、彼此呼應或具有上下位關係的關鍵詞句。\n"
            "4. 將各部分重點往上統整成一個核心概念，再檢查這個概念能否同時涵蓋全文／各資料。\n"
            "5. 逐項比對選項，排除只涵蓋局部資訊、把例子當主旨、過度延伸或與整體文意不符者。\n"
            "6. 請學生用自己的話說出全文核心，再回頭驗證正確選項。"
        ),
        "筆記策略": "採「分段重點 → 共同關鍵詞 → 核心概念 → 正確選項」四層筆記。"
    },
    "推論理解": {
        "教學重點": "學生能根據文本證據建立合理推論鏈，並區辨合理推論與過度推論。",
        "教學步驟": (
            "1. 確認題目要求推論的對象與限制條件，圈出『推論、最可能、可知』等關鍵詞。\n"
            "2. 找出與每一選項相關的文本證據，先整理『文本已知』，再進一步思考『因此可以推得什麼』。\n"
            "3. 以『證據 → 中間判斷 → 結論』建立推論鏈，要求每一步都能回到文本或資料支持。\n"
            "4. 逐項檢查是否有文中未提及、因果顛倒、範圍擴大、把可能說成必然或加入外部常識等過度推論。\n"
            "5. 請學生說明『我是從哪一句推到這個答案』，比較不同選項的推論距離與合理性。"
        ),
        "筆記策略": "使用「文本證據 → 推論過程 → 選項判斷」箭頭筆記。"
    },
    "分析評鑑": {
        "教學重點": "學生能比較多文本、人物或觀點，分析論據與立場，並依證據進行比較與評判。",
        "教學步驟": (
            "1. 確認題目要比較的對象、觀點、主張或判準，避免一開始就混讀多則資料。\n"
            "2. 分別整理各文本／人物的主要主張、理由與關鍵證據；先各自讀懂，再進行比較。\n"
            "3. 建立比較表，標示相同點、相異點、因果關係或立場差異。\n"
            "4. 逐項檢視選項是否準確反映資料關係，排除張冠李戴、偷換概念、只符合單一文本或證據不足者。\n"
            "5. 請學生引用具體文本證據說明評判理由；若有不同答案，可用分組討論或辯證方式再次驗證。"
        ),
        "筆記策略": "使用「文本A／文本B／共同點／差異點／證據」比較表。"
    },
    "其他": {
        "教學重點": "學生能辨識本題的核心作答任務，依題型選用適切策略，並以明確證據完成判斷。",
        "教學步驟": (
            "1. 明確圈出題目要求與限制條件。\n"
            "2. 找出完成作答所需的關鍵資訊或知識點。\n"
            "3. 依題型進行逐項比對、排除、分類或轉換。\n"
            "4. 要求學生說明答案與依據，而非只報答案。\n"
            "5. 整理本題可遷移到相似題目的解題原則。"
        ),
        "筆記策略": "依題型建立「關鍵資訊／判斷依據／結論」簡表。"
    }
}

def _empty_reference_db(year=None):
    return {
        "format_version": "3.1",
        "year": int(year) if year is not None else None,
        "publisher": {"翰林": {}, "康軒": {}, "南一": {}},
        "history_raw": {},
        "strategy": DEFAULT_STRATEGY_LIBRARY,
        "drafts": {}
    }

def _load_bundled_reference_library():
    path = Path(__file__).resolve().parent / "reference_library_v30.json"
    if not path.exists():
        return _empty_reference_db(115)
    try:
        db = json.loads(path.read_text(encoding="utf-8"))
        db.setdefault("format_version", "3.1")
        db.setdefault("year", 115)
        db.setdefault("publisher", {"翰林": {}, "康軒": {}, "南一": {}})
        db.setdefault("history_raw", {})
        db.setdefault("strategy", DEFAULT_STRATEGY_LIBRARY)
        db.setdefault("drafts", {})
        return db
    except Exception:
        return _empty_reference_db(115)

def _load_reference_library():
    if "reference_db" in st.session_state and isinstance(st.session_state.reference_db, dict):
        return st.session_state.reference_db
    db = _load_bundled_reference_library()
    st.session_state.reference_db = db
    return db

def _uploaded_file_text(uploaded):
    """Parse annual reference files in Streamlit Cloud.
    Supported directly: DOCX, PPTX, PDF, TXT.
    Legacy .doc is intentionally not silently converted because cloud conversion is unreliable.
    """
    name = uploaded.name
    ext = Path(name).suffix.lower()
    data = uploaded.getvalue()

    if ext == ".docx":
        doc = Document(io.BytesIO(data))
        parts = []
        for p in doc.paragraphs:
            if p.text.strip():
                parts.append(p.text.strip())
        for table in doc.tables:
            for row in table.rows:
                cells = [c.text.strip() for c in row.cells]
                if any(cells):
                    parts.append(" | ".join(cells))
        return "\n".join(parts)

    if ext == ".pptx":
        prs = Presentation(io.BytesIO(data))
        parts = []
        for i, slide in enumerate(prs.slides, 1):
            slide_text = []
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text and shape.text.strip():
                    slide_text.append(shape.text.strip())
            if slide_text:
                parts.append(f"[投影片{i}]\n" + "\n".join(slide_text))
        return "\n".join(parts)

    if ext == ".pdf":
        return pdf_text(data)

    if ext == ".txt":
        return data.decode("utf-8", errors="ignore")

    if ext == ".doc":
        raise ValueError("舊式 .doc 無法在 Streamlit Cloud 穩定解析，請先在 Word 另存為 .docx 再上傳。")

    raise ValueError(f"目前不支援 {ext}。")

def _normalize_reference_text(s: str) -> str:
    s = (s or "").replace("\x0b", "\n").replace("\r", "\n")
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()

def _group_range_from_slide(ch: str, expected_count=None):
    """Return an announced question range such as 回答24～25題 / 題組(24～42題)."""
    max_q = expected_count or 99
    # Only an explicit "回答X～Y題" passage heading drives sequential inference.
    # Broad section labels such as "題組(24～42題)" or "單題(1～23題)" are NOT
    # treated as one continuous group, otherwise they can mis-number later slides.
    pats = [
        re.compile(r"回答\s*[（(]?\s*(\d{1,2})\s*[～~\-－至]\s*(\d{1,2})\s*題"),
    ]
    for pat in pats:
        m = pat.search(ch)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            if 1 <= a <= b <= max_q:
                return a, b
    return None


def _looks_like_question_slide(ch: str):
    """Question slides usually contain an answer mark and/or A-D options."""
    has_options = sum(token in ch for token in ("(A)", "(B)", "(C)", "(D)", "（A）", "（B）", "（C）", "（D）")) >= 2
    # Publisher PPTs often place the answer as a standalone A/B/C/D line.
    has_answer = bool(re.search(r"(?m)^\s*[A-DＡ-Ｄ]\s*$", ch))
    return has_options or has_answer


def _split_slides_by_question(text: str, expected_count=None):
    """PPT-aware splitter with generic inference for group-question numbering.

    Handles three recurring publisher-export cases:
    1) a group passage announces "回答24～25題", but the first question slide omits "24.";
    2) middle/last slides in a group omit their question number;
    3) a publisher accidentally repeats the previous number (e.g. 41 twice for 41～42).
    The inference is based on the announced range and slide order, not hard-coded question numbers.
    """
    chunks = [c for c in re.split(r"(?=\[投影片\d+\])", text) if c.strip()]
    out = {}
    current = None
    pending_shared = []
    active_range = None
    next_group_q = None

    for ch in chunks:
        normalized = _normalize_reference_text(ch)

        announced = _group_range_from_slide(normalized, expected_count)
        if announced:
            active_range = announced
            next_group_q = announced[0]
            # Shared-passage slides are kept for the first question in the range.
            if len(normalized) >= 60 and not _looks_like_question_slide(normalized):
                pending_shared.append(normalized)
                continue

        nums = _slide_question_numbers(normalized, expected_count)
        explicit_q = nums[0] if nums else None
        q = explicit_q

        if _looks_like_question_slide(normalized) and active_range:
            lo, hi = active_range

            # Accept an explicit number only if it fits the active range and does not move backward.
            if q is not None and not (lo <= q <= hi):
                q = None

            # If the publisher repeats the preceding number or omits the number,
            # use the next expected number in the announced group.
            if next_group_q is not None:
                if q is None or q < next_group_q:
                    q = next_group_q
                elif q > next_group_q:
                    # Respect a valid forward jump, but the unnumbered earlier slide(s)
                    # should already have consumed the missing number(s).
                    pass

            if q is not None and lo <= q <= hi:
                next_group_q = q + 1
                if next_group_q > hi:
                    active_range = None
                    next_group_q = None

        if q is not None and _looks_like_question_slide(normalized):
            prefix = "\n\n".join(pending_shared[-3:]) if pending_shared else ""
            pending_shared = []
            current = q
            out.setdefault(str(q), "")
            merged = (prefix + "\n\n" + normalized).strip() if prefix else normalized
            if merged and merged not in out[str(q)]:
                out[str(q)] += (("\n\n" if out[str(q)] else "") + merged)
            continue

        # Continuation/detail slide for the current question.
        if current is not None and any(k in normalized for k in ("解析", "詳解", "答案", "語譯", "對應教材")):
            if normalized and normalized not in out.get(str(current), ""):
                out.setdefault(str(current), "")
                out[str(current)] += "\n\n" + normalized
        else:
            # Likely shared passage or table/image description.
            if len(normalized) >= 60:
                pending_shared.append(normalized)

    return {k: v.strip() for k, v in out.items() if v.strip()}


def _question_anchor_candidates(text: str, expected_count=None):
    """找出出版社文件中的可能題號位置。

    支援常見格式：
    24.
    24、
    (24)
    第24題
    ( B ) 24.

    這個函式只負責提供候選位置；實際題組中省略題號的情況，
    仍由 PPT 題組範圍推論機制處理。
    """
    text = _normalize_reference_text(text)
    max_q = expected_count or 99

    patterns = [
        # (B) 24. / 24. / 24、
        re.compile(
            r"(?mi)^\s*(?:[（(]\s*[A-DＡ-Ｄ]?\s*[）)]\s*)?"
            r"(\d{1,2})\s*[.．、]\s*(?=\S|\n)"
        ),
        # (24)
        re.compile(
            r"(?mi)^\s*[（(]\s*(\d{1,2})\s*[）)]\s*(?=\S|\n)"
        ),
        # 第24題
        re.compile(
            r"(?mi)^\s*第\s*(\d{1,2})\s*題\s*[:：.．、]?"
        ),
    ]

    found = []
    seen = set()

    for pat in patterns:
        for m in pat.finditer(text):
            qno = int(m.group(1))
            if not (1 <= qno <= max_q):
                continue
            pos = m.start()
            if pos in seen:
                continue
            seen.add(pos)
            found.append((qno, pos))

    return sorted(found, key=lambda x: x[1])


def _recover_missing_question_blocks(text: str, missing_numbers, expected_count=None):
    """Recover irregular publisher headings without hard-coding a year's missing numbers."""
    text = _normalize_reference_text(text)
    if not text or not missing_numbers:
        return {}
    anchors = _question_anchor_candidates(text, expected_count)
    recovered = {}
    for q in sorted(set(int(x) for x in missing_numbers)):
        candidates = [(qq, pos) for qq, pos in anchors if qq == q]
        for _, pos in candidates:
            later_positions = [p for qq, p in anchors if p > pos and qq > q]
            end = min(later_positions, default=len(text))
            block = text[pos:end].strip()
            if len(block) >= 40 and any(k in block for k in ("詳解", "解析", "對應教材", "(A)", "（A）", "答案")):
                recovered[str(q)] = block
                break
    return recovered



def _match_norm(s: str) -> str:
    """Normalize Chinese question text for cross-file matching."""
    s = (s or "").lower()
    s = re.sub(r"[\s\u3000]+", "", s)
    s = re.sub(r"[，。！？；：、,.!?;:「」『』（）()【】\[\]〈〉《》—－…．·_]+", "", s)
    return s


def _publisher_raw_chunks(raw: str):
    """Create candidate chunks while preserving PPT slide boundaries when possible."""
    raw = _normalize_reference_text(raw)
    if "[投影片" in raw:
        return [c.strip() for c in re.split(r"(?=\[投影片\d+\])", raw) if c.strip()]

    # For Word/PDF/TXT, paragraph windows are safer than depending only on question numbering.
    paras = [p.strip() for p in re.split(r"\n+", raw) if p.strip()]
    chunks = []
    for i in range(len(paras)):
        # Use a moving window so a stem and its explanation can be matched even if they
        # occupy adjacent paragraphs/table rows.
        chunks.append("\n".join(paras[i:i+6]))
    return chunks


def _content_match_score(q, chunk: str):
    """Score how strongly a publisher chunk corresponds to an official question."""
    cn = _match_norm(chunk)
    if not cn:
        return 0.0

    stem = _match_norm(getattr(q, "text", "") or "")
    material = _match_norm(getattr(q, "material", "") or "")
    options = [
        _match_norm((getattr(q, "options", {}) or {}).get(k, ""))
        for k in ("A", "B", "C", "D")
    ]

    score = 0.0

    # Exact stem fragments are the strongest signal.
    if stem:
        probes = []
        for n in (28, 22, 16, 12):
            if len(stem) >= n:
                probes.extend([stem[:n], stem[-n:]])
        if any(p and p in cn for p in probes):
            score += 1.2

        # Longest common sequence is robust to minor punctuation/export differences.
        m = difflib.SequenceMatcher(None, stem, cn).find_longest_match(
            0, len(stem), 0, len(cn)
        )
        score += min(1.0, m.size / max(12, min(len(stem), 60)))

    # Shared material helps especially for grouped reading questions.
    if material and len(material) >= 20:
        probe = material[:min(32, len(material))]
        if probe in cn:
            score += 0.35

    # Options are useful when the question stem is short or numbering is omitted.
    opt_hits = 0
    for opt in options:
        if len(opt) >= 6:
            probe = opt[:min(18, len(opt))]
            if probe in cn:
                opt_hits += 1
    score += min(0.8, opt_hits * 0.25)

    # A chunk that contains explanation vocabulary is preferred as a reference block.
    if any(k in chunk for k in ("詳解", "解析", "對應教材", "答案")):
        score += 0.12

    return score


def _recover_by_question_bank(raw_sources, missing_numbers, question_bank):
    """Recover missing publisher questions by matching their actual content.

    This is the key v4.0 change: after the official question bank exists, we no longer
    require the publisher to print a reliable question number. A missing/duplicated
    number in DOCX/PPTX can still be aligned by the stem/options/material.
    """
    if not question_bank or not missing_numbers:
        return {}

    qmap = {int(q.source_no): q for q in question_bank}
    all_candidates = []

    for source_name, raw in raw_sources:
        chunks = _publisher_raw_chunks(raw)
        for idx, ch in enumerate(chunks):
            all_candidates.append((source_name, idx, ch, chunks))

    recovered = {}
    used = set()

    for qno in sorted(set(int(x) for x in missing_numbers)):
        q = qmap.get(qno)
        if q is None:
            continue

        ranked = []
        for source_name, idx, ch, chunks in all_candidates:
            score = _content_match_score(q, ch)
            ranked.append((score, source_name, idx, ch, chunks))

        ranked.sort(key=lambda x: x[0], reverse=True)
        if not ranked:
            continue

        score, source_name, idx, ch, chunks = ranked[0]

        # Conservative threshold: do not fabricate a match just to make the count 42.
        if score < 0.70:
            continue

        block_parts = [ch]

        # In PPT exports the explanation can be on the immediately following slide.
        # Append at most two continuation slides only when they look like explanation pages.
        for j in range(idx + 1, min(idx + 3, len(chunks))):
            nxt = chunks[j]
            if any(k in nxt for k in ("詳解", "解析", "對應教材", "語譯")):
                block_parts.append(nxt)
            else:
                break

        recovered[str(qno)] = "\n\n".join(block_parts).strip()

    return recovered


def _parse_publisher_files(files, expected_count=None, question_bank=None):
    """Parse and merge publisher files.

    v4.0 uses three layers:
    1. normal question-number parsing;
    2. irregular-number recovery;
    3. official-question-bank CONTENT matching for omitted/duplicated numbers.
    """
    combined, errors, raw_sources = {}, [], []

    for uploaded in files or []:
        try:
            raw = _uploaded_file_text(uploaded)
            raw_sources.append((uploaded.name, raw))

            if "[投影片" in raw:
                parsed = _split_slides_by_question(raw, expected_count)
            else:
                # Some older code paths may not have a reliable Word-number splitter.
                # Number-anchor recovery below plus content matching is intentionally
                # sufficient to complement the PPT/other files.
                try:
                    parsed = _split_text_by_question(raw, expected_count)
                except Exception:
                    parsed = {}

            for q, block in parsed.items():
                block = (block or "").strip()
                if not block:
                    continue
                if q in combined and block not in combined[q]:
                    combined[q] += "\n\n" + block
                elif q not in combined:
                    combined[q] = block

        except Exception as e:
            errors.append(f"{uploaded.name}：{e}")

    # Layer 2: generic question-number-anchor recovery.
    if expected_count:
        missing = [q for q in range(1, expected_count + 1) if str(q) not in combined]
        for _, raw in raw_sources:
            if not missing:
                break
            recovered = _recover_missing_question_blocks(raw, missing, expected_count)
            for q, block in recovered.items():
                combined.setdefault(q, block.strip())
            missing = [q for q in range(1, expected_count + 1) if str(q) not in combined]

        # Layer 3: match the actual official question content.
        if missing and question_bank:
            recovered = _recover_by_question_bank(
                raw_sources,
                missing,
                question_bank
            )
            for q, block in recovered.items():
                combined.setdefault(q, block.strip())

    return combined, errors

def _publisher_analysis_only(block: str) -> str:
    block = (block or "").strip()
    for marker in ("試題解析：", "詳解：", "詳解 ", "解析："):
        if marker in block:
            tail = block.split(marker, 1)[1].strip()
            if tail:
                return tail
    return block

def _history_examples_for_category(refdb, category: str, limit=6):
    if not category:
        return []
    out = []
    for source, raw in refdb.get("history_raw", {}).items():
        keys = [category]
        if category == "表層文意理解":
            keys.append("表層文意")
        positions = [raw.find(k) for k in keys if raw.find(k) >= 0]
        if not positions:
            continue
        pos = min(positions)
        teach = raw.find("【教學步驟】", pos)
        if teach >= 0 and teach - pos < 5000:
            start = max(0, pos - 450)
            end = min(len(raw), teach + 2200)
        else:
            start = max(0, pos - 450)
            end = min(len(raw), pos + 2400)
        out.append((source, raw[start:end].strip()))
        if len(out) >= limit:
            break
    return out

def _strategy_for_category(refdb, category: str):
    return refdb.get("strategy", {}).get(category) or DEFAULT_STRATEGY_LIBRARY.get(category) or DEFAULT_STRATEGY_LIBRARY["其他"]

def _annual_package_json(refdb):
    return json.dumps(refdb, ensure_ascii=False, indent=2).encode("utf-8")

def _drafts_from_questions(questions, year):
    drafts = {}
    for q in questions:
        drafts[str(q.source_no)] = {
            "category": q.category,
            "synthesis_notes": q.synthesis_notes,
            "explanation": q.explanation,
            "teaching_focus": q.teaching_focus,
            "teaching": q.teaching,
            "note_strategy": q.note_strategy,
            "reviewed": q.workbench_reviewed,
        }
    return {
        "format_version": "3.1-drafts",
        "year": int(year),
        "drafts": drafts
    }

def _apply_drafts_to_questions(payload, questions):
    drafts = payload.get("drafts", {}) if isinstance(payload, dict) else {}
    qmap = {str(q.source_no): q for q in questions}
    applied = 0
    for no, d in drafts.items():
        q = qmap.get(str(no))
        if not q or not isinstance(d, dict):
            continue
        q.category = d.get("category", q.category)
        q.synthesis_notes = d.get("synthesis_notes", q.synthesis_notes)
        q.explanation = d.get("explanation", q.explanation)
        q.teaching_focus = d.get("teaching_focus", q.teaching_focus)
        q.teaching = d.get("teaching", q.teaching)
        q.note_strategy = d.get("note_strategy", q.note_strategy)
        q.workbench_reviewed = bool(d.get("reviewed", q.workbench_reviewed))
        applied += 1
    return applied


def _json_safe(value):
    """Recursively convert app state to JSON-safe values.

    Bytes are intentionally omitted as empty strings because source-PDF crops are
    reproducible artifacts, not canonical project data.
    """
    if isinstance(value, (bytes, bytearray, memoryview)):
        return ""
    if isinstance(value, dict):
        return {str(k): _json_safe(v) for k, v in value.items()}
    if isinstance(value, (list, tuple, set)):
        return [_json_safe(v) for v in value]
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    return str(value)



def _render_question_review_editor(q, key_prefix="overview"):
    """Inline structure-review editor used by the consolidated overview page."""
    no = q.source_no

    c1, c2 = st.columns([1.0, 1.05], vertical_alignment="top")

    with c1:
        st.markdown(f"**原 PDF｜第 {no} 題**")
        st.caption(
            f"原頁：{q.page_no}｜答案：{q.answer or '—'}｜"
            f"通過率：{q.pass_rate if q.pass_rate is not None else '—'}"
        )
        if q.crop_png:
            st.image(q.crop_png, caption="原 PDF 題目區塊", use_container_width=True)
        else:
            st.info("本題目前沒有原 PDF 裁圖。")

    with c2:
        material_edit = st.text_area(
            "閱讀／共用材料（沒有可留白）",
            value=q.material,
            height=135,
            key=f"{key_prefix}_material_{no}"
        )
        stem_edit = st.text_area(
            "題幹",
            value=q.text,
            height=120,
            key=f"{key_prefix}_stem_{no}"
        )

        o1, o2 = st.columns(2)
        with o1:
            oa = st.text_area(
                "A", value=q.options.get("A", ""), height=76,
                key=f"{key_prefix}_A_{no}"
            )
            oc = st.text_area(
                "C", value=q.options.get("C", ""), height=76,
                key=f"{key_prefix}_C_{no}"
            )
        with o2:
            ob = st.text_area(
                "B", value=q.options.get("B", ""), height=76,
                key=f"{key_prefix}_B_{no}"
            )
            od = st.text_area(
                "D", value=q.options.get("D", ""), height=76,
                key=f"{key_prefix}_D_{no}"
            )

        group_edit = st.text_input(
            "題組 ID（例如 21-23；非題組留白）",
            value=q.group_id,
            key=f"{key_prefix}_group_{no}"
        )

        rc1, rc2 = st.columns(2)
        with rc1:
            render_choices = ["自動", "可編輯文字", "圖文混合", "整題圖像"]
            current_render = q.render_mode if q.render_mode in render_choices else "自動"
            render_edit = st.selectbox(
                "本題輸出模式",
                render_choices,
                index=render_choices.index(current_render),
                key=f"{key_prefix}_render_{no}",
                help="整題圖像：題號與答案括弧維持可編輯，題目本體使用原 PDF 圖片。"
            )

            layout_choices = ["一般直列", "圖片在右", "圖片在上", "選項兩欄"]
            current_layout = q.layout_style if q.layout_style in layout_choices else "一般直列"
            layout_edit = st.selectbox(
                "可編輯 Word 版型",
                layout_choices,
                index=layout_choices.index(current_layout),
                key=f"{key_prefix}_layout_{no}"
            )

        with rc2:
            include_image_edit = st.checkbox(
                "可編輯版輸出獨立圖片",
                value=q.include_image,
                key=f"{key_prefix}_include_image_{no}"
            )
            visual_edit = st.checkbox(
                "保留原 PDF 裁圖作為備用／原版型輸出",
                value=q.visual_mode,
                key=f"{key_prefix}_visual_{no}"
            )
            reviewed_edit = st.checkbox(
                "本題內容已人工確認",
                value=q.reviewed,
                key=f"{key_prefix}_reviewed_{no}"
            )

        if _effective_render_mode(q) == "整題圖像":
            st.info("本題目前採「整題圖像」，A～D 可不必另外重建；Word 會保留可編輯題號與答案括弧。")

        if st.button(
            "💾 儲存本題校對",
            type="primary",
            key=f"{key_prefix}_save_{no}",
            use_container_width=True
        ):
            _apply_structure_edit(
                q, material_edit, stem_edit, oa, ob, oc, od,
                group_edit, visual_edit, reviewed_edit
            )
            q.render_mode = render_edit
            q.layout_style = layout_edit
            q.include_image = include_image_edit
            st.success(f"第 {no} 題已儲存並更新總覽。")
            st.rerun()

# -----------------------------
# Streamlit UI
# -----------------------------
st.set_page_config(page_title="會考教材產製工具", page_icon="📘", layout="wide")
st.title("📘 會考教材產製工具")
st.caption(f"{APP_VERSION}｜建立題庫 → 年度資料 → 考題總覽暨校對 → 篩選組題 → 詳解／教學 → Word")

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

tab1, ref_tab, overview_tab, tab2, tab3, tab4 = st.tabs(["① 建立題庫", "② 年度資料", "③ 考題總覽暨校對", "④ 篩選組題", "⑤ 詳解工作台", "⑥ 產生 Word"])


with ref_tab:
    st.subheader(f"{int(st.session_state.year)} 年度資料與詳解參考庫")
    st.caption(
        "建議先完成「① 建立題庫」，再建立出版社／內部詳解參考庫。這樣 B 區才能立即核對每家出版社是否完整對應本年度所有題目。"
    )

    refdb = _load_reference_library()
    ref_year = refdb.get("year")

    if ref_year == int(st.session_state.year):
        st.success(f"目前載入的參考庫年度：{ref_year}")
    else:
        st.warning(
            f"目前參考庫年度為 {ref_year or '未設定'}，與現在設定的 {int(st.session_state.year)} 年不同。"
            "若要製作本年度教材，請先完成 A 區建立／更新本年度參考庫，或到 D 區載入既有年度參考包。"
        )

    # --------------------------------------------------
    # A. Build
    # --------------------------------------------------
    st.markdown("## A. 建立／更新本年度參考庫")
    st.caption(
        "第一次處理某一年度時，先從這裡開始。上傳三家出版社詳解，以及本團隊歷年教師版參考檔。"
    )

    st.info(
        "建議格式：DOCX、PPTX、PDF、TXT。舊式 .doc 在 Streamlit Cloud 不穩定，"
        "請先用 Word 另存成 .docx 再上傳。"
    )

    pc1, pc2, pc3 = st.columns(3)
    with pc1:
        hanlin_files = st.file_uploader(
            "翰林詳解",
            type=["docx", "pptx", "pdf", "txt", "doc"],
            accept_multiple_files=True,
            key="annual_hanlin"
        )
    with pc2:
        kang_files = st.file_uploader(
            "康軒詳解",
            type=["docx", "pptx", "pdf", "txt", "doc"],
            accept_multiple_files=True,
            key="annual_kang"
        )
    with pc3:
        nanyi_files = st.file_uploader(
            "南一詳解",
            type=["docx", "pptx", "pdf", "txt", "doc"],
            accept_multiple_files=True,
            key="annual_nanyi"
        )

    history_files = st.file_uploader(
        "本團隊歷年教師版參考檔（可一次上傳多份）",
        type=["docx", "pptx", "pdf", "txt", "doc"],
        accept_multiple_files=True,
        key="annual_history"
    )

    expected_for_refs = len(st.session_state.questions) if st.session_state.questions else None

    if st.button(
        "建立／更新本年度參考庫",
        type="primary",
        disabled=not (hanlin_files or kang_files or nanyi_files or history_files),
        key="build_annual_ref"
    ):
        newdb = _empty_reference_db(int(st.session_state.year))
        all_errors = []

        for pub, fileset in [("翰林", hanlin_files), ("康軒", kang_files), ("南一", nanyi_files)]:
            parsed, errs = _parse_publisher_files(fileset, expected_for_refs, st.session_state.questions)
            newdb["publisher"][pub] = parsed
            all_errors.extend(errs)

        for uploaded in history_files or []:
            try:
                newdb["history_raw"][uploaded.name] = _normalize_reference_text(_uploaded_file_text(uploaded))
            except Exception as e:
                all_errors.append(f"{uploaded.name}：{e}")

        olddb = _load_reference_library()
        if olddb.get("year") == int(st.session_state.year):
            newdb["drafts"] = olddb.get("drafts", {})

        st.session_state.reference_db = newdb

        st.success(
            "年度參考庫已建立。出版社辨識："
            + "／".join(f"{p}{len(newdb['publisher'][p])}題" for p in ("翰林","康軒","南一"))
            + f"；內部參考 {len(newdb['history_raw'])} 份。"
        )

        if all_errors:
            st.warning("以下檔案需處理：\n- " + "\n- ".join(all_errors))

        st.rerun()

    st.divider()

    # --------------------------------------------------
    # B. Check
    # --------------------------------------------------
    st.markdown("## B. 檢查本年度參考庫")
    st.caption(
        "建立後先在這裡確認資料是否完整。v4.0 會以「正式題庫內容」比對出版社詳解，不再只依賴出版社題號。確認完整後再下載年度 JSON。"
    )

    active = _load_reference_library()
    expected_count = len(st.session_state.questions) if st.session_state.questions else None

    summary_rows = []
    for p in ("翰林", "康軒", "南一"):
        qdict = active.get("publisher", {}).get(p, {})
        nums = sorted(int(x) for x in qdict.keys() if str(x).isdigit())
        if expected_count:
            missing = [str(i) for i in range(1, expected_count + 1) if i not in nums]
            missing_text = "、".join(missing) if missing else "無"
        else:
            missing_text = "需先建立題庫才能比對缺題"
        summary_rows.append({
            "來源": p,
            "已辨識題數": len(qdict),
            "缺題": missing_text
        })

    summary_rows.append({
        "來源": "內部教師版參考檔",
        "已辨識題數": len(active.get("history_raw", {})),
        "缺題": "—"
    })

    st.dataframe(summary_rows, hide_index=True, use_container_width=True)

    if expected_count:
        all_ok = True
        for p in ("翰林", "康軒", "南一"):
            nums = {
                int(x) for x in active.get("publisher", {}).get(p, {}).keys()
                if str(x).isdigit()
            }
            if any(i not in nums for i in range(1, expected_count + 1)):
                all_ok = False
                break

        if all_ok:
            st.success(f"三家出版社皆已辨識完整 1～{expected_count} 題。")
        else:
            st.warning(
                "至少一家出版社仍有缺題。建議先回 A 區補資料或改用較容易解析的 DOCX／PPTX，再進入 C 區保存。"
            )
    else:
        st.info("若要檢查出版社是否缺題，請先到「① 建立題庫」建立本年度題庫。")

    st.divider()

    # --------------------------------------------------
    # C. Save
    # --------------------------------------------------
    st.markdown("## C. 儲存本年度參考包")
    st.caption(
        "確認 B 區資料正確後，再下載 JSON 保存。下次不必重新上傳全部出版社與內部檔案。"
    )

    st.download_button(
        f"下載 {int(st.session_state.year)} 年度參考包 JSON",
        data=_annual_package_json(active),
        file_name=f"{int(st.session_state.year)}_會考詳解年度參考包.json",
        mime="application/json",
        use_container_width=True
    )

    st.divider()

    # --------------------------------------------------
    # D. Reload
    # --------------------------------------------------
    st.markdown("## D. 下次繼續工作／載入既有年度參考包")
    st.caption(
        "如果這個年度之前已經建立過，不需要再從 A 區上傳全部原始檔；直接載入之前下載的年度 JSON 即可。"
    )

    package_upload = st.file_uploader(
        "上傳以前建立好的年度參考包 JSON",
        type=["json"],
        key="annual_package_upload"
    )

    if st.button("載入年度參考包", disabled=not package_upload, key="load_annual_package"):
        try:
            db = json.loads(package_upload.getvalue().decode("utf-8"))
            if "publisher" not in db or "history_raw" not in db:
                raise ValueError("這不是有效的年度參考包。")
            db.setdefault("strategy", DEFAULT_STRATEGY_LIBRARY)
            db.setdefault("drafts", {})
            st.session_state.reference_db = db
            st.success(f"已載入 {db.get('year', '未標示年度')} 年度參考包。")
            st.rerun()
        except Exception as e:
            st.error(f"載入失敗：{e}")

    st.divider()

    # --------------------------------------------------
    # E. Integrated drafts
    # --------------------------------------------------
    st.markdown("## E. 本年度整合建議稿")
    st.caption(
        "這裡保存的是『你真正編輯過的成果』，和 C 區的『來源參考庫』不同。"
        "包含每題能力類型、三家比較筆記、建議詳解、教學重點、教學步驟與筆記策略。"
    )

    draft_upload = st.file_uploader(
        "上傳整合建議稿 JSON（選填）",
        type=["json"],
        key="annual_draft_upload"
    )

    if st.button(
        "匯入整合建議稿",
        disabled=not (draft_upload and st.session_state.questions),
        key="import_drafts"
    ):
        try:
            payload = json.loads(draft_upload.getvalue().decode("utf-8"))
            if payload.get("year") not in (None, int(st.session_state.year)):
                st.warning(
                    f"此整合稿標示年度為 {payload.get('year')}，"
                    f"目前設定年度為 {int(st.session_state.year)}，請確認是否正確。"
                )
            applied = _apply_drafts_to_questions(payload, st.session_state.questions)
            st.success(f"已套用 {applied} 題整合建議稿。")
        except Exception as e:
            st.error(f"匯入失敗：{e}")

    if st.session_state.questions:
        st.download_button(
            "下載目前已編輯的整合建議稿 JSON",
            data=json.dumps(
                _drafts_from_questions(st.session_state.questions, st.session_state.year),
                ensure_ascii=False,
                indent=2
            ).encode("utf-8"),
            file_name=f"{int(st.session_state.year)}_教師版整合建議稿.json",
            mime="application/json",
            use_container_width=True
        )
    else:
        st.info("建立題庫後，才能匯出每題的整合建議稿。")

    st.divider()
    st.markdown("### 年度資料正確流程")
    st.write(
        "A 建立／更新參考庫 → B 檢查辨識結果與缺題 → C 下載年度參考包保存 → "
        "下次直接從 D 載入 → 詳解工作完成後從 E 保存整合成果。"
    )

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


        st.info(
            "題庫已建立。題目內容的人工檢查、修改與校對狀態，"
            "統一到「③ 考題總覽暨校對」處理，避免同一題在兩個頁面重複操作。"
        )

with overview_tab:
    st.subheader("考題總覽暨校對")
    st.caption(
        "這一頁同時是題庫總覽與人工校對中心。先看完整題目／題組，發現內容需要修正時直接展開「編輯／校對」。"
        "一般單題逐題呈現；題組題則合併「共用閱讀材料／頂端題幹＋全部子題」，避免脫離脈絡。"
    )

    if not st.session_state.questions:
        st.info("請先到「① 建立題庫」上傳題本、官方答案與通過率資料。")
    else:
        questions = st.session_state.questions

        # ---------- filters ----------
        fc1, fc2, fc3, fc4, fc5 = st.columns([0.95, 0.8, 1.15, 1.0, 0.9])
        with fc1:
            rate_filter = st.selectbox(
                "通過率",
                ["全部", "80%以上", "60%～79.9%", "60%以下"],
                key="overview_rate_filter"
            )
        with fc2:
            answer_filter = st.selectbox(
                "答案",
                ["全部", "A", "B", "C", "D"],
                key="overview_answer_filter"
            )
        with fc3:
            keyword = st.text_input(
                "題目／題組關鍵字",
                placeholder="例如：文意、成語、人物…",
                key="overview_keyword"
            )
        with fc4:
            review_filter = st.selectbox(
                "校對狀態",
                ["全部", "待校對", "已校對", "題組題", "圖片／複雜版面"],
                key="overview_review_filter"
            )
        with fc5:
            only_selected = st.checkbox(
                "只看已選",
                value=False,
                key="overview_only_selected"
            )

        def _norm_rate(q):
            if q.pass_rate is None:
                return None
            r = float(q.pass_rate)
            return r * 100 if r <= 1 else r

        def _rate_ok(q):
            r = _norm_rate(q)
            if rate_filter == "全部":
                return True
            if r is None:
                return False
            if rate_filter == "80%以上":
                return r >= 80
            if rate_filter == "60%～79.9%":
                return 60 <= r < 80
            if rate_filter == "60%以下":
                return r < 60
            return True

        def _question_haystack(q):
            return " ".join([
                q.group_intro or "",
                q.material or "",
                q.text or "",
                " ".join((q.options or {}).values())
            ])

        def _question_matches(q):
            if answer_filter != "全部" and q.answer != answer_filter:
                return False
            if not _rate_ok(q):
                return False
            if keyword.strip() and keyword.strip().lower() not in _question_haystack(q).lower():
                return False
            if review_filter == "待校對" and q.reviewed:
                return False
            if review_filter == "已校對" and not q.reviewed:
                return False
            if review_filter == "題組題" and not (q.group_id or "").strip():
                return False
            if review_filter == "圖片／複雜版面" and not (q.visual_mode or q.image_pngs or q.body_crop_png):
                return False
            if only_selected and not q.selected:
                return False
            return True

        # ---------- build display units ----------
        # A group is rendered once with all its children. If any child matches the
        # filter, the whole group is shown so the reader never loses its context.
        units = []
        seen_groups = set()

        for q in questions:
            gid = (q.group_id or "").strip()
            if gid:
                if gid in seen_groups:
                    continue
                seen_groups.add(gid)
                members = sorted(
                    [x for x in questions if (x.group_id or "").strip() == gid],
                    key=lambda x: x.source_no
                )
                if any(_question_matches(x) for x in members):
                    units.append(("group", gid, members))
            else:
                if _question_matches(q):
                    units.append(("single", str(q.source_no), [q]))

        visible_question_nos = {
            q.source_no
            for kind, uid, members in units
            for q in members
        }

        selected_count = sum(1 for q in questions if q.selected)
        reviewed_count = sum(1 for q in questions if q.reviewed)
        group_count = len({q.group_id for q in questions if (q.group_id or "").strip()})
        issue_count = sum(1 for q in questions if not q.reviewed)

        mc1, mc2, mc3, mc4, mc5 = st.columns(5)
        mc1.metric("全部題數", len(questions))
        mc2.metric("目前顯示", len(visible_question_nos))
        mc3.metric("已校對", reviewed_count)
        mc4.metric("待校對", issue_count)
        mc5.metric("已選入題本", selected_count)

        st.progress(reviewed_count / len(questions) if questions else 0)
        st.caption(f"人工校對進度：{reviewed_count}/{len(questions)} 題｜題組：{group_count} 組")

        bc1, bc2, bc3 = st.columns(3)
        if bc1.button("清除全部選取", key="overview_clear_all", use_container_width=True):
            for q in questions:
                q.selected = False
                st.session_state[f"sel_{q.source_no}"] = False
            st.rerun()

        if bc2.button("選取目前顯示結果", key="overview_select_visible", use_container_width=True):
            for q in questions:
                if q.source_no in visible_question_nos:
                    q.selected = True
                    st.session_state[f"sel_{q.source_no}"] = True
            st.rerun()

        if bc3.button("取消目前顯示結果", key="overview_unselect_visible", use_container_width=True):
            for q in questions:
                if q.source_no in visible_question_nos:
                    q.selected = False
                    st.session_state[f"sel_{q.source_no}"] = False
            st.rerun()

        st.divider()

        if not units:
            st.warning("目前篩選條件下沒有題目。")
        else:
            for kind, uid, members in units:
                if kind == "group":
                    first_no = min(x.source_no for x in members)
                    last_no = max(x.source_no for x in members)
                    group_title = uid if uid else f"{first_no}-{last_no}"

                    st.markdown(f"## 題組 {group_title}｜原第 {first_no}～{last_no} 題")

                    # Prefer explicit group_intro. Otherwise take the most complete
                    # shared material available among children.
                    intro_candidates = [
                        (x.group_intro or "").strip()
                        for x in members
                        if (x.group_intro or "").strip()
                    ]
                    material_candidates = [
                        (x.material or "").strip()
                        for x in members
                        if (x.material or "").strip()
                    ]

                    if intro_candidates:
                        shared_intro = max(intro_candidates, key=len)
                    elif material_candidates:
                        shared_intro = max(material_candidates, key=len)
                    else:
                        shared_intro = ""

                    shared_crops = []
                    for x in members:
                        if getattr(x, "group_crop_pngs", None):
                            shared_crops = x.group_crop_pngs
                            break

                    st.markdown("### 共用閱讀材料／頂端題幹")
                    if shared_crops:
                        st.caption("以下先呈現原題本共用材料，確保圖表、版面與特殊符號完整。")
                        for img in shared_crops:
                            st.image(img, use_container_width=True)

                    if shared_intro:
                        with st.expander("查看／複製共用材料文字", expanded=False):
                            st.write(shared_intro)
                    elif not shared_crops:
                        st.warning(
                            "本題組目前沒有辨識到共用閱讀材料。請到「① 建立題庫 → 題目結構校對」"
                            "把題組頂端文章／共用材料填入任一子題的「閱讀／共用材料」。"
                        )

                    # Whole-group selection: groups should not be split.
                    all_selected = all(x.selected for x in members)
                    some_selected = any(x.selected for x in members)
                    if some_selected and not all_selected:
                        st.warning("目前此題組只有部分子題被選取；題組題建議整組保留。")

                    group_select = st.checkbox(
                        f"整組加入本次題本（{len(members)} 題）",
                        value=all_selected,
                        key=f"overview_group_select_{uid}"
                    )
                    if group_select != all_selected:
                        for x in members:
                            x.selected = group_select
                            st.session_state[f"sel_{x.source_no}"] = group_select
                        st.rerun()

                    # Render all subquestions together, each with answer and rate.
                    for x in members:
                        rate = _norm_rate(x)
                        rate_text = "—" if rate is None else f"{rate:.1f}%"

                        st.markdown(f"### 第 {x.source_no} 題　｜答案：**{x.answer or '—'}**　｜通過率：**{rate_text}**")

                        # Avoid repeating the same shared material before every child.
                        child_material = (x.material or "").strip()
                        if child_material and child_material != shared_intro:
                            st.markdown("**本子題附加材料**")
                            st.write(child_material)

                        st.write(x.text.strip() if x.text else "（題幹未辨識，請至建立題庫頁校對）")
                        for letter in ["A", "B", "C", "D"]:
                            val = (x.options or {}).get(letter, "")
                            if val and val.strip():
                                st.write(f"({letter}) {val.strip()}")

                        status_text = "✅ 已校對" if x.reviewed else "⚠️ 待校對"
                        with st.expander(f"✏️ 編輯／校對第 {x.source_no} 題｜{status_text}", expanded=False):
                            _render_question_review_editor(
                                x,
                                key_prefix=f"overview_group_{uid}"
                            )

                    with st.expander("查看此題組原 PDF 裁圖", expanded=False):
                        for x in members:
                            if x.crop_png:
                                st.markdown(f"**原第 {x.source_no} 題**")
                                st.image(x.crop_png, use_container_width=True)

                    st.divider()

                else:
                    q = members[0]
                    rate = _norm_rate(q)
                    rate_text = "—" if rate is None else f"{rate:.1f}%"

                    card_left, card_right = st.columns([0.78, 0.22], vertical_alignment="top")
                    with card_left:
                        st.markdown(f"### 第 {q.source_no} 題")
                        if q.material and q.material.strip():
                            st.markdown("**閱讀／共用材料**")
                            st.write(q.material.strip())

                        st.markdown("**題目**")
                        st.write(q.text.strip() if q.text else "（題幹未辨識，請至建立題庫頁校對）")
                        for letter in ["A", "B", "C", "D"]:
                            val = (q.options or {}).get(letter, "")
                            if val and val.strip():
                                st.write(f"({letter}) {val.strip()}")

                    with card_right:
                        st.markdown("**題目資訊**")
                        st.write(f"答案：**{q.answer or '—'}**")
                        st.write(f"通過率：**{rate_text}**")
                        if q.category:
                            st.write(f"能力：{q.category}")
                        q.selected = st.checkbox(
                            "加入本次題本",
                            value=q.selected,
                            key=f"overview_select_{q.source_no}"
                        )

                    status_text = "✅ 已校對" if q.reviewed else "⚠️ 待校對"
                    with st.expander(f"✏️ 編輯／校對｜{status_text}", expanded=False):
                        _render_question_review_editor(q, key_prefix="overview_single")

                    with st.expander("查看原題裁圖", expanded=False):
                        if q.crop_png:
                            st.image(q.crop_png, use_container_width=True)
                        else:
                            st.caption("目前沒有原題裁圖。")
                    st.divider()

        st.info(
            "建議操作：先利用「待校對」篩選逐題驗收 → 發現問題就地展開修改 → 全部校對完成後，"
            "再利用通過率／關鍵字挑選題目 → 到「④ 篩選組題」做最後確認。"
        )

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
        if refdb.get("year") != int(st.session_state.year):
            st.warning(
                "目前年度參考庫與題本年度不同。建議先到「① 年度資料」建立／載入本年度參考包，"
                "避免誤用其他年度出版社詳解。"
            )

        selected_nums = [q.source_no for q in st.session_state.questions if q.selected]
        choices = selected_nums or [q.source_no for q in st.session_state.questions]
        qno = st.selectbox("選擇題目", choices, key="workbench_qno")
        q = next(x for x in st.session_state.questions if x.source_no == qno)

        reviewed_count = sum(1 for x in st.session_state.questions if x.workbench_reviewed)
        st.caption(f"教師詳解人工確認進度：{reviewed_count}/{len(st.session_state.questions)} 題")

        left, right = st.columns([0.88, 1.12])

        with left:
            st.markdown(f"### 原第 {q.source_no} 題")
            if q.material.strip():
                st.markdown("**閱讀／共用材料**")
                st.write(q.material)
            st.markdown("**題幹**")
            st.write(q.text)
            for k, v in q.options.items():
                st.write(f"({k}) {v}")
            st.caption(
                f"官方答案：{q.answer or '—'}｜通過率："
                f"{q.pass_rate if q.pass_rate is not None else '—'}｜原頁：{q.page_no}"
            )
            if q.crop_png:
                st.image(q.crop_png, caption="原 PDF 題目區塊", use_container_width=True)

        with right:
            st.markdown("### 一、三家出版社原始詳解")
            st.caption("此區只做來源並列，不會自動把不同出版社文字粗糙拼接成『建議詳解』。")
            pub_tabs = st.tabs(["翰林", "康軒", "南一"])
            for ptab, pub in zip(pub_tabs, ["翰林", "康軒", "南一"]):
                with ptab:
                    block = refdb.get("publisher", {}).get(pub, {}).get(str(q.source_no), "")
                    if block:
                        st.text_area(
                            f"{pub}第{q.source_no}題",
                            value=block,
                            height=330,
                            key=f"ref_{pub}_{qno}",
                            label_visibility="collapsed"
                        )
                    else:
                        st.info("目前年度參考庫沒有辨識到這一題，請回「① 年度資料」檢查來源檔。")

            st.markdown("### 二、能力類型與歷年本團隊寫法")
            current = q.category if q.category in ABILITY_OPTIONS else ""
            q.category = st.selectbox(
                "能力類型",
                ABILITY_OPTIONS,
                index=ABILITY_OPTIONS.index(current),
                key=f"cat_{qno}"
            )

            strategy = _strategy_for_category(refdb, q.category or "其他")
            if q.category:
                st.markdown("**依本團隊歷年寫法整理的教學框架**")
                st.info(strategy.get("教學重點", ""))
                st.text_area(
                    "詳細教學步驟框架",
                    value=strategy.get("教學步驟", ""),
                    height=250,
                    key=f"strategy_preview_{qno}"
                )
                examples = _history_examples_for_category(refdb, q.category)
                with st.expander(f"查看歷年原教師版摘錄（{len(examples)} 則）", expanded=False):
                    if not examples:
                        st.caption("目前年度參考包內沒有找到同能力類型文字。")
                    for source, excerpt in examples:
                        st.markdown(f"**{source}**")
                        st.text_area(
                            f"hist_{source}_{qno}",
                            value=excerpt,
                            height=220,
                            key=f"hist_{source}_{qno}",
                            label_visibility="collapsed"
                        )

            st.markdown("### 三、三家比較筆記")
            if f"syn_{qno}" not in st.session_state:
                st.session_state[f"syn_{qno}"] = q.synthesis_notes
            q.synthesis_notes = st.text_area(
                "請先記下：三家共同核心、哪一家解釋較完整、哪些選項理由值得保留、哪些內容可刪。",
                height=180,
                key=f"syn_{qno}"
            )

            st.markdown("### 四、本次整合建議稿（人工可修改）")
            st.caption(
                "這裡不再把某一家出版社直接當成『建議稿』。"
                "建議先完成上面的比較，再將人工／ChatGPT整合後的內容貼入。"
            )

            # Apply saved annual draft if available and the fields are still empty.
            saved = refdb.get("drafts", {}).get(str(q.source_no), {})
            if saved and not any([q.explanation, q.teaching_focus, q.teaching]):
                q.category = saved.get("category", q.category)
                q.synthesis_notes = saved.get("synthesis_notes", q.synthesis_notes)
                q.explanation = saved.get("explanation", "")
                q.teaching_focus = saved.get("teaching_focus", "")
                q.teaching = saved.get("teaching", "")
                q.note_strategy = saved.get("note_strategy", "")

            for key, value in {
                f"exp_{qno}": q.explanation,
                f"focus_{qno}": q.teaching_focus,
                f"teach_{qno}": q.teaching,
                f"note_{qno}": q.note_strategy,
            }.items():
                if key not in st.session_state:
                    st.session_state[key] = value

            if st.button("套用本能力類型的詳細教學框架", key=f"apply_strategy_{qno}", use_container_width=True):
                strategy = _strategy_for_category(refdb, q.category or "其他")
                st.session_state[f"focus_{qno}"] = strategy.get("教學重點", "")
                st.session_state[f"teach_{qno}"] = strategy.get("教學步驟", "")
                st.session_state[f"note_{qno}"] = strategy.get("筆記策略", "")
                q.teaching_focus = st.session_state[f"focus_{qno}"]
                q.teaching = st.session_state[f"teach_{qno}"]
                q.note_strategy = st.session_state[f"note_{qno}"]
                st.rerun()

            q.explanation = st.text_area("建議詳解", height=300, key=f"exp_{qno}")
            q.teaching_focus = st.text_area("教學重點", height=100, key=f"focus_{qno}")
            q.teaching = st.text_area("建議教學步驟", height=300, key=f"teach_{qno}")
            q.note_strategy = st.text_area("筆記策略（選填）", height=130, key=f"note_{qno}")

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
                "跨年度原則：出版社詳解、內部教師版與整合建議稿都放在「年度資料包」，"
                "Python 程式本身不綁定 115 年。隔年只更換資料包即可。"
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
            # Binary image fields cannot be serialized directly to JSON.
            # They are regenerated from the source PDF when the question bank is rebuilt,
            # so project JSON intentionally stores text/metadata only.
            d["crop_png"] = ""
            d["image_pngs"] = []
            d["body_crop_png"] = ""
            d["group_crop_pngs"] = []
            project.append(d)
        project_json = json.dumps(_json_safe({
            "version": APP_VERSION,
            "year": st.session_state.year,
            "questions": project
        }), ensure_ascii=False, indent=2)
        st.download_button(
            "下載題庫/詳解 JSON",
            data=project_json.encode("utf-8"),
            file_name=f"{st.session_state.year}_國文題庫.json",
            mime="application/json"
        )
