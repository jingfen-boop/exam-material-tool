
import io
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

APP_VERSION = "Web v1.2"

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
    teaching: str = ""
    note_strategy: str = ""
    group_id: str = ""
    group_intro: str = ""
    crop_png: bytes = b""
    visual_mode: bool = False
    selected: bool = True

    def __post_init__(self):
        if self.options is None:
            self.options = {}

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

def _candidate_number_lines(page):
    out = []
    d = page.get_text("dict")
    for block in d.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            txt = _line_text(line).strip()
            m = re.fullmatch(r"(\d{1,2})\.\s*", txt)
            if m:
                bbox = line.get("bbox", [0,0,0,0])
                out.append((int(m.group(1)), bbox))
    return sorted(out, key=lambda x: (x[1][1], x[1][0]))

def _sequential_question_starts(doc, expected_count):
    expected = 1
    starts = []
    for pi, page in enumerate(doc, start=1):
        if pi == 1:
            continue
        for qno, bbox in _candidate_number_lines(page):
            if qno == expected:
                starts.append({"qno": qno, "page": pi, "bbox": bbox})
                expected += 1
                if expected > expected_count:
                    return starts
    return starts

def _region_text_and_options(page, y0, y1, qno):
    d = page.get_text("dict")
    rows = []
    for block in d.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            bbox = line.get("bbox", [0,0,0,0])
            cy = (bbox[1]+bbox[3])/2
            if y0 <= cy < y1:
                t = _line_text(line).strip()
                if t and t not in {"請翻頁繼續作答", "請不要翻到次頁！"}:
                    if re.fullmatch(r"\d{1,2}", t) and bbox[1] > page.rect.height-100:
                        continue
                    rows.append((bbox[1], bbox[0], t))
    rows.sort(key=lambda x:(x[0],x[1]))

    stem_lines, options = [], {}
    current_opt = None
    for _, _, t in rows:
        if re.fullmatch(rf"{qno}\.\s*", t):
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

    promptish = [x for x in stem_lines if ("何者" in x or "下列" in x or "根據" in x or "關於" in x)]
    if promptish and stem_lines and stem_lines[0] not in promptish:
        first = promptish[0]
        stem_lines = [first] + [x for x in stem_lines if x != first]

    return "\n".join(stem_lines).strip(), options

def extract_questions(question_pdf: bytes, expected_count: int = 42) -> Tuple[List[Question], Dict[int, bytes]]:
    doc = fitz.open(stream=question_pdf, filetype="pdf")
    page_images = {}
    for pi, page in enumerate(doc, start=1):
        pix = page.get_pixmap(matrix=fitz.Matrix(1.25,1.25), alpha=False)
        page_images[pi] = pix.tobytes("png")

    full_text = "\n".join(page.get_text("text") for page in doc)
    groups = []
    for m in re.finditer(r"回答\s*(\d{1,2})\s*[～~\-至]\s*(\d{1,2})\s*題", full_text):
        a,b=int(m.group(1)),int(m.group(2))
        if a<=b:
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
        cpix = page.get_pixmap(matrix=fitz.Matrix(1.5,1.5), clip=crop_rect, alpha=False)
        crop = cpix.tobytes("png")

        visual = False
        d = page.get_text("dict")
        for block in d.get("blocks", []):
            if block.get("type") == 1 and fitz.Rect(block.get("bbox")).intersects(crop_rect):
                visual = True
                break

        q = Question(source_no=qno, page_no=pi, text=stem, options=options,
                     crop_png=crop, visual_mode=visual)
        for a,b in groups:
            if a <= qno <= b:
                q.group_id=f"{a}-{b}"
                break
        questions.append(q)

    return questions, page_images

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
    doc.save(out)
    return out.getvalue()

# -----------------------------
# AI generation
# -----------------------------
def ai_generate(q: Question, api_key: str, model: str):
    from openai import OpenAI
    client = OpenAI(api_key=api_key)
    opts = "\n".join([f"({k}) {v}" for k,v in q.options.items()])
    prompt = f"""
你是臺灣國中會考國文教材編寫助理。請根據題目、選項與官方答案產出教材初稿。
不得改動官方答案。若資料不足，請明確標示「待人工確認」。

題號：{q.source_no}
題目：{q.text}
選項：
{opts}
官方答案：{q.answer}
通過率：{q.pass_rate}

請只回傳 JSON，不要加 Markdown，格式：
{{
  "category": "從字詞辨識、表層文意理解、推論理解、分析評鑑、其他中擇一",
  "explanation": "逐項解析 A-D，最後明確說明正確答案；若選項資料不完整則據實說明",
  "teaching": "3-5 個可直接在國中課堂操作的教學步驟，強調讀題、關鍵字句、選項比對、討論與遷移",
  "note_strategy": "如適合則提供簡短筆記策略，不適合可留空"
}}
"""
    resp = client.responses.create(model=model, input=prompt)
    txt = resp.output_text.strip()
    txt = re.sub(r"^```(?:json)?\s*|\s*```$", "", txt, flags=re.S)
    return json.loads(txt)

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
    st.subheader("AI（選用）")
    api_key = st.text_input("OpenAI API Key", type="password", help="不填也能建立題庫、篩題與輸出 Word。")
    model = st.text_input("模型", value="gpt-5")
    st.caption("API Key 只存在目前瀏覽器工作階段，不會寫入輸出檔。")

tab1, tab2, tab3, tab4 = st.tabs(["① 建立題庫", "② 篩選組題", "③ 編輯詳解", "④ 產生 Word"])

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
                "有圖/複雜版面": "是" if q.visual_mode else "",
                "題幹預覽": q.text[:55].replace("\n"," ")
            })
        st.dataframe(data, use_container_width=True, hide_index=True)

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

        if st.button("套用篩選"):
            qs = st.session_state.questions
            selected_nos = {q.source_no for q in qs if q.pass_rate is not None and lo <= q.pass_rate <= hi}
            if keep_groups:
                group_ids = {q.group_id for q in qs if q.source_no in selected_nos and q.group_id}
                for q in qs:
                    if q.group_id and q.group_id in group_ids:
                        selected_nos.add(q.source_no)
            for q in qs:
                q.selected = q.source_no in selected_nos
            st.success(f"目前選取 {len(selected_nos)} 題。")

        st.markdown("#### 手動勾選")
        for q in st.session_state.questions:
            rate = "—" if q.pass_rate is None else f"{q.pass_rate:.2f}"
            q.selected = st.checkbox(
                f"原第 {q.source_no} 題｜通過率 {rate}｜答案 {q.answer or '—'}｜{q.category or '未分類'}",
                value=q.selected,
                key=f"sel_{q.source_no}"
            )

with tab3:
    if not st.session_state.questions:
        st.info("請先建立題庫。")
    else:
        selected_nums = [q.source_no for q in st.session_state.questions if q.selected]
        choices = selected_nums or [q.source_no for q in st.session_state.questions]
        qno = st.selectbox("選擇題目", choices)
        q = next(x for x in st.session_state.questions if x.source_no == qno)

        c1,c2 = st.columns([1,1])
        with c1:
            st.markdown(f"### 原第 {q.source_no} 題")
            st.write(q.text)
            for k,v in q.options.items():
                st.write(f"({k}) {v}")
            st.caption(f"官方答案：{q.answer or '—'}｜通過率：{q.pass_rate if q.pass_rate is not None else '—'}｜原頁：{q.page_no}")
            if q.crop_png:
                st.image(q.crop_png, caption="原 PDF 題目區塊（供校對）", use_container_width=True)
        with c2:
            q.category = st.selectbox(
                "能力類型",
                ["", "字詞辨識", "表層文意理解", "推論理解", "分析評鑑", "其他"],
                index=["", "字詞辨識", "表層文意理解", "推論理解", "分析評鑑", "其他"].index(q.category if q.category in ["", "字詞辨識", "表層文意理解", "推論理解", "分析評鑑", "其他"] else "其他"),
                key=f"cat_{qno}"
            )
            q.explanation = st.text_area("解析", value=q.explanation, height=220, key=f"exp_{qno}")
            q.teaching = st.text_area("教學步驟", value=q.teaching, height=180, key=f"teach_{qno}")
            q.note_strategy = st.text_area("筆記策略（選填）", value=q.note_strategy, height=100, key=f"note_{qno}")
            q.visual_mode = st.checkbox("輸出時保留原題裁圖（適合圖片／表格／複雜版面）", value=q.visual_mode, key=f"vis_{qno}")

            if st.button("AI 產生／重寫此題", disabled=not api_key):
                try:
                    with st.spinner("產生教材初稿…"):
                        result = ai_generate(q, api_key, model)
                        q.category = result.get("category", q.category)
                        q.explanation = result.get("explanation", q.explanation)
                        q.teaching = result.get("teaching", q.teaching)
                        q.note_strategy = result.get("note_strategy", q.note_strategy)
                    st.success("已產生。請人工確認後再輸出。")
                    st.rerun()
                except Exception as e:
                    st.error(f"AI 產生失敗：{e}")

with tab4:
    if not st.session_state.questions:
        st.info("請先建立題庫。")
    else:
        selected = [q for q in st.session_state.questions if q.selected]
        st.write(f"目前選取：**{len(selected)} 題**")
        preserve_visual = st.checkbox("圖片／圖表／複雜題優先保留原 PDF 裁圖", value=True)
        suffix = st.text_input("題本標題後綴", value="通過率篩選題本")

        c1,c2 = st.columns(2)
        with c1:
            student_bytes = make_docx(st.session_state.questions, int(st.session_state.year), suffix, teacher=False, preserve_visual=preserve_visual)
            st.download_button(
                "⬇️ 下載學生題本 Word",
                data=student_bytes,
                file_name=f"{st.session_state.year}年會考國文_{suffix}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True
            )
        with c2:
            teacher_bytes = make_docx(st.session_state.questions, int(st.session_state.year), suffix+"(詳解_教學法)", teacher=True, preserve_visual=preserve_visual)
            st.download_button(
                "⬇️ 下載詳解／教學法 Word",
                data=teacher_bytes,
                file_name=f"{st.session_state.year}年會考國文_{suffix}(詳解_教學法).docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

        # Export/import project JSON
        st.divider()
        st.markdown("#### 儲存工作進度")
        project = []
        for q in st.session_state.questions:
            d = asdict(q)
            d["crop_png"] = ""  # JSON 不保存大型圖片，重新匯入 PDF 可再建 crop
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
