#!/usr/bin/env python3
"""Своя игра: отработка возражений (массовый сегмент) — поле, гиперссылки, таймеры."""

from __future__ import annotations

import re
import shutil
from pathlib import Path

from docx import Document

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "Презентация ГК Форус темный шаблон 16х9 (1).pptx"
OUT_DIR = ROOT / "своя игра возражения"
OUT = OUT_DIR / "Своя_игра_Отработка_возражений.pptx"
OUT_COPY = ROOT / "presentation" / "quiz" / "Своя_игра_Отработка_возражений.pptx"
TIMER_DIR = ROOT / "presentation" / "quiz"
SOURCE = ROOT / "своя игра 21 (2).docx"

# Brand
BLUE = RGBColor(0x26, 0xA6, 0xE0)
CARD = RGBColor(0x3F, 0x3F, 0x3F)
CARD_DARK = RGBColor(0x2A, 0x2A, 0x2A)
CARD_LIGHT = RGBColor(0xF2, 0xF2, 0xF2)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
SOFT = RGBColor(0xBF, 0xBF, 0xBF)
NEAR_BLACK = RGBColor(0x1A, 0x1A, 0x1A)
GOLD = RGBColor(0xF0, 0xB4, 0x2E)
RED = RGBColor(0xC0, 0x39, 0x2B)

FONT = "Verdana"
POINTS = [30, 70, 100, 150]
L_TITLE = 0
L_CONTENT = 3
L_BG = 6

# Таймеры: 30 → 1,5 мин; 70/100 → 2 мин; 150 → 3 мин
def think_seconds(points: int) -> int:
    if points <= 50:
        return 90
    if points <= 100:
        return 120
    return 180


def timer_label(points: int) -> str:
    sec = think_seconds(points)
    if sec % 60 == 0:
        return f"{sec // 60} мин"
    return f"{sec // 60}:{sec % 60:02d}"



def load_topics_from_docx(path: Path) -> list[dict]:
    """Читает кейсы ДОСЛОВНО из Word: п.1–5 = case, п.6–8 = key."""
    doc = Document(str(path))
    paras = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    theme_re = re.compile(r"^Тема\s*\d+\.\s*(.+)$", re.I)
    q_re = re.compile(r"^(\d+)\s*баллов?\s*[—\-–]\s*(.+)$", re.I)
    sec_re = re.compile(r"^(\d+)\.\s*(.+)$", re.S)

    themes: list[dict] = []
    cur = None
    q = None
    sections: dict[int, str] = {}
    cur_sec = None

    def flush_q() -> None:
        nonlocal q, sections, cur_sec
        if cur is None or q is None:
            q = None
            sections = {}
            cur_sec = None
            return
        case_parts = [sections[n].strip() for n in range(1, 6) if n in sections]
        key_parts = [sections[n].strip() for n in range(6, 9) if n in sections]
        cur["questions"].append(
            {
                "points": q["points"],
                "title": q["title"],
                "case": "\n\n".join(case_parts),
                "key": "\n\n".join(key_parts),
            }
        )
        q = None
        sections = {}
        cur_sec = None

    def flush_theme() -> None:
        nonlocal cur
        flush_q()
        if cur is not None:
            themes.append(cur)
        cur = None

    for p in paras:
        m_theme = theme_re.match(p)
        if m_theme:
            flush_theme()
            cur = {"name": m_theme.group(1).strip(), "questions": []}
            continue
        m_q = q_re.match(p)
        if m_q and cur is not None:
            flush_q()
            q = {"points": int(m_q.group(1)), "title": m_q.group(2).strip()}
            continue
        if q is None:
            continue
        m_sec = sec_re.match(p)
        if m_sec and 1 <= int(m_sec.group(1)) <= 8:
            cur_sec = int(m_sec.group(1))
            sections[cur_sec] = p  # дословный абзац из файла
            continue
        if cur_sec is not None:
            sections[cur_sec] = sections[cur_sec] + "\n" + p

    flush_theme()
    for t in themes:
        t["questions"].sort(key=lambda x: POINTS.index(x["points"]))
    return themes



TOPICS: list[dict] = []

def emu(inches: float) -> int:
    return int(Inches(inches))


def delete_all_slides(prs: Presentation) -> None:
    sld_id_lst = prs.slides._sldIdLst
    for sld_id in list(sld_id_lst):
        r_id = sld_id.get(qn("r:id"))
        prs.part.drop_rel(r_id)
        sld_id_lst.remove(sld_id)


def set_run(run, text, size_pt, bold=False, color=WHITE, font_name=FONT):
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.color.rgb = color
    r_pr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = r_pr.find(qn(tag))
        if el is None:
            el = etree.SubElement(r_pr, qn(tag))
        el.set("typeface", font_name)


def set_anchor(text_frame, anchor=MSO_ANCHOR.TOP):
    body_pr = text_frame._txBody.find(qn("a:bodyPr"))
    if body_pr is not None:
        body_pr.set(
            "anchor",
            {MSO_ANCHOR.TOP: "t", MSO_ANCHOR.MIDDLE: "ctr", MSO_ANCHOR.BOTTOM: "b"}.get(
                anchor, "t"
            ),
        )


def add_textbox(
    slide,
    left,
    top,
    width,
    height,
    text,
    size_pt=14,
    bold=False,
    color=WHITE,
    align=PP_ALIGN.LEFT,
    anchor=MSO_ANCHOR.TOP,
):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    set_anchor(tf, anchor)
    tf.paragraphs[0].alignment = align
    set_run(tf.paragraphs[0].add_run(), text, size_pt, bold, color)
    return box


def add_card(slide, left, top, width, height, fill=CARD, corner=0.1):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    try:
        shape.adjustments[0] = corner
    except Exception:
        pass
    return shape


def clear_body_placeholders(slide):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx in (13, 14, 1, 2):
            if ph.has_text_frame:
                ph.text_frame.clear()


def fill_title(slide, text, size=26):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0:
            tf = ph.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT
            set_run(p.add_run(), text, size, True, WHITE)
            return ph
    return add_textbox(slide, emu(0.97), emu(0.48), emu(11.4), emu(1.0), text, size, True, WHITE)


def fill_shape_text(shape, text, size_pt=14, bold=False, color=WHITE, align=PP_ALIGN.CENTER):
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    set_anchor(tf, MSO_ANCHOR.MIDDLE)
    p = tf.paragraphs[0]
    p.alignment = align
    set_run(p.add_run(), text, size_pt, bold, color)
    return shape


def link_to_slide(shape, target_slide):
    shape.click_action.target_slide = target_slide


def set_run_scheme_color(run, scheme: str = "hlink"):
    r_pr = run._r.get_or_add_rPr()
    for child in list(r_pr):
        if child.tag == qn("a:solidFill") or child.tag.endswith("}solidFill"):
            r_pr.remove(child)
    solid = etree.SubElement(r_pr, qn("a:solidFill"))
    scheme_el = etree.SubElement(solid, qn("a:schemeClr"))
    scheme_el.set("val", scheme)


def fill_shape_text_hyperlink(shape, text, size_pt=18, bold=True):
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    set_anchor(tf, MSO_ANCHOR.MIDDLE)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.name = FONT
    r_pr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        el = r_pr.find(qn(tag))
        if el is None:
            el = etree.SubElement(r_pr, qn(tag))
        el.set("typeface", FONT)
    set_run_scheme_color(run, "hlink")
    set_run_scheme_color(run, "hlink")
    return shape


def patch_theme_followed_hyperlink_red(prs: Presentation) -> None:
    for part in prs.part.package.iter_parts():
        name = str(getattr(part, "partname", ""))
        if "theme" not in name or not name.endswith(".xml"):
            continue
        try:
            root = etree.fromstring(part.blob)
        except Exception:
            continue
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
        changed = False
        hlink = root.find(".//a:clrScheme/a:hlink", ns)
        fol = root.find(".//a:clrScheme/a:folHlink", ns)
        if hlink is not None:
            for child in list(hlink):
                hlink.remove(child)
            srgb = etree.SubElement(hlink, qn("a:srgbClr"))
            srgb.set("val", "FFFFFF")
            changed = True
        if fol is not None:
            for child in list(fol):
                fol.remove(child)
            srgb = etree.SubElement(fol, qn("a:srgbClr"))
            srgb.set("val", "C0392B")
            changed = True
        if changed:
            part._blob = etree.tostring(
                root, xml_declaration=True, encoding="UTF-8", standalone=True
            )


def bind_text_run_hyperlink(shape) -> None:
    sp = shape._element
    hlink = None
    for el in sp.iter():
        if el.tag == qn("a:hlinkClick"):
            hlink = el
            break
    if hlink is None:
        return
    r_id = hlink.get(qn("r:id"))
    action = hlink.get("action")
    if not r_id:
        return
    for r in sp.iter(qn("a:r")):
        r_pr = r.find(qn("a:rPr"))
        if r_pr is None:
            r_pr = etree.Element(qn("a:rPr"))
            r.insert(0, r_pr)
        for old in list(r_pr):
            if old.tag == qn("a:hlinkClick") or old.tag.endswith("}hlinkClick"):
                r_pr.remove(old)
        for child in list(r_pr):
            if child.tag == qn("a:solidFill") or child.tag.endswith("}solidFill"):
                r_pr.remove(child)
        solid = etree.SubElement(r_pr, qn("a:solidFill"))
        scheme_el = etree.SubElement(solid, qn("a:schemeClr"))
        scheme_el.set("val", "hlink")
        r_pr.set("u", "none")
        hl = etree.SubElement(r_pr, qn("a:hlinkClick"))
        hl.set(qn("r:id"), r_id)
        if action:
            hl.set("action", action)


def fill_shape_key(shape, key_text: str):
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    set_anchor(tf, MSO_ANCHOR.TOP)
    p0 = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.LEFT
    set_run(p0.add_run(), "Ключ для ведущего", 11, True, BLUE)
    p1 = tf.add_paragraph()
    p1.space_before = Pt(4)
    p1.alignment = PP_ALIGN.LEFT
    set_run(p1.add_run(), key_text, 11, False, NEAR_BLACK)
    try:
        tf.margin_left = Inches(0.12)
        tf.margin_right = Inches(0.12)
        tf.margin_top = Inches(0.08)
    except Exception:
        pass
    return shape


def add_appear_after_ms(slide, shape, delay_ms: int) -> None:
    """Автопоявление ключа после таймера без клика (проверенная схема)."""
    spid = str(shape.shape_id)
    sld = slide._element
    for old in list(sld.findall(qn("p:timing"))):
        sld.remove(old)
    timing_xml = f"""
    <p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:tnLst>
        <p:par>
          <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
            <p:childTnLst>
              <p:seq concurrent="1" nextAc="seek">
                <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                  <p:childTnLst>
                    <p:par>
                      <p:cTn id="3" fill="hold">
                        <p:stCondLst>
                          <p:cond delay="0"/>
                        </p:stCondLst>
                        <p:childTnLst>
                          <p:par>
                            <p:cTn id="4" fill="hold">
                              <p:stCondLst>
                                <p:cond delay="{delay_ms}"/>
                              </p:stCondLst>
                              <p:childTnLst>
                                <p:par>
                                  <p:cTn id="5" presetID="1" presetClass="entr" presetSubtype="0"
                                         fill="hold" grpId="0" nodeType="withEffect">
                                    <p:stCondLst>
                                      <p:cond delay="0"/>
                                    </p:stCondLst>
                                    <p:childTnLst>
                                      <p:set>
                                        <p:cBhvr>
                                          <p:cTn id="6" dur="1" fill="hold">
                                            <p:stCondLst>
                                              <p:cond delay="0"/>
                                            </p:stCondLst>
                                          </p:cTn>
                                          <p:tgtEl>
                                            <p:spTgt spid="{spid}"/>
                                          </p:tgtEl>
                                          <p:attrNameLst>
                                            <p:attrName>style.visibility</p:attrName>
                                          </p:attrNameLst>
                                        </p:cBhvr>
                                        <p:to>
                                          <p:strVal val="visible"/>
                                        </p:to>
                                      </p:set>
                                    </p:childTnLst>
                                  </p:cTn>
                                </p:par>
                              </p:childTnLst>
                            </p:cTn>
                          </p:par>
                        </p:childTnLst>
                      </p:cTn>
                    </p:par>
                  </p:childTnLst>
                </p:cTn>
                <p:prevCondLst>
                  <p:cond evt="onPrev" delay="0">
                    <p:tgtEl><p:sldTgt/></p:tgtEl>
                  </p:cond>
                </p:prevCondLst>
                <p:nextCondLst>
                  <p:cond evt="onNext" delay="0">
                    <p:tgtEl><p:sldTgt/></p:tgtEl>
                  </p:cond>
                </p:nextCondLst>
              </p:seq>
            </p:childTnLst>
          </p:cTn>
        </p:par>
      </p:tnLst>
      <p:bldLst>
        <p:bldP spid="{spid}" grpId="0" animBg="1"/>
      </p:bldLst>
    </p:timing>
    """
    sld.append(etree.fromstring(timing_xml))



def build_title(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[L_TITLE])
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == 0:
            tf = ph.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            set_run(p.add_run(), "Клуб продавцов", 20, False, SOFT)
            p2 = tf.add_paragraph()
            p2.space_before = Pt(8)
            set_run(p2.add_run(), "Отработка возражений", 30, True, BLUE)
            p3 = tf.add_paragraph()
            p3.space_before = Pt(14)
            set_run(p3.add_run(), "Формат «Своя игра»  ·  командная работа", 15, True, WHITE)
        elif ph.placeholder_format.idx == 1:
            tf = ph.text_frame
            tf.clear()
            p = tf.paragraphs[0]
            set_run(
                p.add_run(),
                "5 тем  ·  30 кейсов  ·  10 / 30 / 50 / 70 / 100 / 150\n"
                "Таймер: 1,5 мин (10–50)  ·  2 мин (70–100)  ·  3 мин (150)\n"
                "При неверном ответе вопрос может перейти другой команде",
                13,
                False,
                SOFT,
            )
    return slide


def build_board(prs, topics=None):
    slide = prs.slides.add_slide(prs.slide_layouts[L_CONTENT])
    fill_title(slide, "Игровое поле  ·  Своя игра", 24)
    clear_body_placeholders(slide)
    add_textbox(
        slide,
        emu(0.7),
        emu(1.15),
        emu(11.8),
        emu(0.35),
        "Клик по баллам → кейс  ·  открытые ячейки краснеют  ·  таймер по сложности  ·  «К полю» — назад",
        11,
        False,
        SOFT,
    )

    topic_w = emu(3.05)
    cell_w = emu(2.05)
    cell_h = emu(0.88)
    gap_x = emu(0.14)
    gap_y = emu(0.12)
    left0 = emu(0.5)
    top0 = emu(1.55)

    theme_hdr = add_card(slide, left0, top0, topic_w, cell_h, CARD_DARK, 0.08)
    fill_shape_text(theme_hdr, "Тема", 13, True, SOFT)
    for ci, pts in enumerate(POINTS):
        left = left0 + topic_w + gap_x + ci * (cell_w + gap_x)
        hdr = add_card(slide, left, top0, cell_w, cell_h, BLUE, 0.1)
        fill_shape_text(hdr, str(pts), 16, True, WHITE)

    cell_shapes: dict[tuple[int, int], object] = {}
    topics = topics if topics is not None else TOPICS
    for ti, topic in enumerate(topics):
        top = top0 + (ti + 1) * (cell_h + gap_y)
        topic_card = add_card(slide, left0, top, topic_w, cell_h, CARD, 0.08)
        fill_shape_text(topic_card, topic["name"], 11, True, WHITE)
        for qi, q in enumerate(topic["questions"]):
            left = left0 + topic_w + gap_x + qi * (cell_w + gap_x)
            fill = GOLD if q["points"] >= 100 else BLUE
            shape = add_card(slide, left, top, cell_w, cell_h, fill, 0.12)
            fill_shape_text_hyperlink(shape, str(q["points"]), 18, True)
            cell_shapes[(ti, qi)] = shape
    return slide, cell_shapes



def fill_shape_key_exact(shape, key_text: str):
    """Ключ п.6–8 дословно из файла."""
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    set_anchor(tf, MSO_ANCHOR.TOP)
    try:
        tf.margin_left = Inches(0.12)
        tf.margin_right = Inches(0.12)
        tf.margin_top = Inches(0.06)
    except Exception:
        pass
    first = True
    size = 9 if len(key_text) > 900 else 10
    for block in key_text.split("\n\n"):
        lines = block.split("\n")
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_before = Pt(2)
        p.alignment = PP_ALIGN.LEFT
        set_run(p.add_run(), lines[0], size, True, BLUE)
        for line in lines[1:]:
            p2 = tf.add_paragraph()
            p2.space_before = Pt(0)
            p2.alignment = PP_ALIGN.LEFT
            set_run(p2.add_run(), line, size, False, NEAR_BLACK)
    return shape


def set_auto_advance(slide, ms: int, allow_click: bool = True) -> None:
    """Автопереход на следующий слайд через ms мс (таймер = длительность GIF)."""
    sld = slide._element
    for old in list(sld.findall(qn("p:transition"))):
        sld.remove(old)
    transition = etree.Element(qn("p:transition"))
    transition.set("advClick", "1" if allow_click else "0")
    transition.set("advTm", str(int(ms)))
    cSld = sld.find(qn("p:cSld"))
    if cSld is not None:
        cSld.addnext(transition)
    else:
        sld.append(transition)


def fill_multiline_card(shape, title: str, body: str, title_color=GOLD) -> None:
    """Выводит заголовок + текст дословно (по абзацам/строкам файла)."""
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    set_anchor(tf, MSO_ANCHOR.TOP)
    try:
        tf.margin_left = Inches(0.12)
        tf.margin_right = Inches(0.12)
        tf.margin_top = Inches(0.08)
        tf.margin_bottom = Inches(0.08)
    except Exception:
        pass
    p0 = tf.paragraphs[0]
    p0.alignment = PP_ALIGN.LEFT
    set_run(p0.add_run(), title, 12, True, title_color)

    size = 11 if len(body) < 800 else (10 if len(body) < 1200 else 9)
    for block in body.split("\n\n"):
        lines = block.split("\n")
        p = tf.add_paragraph()
        p.space_before = Pt(6)
        p.alignment = PP_ALIGN.LEFT
        set_run(p.add_run(), lines[0], size, True, WHITE)
        for line in lines[1:]:
            p2 = tf.add_paragraph()
            p2.space_before = Pt(1)
            p2.alignment = PP_ALIGN.LEFT
            set_run(p2.add_run(), line, size, False, WHITE)


def build_case_slide(prs, topic, qdata, board_slide):
    """Слайд кейса: только п.1–5 + таймер. После таймера — автопереход на «Ответ»."""
    slide = prs.slides.add_slide(prs.slide_layouts[L_BG])
    fill_title(slide, topic["name"], 20)
    clear_body_placeholders(slide)

    pts = qdata["points"]
    sec = think_seconds(pts)

    badge = add_card(slide, emu(10.55), emu(0.38), emu(1.55), emu(0.5), GOLD, 0.2)
    fill_shape_text(badge, f"{pts}", 16, True, NEAR_BLACK)

    gif = TIMER_DIR / f"timer_{sec}s.gif"
    if gif.exists():
        slide.shapes.add_picture(str(gif), emu(10.4), emu(1.0), width=emu(1.85))
    else:
        tcard = add_card(slide, emu(10.4), emu(1.0), emu(1.85), emu(1.85), CARD, 0.15)
        fill_shape_text(tcard, str(sec), 28, True, GOLD)

    add_textbox(
        slide,
        emu(10.3),
        emu(2.95),
        emu(2.05),
        emu(0.55),
        f"Таймер {timer_label(pts)}\nдалее — Ответ",
        11,
        True,
        SOFT,
        PP_ALIGN.CENTER,
    )

    case_card = add_card(slide, emu(0.45), emu(1.05), emu(9.7), emu(5.5), CARD, 0.08)
    title = qdata.get("title") or ""
    fill_multiline_card(case_card, f"{pts} баллов — {title}", qdata["case"])

    # Назад на поле можно досрочно
    back = add_card(slide, emu(10.4), emu(5.7), emu(1.85), emu(0.9), BLUE, 0.12)
    fill_shape_text(back, "← К полю", 13, True, WHITE)
    link_to_slide(back, board_slide)

    # Автопереход на следующий слайд (Ответ) ровно по длительности таймера
    set_auto_advance(slide, sec * 1000, allow_click=True)
    return slide


def build_answer_slide(prs, topic, qdata, board_slide):
    """Слайд «Ответ»: п.6–8 дословно из файла."""
    slide = prs.slides.add_slide(prs.slide_layouts[L_BG])
    pts = qdata["points"]
    fill_title(slide, f"Ответ  ·  {topic['name']}  ·  {pts}", 18)
    clear_body_placeholders(slide)

    badge = add_card(slide, emu(10.55), emu(0.38), emu(1.55), emu(0.5), GOLD, 0.2)
    fill_shape_text(badge, "ОТВЕТ", 14, True, NEAR_BLACK)

    title = qdata.get("title") or ""
    add_textbox(
        slide,
        emu(0.55),
        emu(1.0),
        emu(9.7),
        emu(0.4),
        f"{pts} баллов — {title}",
        12,
        True,
        GOLD,
        PP_ALIGN.LEFT,
    )

    ans_card = add_card(slide, emu(0.45), emu(1.45), emu(9.7), emu(4.9), CARD_LIGHT, 0.08)
    # На светлой карточке текст тёмный
    tf = ans_card.text_frame
    tf.clear()
    tf.word_wrap = True
    set_anchor(tf, MSO_ANCHOR.TOP)
    try:
        tf.margin_left = Inches(0.14)
        tf.margin_right = Inches(0.14)
        tf.margin_top = Inches(0.1)
        tf.margin_bottom = Inches(0.1)
    except Exception:
        pass
    body = qdata["key"]
    size = 11 if len(body) < 900 else (10 if len(body) < 1300 else 9)
    first = True
    for block in body.split("\n\n"):
        lines = block.split("\n")
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_before = Pt(5)
        p.alignment = PP_ALIGN.LEFT
        set_run(p.add_run(), lines[0], size, True, BLUE)
        for line in lines[1:]:
            p2 = tf.add_paragraph()
            p2.space_before = Pt(1)
            p2.alignment = PP_ALIGN.LEFT
            set_run(p2.add_run(), line, size, False, NEAR_BLACK)

    back = add_card(slide, emu(10.4), emu(5.7), emu(1.85), emu(0.9), BLUE, 0.12)
    fill_shape_text(back, "← К полю", 13, True, WHITE)
    link_to_slide(back, board_slide)
    return slide


def sort_topic_questions(topics: list[dict]) -> None:
    for topic in topics:
        topic["questions"].sort(key=lambda q: POINTS.index(q["points"]))


def verify_package(path: Path) -> None:
    import zipfile

    def slide_key(name: str) -> int:
        return int(re.search(r"slide(\d+)", name).group(1))

    with zipfile.ZipFile(path) as z:
        slides = sorted(
            (n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)),
            key=slide_key,
        )
        # title + board + 20 case + 20 answer = 42
        assert len(slides) == 42, len(slides)
        board = z.read(slides[1]).decode("utf-8")
        h = board.count("hlinkClick")
        assert h >= 20, h

        # Case slides: 2,4,6,... indices 2..41 step? Actually order: case then answer pairs
        # slides[2]=case1, slides[3]=answer1, ...
        delays = set()
        advances = 0
        for i in range(2, len(slides), 2):
            case_xml = z.read(slides[i]).decode("utf-8")
            ans_xml = z.read(slides[i + 1]).decode("utf-8")
            assert "advTm=" in case_xml, slides[i]
            assert "advTm=" not in ans_xml, slides[i + 1]
            m = re.search(r'advTm="(\d+)"', case_xml)
            assert m, slides[i]
            delays.add(int(m.group(1)))
            advances += 1
            # timer gif on case
            rel = slides[i].replace("slides/", "slides/_rels/") + ".rels"
            assert "media/" in z.read(rel).decode("utf-8"), slides[i]
        assert advances == 20, advances
        assert delays <= {90000, 120000, 180000} and delays, delays
        media = [n for n in z.namelist() if n.startswith("ppt/media/") and n.endswith(".gif")]
        assert len(media) >= 3, media
        print(
            "Verify OK: slides=",
            len(slides),
            "board hlinks=",
            h,
            "advTm=",
            sorted(delays),
            "timer gifs=",
            len(media),
        )


def main():
    if not TEMPLATE.exists():
        raise SystemExit(f"Template not found: {TEMPLATE}")
    if not SOURCE.exists():
        raise SystemExit(f"Source not found: {SOURCE}")
    for sec in (90, 120, 180):
        gif = TIMER_DIR / f"timer_{sec}s.gif"
        if not gif.exists():
            raise SystemExit(f"Timer GIF not found: {gif}")

    topics = load_topics_from_docx(SOURCE)
    if len(topics) != 5:
        raise SystemExit(f"Expected 5 themes, got {len(topics)}")
    for t in topics:
        pts = [q["points"] for q in t["questions"]]
        if pts != POINTS:
            raise SystemExit(f"Theme {t['name']}: points {pts}")
        for q in t["questions"]:
            if not q["case"].strip() or not q["key"].strip():
                raise SystemExit(f"Empty case/key: {t['name']} / {q['points']}")

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    OUT_COPY.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(TEMPLATE, OUT)

    prs = Presentation(str(OUT))
    delete_all_slides(prs)

    build_title(prs)
    board_slide, cell_shapes = build_board(prs, topics)

    # Важно: board ссылается на слайд КЕЙСА; ответ идёт следующим и открывается по advTm
    for ti, topic in enumerate(topics):
        for qi, qdata in enumerate(topic["questions"]):
            case_slide = build_case_slide(prs, topic, qdata, board_slide)
            build_answer_slide(prs, topic, qdata, board_slide)
            cell = cell_shapes.get((ti, qi))
            if cell is not None:
                link_to_slide(cell, case_slide)
                bind_text_run_hyperlink(cell)
                try:
                    cell.name = f"Cell_{ti}_{qi}"
                except Exception:
                    pass

    patch_theme_followed_hyperlink_red(prs)
    prs.save(str(OUT))
    shutil.copy2(OUT, OUT_COPY)

    n_q = sum(len(t["questions"]) for t in topics)
    print(f"Saved: {OUT}")
    print(f"Copy:  {OUT_COPY}")
    print(f"Slides: {len(prs.slides)} (title + board + {n_q} cases + {n_q} answers)")
    print("Timers/auto-advance: 90s (30), 120s (70–100), 180s (150)")
    print("Themes:", " | ".join(t["name"] for t in topics))
    verify_package(OUT)


if __name__ == "__main__":
    main()
