#!/usr/bin/env python3
"""
export_service.py — Generate polished PowerPoint reports from extracted document data.

Public API:
    build_pptx(info: dict, source_filename: str) -> bytes   # returns raw .pptx bytes

`info` dict schema (all keys optional except 'summary'):
    summary       str        — main document summary text
    topics        list[str]  — key topics / themes
    entities      list[str]  — named entities (people, orgs, places)
    dates         list[str]  — dates / time references
    actions       list[str]  — action items or recommendations
    sentiment     str        — "positive" | "neutral" | "negative" | free text
    language      str        — detected language label (e.g. "Chinese", "English")
    page_count    int        — source document page count
    char_count    int        — total extracted characters
    sections      list[dict] — optional extra sections: [{title, content, type}]
                               type: "bullets" | "text" | "table"
"""

from __future__ import annotations
import io, re, math, datetime
from typing import Any, List

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
except ImportError:
    raise ImportError("python-pptx required: pip install python-pptx")


# ─── Design tokens ────────────────────────────────────────────────────────────

W = Inches(13.33)
H = Inches(7.5)

C_PURPLE_DARK  = RGBColor(0x2D, 0x17, 0x6B)
C_PURPLE_MID   = RGBColor(0x4A, 0x2C, 0x99)
C_PURPLE_LIGHT = RGBColor(0x7C, 0x5C, 0xC8)
C_TEAL         = RGBColor(0x00, 0x96, 0x88)
C_AMBER        = RGBColor(0xF5, 0x9E, 0x0E)
C_GREEN        = RGBColor(0x16, 0xA3, 0x4A)
C_RED          = RGBColor(0xDC, 0x26, 0x26)
C_GRAY_TEXT    = RGBColor(0x37, 0x41, 0x51)
C_MUTED        = RGBColor(0x6B, 0x72, 0x80)
C_WHITE        = RGBColor(0xFF, 0xFF, 0xFF)
C_BG_LIGHT     = RGBColor(0xF8, 0xF7, 0xFF)
C_CARD_BG      = RGBColor(0xFF, 0xFF, 0xFF)
C_DIVIDER      = RGBColor(0xE5, 0xE7, 0xEB)

FONT_DISPLAY = "Calibri"
FONT_BODY    = "Calibri"

# Layout limits — if content exceeds these, additional slides are auto-created
BULLETS_PER_SLIDE  = 7    # max bullet items per slide
CHARS_PER_SLIDE    = 900  # max characters of body text per slide
PILLS_PER_SLIDE    = 12   # max topic pills per slide (3-col × 4-row)
ITEMS_PER_LIST_SLIDE = 10  # max entities or dates per two-column slide
ACTIONS_PER_SLIDE  = 7    # max action rows per slide


# ─── Markdown stripper ────────────────────────────────────────────────────────

def _strip_md(text: str) -> str:
    if not text:
        return ""
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
    text = re.sub(r"\*(.+?)\*",       r"\1", text)
    text = re.sub(r"`(.+?)`",           r"\1", text)
    text = re.sub(r"^#{1,6}\s+",       "",  text, flags=re.M)
    text = re.sub(r"^[-–—]{3,}$",       "",  text, flags=re.M)
    text = re.sub(r"^[\*\-]\s+",     "",  text, flags=re.M)
    text = re.sub(r"\n{3,}",           "\n\n", text)
    return text.strip()


def _chunk_list(lst: list, size: int) -> list:
    """Split a list into chunks of at most `size` items."""
    return [lst[i:i+size] for i in range(0, max(len(lst), 1), size)]


def _chunk_text(text: str, max_chars: int) -> list:
    """Split long text into chunks that fit within max_chars, respecting paragraphs."""
    paragraphs = [p.strip() for p in text.split("\n") if p.strip()]
    chunks, current, current_len = [], [], 0
    for para in paragraphs:
        if current_len + len(para) > max_chars and current:
            chunks.append("\n".join(current))
            current, current_len = [], 0
        current.append(para)
        current_len += len(para)
    if current:
        chunks.append("\n".join(current))
    return chunks or [text]


# ─── Low-level helpers ────────────────────────────────────────────────────────

def _slide_bg(slide, color: RGBColor):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _rect(slide, x, y, w, h, fill=None, line_color=None, line_width=0):
    shape = slide.shapes.add_shape(1, x, y, w, h)
    shape.line.fill.background()
    if fill:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill
    else:
        shape.fill.background()
    if line_color and line_width:
        shape.line.color.rgb = line_color
        shape.line.width = Pt(line_width)
    else:
        shape.line.fill.background()
    return shape


def _text_box(slide, x, y, w, h, text: str, *,
              font_name=FONT_BODY, font_size=14, bold=False, italic=False,
              color: RGBColor = C_GRAY_TEXT, align=PP_ALIGN.LEFT, word_wrap=True):
    txb = slide.shapes.add_textbox(x, y, w, h)
    txb.word_wrap = word_wrap
    tf = txb.text_frame
    tf.word_wrap = word_wrap
    tf.auto_size = None
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = _strip_md(text)
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txb


def _text_box_multiline(slide, x, y, w, h, lines: list, *,
                         font_name=FONT_BODY, font_size=13,
                         color=C_GRAY_TEXT, bullet_color=C_PURPLE_MID,
                         indent=True):
    """Add a text box with one paragraph per line, each with a bullet marker."""
    txb = slide.shapes.add_textbox(x, y, w, h)
    txb.word_wrap = True
    tf = txb.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        if indent:
            from pptx.util import Pt as _Pt
            p.space_before = _Pt(3)
        run = p.add_run()
        run.text = ("• " if indent else "") + _strip_md(line)
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.font.color.rgb = color
    return txb


def _footer(slide, label: str = "AI Document Analysis"):
    bar_h = Inches(0.28)
    _rect(slide, 0, H - bar_h, W, bar_h, fill=C_PURPLE_DARK)
    _text_box(slide, Inches(0.3), H - bar_h, Inches(6), bar_h,
              label, font_size=8, color=RGBColor(0xCC, 0xBB, 0xFF), align=PP_ALIGN.LEFT)
    ts = datetime.datetime.now().strftime("%Y-%m-%d")
    _text_box(slide, W - Inches(2), H - bar_h, Inches(1.9), bar_h,
              ts, font_size=8, color=RGBColor(0xCC, 0xBB, 0xFF), align=PP_ALIGN.RIGHT)


def _section_heading(slide, title: str, y_pos, subtitle: str = ""):
    _rect(slide, Inches(0.4), y_pos, Inches(0.06), Inches(0.38), fill=C_PURPLE_MID)
    _text_box(slide, Inches(0.6), y_pos - Inches(0.02), Inches(11.5), Inches(0.45),
              title, font_name=FONT_DISPLAY, font_size=18, bold=True,
              color=C_PURPLE_DARK, align=PP_ALIGN.LEFT)
    if subtitle:
        _text_box(slide, Inches(0.6), y_pos + Inches(0.42), Inches(11.5), Inches(0.32),
                  subtitle, font_size=10, color=C_MUTED)


def _content_card(slide, x, y, w, h, accent_color=None):
    """White card with optional colored top bar."""
    _rect(slide, x, y, w, h, fill=C_CARD_BG, line_color=C_DIVIDER, line_width=1)
    if accent_color:
        _rect(slide, x, y, w, Inches(0.055), fill=accent_color)


# ─── Slide builders ───────────────────────────────────────────────────────────

def _slide_cover(prs, info: dict, source_filename: str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _slide_bg(slide, C_PURPLE_DARK)
    _rect(slide, 0, 0, W, Inches(0.08), fill=C_PURPLE_LIGHT)
    _rect(slide, 0, Inches(0.08), W, Inches(0.04), fill=C_AMBER)

    _text_box(slide, Inches(0.7), Inches(0.9), Inches(11.9), Inches(1.1),
              "AI 文档分析报告",
              font_name=FONT_DISPLAY, font_size=40, bold=True,
              color=C_WHITE, align=PP_ALIGN.LEFT)

    clean_src = source_filename.replace("_", " ")
    _text_box(slide, Inches(0.7), Inches(2.0), Inches(11.9), Inches(0.55),
              f"来源文件：{clean_src}",
              font_size=13, color=RGBColor(0xCC, 0xBB, 0xFF))

    ts = datetime.datetime.now().strftime("导出时间：%Y-%m-%d  %H:%M")
    _text_box(slide, Inches(0.7), Inches(2.55), Inches(6), Inches(0.4),
              ts, font_size=11, color=RGBColor(0xAA, 0x99, 0xDD))

    stats = [
        ("📄", str(info.get("page_count", "—")), "页数"),
        ("🔤", f"{info.get('char_count', 0):,}", "字符数"),
        ("🏷", str(len(info.get("topics", []))), "主题"),
        ("✅", str(len(info.get("actions", []))), "行动项"),
    ]
    card_w = Inches(2.8)
    card_h = Inches(1.35)
    gap    = Inches(0.28)
    start_x = (W - (card_w * 4 + gap * 3)) / 2
    y = Inches(3.6)

    for i, (emoji, value, label) in enumerate(stats):
        cx = start_x + i * (card_w + gap)
        _rect(slide, cx, y, card_w, card_h,
              fill=RGBColor(0x3D, 0x27, 0x7B),
              line_color=C_PURPLE_LIGHT, line_width=1)
        _text_box(slide, cx, y + Inches(0.08), card_w, Inches(0.45),
                  emoji, font_size=22, align=PP_ALIGN.CENTER, color=C_WHITE)
        _text_box(slide, cx, y + Inches(0.48), card_w, Inches(0.52),
                  value, font_size=26, bold=True, color=C_WHITE, align=PP_ALIGN.CENTER)
        _text_box(slide, cx, y + Inches(0.98), card_w, Inches(0.32),
                  label, font_size=11, color=RGBColor(0xCC, 0xBB, 0xFF), align=PP_ALIGN.CENTER)

    sentiment = info.get("sentiment", "neutral").lower()
    s_color = C_GREEN if "pos" in sentiment else (C_RED if "neg" in sentiment else C_TEAL)
    s_label = ("🟢 正面" if "pos" in sentiment
               else ("🔴 负面" if "neg" in sentiment else "🔵 中性"))
    _rect(slide, Inches(0.7), Inches(5.35), Inches(2.5), Inches(0.48),
          fill=RGBColor(0x3D, 0x27, 0x7B), line_color=s_color, line_width=2)
    _text_box(slide, Inches(0.7), Inches(5.35), Inches(2.5), Inches(0.48),
              f"情感倾向: {s_label}", font_size=12, bold=True,
              color=C_WHITE, align=PP_ALIGN.CENTER)


def _slides_summary(prs, info: dict):
    """
    Summary — auto-splits into multiple slides if text exceeds CHARS_PER_SLIDE.
    """
    summary = _strip_md(info.get("summary", "暂无摘要。"))
    chunks = _chunk_text(summary, CHARS_PER_SLIDE)

    for idx, chunk in enumerate(chunks):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        _slide_bg(slide, C_BG_LIGHT)
        subtitle = f"({idx+1}/{len(chunks)})" if len(chunks) > 1 else ""
        _section_heading(slide, "文档摘要", Inches(0.38), subtitle)

        card_x, card_y = Inches(0.4), Inches(1.0)
        card_w = W - Inches(0.8)
        card_h = H - Inches(1.7)
        _content_card(slide, card_x, card_y, card_w, card_h, C_PURPLE_MID)
        _text_box(slide, card_x + Inches(0.25), card_y + Inches(0.18),
                  card_w - Inches(0.5), card_h - Inches(0.3),
                  chunk, font_size=15, color=C_GRAY_TEXT, word_wrap=True)
        _footer(slide)


def _slides_topics(prs, info: dict):
    """
    Topics pill grid — auto-adds slides for every PILLS_PER_SLIDE topics.
    """
    topics = info.get("topics", [])
    if not topics:
        return

    chunks = _chunk_list(topics, PILLS_PER_SLIDE)
    pill_colors = [C_PURPLE_MID, C_TEAL, C_AMBER,
                   RGBColor(0x06, 0x82, 0x72), RGBColor(0x15, 0x78, 0xC2)]
    pill_w, pill_h = Inches(3.8), Inches(0.68)
    cols = 3
    gap_x, gap_y = Inches(0.28), Inches(0.32)
    start_x = (W - (pill_w * cols + gap_x * (cols - 1))) / 2

    for idx, chunk in enumerate(chunks):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        _slide_bg(slide, C_BG_LIGHT)
        subtitle = f"({idx+1}/{len(chunks)})" if len(chunks) > 1 else ""
        _section_heading(slide, "核心主题", Inches(0.38), subtitle)

        start_y = Inches(1.15)
        for i, topic in enumerate(chunk):
            row, col = divmod(i, cols)
            px = start_x + col * (pill_w + gap_x)
            py = start_y + row * (pill_h + gap_y)
            c = pill_colors[i % len(pill_colors)]
            _rect(slide, px, py, pill_w, pill_h, fill=c)
            _text_box(slide, px, py, pill_w, pill_h,
                      _strip_md(topic), font_size=14, bold=True,
                      color=C_WHITE, align=PP_ALIGN.CENTER)
        _footer(slide)


def _slides_entities_dates(prs, info: dict):
    """
    Entities + Dates — each list auto-splits into additional slides when long.
    """
    entities = info.get("entities", [])
    dates    = info.get("dates", [])
    if not entities and not dates:
        return

    e_chunks = _chunk_list(entities, ITEMS_PER_LIST_SLIDE) if entities else [[]]
    d_chunks = _chunk_list(dates,    ITEMS_PER_LIST_SLIDE) if dates    else [[]]
    n_slides = max(len(e_chunks), len(d_chunks))

    for idx in range(n_slides):
        e_chunk = e_chunks[idx] if idx < len(e_chunks) else []
        d_chunk = d_chunks[idx] if idx < len(d_chunks) else []

        slide = prs.slides.add_slide(prs.slide_layouts[6])
        _slide_bg(slide, C_BG_LIGHT)
        subtitle = f"({idx+1}/{n_slides})" if n_slides > 1 else ""
        _section_heading(slide, "实体与时间线", Inches(0.38), subtitle)

        half_w = W / 2 - Inches(0.6)
        card_h = H - Inches(1.7)
        card_y = Inches(1.0)

        # Entities card
        _content_card(slide, Inches(0.4), card_y, half_w, card_h, C_PURPLE_MID)
        _text_box(slide, Inches(0.5), card_y + Inches(0.1), half_w - Inches(0.2), Inches(0.42),
                  "📌 关键实体", font_size=13, bold=True, color=C_PURPLE_DARK)
        for i, ent in enumerate(e_chunk):
            ey = card_y + Inches(0.58) + i * Inches(0.5)
            if ey + Inches(0.42) > card_y + card_h - Inches(0.15):
                break
            _rect(slide, Inches(0.5), ey, Inches(0.06), Inches(0.32), fill=C_PURPLE_MID)
            _text_box(slide, Inches(0.68), ey - Inches(0.03),
                      half_w - Inches(0.35), Inches(0.42),
                      _strip_md(ent), font_size=12, color=C_GRAY_TEXT)

        # Dates card
        dx = W / 2 + Inches(0.2)
        _content_card(slide, dx, card_y, half_w, card_h, C_TEAL)
        _text_box(slide, dx + Inches(0.1), card_y + Inches(0.1),
                  half_w - Inches(0.2), Inches(0.42),
                  "📅 时间节点", font_size=13, bold=True, color=C_TEAL)
        for i, dt in enumerate(d_chunk):
            dy_pos = card_y + Inches(0.58) + i * Inches(0.5)
            if dy_pos + Inches(0.42) > card_y + card_h - Inches(0.15):
                break
            _rect(slide, dx + Inches(0.1), dy_pos, Inches(0.06), Inches(0.32), fill=C_TEAL)
            _text_box(slide, dx + Inches(0.28), dy_pos - Inches(0.03),
                      half_w - Inches(0.35), Inches(0.42),
                      _strip_md(dt), font_size=12, color=C_GRAY_TEXT)
        _footer(slide)


def _slides_actions(prs, info: dict):
    """
    Action items — one slide per ACTIONS_PER_SLIDE items.
    """
    actions = info.get("actions", [])
    if not actions:
        return

    chunks = _chunk_list(actions, ACTIONS_PER_SLIDE)
    badge_colors = [C_PURPLE_MID, C_TEAL, C_AMBER, C_GREEN,
                    C_RED, C_PURPLE_LIGHT, C_TEAL]

    for idx, chunk in enumerate(chunks):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        _slide_bg(slide, C_BG_LIGHT)
        subtitle = f"({idx+1}/{len(chunks)})" if len(chunks) > 1 else ""
        _section_heading(slide, "行动项与建议", Inches(0.38), subtitle)

        row_h   = Inches(0.68)
        row_w   = W - Inches(0.8)
        start_y = Inches(1.05)
        global_i = idx * ACTIONS_PER_SLIDE  # keep badge numbers sequential across slides

        for i, action in enumerate(chunk):
            ry = start_y + i * (row_h + Inches(0.1))
            if ry + row_h > H - Inches(0.5):
                break
            bc = badge_colors[(global_i + i) % len(badge_colors)]
            _content_card(slide, Inches(0.4), ry, row_w, row_h)
            _rect(slide, Inches(0.4), ry, Inches(0.07), row_h, fill=bc)
            _rect(slide, Inches(0.57), ry + Inches(0.17), Inches(0.36), Inches(0.36), fill=bc)
            _text_box(slide, Inches(0.57), ry + Inches(0.17), Inches(0.36), Inches(0.36),
                      str(global_i + i + 1), font_size=12, bold=True,
                      color=C_WHITE, align=PP_ALIGN.CENTER)
            _text_box(slide, Inches(1.05), ry + Inches(0.12),
                      row_w - Inches(0.72), Inches(0.50),
                      _strip_md(action), font_size=13, color=C_GRAY_TEXT)
        _footer(slide)


def _slides_extra_sections(prs, info: dict):
    """
    Render any extra sections passed in info['sections'].
    Each section: {title, content, type}
      type = "text"    -> auto-paginated text card
      type = "bullets" -> auto-paginated bullet list
      type = "table"   -> simple 2-column key/value table
                         content = list of [key, value] pairs
    """
    sections = info.get("sections", [])
    if not sections:
        return

    for section in sections:
        title   = _strip_md(section.get("title", "附加内容"))
        content = section.get("content", "")
        stype   = section.get("type", "text")

        if stype == "text":
            text = _strip_md(content) if isinstance(content, str) else ""
            chunks = _chunk_text(text, CHARS_PER_SLIDE)
            for idx, chunk in enumerate(chunks):
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                _slide_bg(slide, C_BG_LIGHT)
                subtitle = f"({idx+1}/{len(chunks)})" if len(chunks) > 1 else ""
                _section_heading(slide, title, Inches(0.38), subtitle)
                card_x, card_y = Inches(0.4), Inches(1.0)
                card_w = W - Inches(0.8)
                card_h = H - Inches(1.7)
                _content_card(slide, card_x, card_y, card_w, card_h, C_PURPLE_MID)
                _text_box(slide, card_x + Inches(0.25), card_y + Inches(0.18),
                          card_w - Inches(0.5), card_h - Inches(0.3),
                          chunk, font_size=14, color=C_GRAY_TEXT, word_wrap=True)
                _footer(slide)

        elif stype == "bullets":
            items = content if isinstance(content, list) else [
                l.strip() for l in str(content).splitlines() if l.strip()
            ]
            chunks = _chunk_list(items, BULLETS_PER_SLIDE)
            for idx, chunk in enumerate(chunks):
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                _slide_bg(slide, C_BG_LIGHT)
                subtitle = f"({idx+1}/{len(chunks)})" if len(chunks) > 1 else ""
                _section_heading(slide, title, Inches(0.38), subtitle)
                card_x, card_y = Inches(0.4), Inches(1.0)
                card_w = W - Inches(0.8)
                card_h = H - Inches(1.7)
                _content_card(slide, card_x, card_y, card_w, card_h, C_AMBER)
                _text_box_multiline(
                    slide,
                    card_x + Inches(0.3), card_y + Inches(0.15),
                    card_w - Inches(0.5), card_h - Inches(0.25),
                    chunk, font_size=14, color=C_GRAY_TEXT,
                    bullet_color=C_AMBER
                )
                _footer(slide)

        elif stype == "table":
            rows = content if isinstance(content, list) else []
            chunks = _chunk_list(rows, 12)  # up to 12 rows per slide
            for idx, chunk in enumerate(chunks):
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                _slide_bg(slide, C_BG_LIGHT)
                subtitle = f"({idx+1}/{len(chunks)})" if len(chunks) > 1 else ""
                _section_heading(slide, title, Inches(0.38), subtitle)

                col_w = (W - Inches(0.8)) / 2
                row_h = Inches(0.48)
                start_y = Inches(1.1)
                for r_idx, row in enumerate(chunk):
                    ry = start_y + r_idx * (row_h + Inches(0.04))
                    bg = C_BG_LIGHT if r_idx % 2 == 0 else C_CARD_BG
                    _rect(slide, Inches(0.4), ry, W - Inches(0.8), row_h, fill=bg,
                          line_color=C_DIVIDER, line_width=1)
                    key_text = _strip_md(str(row[0])) if len(row) > 0 else ""
                    val_text = _strip_md(str(row[1])) if len(row) > 1 else ""
                    _text_box(slide, Inches(0.55), ry + Inches(0.06),
                              col_w - Inches(0.3), row_h - Inches(0.1),
                              key_text, font_size=12, bold=True, color=C_PURPLE_DARK)
                    _text_box(slide, Inches(0.4) + col_w + Inches(0.15), ry + Inches(0.06),
                              col_w - Inches(0.3), row_h - Inches(0.1),
                              val_text, font_size=12, color=C_GRAY_TEXT)
                _footer(slide)


def _slide_sentiment(prs, info: dict):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _slide_bg(slide, C_BG_LIGHT)
    _section_heading(slide, "情感分析", Inches(0.38))

    sentiment = (info.get("sentiment") or "neutral").lower()
    lang      = info.get("language", "")

    if "pos" in sentiment:
        emoji, label, color = "😊", "正面 / Positive", C_GREEN
    elif "neg" in sentiment:
        emoji, label, color = "😟", "负面 / Negative", C_RED
    else:
        emoji, label, color = "😐", "中性 / Neutral", C_TEAL

    card_w, card_h = Inches(7), Inches(3.8)
    card_x = (W - card_w) / 2
    card_y = Inches(1.2)
    _content_card(slide, card_x, card_y, card_w, card_h, color)

    _text_box(slide, card_x, card_y + Inches(0.2), card_w, Inches(1.2),
              emoji, font_size=60, align=PP_ALIGN.CENTER, color=color)
    _text_box(slide, card_x, card_y + Inches(1.45), card_w, Inches(0.7),
              label, font_size=28, bold=True, color=color, align=PP_ALIGN.CENTER)
    if lang:
        _text_box(slide, card_x, card_y + Inches(2.2), card_w, Inches(0.45),
                  f"语言: {lang}", font_size=14, color=C_MUTED, align=PP_ALIGN.CENTER)

    lights = [
        ("正面", "pos" in sentiment, C_GREEN),
        ("中性", sentiment in ("neutral", "中性", ""), C_TEAL),
        ("负面", "neg" in sentiment, C_RED),
    ]
    light_w, light_gap = Inches(1.8), Inches(0.4)
    lx = (W - (light_w * 3 + light_gap * 2)) / 2
    ly = card_y + card_h + Inches(0.3)
    for lbl, active, lc in lights:
        fill = lc if active else RGBColor(0xE5, 0xE7, 0xEB)
        _rect(slide, lx, ly, light_w, Inches(0.52), fill=fill,
              line_color=lc, line_width=1)
        _text_box(slide, lx, ly, light_w, Inches(0.52),
                  lbl, font_size=14, bold=active,
                  color=C_WHITE if active else C_MUTED, align=PP_ALIGN.CENTER)
        lx += light_w + light_gap
    _footer(slide)


def _slide_closing(prs, source_filename: str):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _slide_bg(slide, C_PURPLE_DARK)
    _rect(slide, W - Inches(2.8), Inches(-0.5), Inches(3.5), Inches(3.5),
          fill=RGBColor(0x3D, 0x27, 0x7B))
    _rect(slide, Inches(-0.8), H - Inches(2.2), Inches(3), Inches(3),
          fill=RGBColor(0x3D, 0x27, 0x7B))
    _text_box(slide, Inches(1.2), Inches(2.4), Inches(10.9), Inches(1.3),
              "分析完成",
              font_name=FONT_DISPLAY, font_size=52, bold=True,
              color=C_WHITE, align=PP_ALIGN.CENTER)
    _text_box(slide, Inches(1.2), Inches(3.9), Inches(10.9), Inches(0.6),
              "AI Document Analysis · file-processor",
              font_size=14, color=RGBColor(0xCC, 0xBB, 0xFF), align=PP_ALIGN.CENTER)
    ts = datetime.datetime.now().strftime("%Y-%m-%d")
    _text_box(slide, Inches(1.2), Inches(4.55), Inches(10.9), Inches(0.45),
              ts, font_size=12, color=RGBColor(0xAA, 0x99, 0xDD), align=PP_ALIGN.CENTER)


# ─── Public API ───────────────────────────────────────────────────────────────

def build_pptx(info: dict, source_filename: str = "document") -> bytes:
    """
    Build a dynamic PowerPoint report whose slide count scales with content.

    Slide count is driven entirely by the data in `info`:
      - Long summaries    → multiple summary slides
      - Many topics       → multiple topic pill slides  (PILLS_PER_SLIDE=12 each)
      - Many entities     → multiple entity/date slides (ITEMS_PER_LIST_SLIDE=10 each)
      - Many actions      → multiple action slides      (ACTIONS_PER_SLIDE=7 each)
      - Extra sections    → one or more slides each     (text/bullets/table)
      - Sentiment slide   → always 1
      - Cover + Closing   → always 1 each

    Minimum: 4 slides (Cover + Summary + Sentiment + Closing).
    Maximum: unbounded — scales with document content.

    Args:
        info            : extraction dict (see module docstring for schema)
        source_filename : original filename shown on cover slide

    Returns:
        Raw .pptx bytes — write to disk or stream via HTTP response.
    """
    prs = Presentation()
    prs.slide_width  = W
    prs.slide_height = H

    _slide_cover(prs, info, source_filename)          # always 1
    _slides_summary(prs, info)                        # 1..N
    _slides_topics(prs, info)                         # 0..N
    _slides_entities_dates(prs, info)                 # 0..N
    _slides_actions(prs, info)                        # 0..N
    _slides_extra_sections(prs, info)                 # 0..N (custom sections)
    _slide_sentiment(prs, info)                       # always 1
    _slide_closing(prs, source_filename)              # always 1

    buf = io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


# ─── CLI ─────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import json, sys

    if len(sys.argv) < 2:
        print("Usage: python export_service.py <info.json> [output.pptx]")
        sys.exit(1)

    with open(sys.argv[1], encoding="utf-8") as f:
        info = json.load(f)

    out_path = sys.argv[2] if len(sys.argv) > 2 else "report.pptx"
    raw = build_pptx(info, info.get("source_filename", "document"))
    with open(out_path, "wb") as f:
        f.write(raw)
    print(f"Saved: {out_path}  ({len(raw):,} bytes, {len(info.get('topics',[])) + len(info.get('actions',[])) + len(info.get('sections',[])) + 4} slides approx.)")