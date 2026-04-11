"""
PDF生成モジュール

ReportLab を使ったPDF組み立てロジックをすべてここに集約する。
app.py の /generate-pdf ルートから build_pdf(data) を呼び出すだけでよい。
"""
import io
import base64

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph,
    Spacer, Image as RLImage, PageBreak, Flowable,
)
from reportlab.pdfbase import pdfmetrics

from utils.font_setup import FONT_NAME, FONT_BOLD

# ── 単位定数 ──────────────────────────────────────────────────────────────────
Q13 = 13 * 0.25 * mm   # 13Q ≈ 9.21pt

# ── カラーパレット（固定3色） ──────────────────────────────────────────────────
C_WHITE  = colors.white                 # #ffffff
C_BLACK  = colors.HexColor("#1a1a1a")   # #1a1a1a
C_YELLOW = colors.HexColor("#fffeee")   # #fffeee

# デフォルトカラー（学科未指定時：機械系ピンク）
DEFAULT_MAIN = "#e5809e"
DEFAULT_SUB  = "#fbdbd6"

# 後方互換（_build_styles 内のグローバル参照用）
ACCENT     = colors.HexColor(DEFAULT_MAIN)
ACCENT_BG  = colors.HexColor(DEFAULT_SUB)
PRIMARY    = colors.HexColor(DEFAULT_MAIN)
DARK       = C_BLACK
BODY_COLOR = C_BLACK
MUTED      = C_BLACK
WHITE      = C_WHITE


# ── スタイル ──────────────────────────────────────────────────────────────────
def _build_styles():
    return {
        "title": ParagraphStyle(
            "title", fontName=FONT_NAME, fontSize=18, leading=24,
            textColor=DARK, spaceAfter=3 * mm,
        ),
        "heading": ParagraphStyle(
            "heading", fontName=FONT_BOLD, fontSize=11, leading=15,
            textColor=PRIMARY, spaceBefore=3 * mm, spaceAfter=1.5 * mm,
        ),
        "body": ParagraphStyle(
            "body", fontName=FONT_NAME, fontSize=Q13, leading=Q13 * 1.7,
            textColor=BODY_COLOR,
        ),
        "cell": ParagraphStyle(
            "cell", fontName=FONT_NAME, fontSize=Q13 * 0.9, leading=Q13 * 1.4,
            textColor=BODY_COLOR,
        ),
        "cell_header": ParagraphStyle(
            "cell_header", fontName=FONT_BOLD, fontSize=Q13 * 0.9, leading=Q13 * 1.4,
            textColor=WHITE,
        ),
        # プロフィール用
        "prof_title_main": ParagraphStyle(
            "prof_title_main", fontName=FONT_BOLD, fontSize=13, leading=17,
            textColor=DARK,
        ),
        "prof_title_sub": ParagraphStyle(
            "prof_title_sub", fontName=FONT_NAME, fontSize=Q13, leading=Q13 * 1.5,
            textColor=MUTED,
        ),
        "prof_name_en": ParagraphStyle(
            "prof_name_en", fontName=FONT_NAME, fontSize=8, leading=11,
            textColor=MUTED,
        ),
        "prof_name_ja": ParagraphStyle(
            "prof_name_ja", fontName=FONT_BOLD, fontSize=16, leading=20,
            textColor=DARK,
        ),
        "prof_label": ParagraphStyle(
            "prof_label", fontName=FONT_BOLD, fontSize=7.5, leading=10,
            textColor=WHITE,
        ),
        "prof_label_dark": ParagraphStyle(
            "prof_label_dark", fontName=FONT_BOLD, fontSize=7.5, leading=10,
            textColor=ACCENT,
        ),
        "prof_value": ParagraphStyle(
            "prof_value", fontName=FONT_NAME, fontSize=Q13, leading=Q13 * 1.5,
            textColor=BODY_COLOR,
        ),
        "prof_section_heading": ParagraphStyle(
            "prof_section_heading", fontName=FONT_BOLD, fontSize=Q13 * 0.95, leading=Q13 * 1.3,
            textColor=WHITE, alignment=1,
        ),
        "prof_section_body": ParagraphStyle(
            "prof_section_body", fontName=FONT_NAME, fontSize=Q13, leading=Q13 * 1.65,
            textColor=BODY_COLOR,
        ),
        "prof_badge": ParagraphStyle(
            "prof_badge", fontName=FONT_BOLD, fontSize=7, leading=9,
            textColor=WHITE,
        ),
        "prof_badge_val": ParagraphStyle(
            "prof_badge_val", fontName=FONT_NAME, fontSize=Q13, leading=Q13 * 1.4,
            textColor=BODY_COLOR,
        ),
    }


# ── テーブル ──────────────────────────────────────────────────────────────────
def _build_table(rows, has_header, styles):
    col_count  = max(len(r) for r in rows) if rows else 1
    page_width = A4[0] - 30 * mm
    col_width  = page_width / col_count

    table_data = []
    for i, row in enumerate(rows):
        padded     = row + [""] * (col_count - len(row))
        is_header  = has_header and i == 0
        cell_style = styles["cell_header"] if is_header else styles["cell"]
        table_data.append([Paragraph(str(c), cell_style) for c in padded])

    tbl = Table(table_data, colWidths=[col_width] * col_count, repeatRows=1 if has_header else 0)
    cmd = [
        ("BACKGROUND",    (0, 0), (-1, 0),  PRIMARY if has_header else colors.HexColor("#F8FAFF")),
        ("ROWBACKGROUNDS",(0, 1), (-1, -1), [WHITE, colors.HexColor("#F5F3FF")]),
        ("GRID",          (0, 0), (-1, -1), 0.5, colors.HexColor("#E0E7FF")),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING",   (0, 0), (-1, -1), 6),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 6),
    ]
    tbl.setStyle(TableStyle(cmd))
    return tbl


# ── 画像ユーティリティ ─────────────────────────────────────────────────────────
def _b64_to_image(data_url, max_w, max_h=None):
    """base64 データURLを ReportLab Image に変換（アスペクト比保持）"""
    if not data_url:
        return None
    try:
        _, b64 = data_url.split(",", 1)
        raw = base64.b64decode(b64)
        buf = io.BytesIO(raw)
        img = RLImage(buf)
        iw, ih = img.imageWidth, img.imageHeight
        scale = max_w / iw
        new_h = ih * scale
        if max_h and new_h > max_h:
            scale = max_h / ih
            new_w = iw * scale
            new_h = max_h
        else:
            new_w = max_w
        buf.seek(0)
        return RLImage(buf, width=new_w, height=new_h)
    except Exception as e:
        print(f"[image] load error: {e}")
        return None


# ── Flowable クラス群 ─────────────────────────────────────────────────────────
class _SectionsCard(Flowable):
    """全セクションをひとつの角丸カードにまとめる Flowable"""
    _HEAD_W    = 216
    _HEAD_PAD  = 4
    _HEAD_LPAD = 8
    _INNER_GAP = 3 * mm
    _OUTER_GAP = 4 * mm

    def __init__(self, sections_data, card_w,
                 heading_style, body_style,
                 sec_bg, heading_bg,
                 radius=6 * mm, pad_x=11, pad_tb=11, fixed_h=None):
        Flowable.__init__(self)
        self.sections_data = sections_data
        self.card_w        = card_w
        self.hstyle        = heading_style
        self.bstyle        = body_style
        self.sec_bg        = sec_bg
        self.heading_bg    = heading_bg
        self.radius        = radius
        self.pad_x         = pad_x
        self.pad_tb        = pad_tb
        self.fixed_h       = fixed_h
        self._metrics      = []

    def wrap(self, availWidth, availHeight):
        w       = self.card_w
        inner_w = w - 2 * self.pad_x
        n       = len(self.sections_data)

        self._metrics = []
        total_h = self.pad_tb

        for i, (heading, body_text) in enumerate(self.sections_data):
            hp = Paragraph(heading, self.hstyle)
            _, th = hp.wrap(self._HEAD_W, availHeight)
            hh = th + 2 * self._HEAD_PAD

            bp = Paragraph(body_text, self.bstyle)
            _, bh = bp.wrap(inner_w, availHeight)

            self._metrics.append((hh, bh))
            total_h += hh + self._INNER_GAP + bh
            if i < n - 1:
                total_h += self._OUTER_GAP

        total_h += self.pad_tb
        self.height = self.fixed_h if self.fixed_h is not None else total_h
        return w, self.height

    def draw(self):
        c = self.canv
        w = self.card_w
        h = self.height
        r = self.radius
        n = len(self.sections_data)

        c.saveState()
        c.setFillColor(self.sec_bg)
        c.roundRect(0, 0, w, h, r, fill=1, stroke=0)
        c.restoreState()

        inner_w   = w - 2 * self.pad_x
        current_y = h - self.pad_tb

        for i, (heading, body_text) in enumerate(self.sections_data):
            hh, bh = self._metrics[i]
            capsule_r = hh / 2

            c.saveState()
            c.setFillColor(self.heading_bg)
            c.roundRect(self.pad_x, current_y - hh,
                        self._HEAD_W, hh, capsule_r, fill=1, stroke=0)
            c.restoreState()

            hp = Paragraph(heading, self.hstyle)
            hp.wrap(self._HEAD_W, hh)
            hp.drawOn(c, self.pad_x, current_y - hh + self._HEAD_PAD)

            current_y -= hh + self._INNER_GAP

            bp = Paragraph(body_text, self.bstyle)
            bp.wrap(inner_w, bh + 200)
            bp.drawOn(c, self.pad_x, current_y - bh)

            current_y -= bh
            if i < n - 1:
                current_y -= self._OUTER_GAP


class _KeywordCard(Flowable):
    """キーワード用カード：白い角丸背景＋ピンクバッジ"""
    _BADGE_FS    = 7.5
    _BADGE_LEAD  = 10
    _BADGE_PAD_X = 8
    _BADGE_PAD_Y = 3.6
    _BADGE_R     = 4.2

    def __init__(self, keywords_text, badge_color, body_style,
                 card_radius=8, pad_x=12, pad_tb=8, gap=-6):
        Flowable.__init__(self)
        self.keywords_text = keywords_text
        self.badge_color   = badge_color
        self.body_style    = body_style
        self.card_radius   = card_radius
        self.pad_x         = pad_x
        self.pad_tb        = pad_tb
        self.gap           = gap

    def wrap(self, availWidth, availHeight):
        self._card_w = availWidth

        btw = pdfmetrics.stringWidth("キーワード", FONT_BOLD, self._BADGE_FS)
        self._bw = btw + 2 * self._BADGE_PAD_X
        self._bh = self._BADGE_LEAD + 2 * self._BADGE_PAD_Y

        self._cx     = self.pad_x
        self._cw     = availWidth - self.pad_x - 16
        self._text_x = self.pad_x + self._BADGE_PAD_X
        inner_w      = availWidth - self._text_x - self.pad_x
        bp = Paragraph(self.keywords_text, self.body_style)
        _, self._body_h = bp.wrap(inner_w, availHeight)
        self._card_h = self.pad_tb + self._body_h + self.pad_tb

        self.height = self._bh + self.gap + self._card_h
        return availWidth, self.height

    def draw(self):
        c = self.canv

        c.saveState()
        c.setFillColor(colors.white)
        c.roundRect(self._cx, 0, self._cw, self._card_h,
                    self.card_radius, fill=1, stroke=0)
        c.restoreState()

        bw, bh = self._bw, self._bh
        badge_x = self.pad_x
        badge_y = self._card_h + self.gap
        c.saveState()
        c.setFillColor(self.badge_color)
        c.roundRect(badge_x, badge_y, bw, bh, self._BADGE_R, fill=1, stroke=0)
        c.setFont(FONT_BOLD, self._BADGE_FS)
        c.setFillColor(colors.white)
        text_y = badge_y + (bh - self._BADGE_FS) / 2
        c.drawString(badge_x + self._BADGE_PAD_X, text_y, "キーワード")
        c.restoreState()

        inner_w = self._card_w - self._text_x - self.pad_x
        bp = Paragraph(self.keywords_text, self.body_style)
        bp.wrap(inner_w, self._body_h + 100)
        bp.drawOn(c, self._text_x, self.pad_tb)


class _WhiteCard(Flowable):
    """テキストを白い角丸カードで囲む Flowable（キャッチコピー用）"""
    def __init__(self, text, style, card_radius=8, pad_x=14, pad_tb=10,
                 offset_x=0, fixed_w=None):
        Flowable.__init__(self)
        self.text        = text
        self.style       = style
        self.card_radius = card_radius
        self.pad_x       = pad_x
        self.pad_tb      = pad_tb
        self.offset_x    = offset_x
        self.fixed_w     = fixed_w

    def wrap(self, availWidth, availHeight):
        self._avail = availWidth
        self._cw    = self.fixed_w if self.fixed_w is not None else availWidth
        inner_w     = self._cw - 2 * self.pad_x
        p = Paragraph(self.text, self.style)
        _, self._ph = p.wrap(inner_w, availHeight)
        self.height = self.pad_tb + self._ph + self.pad_tb
        return availWidth, self.height

    def draw(self):
        c = self.canv
        ox = self.offset_x
        cw = self._cw
        h  = self.height
        r  = self.card_radius

        c.saveState()
        c.setFillColor(colors.white)
        c.roundRect(ox, 0, cw, h, r, fill=1, stroke=0)
        c.rect(ox + cw - r, 0, r, r, fill=1, stroke=0)
        c.restoreState()

        inner_w = cw - 2 * self.pad_x
        p = Paragraph(self.text, self.style)
        p.wrap(inner_w, self._ph + 100)
        p.drawOn(c, ox + self.pad_x, self.pad_tb)


class _FaceCard(Flowable):
    """顔写真を角丸カード（メインカラー背景）で囲む Flowable"""
    def __init__(self, image_el, img_w, img_h, pad=12,
                 outer_radius=14, inner_radius=12,
                 color=None, gray_color=None):
        Flowable.__init__(self)
        self.image_el     = image_el
        self.img_w        = img_w
        self.img_h        = img_h
        self.pad          = pad
        self.outer_radius = outer_radius
        self.inner_radius = inner_radius
        self.color        = color
        self.gray_color   = gray_color
        self.width        = img_w + 2 * pad
        self.height       = img_h + 2 * pad

    def wrap(self, availWidth, availHeight):
        return self.width, self.height

    @staticmethod
    def _make_rounded_path(c, x, y, w, h, r):
        k = 0.5523
        p = c.beginPath()
        p.moveTo(x + r, y)
        p.lineTo(x + w - r, y)
        p.curveTo(x + w - r*(1-k), y,           x + w, y + r*(1-k),     x + w, y + r)
        p.lineTo(x + w, y + h - r)
        p.curveTo(x + w, y + h - r*(1-k),       x + w - r*(1-k), y + h, x + w - r, y + h)
        p.lineTo(x + r, y + h)
        p.curveTo(x + r*(1-k), y + h,           x, y + h - r*(1-k),     x, y + h - r)
        p.lineTo(x, y + r)
        p.curveTo(x, y + r*(1-k),               x + r*(1-k), y,         x + r, y)
        p.close()
        return p

    def draw(self):
        c = self.canv
        c.saveState()
        c.setFillColor(self.color)
        c.roundRect(0, 0, self.width, self.height, self.outer_radius, fill=1, stroke=0)
        c.restoreState()
        c.saveState()
        p = self._make_rounded_path(c, self.pad, self.pad,
                                    self.img_w, self.img_h, self.inner_radius)
        c.clipPath(p, stroke=0, fill=0)
        if self.gray_color is not None:
            c.setFillColor(self.gray_color)
            c.rect(self.pad, self.pad, self.img_w, self.img_h, fill=1, stroke=0)
        else:
            self.image_el.wrap(self.img_w, self.img_h)
            self.image_el.drawOn(c, self.pad, self.pad)
        c.restoreState()


class _PillFlowable(Flowable):
    """カプセル型バッジ Flowable"""
    def __init__(self, text, bg_color, style, pad_x=12, pad_y=3.6,
                 radius=None, stroke=True):
        Flowable.__init__(self)
        self.text     = text
        self.bg_color = bg_color
        self.style    = style
        self.pad_x    = pad_x
        self.pad_y    = pad_y
        self.radius   = radius
        self.stroke   = stroke
        self._tw      = 0
        self._th      = 0

    def wrap(self, availWidth, availHeight):
        self._tw    = pdfmetrics.stringWidth(self.text, self.style.fontName, self.style.fontSize)
        self._th    = self.style.leading
        self.width  = self._tw + 2 * self.pad_x
        self.height = self._th + 2 * self.pad_y
        return self.width, self.height

    def draw(self):
        c = self.canv
        w, h = self.width, self.height
        r = self.radius if self.radius is not None else h / 2

        c.saveState()
        c.setFillColor(self.bg_color)
        if self.stroke:
            c.setStrokeColor(colors.white)
            c.setLineWidth(1)
        c.roundRect(0, 0, w, h, r, fill=1, stroke=1 if self.stroke else 0)
        c.restoreState()

        c.saveState()
        c.setFont(self.style.fontName, self.style.fontSize)
        c.setFillColor(self.style.textColor)
        text_y = (h - self.style.fontSize) / 2
        c.drawString(self.pad_x, text_y, self.text)
        c.restoreState()


def _pill(text, color, radius=None, stroke=True):
    style = ParagraphStyle("pl", fontName=FONT_BOLD, fontSize=7.5,
                            leading=10, textColor=WHITE)
    return _PillFlowable(text, color, style, pad_x=12, pad_y=3.6,
                         radius=radius, stroke=stroke)


# ── プロフィールレイアウト ─────────────────────────────────────────────────────
def _build_profile(block, styles, doc):
    pw     = A4[0] - 32 * mm
    CARD_W = 362
    lw     = CARD_W + 2 * mm
    rw     = pw - lw

    _main_hex = block.get("mainColor", DEFAULT_MAIN)
    _sub_hex  = block.get("subColor",  DEFAULT_SUB)
    PINK    = colors.HexColor(_main_hex)
    PINK_BG = colors.HexColor(_sub_hex)
    PINK_HD = colors.HexColor(_main_hex)
    GRAY_PH = C_WHITE

    story = []

    catchcopy       = block.get("catchcopy",      "")
    name_en         = block.get("nameEn",         "")
    name_ja         = block.get("nameJa",         "")
    keywords        = block.get("keywords",       "")
    field_name      = block.get("fieldName",      "")
    email           = block.get("email",          "")
    email_label     = block.get("emailLabel",     "email")
    position        = block.get("position",       "")
    degree          = block.get("degree",         "")
    sections        = block.get("sections",       [])
    face_photo_data = block.get("facePhoto")
    research_images = block.get("researchImages", [])

    cp_s = ParagraphStyle("cp", fontName=FONT_BOLD, fontSize=14, leading=20,
                           textColor=colors.HexColor("#1a1a1a"))
    kw_s = ParagraphStyle("kw", fontName=FONT_NAME, fontSize=Q13, leading=Q13 * 1.4,
                           textColor=BODY_COLOR)

    def _meta_row(badge_txt, val, radius=None, stroke=True):
        pill_style = ParagraphStyle("_tmp", fontName=FONT_BOLD, fontSize=7.5)
        text_w = pdfmetrics.stringWidth(badge_txt, pill_style.fontName, pill_style.fontSize)
        bw = text_w + 24 + 2
        BADGE_OFFSET = 12
        pill = _pill(badge_txt, PINK, radius=radius, stroke=stroke)
        t = Table([[pill, Paragraph(val, kw_s)]],
                  colWidths=[bw + BADGE_OFFSET, lw - bw - BADGE_OFFSET - 4 * mm])
        t.setStyle(TableStyle([
            ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING",   (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING",(0, 0), (-1, -1), 2),
            ("LEFTPADDING",  (0, 0), (-1, -1), 0),
            ("LEFTPADDING",  (0, 0), (0,  0),  BADGE_OFFSET),
            ("LEFTPADDING",  (1, 0), (1,  0),  8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ]))
        return t

    left_hdr = []
    if keywords:
        left_hdr.append(Spacer(1, 56))
        left_hdr.append(_KeywordCard(keywords, PINK, kw_s))
        left_hdr.append(Spacer(1, 8))
    for badge_txt, val in [("分野等", field_name), (email_label, email)]:
        if val:
            left_hdr.append(_meta_row(badge_txt, val))
            left_hdr.append(Spacer(1, 1.5 * mm))

    face_w      = 104
    face_h      = face_w * 4 / 3
    face_card_h = face_h + 2 * 12
    right_hdr   = [Spacer(face_w + 24, face_card_h)]

    hdr_tbl = Table([[left_hdr, right_hdr]], colWidths=[lw, rw])
    hdr_tbl.setStyle(TableStyle([
        ("VALIGN",       (0, 0), (-1, -1), "TOP"),
        ("ALIGN",        (1, 0), (1,  0),  "RIGHT"),
        ("TOPPADDING",   (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 0),
        ("LEFTPADDING",  (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
    ]))
    CARD_R = 14
    story.append(hdr_tbl)
    story.append(Spacer(1, 0))

    body_s = ParagraphStyle("bs", fontName=FONT_NAME, fontSize=Q13, leading=Q13 * 1.65,
                             textColor=BODY_COLOR)
    sec_hs = ParagraphStyle("sh", fontName=FONT_BOLD, fontSize=Q13, leading=Q13 * 1.4,
                             textColor=WHITE, alignment=1)
    SEC_BG = C_YELLOW

    sec_pairs = [
        (sec.get("heading", ""), sec.get("content", "").strip().replace("\n", "<br/>"))
        for sec in sections
        if sec.get("content", "").strip()
    ]

    left_body = []
    if sec_pairs:
        sections_card = _SectionsCard(
            sections_data = sec_pairs,
            card_w        = CARD_W,
            heading_style = sec_hs,
            body_style    = body_s,
            sec_bg        = SEC_BG,
            heading_bg    = PINK_HD,
            radius        = CARD_R,
            pad_x         = 11,
            pad_tb        = 11,
            fixed_h       = int(A4[1] - 2 * 11 * mm - face_h - 24 - 4 - 40),
        )
        left_body.append(sections_card)

    Q11 = 11 * 0.25 * mm
    bdg_pill_style = ParagraphStyle("bdgp", fontName=FONT_BOLD, fontSize=Q11,
                                    leading=Q11 * 1.4, textColor=WHITE)
    bdg_v = ParagraphStyle("bv", fontName=FONT_BOLD, fontSize=Q13, leading=Q13 * 1.4,
                            textColor=BODY_COLOR)
    right_body = [
        Paragraph(name_en,
                  ParagraphStyle("ne", fontName=FONT_NAME, fontSize=9, leading=12,
                                 textColor=colors.HexColor("#1a1a1a"))),
        Paragraph(f"<b>{name_ja}</b>",
                  ParagraphStyle("nj", fontName=FONT_BOLD, fontSize=18, leading=22,
                                 textColor=colors.HexColor("#1a1a1a"))),
        Spacer(1, 3 * mm),
    ]
    for lbl, val in [("職名", position), ("学位", degree)]:
        if not val:
            continue
        text_w = pdfmetrics.stringWidth(lbl, FONT_BOLD, Q11)
        bw = text_w + 2 * 8 + 2
        pill = _PillFlowable(lbl, PINK, bdg_pill_style, pad_x=8, pad_y=3)
        bt = Table(
            [[pill, Paragraph(val, bdg_v)]],
            colWidths=[bw, rw - bw - 2 * mm],
        )
        bt.setStyle(TableStyle([
            ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING",   (0, 0), (-1, -1), 2),
            ("BOTTOMPADDING",(0, 0), (-1, -1), 2),
            ("LEFTPADDING",  (0, 0), (-1, -1), 0),
            ("LEFTPADDING",  (1, 0), (1,  0),  8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ]))
        right_body.append(bt)
        right_body.append(Spacer(1, 1.5 * mm))

    Q10 = 10 * 0.25 * mm
    cap_s = ParagraphStyle("cap", fontName=FONT_NAME, fontSize=Q10, leading=Q10 * 1.4,
                             textColor=colors.HexColor("#1a1a1a"), alignment=0)
    RI_W = 112

    if research_images:
        for ri in research_images:
            img_el = _b64_to_image(ri.get("data"), RI_W)
            if img_el:
                right_body.append(img_el)
            else:
                ph = Table([[""]], colWidths=[RI_W], rowHeights=[RI_W * 3 // 4])
                ph.setStyle(TableStyle([("BACKGROUND", (0,0),(0,0), GRAY_PH),
                                        ("TOPPADDING", (0,0),(0,0), 0),
                                        ("BOTTOMPADDING",(0,0),(0,0), 0)]))
                right_body.append(ph)
            cap_text = ri.get("name", "")
            if cap_text:
                right_body.append(Spacer(1, 0.5 * mm))
                right_body.append(Paragraph(cap_text, cap_s))
            right_body.append(Spacer(1, 2.5 * mm))
    else:
        for _ in range(4):
            ph = Table([[""]], colWidths=[RI_W], rowHeights=[RI_W * 3 // 4])
            ph.setStyle(TableStyle([("BACKGROUND", (0,0),(0,0), GRAY_PH),
                                    ("TOPPADDING", (0,0),(0,0), 0),
                                    ("BOTTOMPADDING",(0,0),(0,0), 0)]))
            right_body.append(ph)
            right_body.append(Spacer(1, 3 * mm))

    body_tbl = Table([[left_body, right_body]], colWidths=[lw, rw])
    body_tbl.setStyle(TableStyle([
        ("VALIGN",       (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING",   (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 0),
        ("LEFTPADDING",  (0, 0), (-1, -1), 0),
        ("LEFTPADDING",  (1, 0), (1,  0),  12),
        ("RIGHTPADDING", (0, 0), (0,  0),  3 * mm),
        ("RIGHTPADDING", (1, 0), (1,  0),  0),
    ]))
    story.append(body_tbl)
    return story


# ── PDF組み立てエントリポイント ────────────────────────────────────────────────
def build_pdf(data):
    """
    JSON データを受け取り、PDF を io.BytesIO で返す。
    app.py の /generate-pdf ルートから呼び出す。
    """
    title  = data.get("title", "レポート")
    blocks = data.get("blocks", [])

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=16 * mm,
        rightMargin=16 * mm,
        topMargin=11 * mm,
        bottomMargin=11 * mm,
    )

    styles = _build_styles()
    story  = []

    nombres_list   = []
    catchcopy_list = []
    face_list      = []
    page_colors    = []
    first_profile  = True

    for block in blocks:
        btype = block.get("type")

        if btype == "heading":
            story.append(Paragraph(block.get("content", ""), styles["heading"]))
            story.append(Spacer(1, 2 * mm))

        elif btype == "text":
            story.append(Paragraph(
                block.get("content", "").replace("\n", "<br/>"), styles["body"]
            ))
            story.append(Spacer(1, 3 * mm))

        elif btype == "table":
            rows       = block.get("rows", [])
            has_header = block.get("hasHeader", True)
            if rows:
                story.append(_build_table(rows, has_header, styles))
                story.append(Spacer(1, 4 * mm))

        elif btype == "spacer":
            story.append(Spacer(1, 6 * mm))

        elif btype == "profile":
            if not first_profile:
                story.append(PageBreak())
            nombres_list.append(str(block.get("pageNum", "") or ""))
            catchcopy_list.append(block.get("catchcopy", "") or "")
            face_list.append(block.get("facePhoto"))
            page_colors.append({
                "main": block.get("mainColor", DEFAULT_MAIN),
                "sub":  block.get("subColor",  DEFAULT_SUB),
            })
            story += _build_profile(block, styles, doc)
            first_profile = False

    def _draw_nombre(canvas, doc):
        pw, ph = A4
        mx, my    = 16 * mm, 11 * mm
        bg_radius = 6 * mm

        page = canvas.getPageNumber()
        _pc    = page_colors[page - 1] if 1 <= page <= len(page_colors) else {"main": DEFAULT_MAIN, "sub": DEFAULT_SUB}
        C_MAIN = colors.HexColor(_pc["main"])
        C_SUB  = colors.HexColor(_pc["sub"])

        # ① ページ背景
        canvas.saveState()
        canvas.setFillColor(C_SUB)
        canvas.roundRect(mx, my, pw - 2 * mx, ph - 2 * my,
                         radius=bg_radius, fill=1, stroke=0)
        canvas.restoreState()

        # ② キャッチコピー白カード
        if 1 <= page <= len(catchcopy_list):
            cp_text = catchcopy_list[page - 1]
            if cp_text:
                cp_fs  = 18
                pad_x  = 12
                pad_tb = 8
                card_r = 14
                cp_style = ParagraphStyle("_cp", fontName=FONT_BOLD, fontSize=cp_fs,
                                          leading=cp_fs * 1.5, textColor=C_BLACK)
                p = Paragraph(cp_text.replace("\n", "<br/>"), cp_style)
                card_w = 353
                _, text_h = p.wrap(card_w - 2 * pad_x, my - 4)
                card_h = text_h + 2 * pad_tb

                canvas.saveState()
                canvas.setFillColor(C_WHITE)
                card_base_y = ph - my + (my - card_h) / 2 - 32
                canvas.roundRect(mx, card_base_y, card_w, card_h, card_r, fill=1, stroke=0)
                canvas.rect(mx, card_base_y, card_r, card_r, fill=1, stroke=0)
                canvas.restoreState()

                p.wrap(card_w - 2 * pad_x, text_h + 10)
                canvas.saveState()
                p.drawOn(canvas, mx + pad_x, card_base_y + pad_tb)
                canvas.restoreState()

        # ③ 顔写真カード
        if 1 <= page <= len(face_list):
            _face_w  = 104
            _face_h  = _face_w * 4 / 3
            _pad     = 12
            _outer_r = 14
            _inner_r = 10
            _card_w  = _face_w + 2 * _pad
            _card_h  = _face_h + 2 * _pad
            _cx = pw - mx - _card_w
            _cy = ph - my - _card_h

            canvas.saveState()
            canvas.setFillColor(C_MAIN)
            canvas.roundRect(_cx, _cy, _card_w, _card_h, _outer_r, fill=1, stroke=0)
            canvas.restoreState()

            _k = 0.5523
            _ix, _iy, _iw, _ih = _cx + _pad, _cy + _pad, _face_w, _face_h
            _r = _inner_r
            canvas.saveState()
            _p = canvas.beginPath()
            _p.moveTo(_ix + _r, _iy)
            _p.lineTo(_ix + _iw - _r, _iy)
            _p.curveTo(_ix + _iw - _r*(1-_k), _iy,           _ix + _iw, _iy + _r*(1-_k),       _ix + _iw, _iy + _r)
            _p.lineTo(_ix + _iw, _iy + _ih - _r)
            _p.curveTo(_ix + _iw, _iy + _ih - _r*(1-_k),     _ix + _iw - _r*(1-_k), _iy + _ih, _ix + _iw - _r, _iy + _ih)
            _p.lineTo(_ix + _r, _iy + _ih)
            _p.curveTo(_ix + _r*(1-_k), _iy + _ih,           _ix, _iy + _ih - _r*(1-_k),        _ix, _iy + _ih - _r)
            _p.lineTo(_ix, _iy + _r)
            _p.curveTo(_ix, _iy + _r*(1-_k),                 _ix + _r*(1-_k), _iy,               _ix + _r, _iy)
            _p.close()
            canvas.clipPath(_p, stroke=0, fill=0)
            _photo = face_list[page - 1]
            if _photo:
                _img = _b64_to_image(_photo, _face_w, _face_h)
                if _img:
                    _img.wrap(_face_w, _face_h)
                    _img.drawOn(canvas, _ix, _iy)
                else:
                    canvas.setFillColor(C_WHITE)
                    canvas.rect(_ix, _iy, _iw, _ih, fill=1, stroke=0)
            else:
                canvas.setFillColor(C_WHITE)
                canvas.rect(_ix, _iy, _iw, _ih, fill=1, stroke=0)
            canvas.restoreState()

        # ④ ノンブル
        if 1 <= page <= len(nombres_list):
            nombre = nombres_list[page - 1]
            if nombre:
                try:
                    num = int(nombre)
                except ValueError:
                    num = 0
                CIRC_R = 10
                cy = 15.6 + CIRC_R
                cx = 14 + CIRC_R if num % 2 == 0 else pw - 14 - CIRC_R

                canvas.saveState()
                canvas.setFillColor(C_MAIN)
                canvas.circle(cx, cy, CIRC_R, fill=1, stroke=0)
                canvas.setFont(FONT_NAME, Q13)
                canvas.setFillColor(C_WHITE)
                canvas.drawCentredString(cx, cy - Q13 * 0.35, nombre)
                canvas.restoreState()

    doc.build(story, onFirstPage=_draw_nombre, onLaterPages=_draw_nombre)
    buf.seek(0)
    return buf
