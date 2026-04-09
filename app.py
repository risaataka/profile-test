import io
import os
import base64
import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph,
    Spacer, HRFlowable, KeepTogether, Image as RLImage, PageBreak, Flowable
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 128 * 1024 * 1024  # 128MB（画像含むJSON対応）

# ── フォント設定（日本語対応）──────────────────────────────────────────────────
FONT_NAME = "Helvetica"
FONT_BOLD = "Helvetica-Bold"

# macOS の TTC フォントは subfontIndex=0 が必要
SYSTEM_FONTS = [
    ("/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc", "HiraginoW3"),
    ("/System/Library/Fonts/ヒラギノ角ゴシック W6.ttc", "HiraginoW6"),
    ("/Library/Fonts/ヒラギノ角ゴ ProN W3.ttc",         "HiraginoProN"),
    ("/Library/Fonts/Arial Unicode MS.ttf",              "ArialUnicode"),
    ("/System/Library/Fonts/Supplemental/Arial Unicode MS.ttf", "ArialUnicode"),
]

for font_path, font_alias in SYSTEM_FONTS:
    if os.path.exists(font_path):
        try:
            kwargs = {"subfontIndex": 0} if font_path.endswith(".ttc") else {}
            pdfmetrics.registerFont(TTFont(font_alias, font_path, **kwargs))
            FONT_NAME = font_alias
            FONT_BOLD = font_alias
            print(f"[font] loaded: {font_path}")
            break
        except Exception as e:
            print(f"[font] failed {font_path}: {e}")
            continue

# システムフォントが見つからない場合は ReportLab 内蔵 CID フォントを使用
if FONT_NAME == "Helvetica":
    try:
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
        pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
        FONT_NAME = "HeiseiKakuGo-W5"
        FONT_BOLD = "HeiseiKakuGo-W5"
        print("[font] using built-in CID font: HeiseiKakuGo-W5")
    except Exception as e:
        print(f"[font] CID font failed: {e}. Japanese may not render.")

# 13Q = 13 × 0.25mm = 3.25mm ≈ 9.21pt
Q13 = 13 * 0.25 * mm  # ポイント単位

# アクセントカラー（ピンク系）
ACCENT     = colors.HexColor("#E8567C")
ACCENT_BG  = colors.HexColor("#FDF2F5")
PRIMARY    = colors.HexColor("#4F46E5")
DARK       = colors.HexColor("#1E1B4B")
BODY_COLOR = colors.HexColor("#1F2937")
MUTED      = colors.HexColor("#6B7280")
WHITE      = colors.white


# ── ルート ────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "ファイルが見つかりません"}), 400
    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "ファイル名が空です"}), 400
    try:
        xl = pd.ExcelFile(file)
        sheets = {}
        for sheet_name in xl.sheet_names:
            df = xl.parse(sheet_name, header=None)
            df = df.fillna("")
            sheets[sheet_name] = df.astype(str).values.tolist()
        return jsonify({"sheets": sheets, "sheetNames": xl.sheet_names})
    except Exception as e:
        return jsonify({"error": f"読み込みエラー: {str(e)}"}), 500


@app.route("/generate-pdf", methods=["POST"])
def generate_pdf():
    data = request.get_json()
    if not data:
        return jsonify({"error": "データがありません"}), 400

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

    # ノンブル用：プロフィールの順番とページ番号文字列を収集
    # （プロフィール1件=1PDFページとして、PageBreakで分離）
    nombres_list = []   # PDF物理ページ順のノンブル文字列
    first_profile = True

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
                story.append(PageBreak())   # プロフィールを1ページ1件に分離
            nombres_list.append(str(block.get("pageNum", "") or ""))
            story += _build_profile(block, styles, doc)
            first_profile = False

    # ── ページ背景＋ノンブル描画コールバック ─────────────────────────────────────
    def _draw_nombre(canvas, doc):
        pw, ph = A4
        mx, my = 16 * mm, 11 * mm    # ページ背景のマージン
        bg_radius = 6 * mm           # 角丸半径（≈24px）

        # ① ページ背景：角丸の塗り（#fbdbd6）
        canvas.saveState()
        canvas.setFillColor(colors.HexColor("#fbdbd6"))
        canvas.roundRect(mx, my, pw - 2 * mx, ph - 2 * my,
                         radius=bg_radius, fill=1, stroke=0)
        canvas.restoreState()

        # ② ノンブル（ピンク円バッジ、ピンク背景の外側）
        page = canvas.getPageNumber()
        if 1 <= page <= len(nombres_list):
            nombre = nombres_list[page - 1]
            if nombre:
                try:
                    num = int(nombre)
                except ValueError:
                    num = 0
                CIRC_R = 10                          # 半径 10pt（直径 20pt ≒ 20px）
                cy     = 15.6 + CIRC_R               # 円の下端がページ下端から 15.6pt
                if num % 2 == 0:                     # 偶数：円の左端がページ左端から 14pt
                    cx = 14 + CIRC_R
                else:                                # 奇数：円の右端がページ右端から 14pt
                    cx = pw - 14 - CIRC_R

                # ピンク円
                canvas.saveState()
                canvas.setFillColor(colors.HexColor("#E8567C"))
                canvas.circle(cx, cy, CIRC_R, fill=1, stroke=0)

                # ノンブルテキスト（白・円の中央）
                canvas.setFont(FONT_NAME, Q13)
                canvas.setFillColor(colors.white)
                text_y = cy - Q13 * 0.35   # ベースラインを縦中央に合わせる
                canvas.drawCentredString(cx, text_y, nombre)
                canvas.restoreState()

    doc.build(story, onFirstPage=_draw_nombre, onLaterPages=_draw_nombre)
    buf.seek(0)
    return send_file(
        buf,
        mimetype="application/pdf",
        as_attachment=False,   # インラインで返してiframeに表示
        download_name="output.pdf",
    )


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
        padded    = row + [""] * (col_count - len(row))
        is_header = has_header and i == 0
        cell_style = styles["cell_header"] if is_header else styles["cell"]
        table_data.append([Paragraph(str(c), cell_style) for c in padded])

    tbl = Table(table_data, colWidths=[col_width] * col_count, repeatRows=1 if has_header else 0)
    cmd = [
        ("BACKGROUND", (0, 0), (-1, 0), PRIMARY if has_header else colors.HexColor("#F8FAFF")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, colors.HexColor("#F5F3FF")]),
        ("GRID",        (0, 0), (-1, -1), 0.5, colors.HexColor("#E0E7FF")),
        ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",  (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING",(0, 0),(-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING",(0, 0), (-1, -1), 6),
    ]
    tbl.setStyle(TableStyle(cmd))
    return tbl


def _b64_to_image(data_url, max_w, max_h=None):
    """base64データURLをReportLab Imageに変換（アスペクト比保持）"""
    if not data_url:
        return None
    try:
        _, b64 = data_url.split(',', 1)
        raw = base64.b64decode(b64)
        buf = io.BytesIO(raw)
        # アスペクト比を保ってmax_wに収める
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


# ── プロフィールレイアウト（3:1カラム）────────────────────────────────────────
class _SectionsCard(Flowable):
    """
    全セクションをひとつの角丸カードにまとめる Flowable
    - 黄色（sec_bg）の角丸背景が全セクションを一括で覆う（上角丸）
    - body_tbl の左セルも同色で塗ることで下方向へ伸張
    - 各セクション見出しは固定幅216pt・カプセル型（border-radius=高さ/2）・左揃え
    """
    _HEAD_W    = 216     # pt: 見出しカプセルの固定幅
    _HEAD_PAD  = 4       # pt: 見出し上下内側パディング
    _HEAD_LPAD = 8       # pt: 見出しテキスト左インデント（カプセル内）
    _INNER_GAP = 3 * mm  # 見出し〜本文の間隔
    _OUTER_GAP = 4 * mm  # セクション間の余白

    def __init__(self, sections_data, card_w,
                 heading_style, body_style,
                 sec_bg, heading_bg,
                 radius=6 * mm, pad_x=11, pad_tb=11, fixed_h=None):
        """
        sections_data: [(heading_str, body_str), ...]  空セクションは除外済み
        pad_x / pad_tb: pt 単位（11pt ≒ 11px @ 72dpi）
        fixed_h: 指定した場合はその高さに固定（pt）
        """
        Flowable.__init__(self)
        self.sections_data = sections_data
        self.card_w        = card_w
        self.hstyle        = heading_style
        self.bstyle        = body_style
        self.sec_bg        = sec_bg
        self.heading_bg    = heading_bg
        self.radius        = radius
        self.pad_x         = pad_x    # pt
        self.pad_tb        = pad_tb   # pt
        self.fixed_h       = fixed_h  # pt（None = 動的計算）
        self._metrics      = []       # [(hh, bh), ...]

    def wrap(self, availWidth, availHeight):
        w       = self.card_w
        inner_w = w - 2 * self.pad_x
        n       = len(self.sections_data)

        self._metrics = []
        total_h = self.pad_tb   # 上端パディング

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

        total_h += self.pad_tb   # 下端パディング
        self.height = self.fixed_h if self.fixed_h is not None else total_h
        return w, self.height

    def draw(self):
        c = self.canv
        w = self.card_w
        h = self.height
        r = self.radius
        n = len(self.sections_data)

        # ① 黄色の角丸背景（上角丸のみ、下は body_tbl 左セル背景色が続く）
        c.saveState()
        c.setFillColor(self.sec_bg)
        c.roundRect(0, 0, w, h, r, fill=1, stroke=0)
        c.restoreState()

        inner_w   = w - 2 * self.pad_x
        current_y = h - self.pad_tb   # 上端から描画開始

        for i, (heading, body_text) in enumerate(self.sections_data):
            hh, bh = self._metrics[i]
            capsule_r = hh / 2   # 完全カプセル（999px 相当）

            # ② ピンク見出しカプセル（固定幅216pt・左揃え）
            c.saveState()
            c.setFillColor(self.heading_bg)
            c.roundRect(self.pad_x, current_y - hh,
                        self._HEAD_W, hh, capsule_r, fill=1, stroke=0)
            c.restoreState()

            # ③ 見出しテキスト（カプセル幅いっぱいで中央揃え）
            hp = Paragraph(heading, self.hstyle)
            hp.wrap(self._HEAD_W, hh)
            hp.drawOn(c, self.pad_x, current_y - hh + self._HEAD_PAD)

            current_y -= hh + self._INNER_GAP

            # ④ 本文テキスト
            bp = Paragraph(body_text, self.bstyle)
            bp.wrap(inner_w, bh + 200)
            bp.drawOn(c, self.pad_x, current_y - bh)

            current_y -= bh
            if i < n - 1:
                current_y -= self._OUTER_GAP


class _KeywordCard(Flowable):
    """
    キーワード用カード：白い角丸背景の中に
    ピンクのキーワードバッジ（上左）＋テキストを配置
    """
    _BADGE_FS    = 7.5   # バッジフォントサイズ (pt)
    _BADGE_LEAD  = 10    # バッジ leading (pt)
    _BADGE_PAD_X = 8     # バッジ左右パディング (pt)
    _BADGE_PAD_Y = 3.6   # バッジ上下パディング (pt)
    _BADGE_R     = 4.2   # バッジ角丸半径 (pt)

    def __init__(self, keywords_text, badge_color, body_style,
                 card_radius=8, pad_x=12, pad_tb=8, gap=-6):
        Flowable.__init__(self)
        self.keywords_text = keywords_text
        self.badge_color   = badge_color
        self.body_style    = body_style
        self.card_radius   = card_radius
        self.pad_x         = pad_x   # バッジ左端 = カード左端のオフセット
        self.pad_tb        = pad_tb
        self.gap           = gap     # 負の値 = バッジがカードに重なる量

    def wrap(self, availWidth, availHeight):
        self._card_w = availWidth

        # バッジ寸法
        btw = pdfmetrics.stringWidth("キーワード", FONT_BOLD, self._BADGE_FS)
        self._bw = btw + 2 * self._BADGE_PAD_X
        self._bh = self._BADGE_LEAD + 2 * self._BADGE_PAD_Y

        # 白いカード（テキストのみ）の高さ・幅
        # カード左端 = pad_x（バッジ左端と揃える）
        self._cx     = self.pad_x
        self._cw     = availWidth - self.pad_x
        # テキスト左端 = バッジ内テキスト左端（pad_x + BADGE_PAD_X）に揃える
        self._text_x = self.pad_x + self._BADGE_PAD_X
        inner_w      = availWidth - self._text_x - self.pad_x
        bp = Paragraph(self.keywords_text, self.body_style)
        _, self._body_h = bp.wrap(inner_w, availHeight)
        self._card_h = self.pad_tb + self._body_h + self.pad_tb

        # 全体高さ（gap 負値 → バッジがカードに重なる）
        self.height = self._bh + self.gap + self._card_h
        return availWidth, self.height

    def draw(self):
        c = self.canv
        w = self._card_w

        # ① 白い角丸カード（バッジ左端に揃えて配置）
        c.saveState()
        c.setFillColor(colors.white)
        c.roundRect(self._cx, 0, self._cw, self._card_h,
                    self.card_radius, fill=1, stroke=0)
        c.restoreState()

        # ② キーワードバッジ（カードの上・gap 分ずらす）
        bw, bh = self._bw, self._bh
        badge_x = self.pad_x               # カード左端と同じ
        badge_y = self._card_h + self.gap  # 負の gap → カードに重なる
        c.saveState()
        c.setFillColor(self.badge_color)
        c.roundRect(badge_x, badge_y, bw, bh, self._BADGE_R, fill=1, stroke=0)
        c.setFont(FONT_BOLD, self._BADGE_FS)
        c.setFillColor(colors.white)
        text_y = badge_y + (bh - self._BADGE_FS) / 2
        c.drawString(badge_x + self._BADGE_PAD_X, text_y, "キーワード")
        c.restoreState()

        # ③ キーワード本文（バッジ内テキスト左端に揃えて配置）
        inner_w = self._card_w - self._text_x - self.pad_x
        bp = Paragraph(self.keywords_text, self.body_style)
        bp.wrap(inner_w, self._body_h + 100)
        bp.drawOn(c, self._text_x, self.pad_tb)


class _WhiteCard(Flowable):
    """テキストを白い角丸カードで囲む Flowable（キャッチコピー用）
    offset_x: カード左端の x オフセット（キーワードカードの左端に揃える）
    fixed_w:  カード幅の固定値（None = availWidth）
    """
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
        # 角丸カード（4隅）
        c.roundRect(ox, 0, cw, h, r, fill=1, stroke=0)
        # 右下コーナーを直角に上書き
        c.rect(ox + cw - r, 0, r, r, fill=1, stroke=0)
        c.restoreState()

        inner_w = cw - 2 * self.pad_x
        p = Paragraph(self.text, self.style)
        p.wrap(inner_w, self._ph + 100)
        p.drawOn(c, ox + self.pad_x, self.pad_tb)


def _build_profile(block, styles, doc):
    """
    ヘッダー行（全幅）:
      左3/4 … キャッチコピー＋キーワード＋分野等＋email
      右1/4 … 氏名カード（英名・漢字名・顔写真）
    本体（左3/4：セクション　右1/4：職名・学位・写真）
    """
    pw      = A4[0] - 32 * mm    # leftMargin 16mm + rightMargin 16mm
    CARD_W  = 362                # 薄黄色カードの幅（pt ≒ px）
    lw      = CARD_W + 2 * mm   # カード幅 + 左右余白
    rw      = pw - lw            # 右カラム幅
    PINK    = colors.HexColor("#E8567C")
    PINK_BG = colors.HexColor("#FDF2F5")
    PINK_HD = colors.HexColor("#F06292")
    GRAY_PH = colors.HexColor("#CCCCCC")
    story   = []

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

    # ─── ① ヘッダー行（3:1）────────────────────────────────────────────────────
    # 左：キャッチコピー＋メタ情報
    cp_s = ParagraphStyle("cp", fontName=FONT_BOLD, fontSize=18, leading=26,
                            textColor=colors.HexColor("#1a1a1a"))
    kw_s = ParagraphStyle("kw", fontName=FONT_NAME, fontSize=Q13, leading=Q13*1.4,
                            textColor=BODY_COLOR)

    def _meta_row(badge_txt, val, radius=None, stroke=True):
        # テキスト実幅 + 左右パディング(12pt×2) で列幅を決定
        pill_style = ParagraphStyle("_tmp", fontName=FONT_BOLD, fontSize=7.5)
        text_w = pdfmetrics.stringWidth(badge_txt, pill_style.fontName, pill_style.fontSize)
        bw = text_w + 24 + 2   # 24pt = pad_x*2, +2pt 余裕
        BADGE_OFFSET = 12   # キーワードの pad_x と揃える
        pill = _pill(badge_txt, PINK, radius=radius, stroke=stroke)
        t = Table([[pill, Paragraph(val, kw_s)]],
                  colWidths=[bw + BADGE_OFFSET, lw - bw - BADGE_OFFSET - 4 * mm])
        t.setStyle(TableStyle([
            ("VALIGN",       (0,0),(-1,-1), "MIDDLE"),
            ("TOPPADDING",   (0,0),(-1,-1), 2),
            ("BOTTOMPADDING",(0,0),(-1,-1), 2),
            ("LEFTPADDING",  (0,0),(-1,-1), 0),
            ("LEFTPADDING",  (0,0),(0,0),   BADGE_OFFSET),  # キーワードに x 位置を揃える
            ("LEFTPADDING",  (1,0),(1,0),   8),             # 見出しとコンテンツの間 8pt
            ("RIGHTPADDING", (0,0),(-1,-1), 3),
        ]))
        return t

    left_hdr = []
    if keywords:
        left_hdr.append(_KeywordCard(keywords, PINK, kw_s))
        left_hdr.append(Spacer(1, 1.5 * mm))
    for badge_txt, val in [("分野等", field_name), (email_label, email)]:
        if val:
            left_hdr.append(_meta_row(badge_txt, val))
            left_hdr.append(Spacer(1, 1.5 * mm))

    # 右：氏名カード（ピンク背景）
    face_el = _b64_to_image(face_photo_data, rw - 6 * mm, 30 * mm)
    if face_el is None:
        face_el = Table([[""]], colWidths=[rw - 6 * mm], rowHeights=[28 * mm])
        face_el.setStyle(TableStyle([
            ("BACKGROUND", (0,0),(0,0), GRAY_PH),
            ("TOPPADDING", (0,0),(0,0), 0),
            ("BOTTOMPADDING", (0,0),(0,0), 0),
        ]))
    right_hdr = [
        Paragraph(name_en,
                  ParagraphStyle("ne", fontName=FONT_NAME, fontSize=7, leading=9, textColor=PINK)),
        Paragraph(f"<b>{name_ja}</b>",
                  ParagraphStyle("nj", fontName=FONT_BOLD, fontSize=14, leading=18,
                                 textColor=colors.HexColor("#111111"))),
        Spacer(1, 2 * mm),
        face_el,
    ]

    hdr_tbl = Table([[left_hdr, right_hdr]], colWidths=[lw, rw])
    hdr_tbl.setStyle(TableStyle([
        ("VALIGN",       (0,0),(-1,-1), "TOP"),
        ("BACKGROUND",   (1,0),(1,0),   PINK_BG),
        ("TOPPADDING",   (0,0),(-1,-1), 3),
        ("BOTTOMPADDING",(0,0),(-1,-1), 4),
        ("LEFTPADDING",  (0,0),(0,0),   0),
        ("LEFTPADDING",  (1,0),(1,0),   5),
        ("RIGHTPADDING", (0,0),(-1,-1), 4),
    ]))
    KW_PAD_X = 12       # _KeywordCard.pad_x と同値
    CARD_R   = 6 * mm  # 角丸半径（薄ピンク背景と共通）
    if catchcopy:
        story.append(_WhiteCard(catchcopy, cp_s,
                                card_radius=CARD_R,   # 薄ピンクと同じ角丸半径
                                offset_x=0,           # 薄ピンクの左端 x と同じ
                                fixed_w=lw))          # 薄ピンクの左端から lw 幅
        story.append(Spacer(1, -20))   # -20pt = 薄ピンク背景と20pt重なる
    story.append(hdr_tbl)
    story.append(Spacer(1, 4 * mm))

    # ─── ② 本体（3:1）──────────────────────────────────────────────────────
    body_s = ParagraphStyle("bs", fontName=FONT_NAME, fontSize=Q13, leading=Q13*1.65,
                             textColor=BODY_COLOR)
    sec_hs = ParagraphStyle("sh", fontName=FONT_BOLD, fontSize=Q13, leading=Q13*1.4,
                             textColor=WHITE, alignment=1)

    # 左：全セクションをひとつの角丸カードにまとめる
    SEC_BG = colors.HexColor("#fffeee")

    # 空セクションを除外してリスト化
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
            pad_x         = 11,      # 11pt ≒ 11px
            pad_tb        = 11,      # 11pt ≒ 11px
            fixed_h       = 636,     # 高さ固定 636pt ≒ 636px
        )
        left_body.append(sections_card)

    # 右：職名・学位バッジ＋写真プレースホルダー
    Q11 = 11 * 0.25 * mm
    bdg_pill_style = ParagraphStyle("bdgp", fontName=FONT_BOLD, fontSize=Q11,
                                    leading=Q11 * 1.4, textColor=WHITE)
    bdg_v = ParagraphStyle("bv", fontName=FONT_BOLD, fontSize=Q13, leading=Q13*1.4,
                            textColor=BODY_COLOR)
    right_body = []
    for lbl, val in [("職名", position), ("学位", degree)]:
        if not val:
            continue
        # カプセルバッジ：14pt・左右8pt・上下3pt・白枠1pt
        text_w = pdfmetrics.stringWidth(lbl, FONT_BOLD, Q11)
        bw = text_w + 2 * 8 + 2    # pad_x×2 + 余裕2pt
        pill = _PillFlowable(lbl, PINK, bdg_pill_style, pad_x=8, pad_y=3)
        bt = Table(
            [[pill, Paragraph(val, bdg_v)]],
            colWidths=[bw, rw - bw - 2 * mm],
        )
        bt.setStyle(TableStyle([
            ("VALIGN",       (0,0),(-1,-1), "MIDDLE"),
            ("TOPPADDING",   (0,0),(-1,-1), 2),
            ("BOTTOMPADDING",(0,0),(-1,-1), 2),
            ("LEFTPADDING",  (0,0),(-1,-1), 0),
            ("LEFTPADDING",  (1,0),(1,0),   8),   # 見出しとコンテンツの間 8pt
            ("RIGHTPADDING", (0,0),(-1,-1), 2),
        ]))
        right_body.append(bt)
        right_body.append(Spacer(1, 3 * mm))

    Q10 = 10 * 0.25 * mm  # 10Q
    cap_s = ParagraphStyle("cap", fontName=FONT_NAME, fontSize=Q10, leading=Q10*1.4,
                             textColor=colors.HexColor("#555555"), alignment=0)  # 0=左揃え
    ph_w = rw - 4 * mm

    if research_images:
        for ri in research_images:
            img_el = _b64_to_image(ri.get("data"), ph_w, 26 * mm)
            if img_el:
                right_body.append(img_el)
            else:
                ph = Table([[""]], colWidths=[ph_w], rowHeights=[24 * mm])
                ph.setStyle(TableStyle([("BACKGROUND",(0,0),(0,0),GRAY_PH),
                                         ("TOPPADDING",(0,0),(0,0),0),("BOTTOMPADDING",(0,0),(0,0),0)]))
                right_body.append(ph)
            cap_text = ri.get("name", "")
            if cap_text:
                right_body.append(Spacer(1, 0.5 * mm))
                right_body.append(Paragraph(cap_text, cap_s))
            right_body.append(Spacer(1, 2.5 * mm))
    else:
        # 画像未設定時：グレープレースホルダー4枚
        for _ in range(4):
            ph = Table([[""]], colWidths=[ph_w], rowHeights=[24 * mm])
            ph.setStyle(TableStyle([("BACKGROUND",(0,0),(0,0),GRAY_PH),
                                     ("TOPPADDING",(0,0),(0,0),0),("BOTTOMPADDING",(0,0),(0,0),0)]))
            right_body.append(ph)
            right_body.append(Spacer(1, 3 * mm))

    body_tbl = Table([[left_body, right_body]], colWidths=[lw, rw])
    body_tbl.setStyle(TableStyle([
        ("VALIGN",       (0,0),(-1,-1), "TOP"),
        ("BACKGROUND",   (0,0),(0,0),   SEC_BG),   # 黄色背景を左セル全体に（カード下まで伸張）
        ("TOPPADDING",   (0,0),(-1,-1), 0),
        ("BOTTOMPADDING",(0,0),(-1,-1), 0),
        ("LEFTPADDING",  (0,0),(-1,-1), 0),
        ("RIGHTPADDING", (0,0),(0,0),   3 * mm),
        ("RIGHTPADDING", (1,0),(1,0),   0),
    ]))
    story.append(body_tbl)
    return story


class _PillFlowable(Flowable):
    """
    カプセル型バッジ Flowable
    - 左右パディング: 12pt（≈12px）
    - 上下パディング: 3.6pt（≈3.6px）
    - 枠線: 白 1pt 実線
    - 角丸: 高さの半分（999px相当の完全カプセル）
    """
    def __init__(self, text, bg_color, style, pad_x=12, pad_y=3.6,
                 radius=None, stroke=True):
        Flowable.__init__(self)
        self.text     = text
        self.bg_color = bg_color
        self.style    = style
        self.pad_x    = pad_x    # pt
        self.pad_y    = pad_y    # pt
        self.radius   = radius   # None → h/2（完全カプセル）、数値 → 固定角丸
        self.stroke   = stroke   # False → 枠線なし
        self._tw      = 0
        self._th      = 0

    def wrap(self, availWidth, availHeight):
        # pdfmetrics.stringWidth でテキストの実際の幅を計測
        # Paragraph.wrap() は availWidth をそのまま返すため使わない
        self._tw = pdfmetrics.stringWidth(
            self.text, self.style.fontName, self.style.fontSize
        )
        self._th = self.style.leading
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

        # テキストを垂直中央に描画
        c.saveState()
        c.setFont(self.style.fontName, self.style.fontSize)
        c.setFillColor(self.style.textColor)
        text_y = (h - self.style.fontSize) / 2
        c.drawString(self.pad_x, text_y, self.text)
        c.restoreState()


def _pill(text, color, radius=None, stroke=True):
    """カプセル型バッジを返す"""
    style = ParagraphStyle("pl", fontName=FONT_BOLD, fontSize=7.5,
                            leading=10, textColor=WHITE)
    return _PillFlowable(text, color, style, pad_x=12, pad_y=3.6,
                         radius=radius, stroke=stroke)


if __name__ == "__main__":
    app.run(debug=True, port=5001)
