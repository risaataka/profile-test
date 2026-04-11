"""
フォント設定モジュール（日本語対応）

FONT_NAME, FONT_BOLD を解決して export する。
app.py から `from utils.font_setup import FONT_NAME, FONT_BOLD` で使用する。
"""
import os
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

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

for _font_path, _font_alias in SYSTEM_FONTS:
    if os.path.exists(_font_path):
        try:
            _kwargs = {"subfontIndex": 0} if _font_path.endswith(".ttc") else {}
            pdfmetrics.registerFont(TTFont(_font_alias, _font_path, **_kwargs))
            FONT_NAME = _font_alias
            FONT_BOLD = _font_alias
            print(f"[font] loaded: {_font_path}")
            break
        except Exception as _e:
            print(f"[font] failed {_font_path}: {_e}")
            continue

# システムフォントが見つからない場合は ReportLab 内蔵 CID フォントを使用
if FONT_NAME == "Helvetica":
    try:
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
        pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
        FONT_NAME = "HeiseiKakuGo-W5"
        FONT_BOLD = "HeiseiKakuGo-W5"
        print("[font] using built-in CID font: HeiseiKakuGo-W5")
    except Exception as _e:
        print(f"[font] CID font failed: {_e}. Japanese may not render.")
