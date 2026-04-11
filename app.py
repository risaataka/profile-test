import pandas as pd
from flask import Flask, render_template, request, jsonify, send_file

from utils.font_setup import FONT_NAME, FONT_BOLD  # noqa: F401（フォント登録を起動時に実行）
from utils.pdf_generator import build_pdf

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 128 * 1024 * 1024  # 128MB（画像含むJSON対応）


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

    buf = build_pdf(data)
    return send_file(
        buf,
        mimetype="application/pdf",
        as_attachment=False,
        download_name="output.pdf",
    )


if __name__ == "__main__":
    app.run(debug=True, port=5001)
