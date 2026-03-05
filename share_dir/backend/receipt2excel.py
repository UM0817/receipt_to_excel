from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import io
import numpy as np
import cv2
import pandas as pd
from datetime import datetime
from paddleocr import PaddleOCR

app = Flask(__name__)
CORS(app)

# OCR初期化（1回のみ）
ocr = PaddleOCR(
    lang="japan",
    use_mkldnn=False,
    use_angle_cls=True,
    det_db_thresh=0.3,
    det_db_box_thresh=0.5,
    det_db_unclip_ratio=2.0,
    rec_batch_num=6,
    show_log=False,
    use_gpu=False
)

@app.route('/api/ocr', methods=['POST'])
def api_ocr():
    if 'image' not in request.files:
        return jsonify({'error': 'no image'}), 400

    file = request.files['image']
    img_bytes = file.read()

    nparr = np.frombuffer(img_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)

    h, w = img.shape[:2]

    result = ocr.ocr(img, cls=True)

    cells = []
    boxes = []

    # box収集
    if result and result[0]:
        for line in result:
            for word_info in line:
                box = word_info[0]
                text = word_info[1][0]

                x = sum(p[0] for p in box) / 4
                y = sum(p[1] for p in box) / 4

                boxes.append((x, y, text))

    # OCR範囲
    xs = [b[0] for b in boxes]
    ys = [b[1] for b in boxes]

    xmin = min(xs)
    xmax = max(xs)
    ymin = min(ys)
    ymax = max(ys)

    xrange = xmax - xmin
    yrange = ymax - ymin

    for x, y, text in boxes:

        # 正規化
        x_norm = (x - xmin) / xrange
        y_norm = (y - ymin) / yrange

        # Excel座標
        col = int(x_norm * 4)
        row = int(y_norm * 29)

        cells.append({
            "text": text,
            "row": row,
            "col": col
        })

    return jsonify({"cells": cells})


@app.route('/api/export_excel', methods=['POST'])
def export_excel():
    data = request.json
    rows = data.get('rows', [])

    sheet = {}

    for r in rows:

        row = r["row"]
        col = r["col"]
        text = r["text"]

        if row not in sheet:
            sheet[row] = {}

        sheet[row][col] = text

    max_row = max(sheet.keys()) + 1
    max_col = 10

    df = pd.DataFrame("", index=range(max_row), columns=range(max_col))

    for r in sheet:
        for c in sheet[r]:
            df.iloc[r, c] = sheet[r][c]

    bio = io.BytesIO()

    with pd.ExcelWriter(bio, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, header=False)

    bio.seek(0)

    filename = f"receipts_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    return send_file(
        bio,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)