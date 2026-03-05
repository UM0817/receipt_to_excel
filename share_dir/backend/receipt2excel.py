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
    use_angle_cls=True,
    lang='japan',
    show_log=False
)

@app.route('/api/ocr', methods=['POST'])
def api_ocr():
    if 'image' not in request.files:
        return jsonify({'error': 'no image'}), 400

    file = request.files['image']
    img_bytes = file.read()

    nparr = np.frombuffer(img_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)

    result = ocr.ocr(img, cls=True)

    lines = []
    if result and result[0]:
        for line in result:
            for word_info in line:
                text = word_info[1][0]
                lines.append(text)

    return jsonify({'lines': lines})


@app.route('/api/export_excel', methods=['POST'])
def export_excel():
    data = request.json
    rows = data.get('rows', [])

    df = pd.DataFrame(rows)

    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='receipts')

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