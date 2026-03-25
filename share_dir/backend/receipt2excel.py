from datetime import datetime
import io
import re

import cv2
import numpy as np
import pandas as pd
from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from paddleocr import PaddleOCR

app = Flask(__name__)
CORS(app)

MIN_CONFIDENCE = 0.45
DEFAULT_ROWS = 18
DEFAULT_COLS = 6
MAX_ROTATION_DEGREES = 12.0
FAST_PATH_MIN_ITEMS = 12
FAST_PATH_MIN_AVG_SCORE = 0.82
SECOND_PASS_MIN_ITEMS = 8
DEFAULT_OCR_ATTEMPTS = 1
MAX_OCR_ATTEMPTS = 5
EXPORT_MODE_SEPARATE = "separate_sheets"
EXPORT_MODE_COMBINED = "combined_sheet"
TEXT_NORMALIZATION_MAP = str.maketrans(
    {
        "０": "0",
        "１": "1",
        "２": "2",
        "３": "3",
        "４": "4",
        "５": "5",
        "６": "6",
        "７": "7",
        "８": "8",
        "９": "9",
        "￥": "¥",
        "，": ",",
        "．": ".",
        "−": "-",
        "ー": "-",
        "―": "-",
    }
)
YEN_SYMBOL_PATTERN = re.compile(r"^[¥YyVv\\/LIl|!]{1,2}$")
LEADING_YEN_PATTERN = re.compile(r"^[¥YyVv\\/LIl|!]+(?=\d)")
AMOUNT_PATTERN = re.compile(r"^\d{1,3}(,\d{3})*(\.\d{1,2})?$|^\d+(\.\d{1,2})?$")
YEN_TEXT_CANDIDATES = {
    "工大",
    "工夫",
    "ギ",
    "エ夫",
    "エ大",
}


def init_ocr():
    return PaddleOCR(
        lang="japan",
        use_mkldnn=False,
        use_angle_cls=True,
        det_db_thresh=0.22,
        det_db_box_thresh=0.45,
        det_db_unclip_ratio=1.8,
        rec_batch_num=8,
        show_log=False,
        use_gpu=False,
    )


ocr = init_ocr()


def normalize_text(text):
    if text is None:
        return ""
    normalized = re.sub(r"\s+", " ", str(text)).strip()
    return normalized.translate(TEXT_NORMALIZATION_MAP)


def normalize_amount_text(text):
    normalized = normalize_text(text).replace(" ", "")
    if normalized in YEN_TEXT_CANDIDATES:
        return "¥"
    normalized = LEADING_YEN_PATTERN.sub("¥", normalized)
    normalized = re.sub(r"¥{2,}", "¥", normalized)
    return normalized


def looks_like_amount(text):
    candidate = normalize_amount_text(text)
    if candidate.startswith("¥"):
        candidate = candidate[1:]
    return bool(candidate) and bool(AMOUNT_PATTERN.fullmatch(candidate))


def is_yen_symbol_candidate(text):
    candidate = normalize_text(text).replace(" ", "")
    return bool(candidate) and (
        candidate in YEN_TEXT_CANDIDATES or bool(YEN_SYMBOL_PATTERN.fullmatch(candidate))
    )


def merge_box_geometry(base_item, next_item):
    base_item["x_right"] = max(base_item["x_right"], next_item["x_right"])
    base_item["y_top"] = min(base_item["y_top"], next_item["y_top"])
    base_item["y_bottom"] = max(base_item["y_bottom"], next_item["y_bottom"])
    base_item["x_center"] = (base_item["x_left"] + base_item["x_right"]) / 2
    base_item["y_center"] = (base_item["y_top"] + base_item["y_bottom"]) / 2
    base_item["width"] = max(1.0, base_item["x_right"] - base_item["x_left"])
    base_item["height"] = max(1.0, base_item["y_bottom"] - base_item["y_top"])
    base_item["score"] = max(base_item["score"], next_item["score"])


def normalize_currency_tokens(row_items):
    if not row_items:
        return []

    normalized_items = []
    index = 0

    while index < len(row_items):
        current = dict(row_items[index])
        current["text"] = normalize_amount_text(current["text"])

        if index + 1 < len(row_items):
            next_item = dict(row_items[index + 1])
            next_text = normalize_amount_text(next_item["text"])

            if is_yen_symbol_candidate(current["text"]) and looks_like_amount(next_text):
                current["text"] = f"¥{normalize_amount_text(next_text).lstrip('¥')}"
                merge_box_geometry(current, next_item)
                normalized_items.append(current)
                index += 2
                continue

        normalized_items.append(current)
        index += 1

    return normalized_items


def resize_for_ocr(image):
    height, width = image.shape[:2]
    longest = max(height, width)

    if longest < 1200:
        scale = 1200 / longest
    elif longest > 1800:
        scale = 1800 / longest
    else:
        scale = 1.0

    if abs(scale - 1.0) < 0.05:
        return image

    return cv2.resize(
        image,
        None,
        fx=scale,
        fy=scale,
        interpolation=cv2.INTER_CUBIC if scale > 1 else cv2.INTER_AREA,
    )


def order_points(points):
    rect = np.zeros((4, 2), dtype="float32")
    sums = points.sum(axis=1)
    diffs = np.diff(points, axis=1)
    rect[0] = points[np.argmin(sums)]
    rect[2] = points[np.argmax(sums)]
    rect[1] = points[np.argmin(diffs)]
    rect[3] = points[np.argmax(diffs)]
    return rect


def four_point_transform(image, points):
    rect = order_points(points)
    top_left, top_right, bottom_right, bottom_left = rect

    width_top = np.linalg.norm(top_right - top_left)
    width_bottom = np.linalg.norm(bottom_right - bottom_left)
    max_width = int(max(width_top, width_bottom))

    height_right = np.linalg.norm(top_right - bottom_right)
    height_left = np.linalg.norm(top_left - bottom_left)
    max_height = int(max(height_right, height_left))

    if max_width < 80 or max_height < 80:
        return image

    destination = np.array(
        [
            [0, 0],
            [max_width - 1, 0],
            [max_width - 1, max_height - 1],
            [0, max_height - 1],
        ],
        dtype="float32",
    )

    matrix = cv2.getPerspectiveTransform(rect, destination)
    return cv2.warpPerspective(image, matrix, (max_width, max_height))


def crop_receipt_region(image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    edges = cv2.Canny(blurred, 60, 180)
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (5, 5))
    edges = cv2.dilate(edges, kernel, iterations=2)
    edges = cv2.morphologyEx(edges, cv2.MORPH_CLOSE, kernel, iterations=2)

    contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    image_area = image.shape[0] * image.shape[1]

    for contour in sorted(contours, key=cv2.contourArea, reverse=True):
        area = cv2.contourArea(contour)
        if area < image_area * 0.2:
            continue

        perimeter = cv2.arcLength(contour, True)
        approximation = cv2.approxPolyDP(contour, 0.02 * perimeter, True)

        if len(approximation) == 4:
            warped = four_point_transform(image, approximation.reshape(4, 2).astype("float32"))
            if warped.shape[0] > 100 and warped.shape[1] > 100:
                return warped, True

    return image, False


def estimate_skew_angle(gray_image):
    inverted = cv2.bitwise_not(gray_image)
    _, threshold = cv2.threshold(
        inverted, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU
    )
    coordinates = np.column_stack(np.where(threshold > 0))

    if len(coordinates) < 100:
        return 0.0

    angle = cv2.minAreaRect(coordinates)[-1]
    if angle < -45:
        angle = 90 + angle
    elif angle > 45:
        angle = angle - 90

    angle = float(np.clip(angle, -MAX_ROTATION_DEGREES, MAX_ROTATION_DEGREES))
    return angle


def rotate_image(image, angle):
    if abs(angle) < 0.25:
        return image

    height, width = image.shape[:2]
    center = (width / 2, height / 2)
    matrix = cv2.getRotationMatrix2D(center, angle, 1.0)

    cos_value = abs(matrix[0, 0])
    sin_value = abs(matrix[0, 1])
    bound_width = int((height * sin_value) + (width * cos_value))
    bound_height = int((height * cos_value) + (width * sin_value))

    matrix[0, 2] += (bound_width / 2) - center[0]
    matrix[1, 2] += (bound_height / 2) - center[1]

    return cv2.warpAffine(
        image,
        matrix,
        (bound_width, bound_height),
        flags=cv2.INTER_CUBIC,
        borderMode=cv2.BORDER_REPLICATE,
    )


def preprocess_receipt_image(image):
    cropped_image, was_cropped = crop_receipt_region(image)
    resized_image = resize_for_ocr(cropped_image)
    gray = cv2.cvtColor(resized_image, cv2.COLOR_BGR2GRAY)
    angle = estimate_skew_angle(gray)
    deskewed_image = rotate_image(resized_image, angle)

    return deskewed_image, {
        "cropped": was_cropped,
        "deskew_angle": round(angle, 2),
    }


def build_primary_ocr_variants(base):
    gray = cv2.cvtColor(base, cv2.COLOR_BGR2GRAY)
    clahe = cv2.createCLAHE(clipLimit=2.2, tileGridSize=(8, 8)).apply(gray)
    enhanced = cv2.cvtColor(clahe, cv2.COLOR_GRAY2BGR)

    binary = cv2.adaptiveThreshold(
        clahe,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        31,
        11,
    )
    binary_bgr = cv2.cvtColor(binary, cv2.COLOR_GRAY2BGR)

    sharpened = cv2.addWeighted(
        enhanced,
        1.5,
        cv2.GaussianBlur(enhanced, (0, 0), 3.0),
        -0.5,
        0,
    )

    return [
        ("base", base),
        ("enhanced", enhanced),
        ("binary", binary_bgr),
    ]


def build_fallback_ocr_variants(base):
    gray = cv2.cvtColor(base, cv2.COLOR_BGR2GRAY)
    clahe = cv2.createCLAHE(clipLimit=2.2, tileGridSize=(8, 8)).apply(gray)
    enhanced = cv2.cvtColor(clahe, cv2.COLOR_GRAY2BGR)

    sharpened = cv2.addWeighted(
        enhanced,
        1.5,
        cv2.GaussianBlur(enhanced, (0, 0), 3.0),
        -0.5,
        0,
    )

    inverse_binary = cv2.adaptiveThreshold(
        clahe,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY_INV,
        31,
        9,
    )
    inverse_binary = cv2.bitwise_not(inverse_binary)
    inverse_binary_bgr = cv2.cvtColor(inverse_binary, cv2.COLOR_GRAY2BGR)

    return [
        ("sharpened", sharpened),
        ("inverse_binary", inverse_binary_bgr),
    ]

def parse_ocr_result(result, min_confidence=MIN_CONFIDENCE):
    items = []

    if not result:
        return items

    for page in result:
        if not page:
            continue

        for word_info in page:
            if not word_info or len(word_info) < 2:
                continue

            box = word_info[0]
            text, score = word_info[1]
            text = normalize_amount_text(text)

            if not text or score < min_confidence:
                continue

            xs = [point[0] for point in box]
            ys = [point[1] for point in box]

            x_left = float(min(xs))
            x_right = float(max(xs))
            y_top = float(min(ys))
            y_bottom = float(max(ys))

            items.append(
                {
                    "text": text,
                    "score": float(score),
                    "x_left": x_left,
                    "x_right": x_right,
                    "y_top": y_top,
                    "y_bottom": y_bottom,
                    "x_center": (x_left + x_right) / 2,
                    "y_center": (y_top + y_bottom) / 2,
                    "width": max(1.0, x_right - x_left),
                    "height": max(1.0, y_bottom - y_top),
                }
            )

    return items


def score_ocr_result(result):
    items = parse_ocr_result(result, min_confidence=0.0)
    if not items:
        return -1

    confidence_sum = sum(item["score"] for item in items)
    return confidence_sum + (len(items) * 0.2)


def average_item_score(items):
    if not items:
        return 0.0
    return sum(item["score"] for item in items) / len(items)


def should_accept_fast_path(items):
    return (
        len(items) >= FAST_PATH_MIN_ITEMS
        and average_item_score(items) >= FAST_PATH_MIN_AVG_SCORE
    )


def normalize_ocr_attempts(value):
    try:
        attempts = int(value)
    except (TypeError, ValueError):
        return DEFAULT_OCR_ATTEMPTS

    return max(1, min(MAX_OCR_ATTEMPTS, attempts))


def normalize_export_mode(value):
    if value == EXPORT_MODE_COMBINED:
        return EXPORT_MODE_COMBINED
    return EXPORT_MODE_SEPARATE


def evaluate_variant(variant_name, variant_image):
    result = ocr.ocr(variant_image, cls=True)
    items = parse_ocr_result(result)

    for item in items:
        item["variant"] = variant_name

    return {
        "name": variant_name,
        "score": score_ocr_result(result),
        "items": items,
    }


def boxes_overlap_ratio(item_a, item_b):
    x_left = max(item_a["x_left"], item_b["x_left"])
    y_top = max(item_a["y_top"], item_b["y_top"])
    x_right = min(item_a["x_right"], item_b["x_right"])
    y_bottom = min(item_a["y_bottom"], item_b["y_bottom"])

    if x_right <= x_left or y_bottom <= y_top:
        return 0.0

    intersection = (x_right - x_left) * (y_bottom - y_top)
    area_a = item_a["width"] * item_a["height"]
    area_b = item_b["width"] * item_b["height"]
    union = area_a + area_b - intersection

    if union <= 0:
        return 0.0

    return intersection / union


def are_similar_items(item_a, item_b):
    y_distance = abs(item_a["y_center"] - item_b["y_center"])
    x_distance = abs(item_a["x_center"] - item_b["x_center"])
    overlap = boxes_overlap_ratio(item_a, item_b)
    similar_text = item_a["text"] == item_b["text"]

    return (
        overlap > 0.35
        or (
            similar_text
            and y_distance <= max(item_a["height"], item_b["height"]) * 0.7
            and x_distance <= max(item_a["width"], item_b["width"]) * 0.8
        )
    )


def merge_item_group(items):
    best = max(items, key=lambda item: item["score"])
    merged = dict(best)
    merged["x_left"] = min(item["x_left"] for item in items)
    merged["x_right"] = max(item["x_right"] for item in items)
    merged["y_top"] = min(item["y_top"] for item in items)
    merged["y_bottom"] = max(item["y_bottom"] for item in items)
    merged["x_center"] = (merged["x_left"] + merged["x_right"]) / 2
    merged["y_center"] = (merged["y_top"] + merged["y_bottom"]) / 2
    merged["width"] = max(1.0, merged["x_right"] - merged["x_left"])
    merged["height"] = max(1.0, merged["y_bottom"] - merged["y_top"])
    merged["score"] = max(item["score"] for item in items)
    merged["variant"] = best.get("variant", "base")
    return merged


def merge_variant_items(variant_items):
    merged_groups = []

    for item in sorted(variant_items, key=lambda entry: (entry["y_center"], entry["x_center"])):
        matched_group = None
        for group in merged_groups:
            if any(are_similar_items(item, existing) for existing in group):
                matched_group = group
                break

        if matched_group is None:
            merged_groups.append([item])
        else:
            matched_group.append(item)

    return [merge_item_group(group) for group in merged_groups]


def run_best_ocr(image, max_attempts=DEFAULT_OCR_ATTEMPTS):
    base_image, preprocess_meta = preprocess_receipt_image(image)
    variant_results = []

    primary_variants = build_primary_ocr_variants(base_image)
    fallback_variants = build_fallback_ocr_variants(base_image)
    all_variants = primary_variants + fallback_variants
    allowed_attempts = min(normalize_ocr_attempts(max_attempts), len(all_variants))

    for index, (variant_name, variant_image) in enumerate(all_variants):
        if index >= allowed_attempts:
            break

        result = evaluate_variant(variant_name, variant_image)
        variant_results.append(result)

        if index == 0 and allowed_attempts == 1 and should_accept_fast_path(result["items"]):
            preprocess_meta["variant_runs"] = 1
            return result["name"], result["items"], preprocess_meta

    base_result = variant_results[0] if variant_results else {"name": "base", "score": -1, "items": []}
    best_variant = max(variant_results, key=lambda result: result["score"], default=base_result)

    merged_items = merge_variant_items(
        [item for result in variant_results for item in result["items"]]
    )

    if not merged_items:
        preprocess_meta["variant_runs"] = len(variant_results)
        return best_variant["name"], best_variant["items"], preprocess_meta

    preprocess_meta["variant_runs"] = len(variant_results)
    return best_variant["name"], merged_items, preprocess_meta


def merge_row_tokens(row_items):
    if not row_items:
        return []

    sorted_items = sorted(row_items, key=lambda item: item["x_left"])
    median_height = float(np.median([item["height"] for item in sorted_items]))
    gap_threshold = max(12.0, median_height * 0.7)
    merged = [dict(sorted_items[0])]

    for item in sorted_items[1:]:
        previous = merged[-1]
        vertical_distance = abs(item["y_center"] - previous["y_center"])
        gap = item["x_left"] - previous["x_right"]

        if gap <= gap_threshold and vertical_distance <= max(item["height"], previous["height"]) * 0.55:
            previous["text"] = normalize_text(f"{previous['text']} {item['text']}")
            merge_box_geometry(previous, item)
        else:
            merged.append(dict(item))

    return normalize_currency_tokens(merged)


def group_rows(items):
    if not items:
        return []

    sorted_items = sorted(items, key=lambda item: item["y_center"])
    median_height = float(np.median([item["height"] for item in sorted_items]))
    row_threshold = max(16.0, median_height * 0.75)
    rows = []

    for item in sorted_items:
        if not rows:
            rows.append([item])
            continue

        current_center = float(np.median([entry["y_center"] for entry in rows[-1]]))
        if abs(item["y_center"] - current_center) <= row_threshold:
            rows[-1].append(item)
        else:
            rows.append([item])

    return [merge_row_tokens(row) for row in rows]


def build_column_anchors(rows):
    all_items = [item for row in rows for item in row]
    if not all_items:
        return []

    median_width = float(np.median([item["width"] for item in all_items]))
    anchor_threshold = max(28.0, median_width * 0.9)
    anchors = []

    for item in sorted(all_items, key=lambda entry: entry["x_center"]):
        if not anchors:
            anchors.append({"x": item["x_center"], "count": 1})
            continue

        nearest_index = min(
            range(len(anchors)),
            key=lambda index: abs(anchors[index]["x"] - item["x_center"]),
        )
        nearest_anchor = anchors[nearest_index]
        distance = abs(nearest_anchor["x"] - item["x_center"])

        if distance <= max(anchor_threshold, item["width"] * 0.75):
            nearest_anchor["count"] += 1
            nearest_anchor["x"] = (
                (nearest_anchor["x"] * (nearest_anchor["count"] - 1)) + item["x_center"]
            ) / nearest_anchor["count"]
        else:
            anchors.append({"x": item["x_center"], "count": 1})

    anchors.sort(key=lambda anchor: anchor["x"])
    return anchors


def rows_to_cells(rows):
    anchors = build_column_anchors(rows)
    cells = []

    if not anchors:
        return cells, 0, 0

    for row_index, row_items in enumerate(rows):
        used_columns = set()

        for item in sorted(row_items, key=lambda entry: entry["x_left"]):
            column_index = min(
                range(len(anchors)),
                key=lambda index: abs(anchors[index]["x"] - item["x_center"]),
            )

            while column_index in used_columns:
                column_index += 1

            used_columns.add(column_index)
            cells.append(
                {
                    "row": row_index,
                    "col": column_index,
                    "text": item["text"],
                    "score": round(item["score"], 4),
                }
            )

    unique_columns = sorted({cell["col"] for cell in cells})
    column_map = {column: index for index, column in enumerate(unique_columns)}

    for cell in cells:
        cell["col"] = column_map[cell["col"]]

    return cells, len(rows), len(unique_columns)


def sanitize_sheet_name(name, fallback):
    cleaned = re.sub(r"[\[\]:*?/\\]", "_", normalize_text(name))
    return (cleaned or fallback)[:31]


@app.route("/api/ocr", methods=["POST"])
def api_ocr():
    if "image" not in request.files:
        return jsonify({"error": "no image"}), 400

    file = request.files["image"]
    ocr_attempts = normalize_ocr_attempts(request.form.get("ocr_attempts"))
    img_bytes = file.read()

    nparr = np.frombuffer(img_bytes, np.uint8)
    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)

    if img is None:
        return jsonify({"error": "invalid image"}), 400

    global ocr

    try:
        variant_name, items, preprocess_meta = run_best_ocr(img, max_attempts=ocr_attempts)
    except RuntimeError as error:
        app.logger.warning("PaddleOCR runtime error, reinitializing OCR: %s", error)
        ocr = init_ocr()
        try:
            variant_name, items, preprocess_meta = run_best_ocr(img, max_attempts=ocr_attempts)
        except Exception as retry_error:
            app.logger.error("PaddleOCR retry failed: %s", retry_error)
            return jsonify({"error": "OCR failed"}), 500
    except Exception as error:
        app.logger.error("OCR general error: %s", error)
        return jsonify({"error": "OCR failed"}), 500

    rows = group_rows(items)
    cells, row_count, column_count = rows_to_cells(rows)

    return jsonify(
        {
            "cells": cells,
            "meta": {
                "variant": variant_name,
                "recognized_tokens": len(items),
                "recognized_rows": max(row_count, DEFAULT_ROWS if cells else 0),
                "recognized_columns": max(column_count, DEFAULT_COLS if cells else 0),
                "cropped": preprocess_meta["cropped"],
                "deskew_angle": preprocess_meta["deskew_angle"],
                "variant_runs": preprocess_meta["variant_runs"],
                "requested_attempts": ocr_attempts,
            },
        }
    )


@app.route("/api/export_excel", methods=["POST"])
def export_excel():
    data = request.json or {}
    rows = data.get("rows", [])
    export_mode = normalize_export_mode(data.get("export_mode"))

    grouped = {}

    for entry in rows:
        receipt_no = entry.get("receipt_no", 1)
        grouped.setdefault(receipt_no, []).append(entry)

    bio = io.BytesIO()

    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        if not grouped:
            pd.DataFrame([[""]]).to_excel(writer, index=False, header=False, sheet_name="Receipt 1")
        elif export_mode == EXPORT_MODE_COMBINED:
            combined_rows = []

            for receipt_no in sorted(grouped.keys()):
                receipt_rows = grouped[receipt_no]
                receipt_name = normalize_text(
                    receipt_rows[0].get("receipt_name", f"Receipt {receipt_no}")
                ) or f"Receipt {receipt_no}"

                filled_cells = []
                for entry in receipt_rows:
                    text = normalize_text(entry.get("text", ""))
                    if not text:
                        continue
                    filled_cells.append(
                        {
                            "row": int(entry.get("row", 0)),
                            "col": int(entry.get("col", 0)),
                            "text": text,
                        }
                    )

                if not filled_cells:
                    combined_rows.append([receipt_name])
                    combined_rows.append([""])
                    continue

                max_row = max(cell["row"] for cell in filled_cells) + 1
                max_col = max(cell["col"] for cell in filled_cells) + 1
                grid = [["" for _ in range(max_col)] for _ in range(max_row)]

                for cell in filled_cells:
                    grid[cell["row"]][cell["col"]] = cell["text"]

                combined_rows.append([receipt_name])
                combined_rows.extend(grid)
                combined_rows.append([""])

            max_width = max((len(row) for row in combined_rows), default=1)
            normalized_rows = [row + [""] * (max_width - len(row)) for row in combined_rows]
            pd.DataFrame(normalized_rows).to_excel(
                writer,
                index=False,
                header=False,
                sheet_name="Receipts",
            )
        else:
            for receipt_no in sorted(grouped.keys()):
                receipt_rows = grouped[receipt_no]
                sheet_name = sanitize_sheet_name(
                    receipt_rows[0].get("receipt_name", f"Receipt {receipt_no}"),
                    f"Receipt {receipt_no}",
                )

                filled_cells = []
                for entry in receipt_rows:
                    text = normalize_text(entry.get("text", ""))
                    if not text:
                        continue
                    filled_cells.append(
                        {
                            "row": int(entry.get("row", 0)),
                            "col": int(entry.get("col", 0)),
                            "text": text,
                        }
                    )

                if not filled_cells:
                    pd.DataFrame([[""]]).to_excel(
                        writer,
                        index=False,
                        header=False,
                        sheet_name=sheet_name,
                    )
                    continue

                max_row = max(cell["row"] for cell in filled_cells) + 1
                max_col = max(cell["col"] for cell in filled_cells) + 1
                df = pd.DataFrame("", index=range(max_row), columns=range(max_col))

                for cell in filled_cells:
                    df.iloc[cell["row"], cell["col"]] = cell["text"]

                df.to_excel(writer, index=False, header=False, sheet_name=sheet_name)

    bio.seek(0)
    filename = f"{datetime.now().strftime('%Y%m%d')}.xlsx"

    return send_file(
        bio,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
