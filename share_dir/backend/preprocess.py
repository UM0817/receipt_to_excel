import cv2
import numpy as np

def detect_receipt(image):

    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    blur = cv2.GaussianBlur(gray, (5,5), 0)

    edge = cv2.Canny(blur, 50, 150)

    contours, _ = cv2.findContours(
        edge,
        cv2.RETR_EXTERNAL,
        cv2.CHAIN_APPROX_SIMPLE
    )

    max_area = 0
    receipt_contour = None

    for cnt in contours:

        area = cv2.contourArea(cnt)

        if area < 10000:
            continue

        peri = cv2.arcLength(cnt, True)

        approx = cv2.approxPolyDP(
            cnt,
            0.02 * peri,
            True
        )

        if len(approx) == 4 and area > max_area:
            receipt_contour = approx
            max_area = area

    if receipt_contour is None:
        return None

    return receipt_contour.reshape(4,2)

def four_point_transform(image, pts):

    rect = order_points(pts)

    (tl, tr, br, bl) = rect

    widthA = np.linalg.norm(br - bl)
    widthB = np.linalg.norm(tr - tl)

    maxWidth = int(max(widthA, widthB))

    heightA = np.linalg.norm(tr - br)
    heightB = np.linalg.norm(tl - bl)

    maxHeight = int(max(heightA, heightB))

    dst = np.array([
        [0,0],
        [maxWidth-1,0],
        [maxWidth-1,maxHeight-1],
        [0,maxHeight-1]
    ], dtype="float32")

    M = cv2.getPerspectiveTransform(rect, dst)

    warped = cv2.warpPerspective(
        image,
        M,
        (maxWidth, maxHeight)
    )

    return warped