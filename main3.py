#!/usr/bin/env python3
"""
OCR to CSV extractor for CAD-style spec sheets.

Rules implemented per request:
- Detect tables, labels, and key-value pairs using EasyOCR with bounding boxes.
- Preserve numeric values, units, and dates exactly as shown by OCR.
- Convert all detected rows into CSV, printing ONLY CSV content to stdout.
- Each output line represents a logical data row (table rows or key-value pairs).
"""

import sys
import argparse
import math
import csv
from typing import List, Tuple, Dict, Any

import cv2
import numpy as np
import easyocr


class OcrCsvExtractor:
    """Extracts CSV-like rows from an image using OCR layout heuristics."""

    def __init__(self) -> None:
        self.reader = easyocr.Reader(['en'], gpu=False)

    # ----------------------------- Image pre-processing -----------------------------
    def preprocess(self, image_path: str) -> np.ndarray:
        image = cv2.imread(image_path)
        if image is None:
            raise FileNotFoundError(f"Cannot read image: {image_path}")
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        denoised = cv2.medianBlur(gray, 3)
        # Adaptive to handle shadows and varying contrast
        binary = cv2.adaptiveThreshold(denoised, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                       cv2.THRESH_BINARY, 31, 5)
        # Slight closing to solidify characters and gridlines
        kernel = np.ones((2, 2), np.uint8)
        cleaned = cv2.morphologyEx(binary, cv2.MORPH_CLOSE, kernel)

        h, w = cleaned.shape
        if max(h, w) < 1400:
            scale = 1400.0 / max(h, w)
            cleaned = cv2.resize(cleaned, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_CUBIC)
        return cleaned

    # ----------------------------- OCR and layout utils -----------------------------
    @staticmethod
    def _box_center(box: List[Tuple[float, float]]) -> Tuple[float, float]:
        xs = [p[0] for p in box]
        ys = [p[1] for p in box]
        return (sum(xs) / 4.0, sum(ys) / 4.0)

    @staticmethod
    def _box_height(box: List[Tuple[float, float]]) -> float:
        # Approximate height from vertical edges
        return (abs(box[0][1] - box[3][1]) + abs(box[1][1] - box[2][1])) / 2.0

    def read_layout(self, image: np.ndarray) -> List[Dict[str, Any]]:
        # EasyOCR expects RGB
        if len(image.shape) == 2:
            rgb = cv2.cvtColor(image, cv2.COLOR_GRAY2RGB)
        else:
            rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)

        results = self.reader.readtext(rgb, detail=1, paragraph=False)
        items: List[Dict[str, Any]] = []
        for box, text, conf in results:
            if not str(text).strip():
                continue
            cx, cy = self._box_center(box)
            height = self._box_height(box)
            x_min = min(p[0] for p in box)
            items.append({
                'text': str(text).strip(),
                'box': box,
                'center_y': cy,
                'center_x': cx,
                'height': max(height, 1.0),
                'x_min': x_min,
            })
        # Sort top-to-bottom then left-to-right
        items.sort(key=lambda d: (d['center_y'], d['x_min']))
        return items

    # ----------------------------- Row reconstruction -----------------------------
    def group_rows(self, items: List[Dict[str, Any]]) -> List[List[Dict[str, Any]]]:
        if not items:
            return []
        rows: List[List[Dict[str, Any]]] = []
        current: List[Dict[str, Any]] = []
        # Use dynamic vertical threshold based on median element height
        heights = [it['height'] for it in items]
        median_h = sorted(heights)[len(heights)//2]
        y_tol = max(6.0, median_h * 0.6)

        last_y = None
        for it in items:
            if last_y is None:
                current = [it]
                last_y = it['center_y']
                continue
            if abs(it['center_y'] - last_y) <= y_tol:
                current.append(it)
                last_y = (last_y + it['center_y']) / 2.0
            else:
                rows.append(sorted(current, key=lambda x: x['x_min']))
                current = [it]
                last_y = it['center_y']
        if current:
            rows.append(sorted(current, key=lambda x: x['x_min']))
        return rows

    # ----------------------------- Cell normalization -----------------------------
    @staticmethod
    def _normalize_cell(text: str) -> str:
        # Keep numeric formats/units/dates as-is; only strip excessive inner spaces
        t = text.replace('\u00a0', ' ').strip()
        t = ' '.join(part for part in t.split())
        return t

    @staticmethod
    def _split_key_value(text: str) -> Tuple[str, str]:
        # Split on the first ':' or ' - ' if it seems like a label
        if ':' in text:
            k, v = text.split(':', 1)
            return k.strip(), v.strip()
        # Handle cases with trailing label followed by value blocks
        for sep in [' - ', ' — ', '–']:
            if sep in text:
                k, v = text.split(sep, 1)
                return k.strip(), v.strip()
        return text.strip(), ''

    # ----------------------------- CSV assembly -----------------------------
    def rows_to_csv(self, rows: List[List[Dict[str, Any]]]) -> List[List[str]]:
        csv_rows: List[List[str]] = []

        for row_items in rows:
            texts = [self._normalize_cell(it['text']) for it in row_items]

            # Heuristic: table header/data if there are 3+ short cells aligned across
            avg_len = sum(len(t) for t in texts) / max(len(texts), 1)
            if len(texts) >= 3 and avg_len <= 12:
                csv_rows.append(texts)
                continue

            # If single long cell and looks like key:value -> split into two columns
            if len(texts) == 1:
                key, val = self._split_key_value(texts[0])
                if key and (val or key.endswith(':')):
                    csv_rows.append([key.rstrip(':'), val])
                else:
                    csv_rows.append([texts[0]])
                continue

            # Two or more cells: if the first cell ends with ':' treat as key/value row
            first = texts[0]
            if first.endswith(':') and len(texts) >= 2:
                csv_rows.append([first.rstrip(':'), ' '.join(texts[1:])])
                continue

            # Default: join as table-like row
            csv_rows.append(texts)

        # Post-process: merge obvious section titles with next line as header if needed
        merged: List[List[str]] = []
        skip_next = False
        for i, r in enumerate(csv_rows):
            if skip_next:
                skip_next = False
                continue
            if len(r) == 1 and i + 1 < len(csv_rows) and len(csv_rows[i + 1]) >= 3:
                title = r[0]
                merged.append([f"[Section] {title}"] + csv_rows[i + 1])
                skip_next = True
            else:
                merged.append(r)
        return merged

    # ----------------------------- Orchestration -----------------------------
    def extract_csv(self, image_path: str) -> List[List[str]]:
        pre = self.preprocess(image_path)
        items = self.read_layout(pre)
        grouped = self.group_rows(items)
        return self.rows_to_csv(grouped)


def main():
    parser = argparse.ArgumentParser(description='Extract tables and key-value pairs from image to CSV (stdout only).')
    parser.add_argument('--input', '-i', required=True, help='Path to input image')
    args = parser.parse_args()

    extractor = OcrCsvExtractor()
    rows = extractor.extract_csv(args.input)

    # Write only CSV content to stdout
    writer = csv.writer(sys.stdout, lineterminator='\n')
    for r in rows:
        writer.writerow(r)


if __name__ == '__main__':
    if len(sys.argv) == 1:
        # Default to the sample image if no args provided
        default_img = r"C:\\Users\\Admin\\Desktop\\CAD_Files\\image2.jpg"
        extractor = OcrCsvExtractor()
        rows = extractor.extract_csv(default_img)
        writer = csv.writer(sys.stdout, lineterminator='\n')
        for r in rows:
            writer.writerow(r)
    else:
        main()


