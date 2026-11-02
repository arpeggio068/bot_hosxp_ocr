"""Matching and coordinates. No OCR runtime required for unit tests."""
import re
import unicodedata


def normalize(text):
    return re.sub(r'\s+', '', unicodedata.normalize('NFC', text))


def match_lines(lines, target, minimum):
    target = normalize(target)
    matches = []
    for line in lines:
        text = normalize(line['text'])
        count = text.count(target) if target else 0
        if count and line['confidence'] >= minimum:
            # Multiple occurrences in one line are also ambiguous.
            matches.extend([line] * count)
    return matches


def restore_box(box, scale, padding, width, height):
    x0, y0, x1, y1 = [float(x) for x in box]
    return [max(0, min(width - 1, (x0 - padding) / scale)),
            max(0, min(height - 1, (y0 - padding) / scale)),
            max(0, min(width - 1, (x1 - padding) / scale)),
            max(0, min(height - 1, (y1 - padding) / scale))]
