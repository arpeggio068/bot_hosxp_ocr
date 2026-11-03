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

# Similarity is independent of the OCR model's confidence.
FUZZY_MINIMUM = 85.0
MIN_FUZZY_LENGTH = 8
THAI_MARK_COST = 0.35
STRICT_NUMBERS = False


def _cost(c):
    return THAI_MARK_COST if '\u0e31' == c or '\u0e34' <= c <= '\u0e3a' or '\u0e47' <= c <= '\u0e4e' else 1.0


def _digits(s):
    return re.findall(r'(?<![a-z0-9])\d+(?![a-z0-9])', s)


def _english_confusions(text):
    # Apply only inside tokens containing English letters; preserve numeric labels.
    return re.sub(r'[a-z01]+', lambda m: m[0].translate(str.maketrans({'l':'i','1':'i','0':'o'})) if re.search('[a-z]', m[0]) else m[0], text)


def _similarity(a, b):
    previous = [0.0]
    for c in b:
        previous.append(previous[-1] + _cost(c))
    for x in a:
        current = [previous[0] + _cost(x)]
        for j, y in enumerate(b, 1):
            substitution = 0 if x == y else max(_cost(x), _cost(y))
            current.append(min(previous[j] + _cost(x), current[-1] + _cost(y), previous[j-1] + substitution))
        previous = current
    return 100 * max(0, 1 - previous[-1] / max(sum(map(_cost, a)), sum(map(_cost, b)), 1))


def fuzzy_lines(lines, target, minimum):
    target = normalize(target).casefold()
    if len(target) < MIN_FUZZY_LENGTH:
        return []
    hits = []
    for line in lines:
        if line['confidence'] < minimum:
            continue
        text = normalize(line['text']).casefold()
        best = 0
        # Sliding spans support labels embedded in a longer recognized line.
        slack = max(3, int(len(target) * .2))
        for start in range(len(text)):
            for length in range(max(1, len(target)-slack), len(target)+slack+1):
                end = start + length
                if end > len(text):
                    break
                candidate = text[start:end]
                # Never allow a span to cut a numeric token in half.
                if start and text[start-1].isdigit() and candidate[0].isdigit():
                    continue
                if end < len(text) and text[end].isdigit() and candidate[-1].isdigit():
                    continue
                if STRICT_NUMBERS and _digits(candidate) != _digits(target):
                    continue
                best = max(best, _similarity(_english_confusions(target), _english_confusions(candidate)))
        if best >= FUZZY_MINIMUM:
            hits.append(dict(line, similarity=round(best, 2), match_method='fuzzy'))
    return hits


def distinct_hits(hits):
    result = []
    for hit in hits:
        x0,y0,x1,y1 = hit['box']
        duplicate = False
        for other in result:
            a,b,c,d = other['box']
            intersection = max(0,min(x1,c)-max(x0,a))*max(0,min(y1,d)-max(y0,b))
            smaller = min(max(1,(x1-x0)*(y1-y0)),max(1,(c-a)*(d-b)))
            if intersection / smaller >= .6:
                duplicate = True
                break
        if not duplicate:
            result.append(hit)
    return result
