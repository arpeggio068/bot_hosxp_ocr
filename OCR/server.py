import argparse
import json
import math
import os
import ctypes
from pathlib import Path
import threading
import time
import traceback
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlsplit, parse_qs
from ocr_core import match_lines, restore_box, fuzzy_lines, distinct_hits

if os.name == "nt":
    ctypes.windll.kernel32.SetConsoleTitleW("AutoIt OCR")

# Skip host connectivity probes. Missing models still require a first download.
os.environ.setdefault('PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK', 'True')
BASE = Path(__file__).resolve().parent
ENGINE = None
LOCK = threading.Lock()
MAX_BYTES = 8 * 1024 * 1024


def load_engine():
    from paddleocr import PaddleOCR
    options = dict(
        text_detection_model_name='PP-OCRv5_mobile_det',
        text_recognition_model_name='th_PP-OCRv5_mobile_rec',
        use_doc_orientation_classify=False,
        use_doc_unwarping=False,
        use_textline_orientation=False,
        device='cpu', enable_mkldnn=False, cpu_threads=4,
    )
    # Explicit local model folders take priority over Paddle's normal cache.
    for name, key in [('PP-OCRv5_mobile_det', 'text_detection_model_dir'),
                      ('th_PP-OCRv5_mobile_rec', 'text_recognition_model_dir')]:
        folder = BASE / 'models' / name
        if folder.is_dir():
            options[key] = str(folder)
    print('Loading PaddleOCR Thai PP-OCRv5 on CPU. First run downloads models.', flush=True)
    return PaddleOCR(**options)


def variants(image):
    import cv2
    # No color thresholds, no fixed UI colors, no cropping by background color.
    yield 'original', image, 1.0, 0
    height, width = image.shape[:2]
    scale = min(2.0, 2400.0 / max(height, width))
    # Use actual resize factors equally on both axes (rounded dimensions corrected below).
    width2, height2 = max(1, round(width * scale)), max(1, round(height * scale))
    enlarged = cv2.resize(image, (width2, height2), interpolation=cv2.INTER_CUBIC)
    gray = cv2.cvtColor(enlarged, cv2.COLOR_BGR2GRAY)
    contrast = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8)).apply(gray)
    padding = 16
    for name, pixels in [('contrast', contrast), ('inverted', 255 - contrast)]:
        pixels = cv2.copyMakeBorder(pixels, padding, padding, padding, padding, cv2.BORDER_CONSTANT, value=255)
        yield name, cv2.cvtColor(pixels, cv2.COLOR_GRAY2BGR), (width2 / width, height2 / height), padding
    threshold = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 9)
    for name, pixels in [('adaptive', threshold), ('adaptive_inverted', 255 - threshold)]:
        pixels = cv2.copyMakeBorder(pixels, padding, padding, padding, padding, cv2.BORDER_CONSTANT, value=255)
        yield name, cv2.cvtColor(pixels, cv2.COLOR_GRAY2BGR), (width2 / width, height2 / height), padding


def read_variant(engine, pixels, scale, padding, width, height):
    lines = []
    for result in engine.predict(pixels):
        texts = result['rec_texts']
        scores = result['rec_scores']
        boxes = result['rec_boxes']
        for text, score, box in zip(texts, scores, boxes):
            sx, sy = scale if isinstance(scale, tuple) else (scale, scale)
            # Separate factors preserve coordinates when resize dimensions were rounded.
            x0, y0, x1, y1 = box
            restored = [(float(x0) - padding) / sx, (float(y0) - padding) / sy,
                        (float(x1) - padding) / sx, (float(y1) - padding) / sy]
            restored = restore_box(restored, 1, 0, width, height)
            lines.append(dict(text=str(text), confidence=round(float(score) * 100, 2), box=restored))
    return lines


# RAM cache only: restarting Python clears learned templates.
CACHE = {}
CACHE_THRESHOLD = 0.97
CACHE_MAX_ENTRIES = 32


def cache_lookup(image, key, minimum):
    import cv2
    import numpy as np
    entry = CACHE.get(key)
    if entry is None or entry['confidence'] < minimum:
        return None
    template = entry['template']
    h, w = template.shape[:2]
    if h > image.shape[0] or w > image.shape[1]:
        return None
    scores = cv2.matchTemplate(image, template, cv2.TM_CCOEFF_NORMED)
    if not np.isfinite(scores).all():
        return None
    _, score, _, location = cv2.minMaxLoc(scores)
    if score < CACHE_THRESHOLD:
        return None
    # Suppress the same location's correlation peak, then reject another match.
    x, y = location
    masked = scores.copy()
    masked[max(0,y-h//2):y+h//2+1, max(0,x-w//2):x+w//2+1] = -1
    if float(masked.max()) >= CACHE_THRESHOLD:
        return None  # let OCR resolve possible duplicates
    return dict(status='MATCH', x=x+entry['offset'][0], y=y+entry['offset'][1],
                confidence=entry['confidence'], text=entry['text'],
                similarity=entry['similarity'], match_method=entry['method'],
                source='opencv_cache', cache_score=round(score, 4), variant='cache', attempts=[])


def cache_store(image, key, hit):
    import numpy as np
    x0,y0,x1,y1 = hit['box']
    left,top = max(0,int(x0)-3),max(0,int(y0)-3)
    right,bottom = min(image.shape[1],math.ceil(x1)+3),min(image.shape[0],math.ceil(y1)+3)
    template = image[top:bottom,left:right].copy()
    if template.size == 0 or float(np.std(template)) < 8:
        return
    if len(CACHE) >= CACHE_MAX_ENTRIES:
        CACHE.pop(next(iter(CACHE)))
    CACHE[key] = dict(template=template, offset=((x0+x1)/2-left,(y0+y1)/2-top),
                      confidence=hit['confidence'], text=hit['text'],
                      similarity=hit['similarity'], method=hit['match_method'])


def recognize(image_bytes, target, minimum, engine):
    import cv2
    import numpy as np
    started = time.perf_counter()
    timing = dict(decode_ms=0, cache_ms=0, preprocess_ms=0, ocr_ms=0, exact_ms=0, fuzzy_ms=0)
    tick = time.perf_counter()
    image = cv2.imdecode(np.frombuffer(image_bytes, dtype=np.uint8), cv2.IMREAD_COLOR)
    if image is None:
        raise ValueError('Cannot decode image')
    height,width = image.shape[:2]
    if height*width > 4_000_000:
        raise ValueError('Image too large; use an area under 4 megapixels')
    timing['decode_ms'] = (time.perf_counter()-tick)*1000
    key = (target, float(minimum), width, height)
    tick = time.perf_counter()
    cached = cache_lookup(image,key,minimum)
    timing['cache_ms'] = (time.perf_counter()-tick)*1000

    def finish(result):
        result['timings_ms'] = {k:round(v,2) for k,v in timing.items()}
        result['seconds'] = round(time.perf_counter()-started,3)
        return result

    if cached:
        return finish(cached)
    # A failed visual match invalidates the old template.
    CACHE.pop(key,None)
    attempts = []
    iterator = iter(variants(image))
    for _ in range(2):
        tick = time.perf_counter()
        try:
            name,pixels,scale,padding = next(iterator)
        except StopIteration:
            break
        timing['preprocess_ms'] += (time.perf_counter()-tick)*1000
        tick = time.perf_counter()
        lines = read_variant(engine,pixels,scale,padding,width,height)
        ocr_ms = (time.perf_counter()-tick)*1000
        timing['ocr_ms'] += ocr_ms
        attempts.append(dict(variant=name,lines=lines,ocr_ms=round(ocr_ms,2)))
        tick = time.perf_counter()
        exact = match_lines(lines,target,minimum)
        timing['exact_ms'] += (time.perf_counter()-tick)*1000
        # Check all fuzzy candidates too so an exact label plus a near duplicate
        # does not silently select one position.
        tick = time.perf_counter()
        fuzzy = fuzzy_lines(lines,target,minimum)
        timing['fuzzy_ms'] += (time.perf_counter()-tick)*1000
        exact_hits = [dict(hit,similarity=100,match_method='exact') for hit in exact]
        hits = distinct_hits(exact_hits + sorted(fuzzy,key=lambda h:h['similarity'],reverse=True))
        repeated_exact = len(exact) > len(distinct_hits(exact))
        if hits:
            hit = hits[0]
            box = hit['box']
            status = 'MATCH' if len(hits)==1 and not repeated_exact else 'AMBIGUOUS'
            if status == 'MATCH':
                cache_store(image,key,hit)
            return finish(dict(status=status,x=(box[0]+box[2])/2,y=(box[1]+box[3])/2,
                               confidence=hit['confidence'],text=hit['text'],
                               similarity=hit['similarity'],match_method=hit['match_method'],
                               source='ocr',variant=name,attempts=attempts))
    return finish(dict(status='NO_MATCH',x=0,y=0,confidence=0,similarity=0,
                       match_method='none',source='ocr',variant='auto',attempts=attempts,
                       text=' | '.join(' '.join(line['text'] for line in a['lines']) for a in attempts)))


def as_tsv(result):
    text = result['text'].replace('\t', ' ').replace('\r', ' ').replace('\n', ' ')
    return '\t'.join(str(result[k]) for k in ['status', 'x', 'y', 'confidence']) + '\t' + text + '\t' + str(result.get('similarity', 0)) + '\t' + result.get('match_method', 'none') + '\t' + json.dumps(result.get('timings_ms', {}), separators=(',', ':')) + '\t' + result.get('source', 'ocr')


class Handler(BaseHTTPRequestHandler):
    def reply(self, status, text):
        data = text.encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'text/plain; charset=utf-8')
        self.send_header('Content-Length', str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        if urlsplit(self.path).path == '/health':
            self.reply(200, 'READY PaddleOCR Thai')
        else:
            self.reply(404, 'NOT_FOUND')

    def do_POST(self):
        url = urlsplit(self.path)
        if url.path != '/ocr':
            return self.reply(404, 'NOT_FOUND')
        try:
            query = parse_qs(url.query)
            target = bytes.fromhex(query.get('target', [''])[0]).decode('utf-8')
            minimum = float(query.get('confidence', ['60'])[0])
            size = int(self.headers.get('Content-Length', '0'))
            if not target.strip() or len(target) > 500 or not math.isfinite(minimum) or not 0 <= minimum <= 100 or not 0 < size <= MAX_BYTES:
                raise ValueError('Invalid target, confidence, or image size')
        except (ValueError, UnicodeError) as error:
            return self.reply(400, str(error))
        if not LOCK.acquire(blocking=False):
            return self.reply(409, 'BUSY: OCR request already running')
        try:
            self.connection.settimeout(30)
            request_started = time.perf_counter()
            image = self.rfile.read(size)
            upload_ms = (time.perf_counter()-request_started)*1000
            if len(image) != size:
                raise ValueError('Incomplete image upload')
            result = recognize(image, target, minimum, ENGINE)
            result['timings_ms']['upload_read_ms'] = round(upload_ms,2)
            result['timings_ms']['request_before_reply_ms'] = round((time.perf_counter()-request_started)*1000,2)
            print(json.dumps({k:v for k,v in result.items() if k != 'attempts'}, ensure_ascii=False), flush=True)
            self.reply(200, as_tsv(result))
        except ValueError as error:
            self.reply(400, str(error))
        except Exception:
            traceback.print_exc()
            self.reply(500, 'OCR_ERROR: see Python terminal')
        finally:
            LOCK.release()


def main():
    global ENGINE
    parser = argparse.ArgumentParser()
    parser.add_argument('--image', help='Test a saved image directly without AutoIt')
    parser.add_argument('--target', default='เคลือบฟลูออไรด์วานิช')
    parser.add_argument('--confidence', type=float, default=60)
    args = parser.parse_args()
    engine_started = time.perf_counter()
    ENGINE = load_engine()
    print("Model initialization_ms=" + str(round((time.perf_counter()-engine_started)*1000, 2)), flush=True)
    if args.image:
        result = recognize(Path(args.image).read_bytes(), args.target, args.confidence, ENGINE)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return
    server = ThreadingHTTPServer(('127.0.0.1', 3078), Handler)
    print('READY: http://127.0.0.1:3078 — press F1 in AutoIt', flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == '__main__':
    main()
