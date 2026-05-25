#!/usr/bin/env python3
"""
Local server for ppt_compositor.html + backend pipeline integration.

Usage:
  pip install flask          (or: venv/bin/pip install flask)
  python server.py           (or: venv/bin/python3 server.py)
  Open http://localhost:5000
"""
from flask import Flask, request, Response, send_file, jsonify
import subprocess, json, os
from pathlib import Path

BASE_DIR = Path(__file__).parent.resolve()
VENV_PY  = str(BASE_DIR / 'venv/bin/python3')
PIPELINE = str(BASE_DIR / 'backend/pipeline.py')

MEDIA_EXTS = {'.mp4', '.mov', '.avi', '.mkv', '.webm', '.pptx'}

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 2 * 1024 * 1024 * 1024  # 2 GB upload limit


@app.route('/')
def index():
    return send_file(str(BASE_DIR / 'ppt_compositor.html'))


@app.route('/api/files')
def list_files():
    files = []
    for p in BASE_DIR.iterdir():
        if p.is_file() and p.suffix.lower() in MEDIA_EXTS:
            files.append({'name': p.name, 'ext': p.suffix.lower()})
    files.sort(key=lambda x: x['name'].lower())
    return jsonify(files)


@app.route('/api/run', methods=['POST'])
def run_pipeline():
    data     = request.get_json(force=True)
    faculty  = data.get('faculty_video', '').strip()
    slides   = data.get('slides_source', '').strip()
    output   = data.get('output', 'output.mp4').strip() or 'output.mp4'

    if not faculty or not slides:
        return jsonify({'error': 'faculty_video and slides_source are required'}), 400

    fac_path = (BASE_DIR / faculty).resolve()
    sli_path = (BASE_DIR / slides).resolve()
    out_path = (BASE_DIR / output).resolve()

    for p, label in ((fac_path, 'Faculty video'), (sli_path, 'Slides source')):
        if not str(p).startswith(str(BASE_DIR)):
            return jsonify({'error': f'{label}: path traversal not allowed'}), 400
        if not p.exists():
            return jsonify({'error': f'{label} not found: {p.name}'}), 400

    # sensitivity: HTML slider is 1-20, pipeline expects float threshold 0.02–0.10
    raw_sens = data.get('sensitivity', 5)
    if isinstance(raw_sens, (int, float)) and 1 <= raw_sens <= 20:
        sensitivity = round(0.02 + (raw_sens - 1) * (0.08 / 19), 4)
    else:
        sensitivity = raw_sens  # already a float from a direct API call

    cmd = [
        VENV_PY, PIPELINE,
        str(fac_path), str(sli_path),
        '-o',             str(out_path),
        '--key',          str(data.get('key', '#00b140')),
        '--sim',          str(int(data.get('sim', 35))),
        '--smooth',       str(int(data.get('smooth', 8))),
        '--spill',        str(round(float(data.get('spill', 0.30)), 3)),
        '--sensitivity',  str(sensitivity),
        '--lerp-speed',   str(round(float(data.get('lerp_speed', 0.06)), 3)),
        '--center-scale', str(round(float(data.get('center_scale', 0.80)), 3)),
    ]
    start_slide = int(data.get('start_slide', 0))
    if start_slide > 0:
        cmd += ['--start-slide', str(start_slide)]

    timings = data.get('timings', '').strip()
    if timings:
        t_path = (BASE_DIR / timings).resolve()
        if str(t_path).startswith(str(BASE_DIR)) and t_path.exists():
            cmd += ['-t', str(t_path)]

    sync_timings = data.get('sync_timings', '').strip()
    if sync_timings:
        s_path = (BASE_DIR / sync_timings).resolve()
        if str(s_path).startswith(str(BASE_DIR)) and s_path.exists():
            cmd += ['--sync', str(s_path)]

    def generate():
        yield f"data: {json.dumps({'type': 'start'})}\n\n"
        try:
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, bufsize=1, cwd=str(BASE_DIR)
            )
            for line in proc.stdout:
                yield f"data: {json.dumps({'type': 'log', 'line': line.rstrip()})}\n\n"
            proc.wait()
            if proc.returncode == 0:
                yield f"data: {json.dumps({'type': 'done', 'output': output})}\n\n"
            else:
                yield f"data: {json.dumps({'type': 'error', 'code': proc.returncode})}\n\n"
        except Exception as exc:
            yield f"data: {json.dumps({'type': 'error', 'msg': str(exc)})}\n\n"

    return Response(
        generate(),
        content_type='text/event-stream',
        headers={'X-Accel-Buffering': 'no', 'Cache-Control': 'no-cache'}
    )


@app.route('/api/sync', methods=['POST'])
def sync_videos():
    """Run speech_sync.py to generate sync timings from faculty audio + VSB scene detection."""
    data    = request.get_json(force=True)
    faculty = data.get('faculty_video', '').strip()
    vsb     = data.get('vsb_video', '').strip()
    output  = data.get('output', 'sync_timings.json').strip() or 'sync_timings.json'

    if not faculty or not vsb:
        return jsonify({'error': 'faculty_video and vsb_video are required'}), 400

    fac_path = (BASE_DIR / faculty).resolve()
    vsb_path = (BASE_DIR / vsb).resolve()
    out_path = (BASE_DIR / output).resolve()

    for p, label in ((fac_path, 'Faculty video'), (vsb_path, 'VSB video')):
        if not str(p).startswith(str(BASE_DIR)):
            return jsonify({'error': f'{label}: path traversal not allowed'}), 400
        if not p.exists():
            return jsonify({'error': f'{label} not found: {p.name}'}), 400

    SYNC_SCRIPT = str(BASE_DIR / 'backend' / 'speech_sync.py')
    cmd = [
        VENV_PY, SYNC_SCRIPT,
        str(fac_path), str(vsb_path),
        '-o', str(out_path),
    ]

    def generate():
        yield f"data: {json.dumps({'type': 'start'})}\n\n"
        try:
            proc = subprocess.Popen(
                cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, bufsize=1, cwd=str(BASE_DIR)
            )
            for line in proc.stdout:
                yield f"data: {json.dumps({'type': 'log', 'line': line.rstrip()})}\n\n"
            proc.wait()
            if proc.returncode == 0:
                yield f"data: {json.dumps({'type': 'done', 'timings_file': output})}\n\n"
            else:
                yield f"data: {json.dumps({'type': 'error', 'code': proc.returncode})}\n\n"
        except Exception as exc:
            yield f"data: {json.dumps({'type': 'error', 'msg': str(exc)})}\n\n"

    return Response(
        generate(),
        content_type='text/event-stream',
        headers={'X-Accel-Buffering': 'no', 'Cache-Control': 'no-cache'}
    )


@app.route('/api/pptx_sync', methods=['POST'])
def pptx_sync():
    """Run pptx_sync.py: match faculty speech to PPTX slide text, write timings.json."""
    data    = request.get_json(force=True)
    faculty = data.get('faculty_video', '').strip()
    pptx    = data.get('pptx_file', '').strip()
    output  = data.get('output', 'timings.json').strip() or 'timings.json'

    if not faculty or not pptx:
        return jsonify({'error': 'faculty_video and pptx_file are required'}), 400

    fac_path  = (BASE_DIR / faculty).resolve()
    pptx_path = (BASE_DIR / pptx).resolve()
    out_path  = (BASE_DIR / output).resolve()

    for p, label in ((fac_path, 'Faculty video'), (pptx_path, 'PPTX file')):
        if not str(p).startswith(str(BASE_DIR)):
            return jsonify({'error': f'{label}: path traversal not allowed'}), 400
        if not p.exists():
            return jsonify({'error': f'{label} not found: {p.name}'}), 400

    SYNC_SCRIPT = str(BASE_DIR / 'backend' / 'pptx_sync.py')
    cmd = [
        VENV_PY, SYNC_SCRIPT,
        str(fac_path), str(pptx_path),
        '-o', str(out_path),
    ]

    def generate():
        yield f"data: {json.dumps({'type': 'start'})}\n\n"
        try:
            proc = subprocess.Popen(
                cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, bufsize=1, cwd=str(BASE_DIR)
            )
            for line in proc.stdout:
                yield f"data: {json.dumps({'type': 'log', 'line': line.rstrip()})}\n\n"
            proc.wait()
            if proc.returncode == 0:
                yield f"data: {json.dumps({'type': 'done', 'timings_file': output})}\n\n"
            else:
                yield f"data: {json.dumps({'type': 'error', 'code': proc.returncode})}\n\n"
        except Exception as exc:
            yield f"data: {json.dumps({'type': 'error', 'msg': str(exc)})}\n\n"

    return Response(
        generate(),
        content_type='text/event-stream',
        headers={'X-Accel-Buffering': 'no', 'Cache-Control': 'no-cache'}
    )


@app.route('/api/convert_pptx', methods=['POST'])
def convert_pptx():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    f = request.files['file']
    if not f.filename.lower().endswith('.pptx'):
        return jsonify({'error': 'File must be .pptx'}), 400

    safe_name = Path(f.filename).name
    pptx_path = BASE_DIR / safe_name
    f.save(str(pptx_path))

    tmp_dir = str(BASE_DIR / '_slides_tmp')

    # Inline conversion: LibreOffice PPTX→PDF, PyMuPDF PDF→PNGs
    converter = str(BASE_DIR / 'backend' / '_pptx_to_png.py')
    result = subprocess.run(
        [VENV_PY, converter, str(pptx_path), tmp_dir],
        capture_output=True, text=True
    )
    if result.returncode != 0:
        return jsonify({'error': result.stderr.strip() or 'Conversion failed'}), 500

    import json as _json
    raw = _json.loads(result.stdout.strip())
    # _pptx_to_png.py outputs {"slides": [...], "person_boxes": [...]}
    if isinstance(raw, list):
        images, person_boxes = raw, [None] * len(raw)
    else:
        images      = raw['slides']
        person_boxes = raw.get('person_boxes', [None] * len(images))
    rel = [str(Path(p).relative_to(BASE_DIR)) for p in images]
    return jsonify({'slides': rel, 'person_boxes': person_boxes})


@app.route('/api/detect_person_frame', methods=['POST'])
def detect_person_frame():
    """Detect a person in a base64-encoded JPEG/PNG frame posted by the browser."""
    import base64, numpy as np, cv2, sys as _sys
    _sys.path.insert(0, str(BASE_DIR / 'backend'))
    from pipeline import detect_person_box

    data  = request.get_json(force=True)
    b64   = data.get('image', '')
    if ',' in b64:
        b64 = b64.split(',', 1)[1]
    try:
        img_bytes = base64.b64decode(b64)
        arr = np.frombuffer(img_bytes, dtype=np.uint8)
        img = cv2.imdecode(arr, cv2.IMREAD_COLOR)
        box = detect_person_box(img) if img is not None else None
    except Exception:
        box = None
    return jsonify({'personBox': box})


@app.route('/api/slide_image/<path:filename>')
def slide_image(filename):
    path = (BASE_DIR / filename).resolve()
    if not str(path).startswith(str(BASE_DIR)) or not path.exists():
        return jsonify({'error': 'Not found'}), 404
    return send_file(str(path))


@app.route('/api/download/<path:filename>')
def download(filename):
    path = (BASE_DIR / filename).resolve()
    if not str(path).startswith(str(BASE_DIR)) or not path.exists():
        return jsonify({'error': 'File not found'}), 404
    return send_file(str(path), as_attachment=True, download_name=path.name)


@app.route('/api/convert', methods=['POST'])
def convert_webm():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    f = request.files['file']

    webm_path = (BASE_DIR / 'recorded_output.webm').resolve()
    mp4_name  = 'recorded_output.mp4'
    mp4_path  = (BASE_DIR / mp4_name).resolve()

    f.save(str(webm_path))

    def generate():
        yield f"data: {json.dumps({'type': 'log', 'line': 'Converting WebM → MP4…'})}\n\n"
        try:
            cmd = [
                'ffmpeg', '-y', '-i', str(webm_path),
                '-c:v', 'libx264', '-preset', 'fast', '-crf', '23',
                '-c:a', 'aac', '-b:a', '128k',
                '-pix_fmt', 'yuv420p',
                str(mp4_path)
            ]
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, bufsize=1, cwd=str(BASE_DIR)
            )
            for line in proc.stdout:
                yield f"data: {json.dumps({'type': 'log', 'line': line.rstrip()})}\n\n"
            proc.wait()
            if proc.returncode == 0:
                yield f"data: {json.dumps({'type': 'done', 'output': mp4_name})}\n\n"
            else:
                yield f"data: {json.dumps({'type': 'error', 'code': proc.returncode})}\n\n"
        except Exception as exc:
            yield f"data: {json.dumps({'type': 'error', 'msg': str(exc)})}\n\n"

    return Response(
        generate(),
        content_type='text/event-stream',
        headers={'X-Accel-Buffering': 'no', 'Cache-Control': 'no-cache'}
    )


@app.route('/api/save_timings', methods=['POST'])
def save_timings():
    data     = request.get_json(force=True)
    filename = data.get('filename', '').strip()
    timings  = data.get('timings')

    if not filename or timings is None:
        return jsonify({'error': 'filename and timings are required'}), 400
    if not filename.endswith('.json'):
        filename += '.json'

    path = (BASE_DIR / filename).resolve()
    if not str(path).startswith(str(BASE_DIR)):
        return jsonify({'error': 'Path traversal not allowed'}), 400

    with open(path, 'w') as f:
        json.dump(timings, f, indent=2)

    return jsonify({'saved': filename})


@app.route('/api/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    f = request.files['file']
    safe_name = Path(f.filename).name
    if not safe_name:
        return jsonify({'error': 'Invalid filename'}), 400
    dest = (BASE_DIR / safe_name).resolve()
    if not str(dest).startswith(str(BASE_DIR)):
        return jsonify({'error': 'Path traversal not allowed'}), 400
    f.save(str(dest))
    return jsonify({'filename': safe_name})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5001))
    print("ppt·key  –  local compositor server")
    print(f"Open  http://localhost:{port}")
    print(f"Base  {BASE_DIR}")
    app.run(debug=False, host='0.0.0.0', port=port, threaded=True)
