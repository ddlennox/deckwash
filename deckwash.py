#!/usr/bin/env python3
"""
DeckWash 🌊
A local web app for rebranding case study PPTX files.
Drop your decks. Get them rebranded.

Usage:
    python3 deckwash.py
Then open http://localhost:5001 in your browser.
"""

import sys
import os
import io
import json
import uuid
import queue
import threading
import tempfile
import webbrowser
from pathlib import Path

from flask import Flask, request, send_file, jsonify, Response, stream_with_context, session, redirect, url_for

# ── App setup ─────────────────────────────────────────────────────────────────
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500 MB max upload
app.secret_key = os.environ.get('SECRET_KEY', 'deckwash-local-secret-key')

# ── Auth helper ───────────────────────────────────────────────────────────────
def get_password():
    return os.environ.get('DECKWASH_PASSWORD', 'deckwash')

# In-memory job store: job_id -> job dict
jobs = {}
jobs_lock = threading.Lock()

# ── Cloud vs local detection ──────────────────────────────────────────────────
# Render sets a RENDER env var automatically. When running locally, we save
# directly to new_case_studies/. On the cloud, we use a temp dir + download.
IS_CLOUD  = bool(os.environ.get('RENDER'))
BASE_DIR  = Path(__file__).parent
OUTPUT_DIR = None if IS_CLOUD else BASE_DIR / 'new_case_studies'
if OUTPUT_DIR:
    OUTPUT_DIR.mkdir(exist_ok=True)

# ── stdout capture ────────────────────────────────────────────────────────────
class QueueWriter:
    """Redirects print() calls into a queue for SSE streaming."""
    def __init__(self, q):
        self.q = q
        self.encoding = 'utf-8'

    def write(self, text):
        text = text.strip()
        if text:
            self.q.put({'type': 'log', 'message': text})

    def flush(self):
        pass


# ── Conversion worker ─────────────────────────────────────────────────────────
def run_conversion(job_id, input_path, output_path):
    job = jobs[job_id]
    q = job['queue']

    old_stdout = sys.stdout
    sys.stdout = QueueWriter(q)

    try:
        # Import here so stdout is already redirected
        sys.path.insert(0, str(BASE_DIR))
        from convert_case_study import convert_pptx
        convert_pptx(str(input_path), str(output_path))
        out_name = Path(output_path).name
        if IS_CLOUD:
            # Read converted file into memory so it survives temp dir cleanup
            with open(str(output_path), 'rb') as fh:
                job['file_bytes'] = fh.read()
            job['out_name'] = out_name
        job['status'] = 'done'
        job['output_path'] = str(output_path)
        q.put({'type': 'done', 'filename': out_name, 'cloud': IS_CLOUD})
    except Exception as e:
        job['status'] = 'error'
        q.put({'type': 'error', 'message': str(e)})
    finally:
        sys.stdout = old_stdout


# ── Login page HTML ───────────────────────────────────────────────────────────
LOGIN_HTML = '''<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>DeckWash — Login</title>
<style>
  @font-face {
    font-family: 'Obviously Narrow';
    src: url('/fonts/ObviouslyNarrow-Bold 1.otf') format('opentype');
    font-weight: 700;
  }
  @font-face {
    font-family: 'Galvji';
    src: url('/fonts/galvji.ttf') format('truetype');
    font-weight: 400;
  }
  @font-face {
    font-family: 'Galvji';
    src: url('/fonts/galvji-bold.ttf') format('truetype');
    font-weight: 700;
  }

  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    background: #231F20;
    min-height: 100vh;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    font-family: 'Galvji', Georgia, serif;
  }

  .logo {
    font-family: 'Obviously Narrow', sans-serif;
    font-size: 3rem;
    font-weight: 700;
    color: #FFFFFF;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-bottom: 6px;
  }

  .wave { font-size: 2rem; margin-right: 8px; }

  .sub {
    font-family: 'Galvji', Georgia, serif;
    font-size: 0.9rem;
    color: #46DE66;
    margin-bottom: 40px;
    letter-spacing: 0.02em;
  }

  .card {
    background: #FFFAF0;
    border-radius: 6px;
    padding: 36px 40px;
    width: 100%;
    max-width: 380px;
  }

  .card h2 {
    font-family: 'Obviously Narrow', sans-serif;
    font-size: 1.3rem;
    font-weight: 700;
    color: #231F20;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    margin-bottom: 20px;
  }

  label {
    display: block;
    font-size: 0.8rem;
    font-weight: 700;
    color: #231F20;
    letter-spacing: 0.05em;
    text-transform: uppercase;
    margin-bottom: 6px;
  }

  input[type="password"] {
    width: 100%;
    padding: 12px 14px;
    border: 2px solid #E0D9CE;
    border-radius: 4px;
    font-family: 'Galvji', Georgia, serif;
    font-size: 1rem;
    color: #231F20;
    background: #FFFFFF;
    outline: none;
    transition: border-color 0.15s;
    margin-bottom: 20px;
  }

  input[type="password"]:focus { border-color: #46DE66; }

  button {
    width: 100%;
    background: #46DE66;
    color: #231F20;
    border: none;
    border-radius: 4px;
    padding: 13px;
    font-family: 'Obviously Narrow', sans-serif;
    font-size: 1rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    cursor: pointer;
    transition: background 0.15s;
  }

  button:hover { background: #30c452; }

  .error {
    background: #fde8e8;
    border: 1px solid #f5b7b7;
    color: #c0392b;
    border-radius: 4px;
    padding: 10px 14px;
    font-size: 0.85rem;
    margin-bottom: 16px;
  }
</style>
</head>
<body>
  <div class="logo"><span class="wave">🌊</span>DeckWash</div>
  <div class="sub">A Sense rebrand tool</div>
  <div class="card">
    <h2>Sign in</h2>
    {% if error %}<div class="error">Incorrect password — try again.</div>{% endif %}
    <form method="POST" action="/login">
      <label for="password">Password</label>
      <input type="password" id="password" name="password" autofocus placeholder="Enter password">
      <button type="submit">Let me in</button>
    </form>
  </div>
</body>
</html>'''


# ── Routes ────────────────────────────────────────────────────────────────────
@app.route('/login', methods=['GET', 'POST'])
def login():
    error = False
    if request.method == 'POST':
        if request.form.get('password') == get_password():
            session['logged_in'] = True
            return redirect(url_for('index'))
        error = True
    from flask import render_template_string
    return render_template_string(LOGIN_HTML, error=error)


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


@app.route('/')
def index():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    return HTML


@app.route('/convert', methods=['POST'])
def convert():
    if not session.get('logged_in'):
        return jsonify({'error': 'Not authenticated'}), 401
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400

    f = request.files['file']
    if not f.filename.lower().endswith('.pptx'):
        return jsonify({'error': 'Only .pptx files are supported'}), 400

    job_id = str(uuid.uuid4())
    tmpdir = Path(tempfile.mkdtemp())
    input_path = tmpdir / f.filename
    stem = Path(f.filename).stem
    output_path = (tmpdir if IS_CLOUD else OUTPUT_DIR) / f'{stem}_rebranded.pptx'

    f.save(str(input_path))

    with jobs_lock:
        jobs[job_id] = {
            'status': 'running',
            'queue': queue.Queue(),
            'output_path': None,
            'filename': f.filename,
        }

    t = threading.Thread(target=run_conversion, args=(job_id, input_path, output_path), daemon=True)
    t.start()

    return jsonify({'job_id': job_id})


@app.route('/status/<job_id>')
def status(job_id):
    if not session.get('logged_in'):
        return Response("data: {\"type\":\"error\",\"message\":\"Not authenticated\"}\n\n",
                        mimetype='text/event-stream')
    job = jobs.get(job_id)
    if not job:
        return Response("data: {\"type\":\"error\",\"message\":\"Job not found\"}\n\n",
                        mimetype='text/event-stream')

    def generate():
        q = job['queue']
        while True:
            try:
                msg = q.get(timeout=60)
                yield f"data: {json.dumps(msg)}\n\n"
                if msg['type'] in ('done', 'error'):
                    break
            except queue.Empty:
                yield "data: {\"type\":\"ping\"}\n\n"

    return Response(
        stream_with_context(generate()),
        mimetype='text/event-stream',
        headers={'Cache-Control': 'no-cache', 'X-Accel-Buffering': 'no'}
    )


@app.route('/download/<job_id>')
def download(job_id):
    if not session.get('logged_in'):
        return 'Not authenticated', 401
    job = jobs.get(job_id)
    if not job:
        return 'File not found', 404
    if IS_CLOUD:
        # Serve from in-memory bytes — reliable on ephemeral filesystems
        file_bytes = job.get('file_bytes')
        out_name   = job.get('out_name', 'rebranded.pptx')
        if not file_bytes:
            return 'File not found', 404
        return send_file(
            io.BytesIO(file_bytes),
            as_attachment=True,
            download_name=out_name,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
    # Local: serve from disk
    path = job.get('output_path')
    if not path or not os.path.exists(path):
        return 'File not found', 404
    return send_file(path, as_attachment=True, download_name=Path(path).name)


# ── Font serving ─────────────────────────────────────────────────────────────
@app.route('/fonts/<path:filename>')
def serve_font(filename):
    return send_file(str(BASE_DIR / filename), mimetype='font/truetype')


# ── HTML template ─────────────────────────────────────────────────────────────
HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>DeckWash</title>
<style>
  @font-face {
    font-family: 'Obviously Narrow';
    src: url('/fonts/ObviouslyNarrow-Bold 1.otf') format('opentype');
    font-weight: 700;
    font-style: normal;
  }
  @font-face {
    font-family: 'Galvji';
    src: url('/fonts/galvji.ttf') format('truetype');
    font-weight: 400;
    font-style: normal;
  }
  @font-face {
    font-family: 'Galvji';
    src: url('/fonts/galvji-bold.ttf') format('truetype');
    font-weight: 700;
    font-style: normal;
  }
  @font-face {
    font-family: 'Galvji';
    src: url('/fonts/galvji-oblique.ttf') format('truetype');
    font-weight: 400;
    font-style: italic;
  }

  * { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    font-family: 'Galvji', Georgia, serif;
    background: #FFFAF0;
    min-height: 100vh;
    display: flex;
    flex-direction: column;
  }

  /* ── Header ── */
  header {
    background: #231F20;
    padding: 20px 48px 16px;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: space-between;
    text-align: center;
    min-height: 90px;
  }

  .header-left {
    display: flex;
    align-items: baseline;
    gap: 12px;
    justify-content: center;
  }

  header h1 {
    font-family: 'Obviously Narrow', sans-serif;
    font-size: 2.2rem;
    font-weight: 700;
    color: #FFFFFF;
    letter-spacing: 0.5px;
    text-transform: uppercase;
  }

  header .wave {
    font-size: 1.4rem;
    line-height: 1;
  }


  /* ── Hero strip ── */
  .hero {
    background: #231F20;
    padding: 48px 48px 56px;
    border-bottom: 4px solid #46DE66;
    text-align: center;
  }

  .hero h2 {
    font-family: 'Obviously Narrow', sans-serif;
    font-size: 3.2rem;
    font-weight: 700;
    color: #FFFFFF;
    text-transform: uppercase;
    line-height: 1;
    margin-bottom: 12px;
    letter-spacing: 0.5px;
  }

  .hero p {
    font-family: 'Galvji', Georgia, serif;
    font-size: 1rem;
    color: rgba(255,255,255,0.65);
    max-width: 480px;
    line-height: 1.6;
    margin: 0 auto;
  }

  /* ── Main content ── */
  main {
    flex: 1;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 52px 24px 60px;
    gap: 32px;
  }

  /* ── Drop zone ── */
  #drop-zone {
    width: 100%;
    max-width: 680px;
    border: 2px dashed #C8C4B0;
    border-radius: 4px;
    background: #FFFFFF;
    padding: 56px 40px;
    text-align: center;
    cursor: pointer;
    transition: border-color 0.2s, background 0.2s, transform 0.15s;
  }

  #drop-zone.drag-over {
    border-color: #46DE66;
    background: #F5FFF8;
    transform: scale(1.01);
  }

  #drop-zone .drop-icon {
    width: 52px;
    height: 52px;
    background: #231F20;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    margin: 0 auto 20px;
    font-size: 1.4rem;
  }

  #drop-zone h3 {
    font-family: 'Obviously Narrow', sans-serif;
    font-size: 1.5rem;
    font-weight: 700;
    color: #231F20;
    text-transform: uppercase;
    letter-spacing: 0.3px;
    margin-bottom: 8px;
  }

  #drop-zone p {
    font-family: 'Galvji', Georgia, serif;
    color: #888;
    font-size: 0.92rem;
    line-height: 1.6;
    margin-bottom: 24px;
  }

  .browse-btn {
    display: inline-block;
    padding: 11px 32px;
    background: #46DE66;
    color: #231F20;
    font-family: 'Obviously Narrow', sans-serif;
    font-size: 1rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    border-radius: 2px;
    cursor: pointer;
    transition: background 0.2s, transform 0.1s;
  }

  .browse-btn:hover { background: #35c455; transform: scale(1.02); }

  #file-input { display: none; }

  /* ── Job list ── */
  #job-list {
    width: 100%;
    max-width: 680px;
    display: flex;
    flex-direction: column;
    gap: 16px;
  }

  .job-card {
    background: #FFFFFF;
    border-radius: 4px;
    border-left: 4px solid #231F20;
    padding: 24px 28px;
    box-shadow: 0 2px 16px rgba(35,31,32,0.08);
    animation: slideIn 0.25s ease;
  }

  @keyframes slideIn {
    from { opacity: 0; transform: translateY(10px); }
    to   { opacity: 1; transform: translateY(0); }
  }

  .job-card .job-name {
    font-family: 'Obviously Narrow', sans-serif;
    font-size: 1.05rem;
    font-weight: 700;
    color: #231F20;
    text-transform: uppercase;
    letter-spacing: 0.3px;
    margin-bottom: 4px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }

  .job-card .job-status {
    font-family: 'Galvji', Georgia, serif;
    font-size: 0.85rem;
    color: #999;
    margin-bottom: 14px;
    min-height: 18px;
    transition: color 0.2s;
  }

  .job-card .job-status.done-text  { color: #1a7a36; font-weight: 700; }
  .job-card .job-status.error-text { color: #c0392b; font-weight: 700; }

  /* Progress bar */
  .progress-track {
    height: 6px;
    background: #EDE9DE;
    border-radius: 100px;
    overflow: hidden;
    margin-bottom: 18px;
  }

  .progress-bar {
    height: 100%;
    background: #46DE66;
    border-radius: 100px;
    width: 0%;
    transition: width 0.4s ease;
  }

  .progress-bar.indeterminate {
    animation: indeterminate 1.4s ease infinite;
    width: 38%;
  }

  @keyframes indeterminate {
    0%   { transform: translateX(-130%); }
    100% { transform: translateX(400%); }
  }

  .progress-bar.complete { width: 100% !important; animation: none; }
  .progress-bar.error    { background: #c0392b; width: 100% !important; animation: none; }

  /* Action badge / download */
  .saved-badge {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    padding: 10px 24px;
    background: #46DE66;
    color: #231F20;
    font-family: 'Obviously Narrow', sans-serif;
    font-size: 0.95rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.4px;
    border-radius: 2px;
    text-decoration: none;
    cursor: pointer;
    transition: background 0.2s;
  }

  .saved-badge:hover { background: #35c455; }

  /* Log drawer */
  .log-toggle {
    background: none;
    border: none;
    font-family: 'Galvji', Georgia, serif;
    color: #bbb;
    font-size: 0.78rem;
    cursor: pointer;
    margin-top: 12px;
    padding: 0;
    text-decoration: underline;
    display: block;
  }

  .log-toggle:hover { color: #888; }

  .log-box {
    margin-top: 10px;
    background: #F5F3EE;
    border-radius: 2px;
    padding: 12px 14px;
    font-size: 0.73rem;
    font-family: 'SF Mono', 'Menlo', monospace;
    color: #666;
    max-height: 150px;
    overflow-y: auto;
    display: none;
    line-height: 1.7;
  }

  .log-box.visible { display: block; }

  /* ── Footer ── */
  footer {
    background: #231F20;
    text-align: center;
    padding: 20px 24px;
    font-family: 'Galvji', Georgia, serif;
    font-size: 0.8rem;
    color: rgba(255,255,255,0.35);
    letter-spacing: 0.03em;
  }

  footer strong {
    font-family: 'Obviously Narrow', sans-serif;
    color: rgba(255,255,255,0.6);
    font-size: 0.85rem;
    letter-spacing: 0.5px;
  }
</style>
</head>
<body>

<header>
  <div class="header-left">
    <span class="wave">🌊</span>
    <h1>DeckWash</h1>
  </div>
</header>

<div class="hero">
  <h2>Rebrand in seconds</h2>
  <p>Drop your old case study decks below. DeckWash applies the new brand — fonts, colours, backgrounds — instantly.</p>
</div>

<main>
  <div id="drop-zone">
    <div class="drop-icon">📂</div>
    <h3>Drop PPTX files here</h3>
    <p>Drag one or more files onto this area, or click the button below.<br>
       Your rebranded files will be ready to download in moments.</p>
    <label class="browse-btn" for="file-input">Browse files</label>
    <input type="file" id="file-input" accept=".pptx" multiple>
  </div>

  <div id="job-list"></div>
</main>

<footer><strong>DECKWASH</strong> &nbsp;·&nbsp; A Sense rebrand tool &nbsp;·&nbsp; Drop your decks. Get them rebranded.</footer>

<script>
const dropZone   = document.getElementById('drop-zone');
const fileInput  = document.getElementById('file-input');
const jobList    = document.getElementById('job-list');

['dragenter','dragover'].forEach(e =>
  dropZone.addEventListener(e, ev => { ev.preventDefault(); dropZone.classList.add('drag-over'); })
);
['dragleave','drop'].forEach(e =>
  dropZone.addEventListener(e, ev => { ev.preventDefault(); dropZone.classList.remove('drag-over'); })
);
dropZone.addEventListener('drop', ev => {
  const files = Array.from(ev.dataTransfer.files).filter(f => f.name.endsWith('.pptx'));
  if (files.length === 0) return alert('Please drop .pptx files only.');
  files.forEach(startJob);
});
fileInput.addEventListener('change', () => {
  Array.from(fileInput.files).forEach(startJob);
  fileInput.value = '';
});

function startJob(file) {
  const card = createCard(file.name);
  const formData = new FormData();
  formData.append('file', file);
  setStatus(card, 'Uploading…');
  fetch('/convert', { method: 'POST', body: formData })
    .then(r => r.json())
    .then(data => {
      if (data.error) { showError(card, data.error); return; }
      setStatus(card, 'Starting conversion…');
      streamStatus(card, data.job_id);
    })
    .catch(err => showError(card, err.message));
}

function streamStatus(card, jobId) {
  const es = new EventSource(`/status/${jobId}`);
  let slideCount = 0;

  es.onmessage = ev => {
    const msg = JSON.parse(ev.data);
    if (msg.type === 'ping') return;

    if (msg.type === 'log') {
      appendLog(card, msg.message);
      const foundMatch = msg.message.match(/Found (\\d+) slides/);
      if (foundMatch) slideCount = parseInt(foundMatch[1]);
      const slideMatch = msg.message.match(/Slide (\\d+):/);
      if (slideMatch && slideCount > 0) {
        const pct = Math.round((parseInt(slideMatch[1]) / slideCount) * 85);
        setProgress(card, pct);
        setStatus(card, `Processing slide ${slideMatch[1]} of ${slideCount}…`);
      }
      if (msg.message.toLowerCase().includes('embedding font')) {
        setProgress(card, 90);
        setStatus(card, 'Embedding fonts…');
      }
    }

    if (msg.type === 'done') {
      setProgress(card, 100, 'complete');
      setStatus(card, msg.cloud ? '✓ Ready to download' : '✓ Saved to new case studies folder', 'done-text');
      showDownload(card, jobId, msg.filename, msg.cloud);
      es.close();
    }

    if (msg.type === 'error') {
      setProgress(card, 100, 'error');
      showError(card, msg.message);
      es.close();
    }
  };

  es.onerror = () => { showError(card, 'Connection lost — please try again.'); es.close(); };
}

function createCard(filename) {
  const card = document.createElement('div');
  card.className = 'job-card';
  card.innerHTML = `
    <div class="job-name">${escHtml(filename)}</div>
    <div class="job-status">Preparing…</div>
    <div class="progress-track"><div class="progress-bar indeterminate"></div></div>
    <div class="job-actions"></div>
    <button class="log-toggle" onclick="toggleLog(this)">Show details</button>
    <div class="log-box"></div>
  `;
  jobList.prepend(card);
  return card;
}

function setStatus(card, text, cls) {
  const el = card.querySelector('.job-status');
  el.textContent = text;
  el.className = 'job-status' + (cls ? ' ' + cls : '');
}

function setProgress(card, pct, state) {
  const bar = card.querySelector('.progress-bar');
  bar.classList.remove('indeterminate', 'complete', 'error');
  if (state) bar.classList.add(state);
  bar.style.width = pct + '%';
}

function showDownload(card, jobId, filename, isCloud) {
  const actions = card.querySelector('.job-actions');
  if (isCloud) {
    const btn = document.createElement('a');
    btn.className = 'saved-badge';
    btn.href = `/download/${jobId}`;
    btn.download = filename;
    btn.innerHTML = '⬇ Download ' + escHtml(filename);
    actions.appendChild(btn);
  } else {
    const badge = document.createElement('div');
    badge.className = 'saved-badge';
    badge.innerHTML = '✓ Saved to new case studies folder';
    actions.appendChild(badge);
  }
}

function showError(card, msg) {
  setProgress(card, 100, 'error');
  setStatus(card, '✗ ' + msg, 'error-text');
}

function appendLog(card, text) {
  const box = card.querySelector('.log-box');
  box.textContent += text + '\\n';
  box.scrollTop = box.scrollHeight;
}

function toggleLog(btn) {
  const box = btn.nextElementSibling;
  btn.textContent = box.classList.toggle('visible') ? 'Hide details' : 'Show details';
}

function escHtml(s) {
  return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}
</script>
</body>
</html>"""


# ── Launch ────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    # ── Local-only: auto-install missing packages then restart ────────────────
    import subprocess
    missing = []
    try:
        import flask
    except ImportError:
        missing.append('flask')
    try:
        from lxml import etree
    except ImportError:
        missing.append('lxml')
    if missing:
        print(f"\n❌  Installing missing packages: {' '.join(missing)}…\n")
        r = subprocess.run([sys.executable, '-m', 'pip', 'install',
                            '--break-system-packages', '-q'] + missing, capture_output=True)
        if r.returncode != 0:
            subprocess.run([sys.executable, '-m', 'pip', 'install', '-q'] + missing)
        print("    ✓ Done! Restarting…\n")
        os.execv(sys.executable, [sys.executable] + sys.argv)

    PORT = int(os.environ.get('PORT', 5001))

    if IS_CLOUD:
        print(f"\n🌊  DeckWash starting on port {PORT} (cloud mode)…")
    else:
        print("\n🌊  DeckWash is starting up…")
        print(f"    Opening http://localhost:{PORT} in your browser\n")
        print("    (Keep this window open while you're using DeckWash)")
        print("    Press Ctrl+C to stop.\n")

        def open_browser():
            import time
            time.sleep(1.2)
            webbrowser.open(f'http://localhost:{PORT}')

        threading.Thread(target=open_browser, daemon=True).start()

    app.run(host='0.0.0.0', port=PORT, debug=False, threaded=True)
