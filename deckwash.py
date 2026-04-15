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

from flask import Flask, request, send_file, jsonify, Response, stream_with_context

# ── App setup ─────────────────────────────────────────────────────────────────
app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500 MB max upload

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
        job['status'] = 'done'
        job['output_path'] = str(output_path)
        q.put({'type': 'done', 'filename': Path(output_path).name, 'cloud': IS_CLOUD})
    except Exception as e:
        job['status'] = 'error'
        q.put({'type': 'error', 'message': str(e)})
    finally:
        sys.stdout = old_stdout


# ── Routes ────────────────────────────────────────────────────────────────────
@app.route('/')
def index():
    return HTML


@app.route('/convert', methods=['POST'])
def convert():
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
    job = jobs.get(job_id)
    if not job or not job.get('output_path'):
        return 'File not found', 404
    path = job['output_path']
    if not os.path.exists(path):
        return 'File not found', 404
    return send_file(path, as_attachment=True, download_name=Path(path).name)


# ── HTML template ─────────────────────────────────────────────────────────────
HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>DeckWash 🌊</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
    background: #FFFAF0;
    min-height: 100vh;
    display: flex;
    flex-direction: column;
  }

  header {
    background: #231F20;
    color: #FFFAF0;
    padding: 28px 40px;
    display: flex;
    align-items: baseline;
    gap: 16px;
  }

  header h1 {
    font-size: 2rem;
    font-weight: 800;
    letter-spacing: -0.5px;
  }

  header span.wave { font-size: 1.8rem; }

  header p {
    color: #46DE66;
    font-size: 0.95rem;
    font-weight: 500;
    margin-left: auto;
    opacity: 0.9;
  }

  main {
    flex: 1;
    display: flex;
    flex-direction: column;
    align-items: center;
    padding: 60px 24px;
    gap: 40px;
  }

  /* ── Drop zone ── */
  #drop-zone {
    width: 100%;
    max-width: 640px;
    border: 3px dashed #ccc;
    border-radius: 20px;
    background: #fff;
    padding: 60px 40px;
    text-align: center;
    cursor: pointer;
    transition: border-color 0.2s, background 0.2s, transform 0.15s;
    position: relative;
  }

  #drop-zone.drag-over {
    border-color: #46DE66;
    background: #f0fff5;
    transform: scale(1.01);
  }

  #drop-zone .icon {
    font-size: 3.5rem;
    margin-bottom: 16px;
    display: block;
  }

  #drop-zone h2 {
    font-size: 1.3rem;
    color: #231F20;
    margin-bottom: 8px;
    font-weight: 700;
  }

  #drop-zone p {
    color: #888;
    font-size: 0.9rem;
    line-height: 1.5;
  }

  #drop-zone .browse-btn {
    display: inline-block;
    margin-top: 20px;
    padding: 10px 28px;
    background: #231F20;
    color: #FFFAF0;
    border-radius: 100px;
    font-size: 0.9rem;
    font-weight: 600;
    cursor: pointer;
    transition: background 0.2s;
  }

  #drop-zone .browse-btn:hover { background: #3a3536; }

  #file-input { display: none; }

  /* ── Job list ── */
  #job-list {
    width: 100%;
    max-width: 640px;
    display: flex;
    flex-direction: column;
    gap: 16px;
  }

  .job-card {
    background: #fff;
    border-radius: 16px;
    padding: 24px 28px;
    box-shadow: 0 2px 12px rgba(0,0,0,0.07);
    animation: slideIn 0.3s ease;
  }

  @keyframes slideIn {
    from { opacity: 0; transform: translateY(12px); }
    to   { opacity: 1; transform: translateY(0); }
  }

  .job-card .job-name {
    font-size: 1rem;
    font-weight: 700;
    color: #231F20;
    margin-bottom: 6px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }

  .job-card .job-status {
    font-size: 0.82rem;
    color: #888;
    margin-bottom: 14px;
    min-height: 18px;
    transition: color 0.2s;
  }

  .job-card .job-status.done-text { color: #2aaa50; font-weight: 600; }
  .job-card .job-status.error-text { color: #e04040; font-weight: 600; }

  /* Progress bar */
  .progress-track {
    height: 8px;
    background: #f0f0f0;
    border-radius: 100px;
    overflow: hidden;
    margin-bottom: 16px;
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
    width: 40%;
  }

  @keyframes indeterminate {
    0%   { transform: translateX(-100%); }
    100% { transform: translateX(350%); }
  }

  .progress-bar.complete { width: 100% !important; animation: none; }
  .progress-bar.error    { background: #e04040; width: 100% !important; animation: none; }

  /* Saved badge */
  .saved-badge {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    padding: 10px 24px;
    background: #46DE66;
    color: #231F20;
    border-radius: 100px;
    font-size: 0.9rem;
    font-weight: 700;
  }

  /* Log drawer */
  .log-toggle {
    background: none;
    border: none;
    color: #aaa;
    font-size: 0.78rem;
    cursor: pointer;
    margin-top: 10px;
    padding: 0;
    text-decoration: underline;
    display: block;
  }

  .log-toggle:hover { color: #666; }

  .log-box {
    margin-top: 10px;
    background: #f7f7f7;
    border-radius: 10px;
    padding: 12px 14px;
    font-size: 0.75rem;
    font-family: 'SF Mono', 'Menlo', monospace;
    color: #555;
    max-height: 160px;
    overflow-y: auto;
    display: none;
    line-height: 1.6;
  }

  .log-box.visible { display: block; }

  footer {
    text-align: center;
    padding: 24px;
    color: #bbb;
    font-size: 0.8rem;
  }
</style>
</head>
<body>

<header>
  <span class="wave">🌊</span>
  <h1>DeckWash</h1>
  <p>Drop your decks. Get them rebranded.</p>
</header>

<main>
  <div id="drop-zone">
    <span class="icon">📂</span>
    <h2>Drop your PPTX files here</h2>
    <p>Drag one or more files onto this area, or click the button below.<br>
       Your rebranded files will be ready to download in moments.</p>
    <label class="browse-btn" for="file-input">Browse files…</label>
    <input type="file" id="file-input" accept=".pptx" multiple>
  </div>

  <div id="job-list"></div>
</main>

<footer>DeckWash — Rebrand tool by Sense</footer>

<script>
const dropZone   = document.getElementById('drop-zone');
const fileInput  = document.getElementById('file-input');
const jobList    = document.getElementById('job-list');

// ── Drag & drop handlers ────────────────────────────────────────────────────
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

// ── Start a conversion job ──────────────────────────────────────────────────
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

// ── Stream SSE status updates ───────────────────────────────────────────────
function streamStatus(card, jobId) {
  const es = new EventSource(`/status/${jobId}`);
  let slideCount = 0;
  let slideDone  = 0;

  es.onmessage = ev => {
    const msg = JSON.parse(ev.data);

    if (msg.type === 'ping') return;

    if (msg.type === 'log') {
      appendLog(card, msg.message);

      // Parse slide count
      const foundMatch = msg.message.match(/Found (\\d+) slides/);
      if (foundMatch) slideCount = parseInt(foundMatch[1]);

      // Parse current slide progress
      const slideMatch = msg.message.match(/Slide (\\d+):/);
      if (slideMatch && slideCount > 0) {
        slideDone = parseInt(slideMatch[1]);
        const pct = Math.round((slideDone / slideCount) * 85); // reserve 15% for fonts
        setProgress(card, pct);
        setStatus(card, `Processing slide ${slideDone} of ${slideCount}…`);
      }

      if (msg.message.toLowerCase().includes('embedding font')) {
        setProgress(card, 90);
        setStatus(card, 'Embedding fonts…');
      }
    }

    if (msg.type === 'done') {
      setProgress(card, 100, 'complete');
      setStatus(card, msg.cloud ? '✓ Done! Click to download.' : '✓ Saved to new case studies folder.', 'done-text');
      showDownload(card, jobId, msg.filename, msg.cloud);
      es.close();
    }

    if (msg.type === 'error') {
      setProgress(card, 100, 'error');
      showError(card, msg.message);
      es.close();
    }
  };

  es.onerror = () => {
    showError(card, 'Connection lost — please try again.');
    es.close();
  };
}

// ── Card helpers ────────────────────────────────────────────────────────────
function createCard(filename) {
  const card = document.createElement('div');
  card.className = 'job-card';
  card.innerHTML = `
    <div class="job-name">${escHtml(filename)}</div>
    <div class="job-status">Preparing…</div>
    <div class="progress-track">
      <div class="progress-bar indeterminate"></div>
    </div>
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
    btn.style.cssText = 'text-decoration:none;cursor:pointer;';
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
  setStatus(card, '✗ Error: ' + msg, 'error-text');
}

function appendLog(card, text) {
  const box = card.querySelector('.log-box');
  box.textContent += text + '\\n';
  box.scrollTop = box.scrollHeight;
}

function toggleLog(btn) {
  const box = btn.nextElementSibling;
  const visible = box.classList.toggle('visible');
  btn.textContent = visible ? 'Hide details' : 'Show details';
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
