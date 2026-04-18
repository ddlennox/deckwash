#!/usr/bin/env python3
"""Build a single self-contained HTML preview gallery that shows every
case-study deck slide-by-slide with the original on the left and the
rebranded version on the right.

Assumes JPEG thumbnails already exist under:
  preview_gallery/old_jpgs/<deck>/slide-*.jpg
  preview_gallery/new_jpgs/<deck>_rebranded/slide-*.jpg

Writes: preview_gallery/index.html
"""
from __future__ import annotations
import html
import os
import re
from pathlib import Path

ROOT    = Path(__file__).resolve().parent / 'preview_gallery'
OLD_DIR = ROOT / 'old_jpgs'
NEW_DIR = ROOT / 'new_jpgs'
IMG_EXT = {'.jpg', '.jpeg', '.png'}
OUT     = ROOT / 'index.html'


# ── Manual overrides: new folder name → old folder name ───────────────
# All decks that don't simply drop "_rebranded" to match.
SPECIAL_MATCH = {
    'Madri_rebranded':                              'Old Case Study Format - Madri',
    'Madri (1)_rebranded':                          'Old Case Study Format - Madri (1)',
    'Pokemon_rebranded':                            'Old Case Study Format - Pokemon',
    'Too Faced Kiss Cafe_rebranded':                'Too Faced Kiss Cafe ',
    'Too Faced Ribbon Lash Wrapped Launch Event_rebranded':
                                                    'Too Faced Ribbon Lash Wrapped Launch Event ',
    'Westfield x Lilo & Stitch_rebranded':          'Westfield x Lilo & Stitch ',
}


def natural_key(s: str):
    """Sort slide-1.png, slide-2.png, …, slide-10.png in numeric order."""
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', s)]


def find_old_dir(new_name: str) -> Path | None:
    if new_name in SPECIAL_MATCH:
        candidate = OLD_DIR / SPECIAL_MATCH[new_name]
        return candidate if candidate.is_dir() else None
    # Strip the "_rebranded" suffix
    stem = new_name[: -len('_rebranded')] if new_name.endswith('_rebranded') else new_name
    candidate = OLD_DIR / stem
    return candidate if candidate.is_dir() else None


def collect_slide_pairs(old_dir: Path | None, new_dir: Path):
    old_slides = sorted(
        (old_dir.iterdir() if old_dir else []),
        key=lambda p: natural_key(p.name),
    ) if old_dir else []
    new_slides = sorted(new_dir.iterdir(), key=lambda p: natural_key(p.name))
    old_slides = [p for p in old_slides if p.suffix.lower() in IMG_EXT]
    new_slides = [p for p in new_slides if p.suffix.lower() in IMG_EXT]
    n = max(len(old_slides), len(new_slides))
    pairs = []
    for i in range(n):
        o = old_slides[i] if i < len(old_slides) else None
        nw = new_slides[i] if i < len(new_slides) else None
        pairs.append((o, nw))
    return pairs


def rel(p: Path | None) -> str:
    if p is None:
        return ''
    return str(p.relative_to(ROOT)).replace('\\', '/')


def build_html() -> str:
    new_decks = sorted(
        [d for d in NEW_DIR.iterdir() if d.is_dir()],
        key=lambda p: natural_key(p.name),
    )

    # Build nav + sections
    nav_items = []
    sections  = []
    for new_dir in new_decks:
        new_name    = new_dir.name
        old_dir     = find_old_dir(new_name)
        display     = new_name[: -len('_rebranded')] if new_name.endswith('_rebranded') else new_name
        slug        = re.sub(r'[^a-z0-9]+', '-', display.lower()).strip('-')
        pairs       = collect_slide_pairs(old_dir, new_dir)

        nav_items.append(
            f'<li><a href="#deck-{slug}">{html.escape(display)}</a></li>'
        )

        # Per-slide rows
        rows = []
        for idx, (o, nw) in enumerate(pairs, 1):
            old_cell = (
                f'<img loading="lazy" src="{html.escape(rel(o))}" alt="old slide {idx}">'
                if o else
                '<div class="missing">original not available</div>'
            )
            new_cell = (
                f'<img loading="lazy" src="{html.escape(rel(nw))}" alt="rebranded slide {idx}">'
                if nw else
                '<div class="missing">rebrand missing</div>'
            )
            rows.append(
                f'<div class="row">'
                f'<div class="slide-num">Slide {idx}</div>'
                f'<div class="pair">'
                f'  <figure><figcaption>Before</figcaption>{old_cell}</figure>'
                f'  <figure><figcaption>After</figcaption>{new_cell}</figure>'
                f'</div></div>'
            )

        if not pairs:
            rows.append('<p class="missing">no slides rendered</p>')

        sections.append(
            f'<section id="deck-{slug}" class="deck">'
            f'  <h2>{html.escape(display)}</h2>'
            + '\n'.join(rows) +
            f'</section>'
        )

    nav_html     = '<ul>' + '\n'.join(nav_items) + '</ul>'
    sections_html = '\n'.join(sections)

    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Case-study rebrand — before / after gallery</title>
<style>
  :root {{
    --cream:#FFFAF0;
    --ink:#111;
    --muted:#777;
    --rule:#d9d2c4;
    --sand:#f3ecd9;
  }}
  * {{ box-sizing:border-box; }}
  html, body {{ margin:0; padding:0; background:var(--cream); color:var(--ink);
                font:14px/1.45 -apple-system,"SF Pro Text","Helvetica Neue",Arial,sans-serif; }}
  header.top {{
    position:sticky; top:0; z-index:10;
    background:var(--cream); border-bottom:1px solid var(--rule);
    padding:14px 24px; display:flex; gap:18px; align-items:baseline;
  }}
  header.top h1 {{ margin:0; font-size:18px; font-weight:700; letter-spacing:-0.01em; }}
  header.top .meta {{ color:var(--muted); font-size:12px; margin-left:auto; }}
  header.top .back {{ color:var(--muted); text-decoration:none; font-size:13px;
                      padding:4px 10px; border:1px solid var(--rule); border-radius:4px;
                      transition:background 0.15s, color 0.15s; }}
  header.top .back:hover {{ background:var(--sand); color:var(--ink); }}
  .layout {{ display:grid; grid-template-columns:260px 1fr; gap:24px; padding:24px; }}
  aside.nav {{
    position:sticky; top:64px; align-self:start;
    max-height:calc(100vh - 80px); overflow:auto;
    border-right:1px solid var(--rule); padding-right:16px;
  }}
  aside.nav h3 {{ font-size:11px; text-transform:uppercase; letter-spacing:0.08em;
                  margin:0 0 8px; color:var(--muted); }}
  aside.nav ul {{ list-style:none; padding:0; margin:0; }}
  aside.nav li {{ margin:2px 0; }}
  aside.nav a {{
    display:block; padding:6px 8px; border-radius:4px;
    text-decoration:none; color:var(--ink); font-size:13px;
  }}
  aside.nav a:hover {{ background:var(--sand); }}
  main {{ min-width:0; }}
  section.deck {{ scroll-margin-top:80px; margin-bottom:56px;
                 padding-bottom:24px; border-bottom:1px solid var(--rule); }}
  section.deck:last-child {{ border-bottom:none; }}
  section.deck h2 {{ font-size:20px; font-weight:700; margin:0 0 16px;
                     letter-spacing:-0.01em; }}
  .row {{ margin-bottom:28px; }}
  .slide-num {{ font-size:11px; text-transform:uppercase; letter-spacing:0.1em;
                color:var(--muted); margin-bottom:6px; }}
  .pair {{ display:grid; grid-template-columns:1fr 1fr; gap:16px; }}
  figure {{ margin:0; }}
  figcaption {{ font-size:11px; text-transform:uppercase; letter-spacing:0.08em;
                color:var(--muted); margin-bottom:4px; }}
  figure img {{
    width:100%; height:auto; display:block;
    border:1px solid var(--rule); border-radius:4px;
    box-shadow:0 1px 3px rgba(0,0,0,0.04); background:#fff;
  }}
  .missing {{ background:#f6efe0; color:var(--muted);
             border:1px dashed var(--rule); border-radius:4px;
             padding:28px 12px; text-align:center; font-style:italic; font-size:12px; }}
  @media (max-width: 900px) {{
    .layout {{ grid-template-columns:1fr; }}
    aside.nav {{ position:static; max-height:none; border-right:none;
                 border-bottom:1px solid var(--rule); padding:0 0 12px; }}
  }}
</style>
</head>
<body>

<header class="top">
  <a href="/" class="back">&larr; DeckWash</a>
  <h1>Before / After gallery</h1>
  <span class="meta">{len(new_decks)} decks · originals on the left, rebranded on the right</span>
</header>

<div class="layout">
  <aside class="nav">
    <h3>Decks</h3>
    {nav_html}
  </aside>
  <main>
    {sections_html}
  </main>
</div>

</body>
</html>
"""


def main():
    html_str = build_html()
    OUT.write_text(html_str, encoding='utf-8')
    size_kb = OUT.stat().st_size / 1024
    print(f"Wrote {OUT} ({size_kb:.1f} KB)")


if __name__ == '__main__':
    main()
