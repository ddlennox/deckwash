#!/usr/bin/env python3
"""
Case Study PPTX Rebrand Converter
Converts old case study PPTX files to new design aesthetic:
- Dark #231F20 backgrounds for cover and testimonial slides
- Cream #FFFAF0 backgrounds for content and photo-grid slides
- Obviously Narrow Bold for titles and section headers
- Galvji for body text and subtitles
- Removes old decorative elements (ovals, agency logos, thin separator lines)
- Section headers uppercased
"""

import sys
import os
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from lxml import etree

# ── Design constants ──────────────────────────────────────────────────────────
DARK_BG    = "231F20"   # cover + testimonial slides
CREAM_BG   = "FFFAF0"   # content + photo-grid slides
CREAM_TEXT = "FFFFFF"   # text colour on dark slides (white)
DARK_TEXT  = "231F20"   # text colour on cream slides

# Fonts
FONT_HEADER = "Obviously Narrow"  # +mj-lt in the new theme
FONT_BODY   = "Galvji"            # +mn-lt in the new theme
TESTIMONIAL_NAME_COLOR = "46DE66"  # green accent from new template (bg2/lt2)

# ── Quote mark image (from new testimonial template) ────────────────────────
QUOTE_MARK_TEMPLATE = os.path.join(os.path.dirname(__file__), 'New testimonial page - example.pptx')
QUOTE_MARK_ZIP_PATH = 'ppt/media/image7.png'
QUOTE_MARK_MEDIA    = 'ppt/media/sense_quote_mark.png'
QUOTE_MARK_RID      = 'rId_sense_qm'
QUOTE_MARK_RTYPE    = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

# EMU positions / sizes from template layout (slideLayout33)
QMARK_TOP = dict(x=562019,   y=570678,  cx=1325248, cy=1319748)
QMARK_BOT = dict(x=10304732, y=4773360, cx=1325248, cy=1319748)

# ── XML namespaces ───────────────────────────────────────────────────────────
NSP = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
A  = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PP = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

# Keywords that identify section-header bold runs (any variant)
HEADER_KEYWORDS = {
    'challenge', 'execution', 'results', 'solution', 'approach',
    'objective', 'objectives', 'outcome', 'outcomes', 'background',
    'overview', 'summary', 'insight', 'insights', 'strategy',
    'deliverables', 'impact', 'key results',
}

# Canonical names — all variants map to one of these three
HEADER_NORMALIZE = {
    'challenge':    'CHALLENGE',
    'objective':    'CHALLENGE',
    'objectives':   'CHALLENGE',
    'background':   'CHALLENGE',
    'overview':     'CHALLENGE',
    'insight':      'CHALLENGE',
    'insights':     'CHALLENGE',
    'execution':    'EXECUTION',
    'solution':     'EXECUTION',
    'approach':     'EXECUTION',
    'strategy':     'EXECUTION',
    'deliverables': 'EXECUTION',
    'results':      'RESULTS',
    'outcome':      'RESULTS',
    'outcomes':     'RESULTS',
    'impact':       'RESULTS',
    'summary':      'RESULTS',
    'key results':  'RESULTS',
}


# ═════════════════════════════════════════════════════════════════════════════
# Utility helpers
# ═════════════════════════════════════════════════════════════════════════════

def qn(ns_prefix, local):
    """Build a Clark-notation qualified name like {http://...}tag"""
    return f"{{{NSP[ns_prefix]}}}{local}"


def solid_fill_el(hex_color):
    """Return an <a:solidFill><a:srgbClr val="HEX"/></a:solidFill> element."""
    fill = etree.Element(qn('a', 'solidFill'))
    clr  = etree.SubElement(fill, qn('a', 'srgbClr'))
    clr.set('val', hex_color)
    return fill


def bg_element(hex_color):
    """Return a <p:bg> element with the given solid colour."""
    bg    = etree.Element(qn('p', 'bg'))
    bgPr  = etree.SubElement(bg, qn('p', 'bgPr'))
    bgPr.append(solid_fill_el(hex_color))
    etree.SubElement(bgPr, qn('a', 'effectLst'))
    return bg


def latin_element(typeface):
    """Return an <a:latin typeface="..."/> element."""
    el = etree.Element(qn('a', 'latin'))
    el.set('typeface', typeface)
    return el


def get_all_text(node):
    """Concatenate all <a:t> text inside node."""
    parts = [t.text or '' for t in node.iter(qn('a', 't'))]
    return ''.join(parts).strip()


def run_is_bold(rpr_el):
    return rpr_el is not None and rpr_el.get('b') in ('1', 'true')


def is_header_text(text):
    """True if cleaned text matches a known section-header keyword."""
    return text.strip().lower().rstrip('.:') in HEADER_KEYWORDS


# ═════════════════════════════════════════════════════════════════════════════
# Slide-type detection
# ═════════════════════════════════════════════════════════════════════════════

def classify_slide(root, slide_index, total_slides):
    """
    Returns one of: 'cover', 'content', 'grid', 'testimonial'

    Testimonial detection runs BEFORE keyword scanning so that quotes which
    happen to contain words like "execution" aren't mislabelled as content.
    Multiple testimonial slides per file are supported.
    """
    # Cover: always the first slide
    if slide_index == 0:
        return 'cover'

    all_text = get_all_text(root)
    pics = root.findall('.//' + qn('p', 'pic'))
    text_sps = [
        sp for sp in root.findall('.//' + qn('p', 'sp'))
        if sp.find('.//' + qn('a', 't')) is not None
    ]

    # Testimonial: no embedded pictures + substantial text.
    # Checked FIRST so keyword-in-quote doesn't push it into 'content'.
    if len(pics) == 0 and len(all_text.strip()) > 150 and text_sps:
        return 'testimonial'

    # Content: has Challenge / Execution / Results section headers
    has_header_kw = any(k in all_text.lower()
                        for k in ['challenge', 'execution', 'results', 'solution'])
    if has_header_kw:
        return 'content'

    # Photo grid: everything else (many pics, or minimal text)
    return 'grid'


# ═════════════════════════════════════════════════════════════════════════════
# Background helpers
# ═════════════════════════════════════════════════════════════════════════════

def set_slide_background(root, hex_color):
    """
    Add/replace explicit <p:bg> on the slide's <p:cSld> element.
    """
    cSld = root.find(qn('p', 'cSld'))
    if cSld is None:
        return

    # Remove existing background element
    existing = cSld.find(qn('p', 'bg'))
    if existing is not None:
        cSld.remove(existing)

    # Insert right after the opening of cSld (before spTree)
    spTree = cSld.find(qn('p', 'spTree'))
    idx = list(cSld).index(spTree) if spTree is not None else 0
    cSld.insert(idx, bg_element(hex_color))


# ═════════════════════════════════════════════════════════════════════════════
# Decorative-element removal
# ═════════════════════════════════════════════════════════════════════════════

SMALL_EMU = 600_000   # ~0.66 inches — pictures smaller than this are logos

def should_remove_shape(element, tag_local):
    """True if this shape is an old decorative element that should be deleted."""
    ns_a = NSP['a']
    ns_p = NSP['p']

    name_el = element.find(f'{{{ns_p}}}nvSpPr/{{{ns_p}}}cNvPr')
    if name_el is None:
        name_el = element.find(f'{{{ns_p}}}nvPicPr/{{{ns_p}}}cNvPr')
    name = (name_el.get('name') or '') if name_el is not None else ''

    # ── Oval auto-shapes ──────────────────────────────────────────────────
    if tag_local == 'sp':
        prstGeom = element.find(f'.//{{{ns_a}}}prstGeom')
        if prstGeom is not None and prstGeom.get('prst') == 'ellipse':
            return True
        # Thin separator rectangles (w < 0.12" OR h < 0.12")
        if prstGeom is not None and prstGeom.get('prst') == 'rect':
            xfrm = element.find(f'.//{{{ns_a}}}xfrm')
            if xfrm is not None:
                ext = xfrm.find(f'{{{ns_a}}}ext')
                if ext is not None:
                    cx = int(ext.get('cx', 999999))
                    cy = int(ext.get('cy', 999999))
                    if cx < 110_000 or cy < 110_000:
                        return True

    # ── Small/logo pictures ───────────────────────────────────────────────
    if tag_local == 'pic':
        # Skip pictures that have a placeholder (content images)
        ph = element.find(f'.//{{{ns_p}}}ph')
        if ph is not None:
            return False  # placeholder image — keep it

        # Remove small non-placeholder images (agency logos, dots, footer strips)
        xfrm = element.find(f'.//{{{ns_a}}}xfrm')
        if xfrm is not None:
            ext = xfrm.find(f'{{{ns_a}}}ext')
            if ext is not None:
                cx = int(ext.get('cx', 999999))
                cy = int(ext.get('cy', 999999))
                # Both dimensions small (square logos) OR very thin strip (footer bars)
                if (cx < SMALL_EMU and cy < SMALL_EMU) or cy < 400_000:
                    return True

        # Remove pictures named "Graphic *" (old agency overlay strips)
        if name.startswith('Graphic') or name.lower().startswith('graphic'):
            return True

    return False


def remove_decorative_elements(root):
    """Remove old brand/decorative shapes from the slide shape tree."""
    spTree = root.find(f'.//{{{PP}}}spTree')
    if spTree is None:
        return

    to_remove = []
    for child in spTree:
        local = child.tag.split('}')[-1] if '}' in child.tag else child.tag
        if local in ('sp', 'pic') and should_remove_shape(child, local):
            to_remove.append(child)

    for el in to_remove:
        spTree.remove(el)
        name_el = el.find(f'.//{{{PP}}}cNvPr')
        name = name_el.get('name', '?') if name_el is not None else '?'
        print(f"    Removed decorative: {name}")


# ═════════════════════════════════════════════════════════════════════════════
# Text / font transformation
# ═════════════════════════════════════════════════════════════════════════════

def clear_latin(rpr):
    """Remove any explicit <a:latin> from a run-properties element."""
    for lat in rpr.findall(qn('a', 'latin')):
        rpr.remove(lat)


BODY_SZ         = '1200'   # 12 pt — grid slide body text
CONTENT_BODY_SZ = '1100'   # 11 pt — bullet points on content slides
SECTION_HDR_SZ  = '1600'   # 16 pt — Challenge / Execution / Results headers
COVER_TITLE_SZ  = '3200'   # 32 pt — cover slide main title
COVER_BODY_SZ   = '1600'   # 16 pt — cover slide body/subtitle line

def strip_sz(rpr):
    """Remove explicit sz so the run inherits size from the slide layout."""
    rpr.attrib.pop('sz', None)


def set_body_sz(rpr):
    """Set grid body-text size (12 pt)."""
    rpr.set('sz', BODY_SZ)


def set_content_body_sz(rpr):
    """Set content-slide bullet/body size (11 pt)."""
    rpr.set('sz', CONTENT_BODY_SZ)


def set_section_hdr_sz(rpr):
    """Set section header size (16 pt) for Challenge/Execution/Results."""
    rpr.set('sz', SECTION_HDR_SZ)


def set_color(rpr, hex_color):
    """Set explicit solid fill colour on a run-properties element."""
    # Remove existing color settings
    for sf in rpr.findall(qn('a', 'solidFill')):
        rpr.remove(sf)
    for sc in rpr.findall(qn('a', 'schemeClr')):
        rpr.remove(sc)

    fill = etree.SubElement(rpr, qn('a', 'solidFill'))
    clr  = etree.SubElement(fill, qn('a', 'srgbClr'))
    clr.set('val', hex_color)


def process_cover_slide(root):
    """Cover: dark bg, title in Obviously Narrow Bold white, subtitle in Galvji Bold white."""
    for sp in root.findall(f'.//{{{PP}}}sp'):
        ph = sp.find(f'.//{{{PP}}}ph')
        ph_type = ph.get('type') if ph is not None else None
        txBody = sp.find(qn('p', 'txBody'))
        if txBody is None:
            continue

        for rpr in txBody.iter(qn('a', 'rPr')):
            clear_latin(rpr)
            set_color(rpr, CREAM_TEXT)

            if ph_type == 'title':
                rpr.set('sz', COVER_TITLE_SZ)   # 32 pt
                rpr.set('b', '1')
                rpr.append(latin_element(FONT_HEADER))
                # Uppercase the run text
                t_el = rpr.getnext()
                if t_el is not None and t_el.tag == qn('a', 't') and t_el.text:
                    t_el.text = t_el.text.upper()
            else:
                rpr.set('sz', COVER_BODY_SZ)    # 16 pt
                rpr.set('b', '1')
                rpr.append(latin_element(FONT_BODY))

        # Also handle endParaRPr
        for endrpr in txBody.iter(qn('a', 'endParaRPr')):
            clear_latin(endrpr)
            set_color(endrpr, CREAM_TEXT)
            if ph_type == 'title':
                endrpr.set('sz', COVER_TITLE_SZ)
                endrpr.set('b', '1')
                endrpr.append(latin_element(FONT_HEADER))
            else:
                endrpr.set('sz', COVER_BODY_SZ)
                endrpr.set('b', '1')
                endrpr.append(latin_element(FONT_BODY))


def _run_text(rpr_el):
    """Get the text of the <a:t> that follows this <a:rPr>."""
    t = rpr_el.getnext()
    if t is not None and t.tag == qn('a', 't'):
        return (t.text or '').strip()
    return ''


def para_has_content(p_el):
    """True if the paragraph contains at least one non-empty <a:t>."""
    return any(
        (t.text or '').strip()
        for t in p_el.iter(qn('a', 't'))
    )


def para_is_header(p_el):
    """True if the paragraph is a bold section header (Challenge, Execution, etc.)."""
    for r in p_el.findall(qn('a', 'r')):
        rpr = r.find(qn('a', 'rPr'))
        t   = r.find(qn('a', 't'))
        text = (t.text or '').strip() if t is not None else ''
        if rpr is not None and run_is_bold(rpr) and is_header_text(text):
            return True
    return False


def fix_content_spacing(txBody):
    """
    Two spacing fixes for the content body text box:

    1. SEPARATE-PARAGRAPH style (e.g. Marriott):
       Remove empty <a:p> elements that immediately follow a header paragraph —
       they create an unwanted blank line between CHALLENGE and its first bullet.

    2. LINE-BREAK style (e.g. Lime):
       When an <a:br> immediately follows a header run, collapse its line height
       by setting sz="100" on its rPr so the break takes minimal vertical space.
    """
    paras = txBody.findall(qn('a', 'p'))

    # ── Fix 1: remove empty paragraphs after headers ──────────────────────
    to_remove = []
    for i, para in enumerate(paras):
        if not para_has_content(para) and i > 0:
            # Check the previous non-empty paragraph
            prev_idx = i - 1
            while prev_idx >= 0 and not para_has_content(paras[prev_idx]):
                prev_idx -= 1
            if prev_idx >= 0 and para_is_header(paras[prev_idx]):
                to_remove.append(para)

    for p in to_remove:
        txBody.remove(p)

    # ── Fix 2: collapse <a:br> line-height after header runs ─────────────
    for para in txBody.findall(qn('a', 'p')):
        children = list(para)
        for idx, child in enumerate(children):
            if child.tag != qn('a', 'r'):
                continue
            # Is this run a header?
            rpr  = child.find(qn('a', 'rPr'))
            t_el = child.find(qn('a', 't'))
            text = (t_el.text or '').strip() if t_el is not None else ''
            if not (rpr is not None and run_is_bold(rpr) and is_header_text(text)):
                continue
            # Look at the next sibling — if it's a <a:br>, shrink it
            if idx + 1 < len(children) and children[idx + 1].tag == qn('a', 'br'):
                br_el  = children[idx + 1]
                br_rpr = br_el.find(qn('a', 'rPr'))
                if br_rpr is None:
                    br_rpr = etree.SubElement(br_el, qn('a', 'rPr'))
                br_rpr.set('sz', '100')  # near-zero line height
                clear_latin(br_rpr)
                br_rpr.append(latin_element(FONT_BODY))


def add_separator_line(root):
    """
    Add a thin dark horizontal line between the intro paragraph and CHALLENGE
    on content slides. Matches the 'Line Placeholder' in the new Pokemon layout.
    We position it at x of the content column, y ≈ 1870000 (~2.04"), h = 12700.
    """
    spTree = root.find(f'.//{{{PP}}}spTree')
    if spTree is None:
        return

    # Find the content placeholder to get its x position and width
    content_x  = 6235700   # default from new Pokemon layout 47
    content_cx = 4788000

    for sp in spTree.findall(qn('p', 'sp')):
        ph = sp.find(f'.//{{{PP}}}ph')
        if ph is not None and ph.get('idx') == '1':   # idx=1 is the generic content ph
            xfrm = sp.find(f'.//{{{A}}}xfrm')
            if xfrm is not None:
                off = xfrm.find(qn('a', 'off'))
                ext = xfrm.find(qn('a', 'ext'))
                if off is not None:
                    content_x = int(off.get('x', content_x))
                if ext is not None:
                    content_cx = int(ext.get('cx', content_cx))
            break

    # Build the separator shape XML
    sep_id = 9001   # unique enough for a new shape
    sep_xml = f'''<p:sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:nvSpPr>
    <p:cNvPr id="{sep_id}" name="SeparatorLine"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="{content_x}" y="1870000"/>
      <a:ext cx="{content_cx}" cy="12700"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="{DARK_BG}"/></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p><a:endParaRPr lang="en-US"/></a:p>
  </p:txBody>
</p:sp>'''
    sep_el = etree.fromstring(sep_xml)
    spTree.append(sep_el)


def process_content_slide(root):
    """Content: cream bg, title UPPERCASED in Obviously Narrow Bold dark, body in Galvji dark.
       Section headers (bold+keyword) → Obviously Narrow Bold uppercase.
       Separator line injected above CHALLENGE. Empty post-header paragraphs removed."""
    for sp in root.findall(f'.//{{{PP}}}sp'):
        ph = sp.find(f'.//{{{PP}}}ph')
        ph_type = ph.get('type') if ph is not None else None
        txBody = sp.find(qn('p', 'txBody'))
        if txBody is None:
            continue

        for rpr in txBody.iter(qn('a', 'rPr')):
            clear_latin(rpr)
            set_color(rpr, DARK_TEXT)

            run_text = _run_text(rpr)
            is_bold  = run_is_bold(rpr)
            is_hdr   = is_bold and is_header_text(run_text)

            if ph_type == 'title':
                strip_sz(rpr)   # title placeholder inherits its own correct size
                rpr.set('b', '1')
                rpr.append(latin_element(FONT_HEADER))
                # Uppercase title text
                t_el = rpr.getnext()
                if t_el is not None and t_el.tag == qn('a', 't') and t_el.text:
                    t_el.text = t_el.text.upper()
            elif is_hdr:
                set_section_hdr_sz(rpr)   # 16 pt — Challenge / Execution / Results
                rpr.set('b', '1')
                rpr.append(latin_element(FONT_HEADER))
                # Normalise to canonical CHALLENGE / EXECUTION / RESULTS
                t_el = rpr.getnext()
                if t_el is not None and t_el.tag == qn('a', 't') and t_el.text:
                    key = t_el.text.strip().lower().rstrip('.:')
                    t_el.text = HEADER_NORMALIZE.get(key, t_el.text.upper())
            else:
                set_content_body_sz(rpr)   # 11 pt — bullet points / body text
                rpr.append(latin_element(FONT_BODY))

        for endrpr in txBody.iter(qn('a', 'endParaRPr')):
            clear_latin(endrpr)
            set_content_body_sz(endrpr)
            set_color(endrpr, DARK_TEXT)
            endrpr.append(latin_element(FONT_BODY))

        # Fix spacing: remove empty paras after headers + collapse br line heights
        if ph_type != 'title':
            fix_content_spacing(txBody)

    # ── Table cells on content slides (e.g. stats/comparison tables) ──
    # Update font + colour but preserve intentional sizes, matching the
    # grid-slide table handling.
    for tc in root.iter(qn('a', 'tc')):
        for rpr in tc.iter(qn('a', 'rPr')):
            clear_latin(rpr)
            set_color(rpr, DARK_TEXT)
            rpr.append(latin_element(FONT_BODY))
        for endrpr in tc.iter(qn('a', 'endParaRPr')):
            clear_latin(endrpr)
            set_color(endrpr, DARK_TEXT)
            endrpr.append(latin_element(FONT_BODY))

    # Add the separator line above CHALLENGE
    add_separator_line(root)


def process_grid_slide(root):
    """Photo grid: cream bg, text updated to Galvji 12 pt dark.

    Regular shapes (<p:sp>) get full font + size normalisation (12 pt).
    Table cells (<a:tc>) get font + colour updated but keep their original
    sizes, since stats tables use intentionally large numbers (36 pt etc.).
    """
    # Regular shapes — normalise font AND size
    for sp in root.findall(f'.//{{{PP}}}sp'):
        for rpr in sp.iter(qn('a', 'rPr')):
            clear_latin(rpr)
            set_body_sz(rpr)
            set_color(rpr, DARK_TEXT)
            rpr.append(latin_element(FONT_BODY))
        for endrpr in sp.iter(qn('a', 'endParaRPr')):
            clear_latin(endrpr)
            set_body_sz(endrpr)
            set_color(endrpr, DARK_TEXT)
            endrpr.append(latin_element(FONT_BODY))

    # Table cells — update font + colour, preserve intentional sizes
    for tc in root.iter(qn('a', 'tc')):
        for rpr in tc.iter(qn('a', 'rPr')):
            clear_latin(rpr)
            set_color(rpr, DARK_TEXT)
            rpr.append(latin_element(FONT_BODY))
        for endrpr in tc.iter(qn('a', 'endParaRPr')):
            clear_latin(endrpr)
            set_color(endrpr, DARK_TEXT)
            endrpr.append(latin_element(FONT_BODY))


def _make_rpr(bold=False, italic=False, font=FONT_BODY, color=CREAM_TEXT):
    """Build a fresh <a:rPr> element with the given style."""
    rpr = etree.Element(qn('a', 'rPr'))
    rpr.set('lang', 'en-GB')
    rpr.set('dirty', '0')
    if bold:
        rpr.set('b', '1')
    else:
        rpr.set('b', '0')
    if italic:
        rpr.set('i', '1')
    else:
        rpr.set('i', '0')
    fill = etree.SubElement(rpr, qn('a', 'solidFill'))
    clr  = etree.SubElement(fill, qn('a', 'srgbClr'))
    clr.set('val', color)
    lat = etree.SubElement(rpr, qn('a', 'latin'))
    lat.set('typeface', font)
    return rpr


def _make_run(text, bold=False, italic=False, font=FONT_BODY, color=CREAM_TEXT):
    """Build a <a:r> element with the given style and text."""
    r = etree.Element(qn('a', 'r'))
    r.append(_make_rpr(bold=bold, italic=italic, font=font, color=color))
    t = etree.SubElement(r, qn('a', 't'))
    t.text = text
    return r


def _make_para(*runs):
    """Build a <a:p> containing the given run elements."""
    p = etree.Element(qn('a', 'p'))
    for r in runs:
        p.append(r)
    endrpr = etree.SubElement(p, qn('a', 'endParaRPr'))
    endrpr.set('lang', 'en-GB')
    fill = etree.SubElement(endrpr, qn('a', 'solidFill'))
    clr  = etree.SubElement(fill, qn('a', 'srgbClr'))
    clr.set('val', CREAM_TEXT)
    lat = etree.SubElement(endrpr, qn('a', 'latin'))
    lat.set('typeface', FONT_BODY)
    return p


def _pic_element(rid, x, y, cx, cy, name, shape_id):
    """Build a <p:pic> element for an image at the given EMU position."""
    pic = etree.Element(qn('p', 'pic'))

    nvPicPr = etree.SubElement(pic, qn('p', 'nvPicPr'))
    cNvPr = etree.SubElement(nvPicPr, qn('p', 'cNvPr'))
    cNvPr.set('id', str(shape_id))
    cNvPr.set('name', name)
    cNvPicPr = etree.SubElement(nvPicPr, qn('p', 'cNvPicPr'))
    locks = etree.SubElement(cNvPicPr, qn('a', 'picLocks'))
    locks.set('noChangeAspect', '1')
    etree.SubElement(nvPicPr, qn('p', 'nvPr'))

    blipFill = etree.SubElement(pic, qn('p', 'blipFill'))
    blip = etree.SubElement(blipFill, qn('a', 'blip'))
    blip.set(f'{{{R}}}embed', rid)
    stretch = etree.SubElement(blipFill, qn('a', 'stretch'))
    etree.SubElement(stretch, qn('a', 'fillRect'))

    spPr = etree.SubElement(pic, qn('p', 'spPr'))
    xfrm = etree.SubElement(spPr, qn('a', 'xfrm'))
    off = etree.SubElement(xfrm, qn('a', 'off'))
    off.set('x', str(x)); off.set('y', str(y))
    ext = etree.SubElement(xfrm, qn('a', 'ext'))
    ext.set('cx', str(cx)); ext.set('cy', str(cy))
    prstGeom = etree.SubElement(spPr, qn('a', 'prstGeom'))
    prstGeom.set('prst', 'rect')
    etree.SubElement(prstGeom, qn('a', 'avLst'))

    return pic


def inject_quote_marks(root, rid):
    """Add top and bottom quote mark images to the testimonial slide."""
    spTree = root.find(f'.//{{{PP}}}spTree')
    if spTree is None:
        return
    spTree.append(_pic_element(rid, name='Top Quote Mark',    shape_id=901, **QMARK_TOP))
    spTree.append(_pic_element(rid, name='Bottom Quote Mark', shape_id=902, **QMARK_BOT))


def inject_qmark_rel(rels_bytes):
    """Inject the quote mark image relationship into a slide .rels file."""
    CT_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships'
    root = etree.fromstring(rels_bytes)
    # Check if already present
    for rel in root.findall(f'{{{CT_RELS}}}Relationship'):
        if rel.get('Id') == QUOTE_MARK_RID:
            return rels_bytes  # already there
    rel_el = etree.SubElement(root, f'{{{CT_RELS}}}Relationship')
    rel_el.set('Id',     QUOTE_MARK_RID)
    rel_el.set('Type',   QUOTE_MARK_RTYPE)
    rel_el.set('Target', f'../media/sense_quote_mark.png')
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


def _split_name_attr(para):
    """Given a name/attribution paragraph, return (name_text, attr_text).

    Tries, in order:
      1. En/em dash separators in the full text  →  before / after
      2. Last run starts with a space (separate component)  →  rest / last run
      3. Everything as the name, no attribution.
    """
    full = get_all_text(para).strip()
    for sep in (' \u2013 ', ' \u2014 ', '\u2013', '\u2014', ' - '):
        if sep in full:
            idx = full.index(sep)
            return full[:idx].strip(), full[idx + len(sep):].strip()
    runs = para.findall(qn('a', 'r'))
    if len(runs) >= 2:
        last_t = runs[-1].find(qn('a', 't'))
        last_text = last_t.text if last_t is not None else ''
        if last_text.startswith(' ') and len(last_text.strip()) > 2:
            name_parts = [
                (r.find(qn('a', 't')).text or '')
                for r in runs[:-1]
                if r.find(qn('a', 't')) is not None
            ]
            return ''.join(name_parts).strip(), last_text.strip()
    return full, ''


def process_testimonial_slide(root, qmark_rid=None):
    """Testimonial: dark bg, cream text, rebuilt to match new template layout.

    Detection strategy (in priority order):
      1. Any paragraph with lvl >= 1  →  that paragraph is name/attribution;
         everything before it (lvl 0) is the quote.
      2. First bold paragraph  →  name; paragraphs after it  →  attribution.
      3. Fallback: last 2 non-empty paragraphs  →  name + attribution.

    Name/attribution paragraph: tries to split on em/en dash or run boundary.
    Output:
      • Quote paragraphs  →  italic Galvji, cream, sentence case
      • Name             →  Obviously Narrow Bold, green, UPPER CASE
      • Attribution      →  Galvji regular, cream, UPPER CASE
    """
    sps = [sp for sp in root.findall(f'.//{{{PP}}}sp')
           if sp.find(f'.//{{{A}}}t') is not None]

    for sp in sps:
        txBody = sp.find(qn('p', 'txBody'))
        if txBody is None:
            continue

        paras = txBody.findall(qn('a', 'p'))

        # ── Strategy 1: paragraph with lvl >= 1 is the name/attribution para ──
        name_para_idx = None
        for i, para in enumerate(paras):
            pPr = para.find(qn('a', 'pPr'))
            lvl = int(pPr.get('lvl', '0')) if pPr is not None else 0
            if lvl >= 1:
                name_para_idx = i
                break

        if name_para_idx is not None:
            quote_texts = [
                get_all_text(p).strip()
                for p in paras[:name_para_idx]
                if get_all_text(p).strip()
            ]
            name_text, attr_text = _split_name_attr(paras[name_para_idx])
            # any further lvl>=1 paras (edge case)
            for p in paras[name_para_idx + 1:]:
                extra = get_all_text(p).strip()
                if extra and not attr_text:
                    attr_text = extra
        else:
            # ── Strategy 2: bold detection ──────────────────────────────
            first_bold_idx = None
            for i, para in enumerate(paras):
                for r in para.findall(qn('a', 'r')):
                    rpr = r.find(qn('a', 'rPr'))
                    if rpr is not None and run_is_bold(rpr):
                        first_bold_idx = i
                        break
                if first_bold_idx is not None:
                    break

            # ── Strategy 3: fallback — last 2 non-empty = name + attribution
            if first_bold_idx is None:
                non_empty = [i for i, p in enumerate(paras) if get_all_text(p).strip()]
                if len(non_empty) >= 2:
                    first_bold_idx = non_empty[-2]

            # ── Special case: bold para is FIRST → it's a section label/heading,
            #    content after it = the quotes, label goes at bottom as "name"
            after_bold = [
                get_all_text(p).strip()
                for p in paras[(first_bold_idx or 0) + 1:]
                if get_all_text(p).strip()
            ] if first_bold_idx == 0 else []

            if first_bold_idx == 0 and after_bold:
                quote_texts = after_bold
                name_text = get_all_text(paras[0]).strip()
                attr_text = ''
            else:
                quote_texts = [
                    get_all_text(p).strip()
                    for i, p in enumerate(paras)
                    if (first_bold_idx is None or i < first_bold_idx) and get_all_text(p).strip()
                ]

                if first_bold_idx is not None:
                    name_text = get_all_text(paras[first_bold_idx]).strip()
                    attr_parts = [
                        get_all_text(p).strip()
                        for p in paras[first_bold_idx + 1:]
                        if get_all_text(p).strip()
                    ]
                    attr_text = ', '.join(attr_parts)
                else:
                    name_text = ''
                    attr_text = ''

        # ── Rebuild txBody paragraphs ───────────────────────────────────
        for p in list(txBody.findall(qn('a', 'p'))):
            txBody.remove(p)

        # Quote paragraphs — italic Galvji, sentence case (NOT uppercased)
        for qt in quote_texts:
            txBody.append(_make_para(_make_run(qt, italic=True)))

        # Name + attribution paragraph — name in green caps, attribution in cream caps
        if name_text:
            runs = [_make_run(name_text.upper(), bold=True, font=FONT_HEADER,
                              color=TESTIMONIAL_NAME_COLOR)]
            if attr_text:
                runs.append(_make_run(', ' + attr_text.upper()))
            txBody.append(_make_para(*runs))

    # ── Inject quote mark images ────────────────────────────────────────
    if qmark_rid:
        inject_quote_marks(root, qmark_rid)


# ═════════════════════════════════════════════════════════════════════════════
# Layout / Master font replacement
# ═════════════════════════════════════════════════════════════════════════════

# Old fonts that should become Galvji (body)
_OLD_BODY_FONTS = {'Georgia', 'Century Gothic', 'Gill Sans MT', 'Gill Sans',
                   'Calibri', 'Arial', 'Helvetica'}

def replace_layout_fonts(xml_bytes):
    """Replace legacy body fonts with Galvji in a slide layout or master XML.

    Only touches explicit <a:latin typeface="..."> elements — does NOT touch
    '+mn-lt' / '+mj-lt' theme references, which are already correct after the
    theme update.  Returns new bytes.
    """
    root = etree.fromstring(xml_bytes)
    changed = False
    for elem in root.iter(f'{{{A}}}latin'):
        tf = elem.get('typeface', '')
        if tf in _OLD_BODY_FONTS:
            elem.set('typeface', FONT_BODY)
            changed = True
    if not changed:
        return xml_bytes
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


# ═════════════════════════════════════════════════════════════════════════════
# Theme transformation
# ═════════════════════════════════════════════════════════════════════════════

def update_theme(theme_xml_bytes):
    """
    Replace major/minor fonts in the theme to Obviously Narrow / Galvji.
    Also update dk2 to #231F20.
    """
    root = etree.fromstring(theme_xml_bytes)

    # Font scheme
    for maj in root.iter(f'{{{A}}}majorFont'):
        lat = maj.find(f'{{{A}}}latin')
        if lat is not None:
            lat.set('typeface', 'Obviously Narrow')
    for mn in root.iter(f'{{{A}}}minorFont'):
        lat = mn.find(f'{{{A}}}latin')
        if lat is not None:
            lat.set('typeface', 'Galvji')

    # Colour: set dk2 to #231F20
    for dk2 in root.iter(f'{{{A}}}dk2'):
        # Remove children and set srgbClr
        for child in list(dk2):
            dk2.remove(child)
        srgb = etree.SubElement(dk2, f'{{{A}}}srgbClr')
        srgb.set('val', DARK_BG)

    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


# ═════════════════════════════════════════════════════════════════════════════
# Font embedding
# ═════════════════════════════════════════════════════════════════════════════

# ── Font embedding via PowerPoint's .fntdata format ─────────────────────────
# Source: New Case Study Example - Pokemon.pptx — fonts already embedded by
# PowerPoint in its native fntdata format. We reuse them verbatim.
FONT_TEMPLATE_PPTX = os.path.join(
    os.path.dirname(__file__), 'New Case Study Example - Pokemon.pptx'
)

FNTDATA_FILES = {
    'ppt/fonts/font_galvji_r.fntdata':  'ppt/fonts/font1.fntdata',
    'ppt/fonts/font_galvji_b.fntdata':  'ppt/fonts/font2.fntdata',
    'ppt/fonts/font_galvji_i.fntdata':  'ppt/fonts/font3.fntdata',
    'ppt/fonts/font_galvji_bi.fntdata': 'ppt/fonts/font4.fntdata',
    'ppt/fonts/font_obv_b.fntdata':     'ppt/fonts/font5.fntdata',
}

FONT_RIDS = {
    'ppt/fonts/font_galvji_r.fntdata':  'rId_font_gv_r',
    'ppt/fonts/font_galvji_b.fntdata':  'rId_font_gv_b',
    'ppt/fonts/font_galvji_i.fntdata':  'rId_font_gv_i',
    'ppt/fonts/font_galvji_bi.fntdata': 'rId_font_gv_bi',
    'ppt/fonts/font_obv_b.fntdata':     'rId_font_ob_b',
}

FONT_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font'


def embed_fonts(zip_out):
    """Embed Galvji + Obviously Narrow using PowerPoint's native .fntdata format.
    Writes the font files into the output zip and returns the <p:embeddedFontLst>
    XML element to be injected into presentation.xml.
    """
    if not os.path.exists(FONT_TEMPLATE_PPTX):
        print("    WARNING: font template not found — skipping font embedding")
        return None

    with zipfile.ZipFile(FONT_TEMPLATE_PPTX, 'r') as src:
        for dest, src_name in FNTDATA_FILES.items():
            zip_out.writestr(dest, src.read(src_name))

    print(f"    Embedded {len(FNTDATA_FILES)} font variants (.fntdata)")

    P  = 'http://schemas.openxmlformats.org/presentationml/2006/main'
    Rn = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    lst = etree.Element(f'{{{P}}}embeddedFontLst')

    def add_font(typeface, panose, pitchFamily, charset,
                 regular=None, bold=None, italic=None, boldItalic=None):
        ef  = etree.SubElement(lst, f'{{{P}}}embeddedFont')
        fnt = etree.SubElement(ef,  f'{{{P}}}font')
        fnt.set('typeface', typeface)
        if panose:
            fnt.set('panose', panose)
        fnt.set('pitchFamily', str(pitchFamily))
        fnt.set('charset',     str(charset))
        for tag, rid in [('regular', regular), ('bold', bold),
                         ('italic', italic),   ('boldItalic', boldItalic)]:
            if rid:
                el = etree.SubElement(ef, f'{{{P}}}{tag}')
                el.set(f'{{{Rn}}}id', rid)

    r = FONT_RIDS
    add_font('Galvji', panose='020B0504020202020204', pitchFamily=34, charset=77,
             regular=r['ppt/fonts/font_galvji_r.fntdata'],
             bold=r['ppt/fonts/font_galvji_b.fntdata'],
             italic=r['ppt/fonts/font_galvji_i.fntdata'],
             boldItalic=r['ppt/fonts/font_galvji_bi.fntdata'])
    add_font('Obviously Narrow', panose=None, pitchFamily=2, charset=77,
             bold=r['ppt/fonts/font_obv_b.fntdata'])
    return lst


# ═════════════════════════════════════════════════════════════════════════════
# Main converter
# ═════════════════════════════════════════════════════════════════════════════

SLIDE_PATTERN = re.compile(r'^ppt/slides/slide(\d+)\.xml$')


def apply_text_replacements(root, replacements):
    """Replace specific run text values across all slides.
    replacements: dict of {old_text: new_text} — set new_text=None to blank the run.
    """
    for t_el in root.iter(qn('a', 't')):
        txt = t_el.text or ''
        # Try full-string match first, then strip
        for old, new in replacements.items():
            if txt == old or txt.strip() == old.strip():
                t_el.text = '' if new is None else new
                break


def convert_pptx(input_path, output_path, text_replacements=None):
    """Convert a single old-style case study PPTX to the new design.
    text_replacements: optional dict {old_run_text: new_run_text} applied before styling.
    """
    input_path  = str(input_path)
    output_path = str(output_path)

    print(f"\n{'='*60}")
    print(f"Converting: {os.path.basename(input_path)}")
    print(f"        to: {os.path.basename(output_path)}")

    with zipfile.ZipFile(input_path, 'r') as zin:
        names = zin.namelist()

        # ── Count slides ──────────────────────────────────────────────────
        slide_names = sorted(
            [n for n in names if SLIDE_PATTERN.match(n)],
            key=lambda n: int(SLIDE_PATTERN.match(n).group(1))
        )
        total_slides = len(slide_names)
        print(f"  Found {total_slides} slides")

        # ── Pre-scan: identify testimonial slide numbers + their layouts ──
        testimonial_nums = set()
        testimonial_layouts = set()   # layout zip-paths used by testimonial slides
        RELS_LAYOUT_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
        for sname in slide_names:
            snum = int(SLIDE_PATTERN.match(sname).group(1))
            root_tmp = etree.fromstring(zin.read(sname))
            if classify_slide(root_tmp, snum - 1, total_slides) == 'testimonial':
                testimonial_nums.add(snum)
                # Find which layout this slide uses
                rels_path = f'ppt/slides/_rels/slide{snum}.xml.rels'
                if rels_path in zin.namelist():
                    rels_root = etree.fromstring(zin.read(rels_path))
                    CT_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships'
                    for rel in rels_root.findall(f'{{{CT_RELS}}}Relationship'):
                        if rel.get('Type') == RELS_LAYOUT_TYPE:
                            # Target is relative to ppt/slides/ e.g. ../slideLayouts/slideLayout47.xml
                            target = rel.get('Target', '')
                            # Resolve to zip path
                            layout_zip = 'ppt/slideLayouts/' + target.split('/')[-1]
                            testimonial_layouts.add(layout_zip)
        if testimonial_layouts:
            print(f"  Testimonial layouts to clean: {testimonial_layouts}")

        # ── Load quote mark image bytes (from the new template) ───────────
        quote_mark_bytes = None
        if testimonial_nums and os.path.exists(QUOTE_MARK_TEMPLATE):
            try:
                with zipfile.ZipFile(QUOTE_MARK_TEMPLATE, 'r') as qtmpl:
                    quote_mark_bytes = qtmpl.read(QUOTE_MARK_ZIP_PATH)
                print(f"  Loaded quote mark image ({len(quote_mark_bytes)} bytes)")
            except Exception as e:
                print(f"  WARNING: could not load quote mark: {e}")

        RELS_SLIDE_PATTERN = re.compile(r'^ppt/slides/_rels/slide(\d+)\.xml\.rels$')

        # ── Process and write ─────────────────────────────────────────────
        content_types_bytes = zin.read('[Content_Types].xml')

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:

            # Embed quote mark image once (before slide processing)
            if quote_mark_bytes is not None:
                zout.writestr(QUOTE_MARK_MEDIA, quote_mark_bytes)

            prs_root_holder = [None]   # holds presentation.xml root until fonts are embedded

            for name in names:
                data = zin.read(name)

                # ── Drop old embedded font files (we re-embed below) ───
                if name.startswith('ppt/fonts/'):
                    continue

                # ── presentation.xml — strip old fonts, inject new later
                if name == 'ppt/presentation.xml':
                    prs_root = etree.fromstring(data)
                    P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
                    for old in prs_root.findall(f'{{{P}}}embeddedFontLst'):
                        prs_root.remove(old)
                    prs_root_holder[0] = prs_root
                    data = None   # written after embed_fonts below

                # ── presentation.xml.rels — inject font relationships ──
                elif name == 'ppt/_rels/presentation.xml.rels':
                    CT_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships'
                    rels_root = etree.fromstring(data)
                    for rel in list(rels_root):
                        if rel.get('Type') == FONT_REL_TYPE:
                            rels_root.remove(rel)
                    for dest, rid in FONT_RIDS.items():
                        rel = etree.SubElement(rels_root, f'{{{CT_RELS}}}Relationship')
                        rel.set('Id',     rid)
                        rel.set('Type',   FONT_REL_TYPE)
                        rel.set('Target', dest.replace('ppt/', ''))
                    data = etree.tostring(rels_root, xml_declaration=True,
                                         encoding='UTF-8', standalone=True)

                # ── Theme ──────────────────────────────────────────────
                elif name.startswith('ppt/theme/') and name.endswith('.xml'):
                    data = update_theme(data)
                    print(f"  Updated theme: {name}")

                # ── Slide layouts + slide masters: replace legacy fonts ────
                elif (name.startswith('ppt/slideLayouts/') or
                      name.startswith('ppt/slideMasters/')) and name.endswith('.xml'):
                    if name in testimonial_layouts:
                        layout_root = etree.fromstring(data)
                        remove_decorative_elements(layout_root)
                        data = etree.tostring(layout_root, xml_declaration=True,
                                             encoding='UTF-8', standalone=True)
                        print(f"  Cleaned testimonial layout: {name}")
                    data = replace_layout_fonts(data)

                # ── Testimonial slide .rels — inject quote mark rId ────
                elif RELS_SLIDE_PATTERN.match(name):
                    rnum = int(RELS_SLIDE_PATTERN.match(name).group(1))
                    if rnum in testimonial_nums and quote_mark_bytes is not None:
                        data = inject_qmark_rel(data)

                # ── Slides ─────────────────────────────────────────────
                elif SLIDE_PATTERN.match(name):
                    slide_num = int(SLIDE_PATTERN.match(name).group(1))
                    slide_idx = slide_num - 1

                    root = etree.fromstring(data)
                    stype = classify_slide(root, slide_idx, total_slides)
                    print(f"  Slide {slide_num}: [{stype}]")

                    # Apply client-specific text rewrites before styling
                    if text_replacements:
                        apply_text_replacements(root, text_replacements)

                    # Remove decorative elements
                    remove_decorative_elements(root)

                    # Set background
                    if stype in ('cover', 'testimonial'):
                        set_slide_background(root, DARK_BG)
                    else:
                        set_slide_background(root, CREAM_BG)

                    # Apply text/font styling
                    if stype == 'cover':
                        process_cover_slide(root)
                    elif stype == 'content':
                        process_content_slide(root)
                    elif stype == 'grid':
                        process_grid_slide(root)
                    elif stype == 'testimonial':
                        qrid = QUOTE_MARK_RID if quote_mark_bytes is not None else None
                        process_testimonial_slide(root, qmark_rid=qrid)

                    # Catchall: sweep any residual legacy fonts that weren't in
                    # rPr/endParaRPr (e.g. defRPr inside lstStyle defaults).
                    for lat in root.iter(qn('a', 'latin')):
                        if lat.get('typeface', '') in _OLD_BODY_FONTS:
                            lat.set('typeface', FONT_BODY)

                    data = etree.tostring(
                        root,
                        xml_declaration=True,
                        encoding='UTF-8',
                        standalone=True
                    )

                # ── Write entry ────────────────────────────────────────
                if data is not None and name != '[Content_Types].xml':
                    zout.writestr(name, data)

            # ── Embed fonts (.fntdata) ─────────────────────────────────
            print("  Embedding fonts…")
            font_lst = embed_fonts(zout)

            # Inject embeddedFontLst into presentation.xml and write it
            if font_lst is not None and prs_root_holder[0] is not None:
                prs_root_holder[0].append(font_lst)
                zout.writestr('ppt/presentation.xml',
                               etree.tostring(prs_root_holder[0],
                                              xml_declaration=True,
                                              encoding='UTF-8', standalone=True))

            # ── Update Content_Types ──────────────────────────────────
            CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
            ct_root = etree.fromstring(content_types_bytes)
            existing_exts = {el.get('Extension','') for el in ct_root.findall(f'{{{CT_NS}}}Default')}
            for ext, ct in [('fntdata', 'application/x-fontdata'), ('png', 'image/png')]:
                if ext not in existing_exts:
                    el = etree.SubElement(ct_root, f'{{{CT_NS}}}Default')
                    el.set('Extension', ext)
                    el.set('ContentType', ct)
            zout.writestr('[Content_Types].xml',
                          etree.tostring(ct_root, xml_declaration=True,
                                         encoding='UTF-8', standalone=True))

    print(f"  Done → {output_path}")


# ═════════════════════════════════════════════════════════════════════════════
# CLI
# ═════════════════════════════════════════════════════════════════════════════

if __name__ == '__main__':
    # ── Folder layout ────────────────────────────────────────────────────────
    # By default the script looks for PPTX files in   ./old_case_studies/
    # and writes rebranded outputs to                 ./new_case_studies/
    # Both folders are resolved relative to this script file.
    SCRIPT_DIR = Path(__file__).parent
    OLD_DIR    = SCRIPT_DIR / 'old_case_studies'
    NEW_DIR    = SCRIPT_DIR / 'new_case_studies'
    NEW_DIR.mkdir(exist_ok=True)

    # ── Explicit args override the folder defaults ───────────────────────────
    if len(sys.argv) >= 2:
        args = sys.argv[1:]

        # Two-arg mode: explicit input → output
        if len(args) == 2 and args[1].lower().endswith('.pptx') and not Path(args[1]).exists():
            convert_pptx(args[0], args[1])
            sys.exit(0)

        # Batch mode: files given on command line → write to new_case_studies/
        for inp in args:
            p = Path(inp)
            if not p.exists():
                print(f"SKIP (not found): {inp}")
                continue
            stem = p.stem
            skip_markers = ('_rebranded', 'new case study', 'new ')
            if any(m in stem.lower() for m in skip_markers):
                print(f"SKIP: {inp}")
                continue
            clean = stem.replace('Old Case Study Format - ', '').strip()
            convert_pptx(p, NEW_DIR / f"{clean}_rebranded.pptx")
        sys.exit(0)

    # ── No args: auto-convert everything in old_case_studies/ ────────────────
    if not OLD_DIR.exists():
        print(f"Nothing to do — drop PPTX files into:  {OLD_DIR}")
        sys.exit(0)

    pptx_files = sorted(OLD_DIR.glob('*.pptx'))
    skip_markers = ('_rebranded', 'new case study', 'new ')
    pptx_files = [f for f in pptx_files
                  if not any(m in f.stem.lower() for m in skip_markers)]

    if not pptx_files:
        print(f"No PPTX files found in {OLD_DIR}")
        sys.exit(0)

    print(f"Found {len(pptx_files)} file(s) in old_case_studies/")
    for p in pptx_files:
        clean = p.stem.replace('Old Case Study Format - ', '').strip()
        out   = NEW_DIR / f"{clean}_rebranded.pptx"
        if out.exists():
            print(f"SKIP (already converted): {p.name}")
            continue
        try:
            convert_pptx(p, out)
        except zipfile.BadZipFile:
            print(f"  ⚠️  SKIP (still downloading / unreadable): {p.name}")
        except Exception as e:
            print(f"  ❌  ERROR converting {p.name}: {e}")

    print(f"\nAll done — rebranded files are in:  {NEW_DIR}")
