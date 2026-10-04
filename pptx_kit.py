#!/usr/bin/env python3
"""pptx-kit — restyle a presentation by cloning slides of a reference deck.

Commands:
  doctor                                   check runtime (python-pptx, LibreOffice, poppler, fonts)
  profile REF.pptx --out DIR [--fonts-of]  style passport + slide-template catalogue + renders
  outline SRC.pptx --out DIR               content outline of the source deck (+ renders, images)
  build   REF.pptx PLAN.json --out OUT.pptx   clone reference slides and fill them with content
  qa      OUT.pptx --plan PLAN.json --ref-dir DIR --out DIR   render, side-by-side pairs, checks

The creative step (choosing a template slide for every source slide and fitting the text)
is done by the model in PLAN.json. Everything mechanical is done here.
"""
import argparse
import copy
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from collections import Counter, defaultdict
from pathlib import Path

try:
    from pptx import Presentation
    from pptx.util import Emu
    from pptx.opc.constants import RELATIONSHIP_TYPE as RT
    from lxml import etree
except ImportError as e:  # pragma: no cover
    sys.exit(f'python-pptx/lxml not available: {e}. pip install python-pptx')

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
A = '{%s}' % NS['a']
P = '{%s}' % NS['p']
R = '{%s}' % NS['r']
VERSION = '0.1.0'


# ----------------------------------------------------------------------------- utils

def die(msg):
    print('ERROR: ' + msg, file=sys.stderr)
    sys.exit(2)


def q(tag):
    pfx, name = tag.split(':')
    return '{%s}%s' % (NS[pfx], name)


def para_text(p):
    out = []
    for el in p.iter():
        if el.tag == A + 't' and el.text:
            out.append(el.text)
        elif el.tag == A + 'br':
            out.append('\n')
    return ''.join(out)


def body_text(txBody):
    return '\n'.join(para_text(p) for p in txBody.findall(A + 'p'))


def norm(s):
    return re.sub(r'[\W_]+', ' ', (s or '').lower(), flags=re.U).strip()


def run(cmd, timeout=300):
    return subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)


def which(*names):
    for n in names:
        p = shutil.which(n)
        if p:
            return p
    return None


def emu2pt(v):
    return v / 12700.0


# ----------------------------------------------------------------------------- geometry

def _xfrm(el):
    """Return (off_x, off_y, ext_cx, ext_cy, chOff_x, chOff_y, chExt_cx, chExt_cy) for an element."""
    sppr = None
    for tag in ('p:spPr', 'p:grpSpPr'):
        sppr = el.find(q(tag))
        if sppr is not None:
            break
    xfrm = None
    if sppr is not None:
        xfrm = sppr.find(A + 'xfrm')
    if xfrm is None:  # graphicFrame
        xfrm = el.find(q('p:xfrm'))
    if xfrm is None:
        return None
    off, ext = xfrm.find(A + 'off'), xfrm.find(A + 'ext')
    if off is None or ext is None:
        return None
    res = [int(off.get('x')), int(off.get('y')), int(ext.get('cx')), int(ext.get('cy'))]
    choff, chext = xfrm.find(A + 'chOff'), xfrm.find(A + 'chExt')
    if choff is not None and chext is not None:
        res += [int(choff.get('x')), int(choff.get('y')), int(chext.get('cx')), int(chext.get('cy'))]
    else:
        res += [res[0], res[1], res[2], res[3]]
    return res


SHAPE_TAGS = {P + 'sp': 'sp', P + 'pic': 'pic', P + 'grpSp': 'grp', P + 'graphicFrame': 'frame', P + 'cxnSp': 'cxn'}


def walk(spTree, tf=(0, 0, 1.0, 1.0), parent=None, depth=0):
    """Yield dicts for every shape element with absolute bbox in EMU."""
    for el in spTree:
        kind = SHAPE_TAGS.get(el.tag)
        if el.tag == '{http://schemas.openxmlformats.org/markup-compatibility/2006}AlternateContent':
            ch = el.find('{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback')
            if ch is None:
                ch = el.find('{http://schemas.openxmlformats.org/markup-compatibility/2006}Choice')
            if ch is not None:
                yield from walk(ch, tf, parent, depth)
            continue
        if not kind:
            continue
        nv = el.find('.//' + P + 'cNvPr')
        sid = nv.get('id') if nv is not None else None
        name = nv.get('name') if nv is not None else ''
        x = _xfrm(el)
        bbox = None
        if x:
            ox, oy, sx, sy = tf
            bbox = (ox + x[0] * sx, oy + x[1] * sy, x[2] * sx, x[3] * sy)
        info = {'el': el, 'id': sid, 'name': name, 'kind': kind, 'bbox': bbox, 'parent': parent, 'depth': depth}
        yield info
        if kind == 'grp' and x:
            ox, oy, sx, sy = tf
            gx, gy = ox + x[0] * sx, oy + x[1] * sy
            ksx = (x[2] / x[6]) if x[6] else 1.0
            ksy = (x[3] / x[7]) if x[7] else 1.0
            ntf = (gx - x[4] * ksx * sx, gy - x[5] * ksy * sy, ksx * sx, ksy * sy)
            yield from walk(el, ntf, sid, depth + 1)


def find_shape(slide_el, sid):
    for el in slide_el.iter():
        if el.tag in SHAPE_TAGS:
            nv = el.find('.//' + P + 'cNvPr')
            if nv is not None and nv.get('id') == str(sid):
                return el
    return None


# ----------------------------------------------------------------------------- text props

def first_rpr(txBody):
    for r in txBody.iter(A + 'r'):
        rpr = r.find(A + 'rPr')
        if rpr is not None:
            return rpr
    e = txBody.find('.//' + A + 'endParaRPr')
    return e


def font_size_pt(txBody, fallback=None):
    for el in txBody.iter():
        if el.tag in (A + 'rPr', A + 'endParaRPr', A + 'defRPr') and el.get('sz'):
            return int(el.get('sz')) / 100.0
    return fallback


def run_color(rpr):
    if rpr is None:
        return None
    sf = rpr.find(A + 'solidFill')
    if sf is None:
        return None
    c = sf.find(A + 'srgbClr')
    if c is not None:
        return '#' + c.get('val')
    c = sf.find(A + 'schemeClr')
    if c is not None:
        return 'scheme:' + c.get('val')
    return None


def run_font(rpr):
    if rpr is None:
        return None
    lat = rpr.find(A + 'latin')
    return lat.get('typeface') if lat is not None else None


def para_styles(txBody):
    """Style signature of each non-empty paragraph: '12pt b #FFFFFF'."""
    out = []
    for p in txBody.findall(A + 'p'):
        if not para_text(p).strip():
            continue
        rpr = None
        for r in p.iter(A + 'r'):
            rpr = r.find(A + 'rPr')
            if rpr is not None:
                break
        sz = rpr.get('sz') if rpr is not None else None
        sty = (f'{int(sz) / 100:g}pt' if sz else '?pt') + (' b' if rpr is not None and rpr.get('b') == '1' else '')
        c = run_color(rpr)
        f = run_font(rpr)
        out.append(sty + (f' {c}' if c else '') + (f' {f}' if f else ''))
    return out


def _sig(p):
    """Style signature of a template paragraph: size, bold, color, font, level."""
    rpr = next((r.find(A + 'rPr') for r in p.iter(A + 'r') if r.find(A + 'rPr') is not None), None)
    if rpr is None:
        return None
    ppr = p.find(A + 'pPr')
    return (rpr.get('sz'), rpr.get('b'), run_color(rpr), run_font(rpr), ppr.get('lvl') if ppr is not None else None)


def capacity(bbox, fs, txBody):
    """Rough number of characters that fit in the box at font size fs (pt)."""
    if not bbox or not fs:
        return None, False
    body = txBody.find(A + 'bodyPr')
    li = int(body.get('lIns', 91440)) if body is not None else 91440
    ri = int(body.get('rIns', 91440)) if body is not None else 91440
    ti = int(body.get('tIns', 45720)) if body is not None else 45720
    bi = int(body.get('bIns', 45720)) if body is not None else 45720
    grows = body is not None and body.find(A + 'spAutoFit') is not None
    w = max(emu2pt(bbox[2] - li - ri), fs)
    h = max(emu2pt(bbox[3] - ti - bi), fs * 1.2)
    per_line = max(1, int(w / (fs * 0.55)))
    lines = max(1, int(h / (fs * 1.2)))
    return per_line * lines, grows


# ----------------------------------------------------------------------------- placeholders

META_PH = ('dt', 'ftr', 'sldNum')
MAIN_FONT = [None]
TITLE_PH = ('title', 'ctrTitle')


def ph_of(el):
    ph = el.find('.//' + P + 'ph')
    if ph is None:
        return None
    return ph.get('type', 'body'), int(ph.get('idx', 0))


def _find_base(container_el, ptype, pidx):
    by_type = None
    for el in container_el.iter(P + 'sp'):
        info = ph_of(el)
        if not info:
            continue
        t, i = info
        if pidx and i == pidx:
            return el
        if by_type is None and ((t in TITLE_PH) == (ptype in TITLE_PH)) and (t == ptype or (t in ('body', 'obj') and ptype in ('body', 'obj'))):
            by_type = el
    return by_type


def ph_inherited(layout, el):
    """(bbox_emu, font_size_pt) for a placeholder, resolving layout -> master."""
    info = ph_of(el)
    if not info:
        return None, None
    ptype, pidx = info
    bbox = None
    x = _xfrm(el)
    if x:
        bbox = tuple(x[:4])
    fs = None
    master = layout.slide_master
    for cont in (layout._element, master._element):
        base = _find_base(cont, ptype, pidx)
        if base is None:
            continue
        if bbox is None:
            bx = _xfrm(base)
            if bx:
                bbox = tuple(bx[:4])
        if fs is None:
            tb = base.find(P + 'txBody')
            if tb is not None:
                fs = font_size_pt(tb)
    if fs is None:
        st = master._element.find(P + 'txStyles')
        if st is not None:
            node = st.find(P + ('titleStyle' if ptype in TITLE_PH else 'bodyStyle'))
            if node is not None:
                d = node.find('.//' + A + 'lvl1pPr/' + A + 'defRPr')
                if d is not None and d.get('sz'):
                    fs = int(d.get('sz')) / 100.0
    return bbox, fs


def ph_keys(elements):
    """Stable names for placeholders of a slide: title, body, body2, subTitle, pic …"""
    out, cnt = [], Counter()
    for el in elements:
        t, _ = ph_of(el)
        t = 'title' if t in TITLE_PH else 'body' if t in ('body', 'obj') else t
        cnt[t] += 1
        out.append(t if cnt[t] == 1 else f'{t}{cnt[t]}')
    return out


# ----------------------------------------------------------------------------- rendering

def render(pptx_path, outdir, dpi=72, prefix='slide'):
    """pptx -> pdf -> png per slide. Returns (pdf_path, [png paths])."""
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    soffice = which('soffice', 'libreoffice')
    if not soffice:
        print('WARN: LibreOffice not found, skipping render')
        return None, []
    with tempfile.TemporaryDirectory() as td:
        prof = Path(td) / 'lo'
        src = Path(td) / 'in.pptx'
        shutil.copy(pptx_path, src)
        r = run([soffice, f'-env:UserInstallation=file://{prof}', '--headless', '--convert-to', 'pdf',
                 '--outdir', td, str(src)], timeout=600)
        pdf = Path(td) / 'in.pdf'
        if not pdf.exists():
            print('WARN: render failed: ' + (r.stderr or r.stdout)[-500:])
            return None, []
        final_pdf = outdir / f'{prefix}.pdf'
        shutil.copy(pdf, final_pdf)
    for old in outdir.glob(f'{prefix}-*.png'):
        old.unlink()
    run(['pdftoppm', '-r', str(dpi), '-png', str(final_pdf), str(outdir / f'{prefix}-tmp')], timeout=600)
    pngs = sorted(outdir.glob(f'{prefix}-tmp-*.png'), key=lambda p: int(re.findall(r'(\d+)\.png$', p.name)[0]))
    out = []
    for i, pth in enumerate(pngs, 1):
        dst = outdir / f'{prefix}-{i:02d}.png'
        pth.rename(dst)
        out.append(dst)
    return final_pdf, out


def contact_sheet(pngs, out, cols=3, width=480, labels=None):
    from PIL import Image, ImageDraw
    if not pngs:
        return None
    ims = []
    for p in pngs:
        im = Image.open(p).convert('RGB')
        im.thumbnail((width, width * 2))
        ims.append(im)
    h = max(i.height for i in ims)
    rows = (len(ims) + cols - 1) // cols
    sheet = Image.new('RGB', (cols * (width + 10) + 10, rows * (h + 34) + 10), 'white')
    d = ImageDraw.Draw(sheet)
    for i, im in enumerate(ims):
        x = 10 + (i % cols) * (width + 10)
        y = 10 + (i // cols) * (h + 34)
        d.text((x, y), labels[i] if labels else f'#{i + 1}', fill='black')
        sheet.paste(im, (x, y + 16))
        d.rectangle([x - 1, y + 15, x + im.width, y + 16 + im.height], outline=(180, 180, 180))
    sheet.save(out, quality=88)
    return out


def annotate(png, shapes, slide_w, slide_h, out):
    """Draw shape bboxes with ids on top of a render (text=blue, picture=red, group=green)."""
    from PIL import Image, ImageDraw
    im = Image.open(png).convert('RGB')
    d = ImageDraw.Draw(im)
    sx, sy = im.width / slide_w, im.height / slide_h
    colors = {'text': (0, 90, 255), 'pic': (230, 0, 0), 'grp': (0, 150, 60), 'frame': (150, 0, 200)}
    for s in shapes:
        if not s.get('bbox_emu') or s.get('inside_complex'):
            continue
        x, y, w, h = s['bbox_emu']
        col = (255, 140, 0) if s.get('complex') else colors.get(s['role_kind'])
        if not col:
            continue
        box = [x * sx, y * sy, (x + w) * sx, (y + h) * sy]
        d.rectangle(box, outline=col, width=2 if s['role_kind'] != 'grp' else 1)
        lab = '#' + s['id']
        tx, ty = box[0] + 2, max(0, box[1] + 1) if s['role_kind'] != 'grp' else max(0, box[3] - 12)
        d.rectangle([tx - 1, ty - 1, tx + 7 * len(lab), ty + 11], fill=col)
        d.text((tx, ty), lab, fill='white')
    im.save(out)


# ----------------------------------------------------------------------------- catalogue

def shape_catalogue(prs):
    """Per-slide list of shapes with roles; detects chrome (repeated elements)."""
    W, H = prs.slide_width, prs.slide_height
    slides = []
    sig_count = Counter()
    for idx, slide in enumerate(prs.slides, 1):
        rels = slide.part.rels
        items = []
        for s in walk(slide._element.cSld.spTree):
            el, kind, bbox = s['el'], s['kind'], s['bbox']
            area = (bbox[2] * bbox[3]) / float(W * H) if bbox else 0
            item = {'id': s['id'], 'name': s['name'], 'kind': kind, 'parent': s['parent'], 'depth': s['depth'],
                    'bbox_emu': [int(v) for v in bbox] if bbox else None,
                    'bbox': [round(bbox[0] / W, 3), round(bbox[1] / H, 3), round(bbox[2] / W, 3),
                             round(bbox[3] / H, 3)] if bbox else None,
                    'area': round(area, 4)}
            txBody = el.find(P + 'txBody')
            text = body_text(txBody).strip() if txBody is not None else ''
            sig = None
            if kind == 'pic':
                blip = el.find('.//' + A + 'blip')
                rid = blip.get(R + 'embed') if blip is not None else None
                digest = ''
                if rid and rid in rels:
                    try:
                        digest = hashlib.sha1(rels[rid].target_part.blob).hexdigest()[:10]
                    except Exception:
                        pass
                item['image'] = digest
                sig = ('pic', digest, tuple(item['bbox'] or ()))
                item['role_kind'] = 'pic'
            elif kind == 'grp':
                item['role_kind'] = 'grp'
            elif kind == 'frame':
                gd = el.find('.//' + A + 'graphicData')
                uri = gd.get('uri', '') if gd is not None else ''
                item['frame'] = 'table' if uri.endswith('/table') else 'chart' if 'chart' in uri else 'diagram' if 'diagram' in uri else 'other'
                if item['frame'] == 'table':
                    tbl = el.find('.//' + A + 'tbl')
                    rows = tbl.findall(A + 'tr')
                    item['table'] = [[body_text(tc.find(A + 'txBody')).strip() if tc.find(A + 'txBody') is not None else ''
                                      for tc in tr.findall(A + 'tc')] for tr in rows]
                item['role_kind'] = 'frame'
            ph = el.find('.//' + P + 'ph')
            if ph is not None and txBody is not None:
                ibbox, ifs = ph_inherited(slide.slide_layout, el)
                if bbox is None and ibbox:
                    bbox = ibbox
                    area = (bbox[2] * bbox[3]) / float(W * H)
                    item.update({'bbox_emu': [int(v) for v in bbox], 'area': round(area, 4),
                                 'bbox': [round(bbox[0] / W, 3), round(bbox[1] / H, 3), round(bbox[2] / W, 3), round(bbox[3] / H, 3)]})
            if txBody is not None and (text or ph is not None):
                rpr = first_rpr(txBody)
                fs = font_size_pt(txBody)
                if fs is None and ph is not None:
                    fs = ifs
                cap, grows = capacity(bbox, fs or 18, txBody)
                item.update({'text': text, 'size': fs, 'bold': (rpr is not None and rpr.get('b') == '1'),
                             'color': run_color(rpr), 'font': run_font(rpr), 'chars': len(text),
                             'capacity': cap, 'grows': grows,
                             'placeholder': ph.get('type', 'body') if ph is not None else None,
                             'paragraphs': len([p for p in txBody.findall(A + 'p') if para_text(p).strip()]),
                             'para_styles': para_styles(txBody)})
                item['role_kind'] = 'text'
                if ph is not None and ph.get('type') in META_PH:
                    item['meta'] = True
                if text:
                    sig = ('text', norm(text)[:60], tuple(item['bbox'] or ()))
            elif kind == 'sp':
                item['role_kind'] = 'shape'
                sppr = el.find(P + 'spPr')
                fill = sppr.find(A + 'solidFill') if sppr is not None else None
                if fill is not None:
                    c = fill.find(A + 'srgbClr')
                    c2 = fill.find(A + 'schemeClr')
                    item['fill'] = ('#' + c.get('val')) if c is not None else ('scheme:' + c2.get('val')) if c2 is not None else None
                elif sppr is not None and sppr.find(A + 'noFill') is not None:
                    item['fill'] = 'none'
                geom = sppr.find(A + 'prstGeom') if sppr is not None else None
                item['geom'] = geom.get('prst') if geom is not None else 'custom'
            if sig:
                item['_sig'] = sig
                sig_count[sig] += 1
            items.append(item)
        slides.append({'index': idx, 'layout': slide.slide_layout.name, 'shapes': items})
    n = len(slides)
    for sl in slides:
        for it in sl['shapes']:
            sig = it.pop('_sig', None)
            it['chrome'] = bool(it.get('meta') or (sig and n >= 3 and sig_count[sig] >= max(2, round(n * 0.4))))
            if it['kind'] == 'pic' and not it['chrome']:
                it['pic_role'] = 'background' if it['area'] > 0.85 else 'icon' if it['area'] < 0.02 else 'content'
    # complex groups (UI mockups, diagrams): collapse
    for sl in slides:
        by_id = {it['id']: it for it in sl['shapes']}
        desc = Counter()
        for it in sl['shapes']:
            par = it['parent']
            while par:
                desc[par] += 1
                par = by_id[par]['parent'] if par in by_id else None
        complex_ids = {g for g, n in desc.items() if n > 12}
        for it in sl['shapes']:
            if it['id'] in complex_ids:
                it['complex'] = desc[it['id']]
            par = it['parent']
            while par:
                if par in complex_ids:
                    it['inside_complex'] = par
                par = by_id[par]['parent'] if par in by_id else None
        # containment: which text/pic shapes sit inside a decorative shape (card)
        for it in sl['shapes']:
            if it.get('role_kind') == 'shape' and it['bbox'] and 0.002 < it['area'] < 0.5:
                x, y, w, h = it['bbox']
                inside = []
                for o in sl['shapes']:
                    if o is it or not o['bbox'] or o.get('role_kind') not in ('text', 'pic') or o.get('chrome'):
                        continue
                    cx, cy = o['bbox'][0] + o['bbox'][2] / 2, o['bbox'][1] + o['bbox'][3] / 2
                    if x <= cx <= x + w and y <= cy <= y + h and o['area'] < it['area']:
                        inside.append(o['id'])
                if inside:
                    it['contains'] = inside
        # picture markers of a list: small pics just left of a multi-paragraph text box, inside its y-range
        for t in sl['shapes']:
            if t.get('role_kind') != 'text' or not t['bbox'] or t.get('paragraphs', 0) < 2:
                continue
            tx, ty, tw, th = t['bbox']
            marks = []
            for o in sl['shapes']:
                if o.get('kind') == 'pic' and o['bbox'] and o['area'] < 0.004 and not o.get('chrome'):
                    cx, cy = o['bbox'][0] + o['bbox'][2] / 2, o['bbox'][1] + o['bbox'][3] / 2
                    if tx - 0.05 <= cx <= tx + 0.02 and ty <= cy <= ty + th:
                        marks.append((cy, o['id']))
            if len(marks) >= 2:
                t['markers'] = [m[1] for m in sorted(marks)]
                fs = t.get('size') or 12
                t['line_chars'] = max(1, int(emu2pt(t['bbox_emu'][2]) / (fs * 0.55)))
        has_text = any(t.get('role_kind') == 'text' for t in sl['shapes'])
        full_pic = any(t.get('kind') == 'pic' and t['area'] > 0.85 for t in sl['shapes'])
        sl['image_only'] = full_pic and not has_text
    # role guess for text: title = largest font among top half, non-chrome
    for sl in slides:
        texts = [t for t in sl['shapes'] if t.get('role_kind') == 'text' and not t['chrome']]
        if texts:
            ph = [t for t in texts if t.get('placeholder') in ('title', 'ctrTitle')]
            top = sorted([t for t in texts if t['bbox'] and t['bbox'][1] < 0.2 and (t['size'] or 0) >= 10],
                         key=lambda t: t['bbox'][1])
            cand = [t for t in texts if t['bbox'] and t['bbox'][1] < 0.45] or texts
            title = ph[0] if ph else top[0] if top else max(cand, key=lambda t: (t['size'] or 0))
            title['role'] = 'title'
    return slides


def layout_catalogue(prs):
    W, H = prs.slide_width, prs.slide_height
    out = []
    for li, layout in enumerate(prs.slide_layouts, 1):
        els = [el for el in layout._element.cSld.spTree.iter(P + 'sp') if ph_of(el) and ph_of(el)[0] not in META_PH]
        keys = ph_keys(els)
        phs = []
        for key, el in zip(keys, els):
            x = _xfrm(el)
            bbox, fs = ph_inherited(layout, el) if True else (None, None)
            if x:
                bbox = tuple(x[:4])
            tb = el.find(P + 'txBody')
            cap = capacity(bbox, fs or 18, tb)[0] if (bbox and tb is not None) else None
            phs.append({'key': key, 'type': ph_of(el)[0], 'size': fs, 'capacity': cap,
                        'bbox_emu': [int(v) for v in bbox] if bbox else None,
                        'bbox': [round(bbox[0] / W, 3), round(bbox[1] / H, 3), round(bbox[2] / W, 3), round(bbox[3] / H, 3)] if bbox else None})
        deco = sum(1 for el in layout._element.cSld.spTree if el.tag in SHAPE_TAGS and not ph_of(el))
        out.append({'index': li, 'name': layout.name, 'placeholders': phs, 'decor_shapes': deco,
                    'used_by_slides': [i for i, sl in enumerate(prs.slides, 1) if sl.slide_layout is layout]})
    return out


def sample_fill(ph):
    if ph['type'] in TITLE_PH or ph['key'].startswith('title'):
        return 'Заголовок слайда'
    if ph['type'] == 'subTitle':
        return 'Подзаголовок'
    if ph['type'] in ('body', 'obj'):
        return ['Первый пункт списка', 'Второй пункт списка', 'Третий пункт']
    return None


def render_layouts(ref_path, layouts, outdir):
    """Deck with one sample slide per layout -> render/layout-NN.png"""
    prs = Presentation(ref_path)
    originals = list(prs.slides)
    for lc in layouts:
        sl = prs.slides.add_slide(prs.slide_layouts[lc['index'] - 1])
        els = [ph._element for ph in sl.placeholders if ph_of(ph._element) and ph_of(ph._element)[0] not in META_PH]
        for key, el in zip(ph_keys(els), els):
            spec = next((p for p in lc['placeholders'] if p['key'] == key), None)
            val = sample_fill(spec) if spec else None
            if val and el.find(P + 'txBody') is not None:
                fill_text(el, val if isinstance(val, list) else f'{val}')
    for s in originals:
        delete_slide(prs, s)
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td) / 'layouts.pptx'
        prs.save(tmp)
        _, pngs = render(tmp, outdir, dpi=60, prefix='layout')
    return pngs


def palette_and_fonts(pptx_path):
    import zipfile
    colors, fonts, theme = Counter(), Counter(), {}
    with zipfile.ZipFile(pptx_path) as z:
        for name in z.namelist():
            if not name.endswith('.xml'):
                continue
            data = z.read(name).decode('utf8', 'ignore')
            if name.startswith('ppt/theme/theme1'):
                for m in re.finditer(r'<a:(dk1|lt1|dk2|lt2|accent\d|hlink)>.*?(?:val|lastClr)="([0-9A-Fa-f]{6})"', data, re.S):
                    theme[m.group(1)] = '#' + m.group(2).upper()
                for m in re.finditer(r'<a:(majorFont|minorFont)><a:latin typeface="([^"]*)"', data):
                    theme[m.group(1)] = m.group(2)
            if name.startswith(('ppt/slides/', 'ppt/slideLayouts/', 'ppt/slideMasters/')):
                w = 3 if name.startswith('ppt/slides/') else 1
                for c in re.findall(r'srgbClr val="([0-9A-Fa-f]{6})"', data):
                    colors['#' + c.upper()] += w
                for f in re.findall(r'<a:latin typeface="([^"]+)"', data):
                    if not f.startswith('+'):
                        fonts[f] += w
    return colors, fonts, theme


METRIC_ALIASES = {'calibri': 'carlito', 'cambria': 'caladea', 'arial': 'liberation sans',
                  'helvetica': 'liberation sans', 'times new roman': 'liberation serif', 'courier new': 'liberation mono'}


def font_available(family):
    if not which('fc-match'):
        return True
    r = run(['fc-match', '-f', '%{family}', family])
    got = {f.strip().lower() for f in r.stdout.split(',')}
    want = family.lower()
    return want in got or METRIC_ALIASES.get(want) in got


# ----------------------------------------------------------------------------- commands

def cmd_doctor(a):
    ok = True
    import pptx
    print(f'pptx-kit {VERSION}')
    print(f'python-pptx {pptx.__version__}')
    for tool in ('soffice', 'pdftoppm', 'pdftotext', 'fc-list'):
        p = which(tool) or (which('libreoffice') if tool == 'soffice' else None)
        print(f'{"OK " if p else "MISSING"} {tool}: {p}')
        ok &= bool(p)
    try:
        import PIL
        print(f'OK  Pillow {PIL.__version__}')
    except ImportError:
        print('MISSING Pillow (pip install pillow)')
        ok = False
    print('fonts dir for custom fonts: ~/.fonts (then run: fc-cache -f)')
    print('READY' if ok else 'NOT READY')


def cmd_profile(a):
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    prs = Presentation(a.ref)
    W, H = prs.slide_width, prs.slide_height
    cat = shape_catalogue(prs)
    colors, fonts, theme = palette_and_fonts(a.ref)
    missing = [f for f, _ in fonts.most_common() if not font_available(f)]
    pdf, pngs = render(a.ref, out / 'render', dpi=a.dpi)
    if pngs:
        contact_sheet(pngs, out / 'contact.jpg')
        for sl, png in zip(cat, pngs):
            annotate(png, sl['shapes'], W, H, out / 'render' / f'annot-{sl["index"]:02d}.png')
    lays = layout_catalogue(prs)
    lpngs = render_layouts(a.ref, lays, out / 'render') if pngs else []
    prof = {'source': str(a.ref), 'slide_size_emu': [W, H], 'aspect': f'{W / H:.3f}', 'layouts_catalogue': lays,
            'aspect_name': '16:9' if abs(W / H - 16 / 9) < 0.02 else '4:3' if abs(W / H - 4 / 3) < 0.02 else 'custom',
            'theme': theme, 'colors': colors.most_common(16), 'fonts': fonts.most_common(10),
            'fonts_missing_in_renderer': missing, 'layouts': [l.name for l in prs.slide_layouts], 'slides': cat}
    (out / 'profile.json').write_text(json.dumps(prof, ensure_ascii=False, indent=1))
    md = [f'# Профиль референса: {Path(a.ref).name}', '',
          f'- Формат: {prof["aspect_name"]} ({W}x{H} EMU), слайдов: {len(cat)}',
          f'- Палитра (по частоте): ' + ', '.join(f'{c}×{n}' for c, n in colors.most_common(12)),
          f'- Тема: ' + ', '.join(f'{k}={v}' for k, v in theme.items()),
          f'- Шрифты: ' + ', '.join(f'{f}×{n}' for f, n in fonts.most_common(6))]
    if missing:
        md.append(f'- ⚠ Нет в рендерере (рендеры будут с заменой шрифта): {", ".join(missing)}')
    md += ['', 'Обозначения: TEXT — текстовый слот (можно заполнить), PIC — картинка, GRP — группа (карточка), '
           'SHAPE — декор без текста. [chrome] — повторяется на многих слайдах (лого, колонтитул), сохраняется всегда. '
           'cap≈N — примерно сколько знаков влезет.', '',
           'Картинки с разметкой id: render/annot-NN.png; чистые рендеры: render/slide-NN.png; обзор: contact.jpg', '']
    for sl in cat:
        md.append(f'## Шаблон {sl["index"]} (layout «{sl["layout"]}»)')
        small = []
        if sl.get('image_only'):
            md.append('- ⚠ Слайд целиком картинка (текст «запечён» в изображение). Как шаблон для текста НЕ использовать.')
        for it in sl['shapes']:
            if it.get('inside_complex'):
                continue
            if it.get('complex'):
                bb = it['bbox']
                pos = f'@{bb[0]:.2f},{bb[1]:.2f} {bb[2]:.2f}x{bb[3]:.2f}' if bb else ''
                md.append(f'{"  " * it["depth"]}- #{it["id"]} COMPLEX-GROUP ({it["complex"]} фигур: макет/схема/скриншот) '
                          f'area={it["area"]:.0%} {pos} — удалить целиком через "delete" или оставить как иллюстрацию')
                continue
            ind = '  ' * it['depth']
            bb = it['bbox']
            pos = f'@{bb[0]:.2f},{bb[1]:.2f} {bb[2]:.2f}x{bb[3]:.2f}' if bb else ''
            ch = ' [chrome]' if it['chrome'] else ''
            rk = it.get('role_kind')
            if rk == 'text':
                t = it['text'].replace('\n', ' / ')
                t = t[:90] + ('…' if len(t) > 90 else '')
                role = ' TITLE' if it.get('role') == 'title' else ''
                if it.get('placeholder'):
                    role += f' ph:{it["placeholder"]}' + (' (пустой)' if not it['text'] else '')
                sty = f'{it["size"] or "?"}pt{" b" if it["bold"] else ""} {it["color"] or ""} {it["font"] or ""}'.strip()
                capt = f'cap≈{it["capacity"]}{"+" if it["grows"] else ""}'
                md.append(f'{ind}- #{it["id"]} TEXT{role}{ch} «{t}» ({it["chars"]} зн., {it["paragraphs"]} абз.) {sty} {capt} {pos}')
                if it.get('markers'):
                    md.append(f'{ind}    список с маркерами-картинками: {len(it["markers"])} шт. ({" ".join("#" + m for m in it["markers"])}), '
                              f'по одному на строку → каждый пункт в одну строку (≤{it["line_chars"]} зн.); лишние маркеры build удалит сам')
                if it['paragraphs'] > 1 and len(set(it['para_styles'])) > 1:
                    md.append(f'{ind}    стили абзацев: ' + ' | '.join(f'{i}: {st}' for i, st in enumerate(it['para_styles'])))
            elif rk == 'pic':
                if it['area'] < 0.004 and not it['chrome']:
                    small.append(f'#{it["id"]}@{bb[0]:.2f},{bb[1]:.2f}' if bb else f'#{it["id"]}')
                    continue
                md.append(f'{ind}- #{it["id"]} PIC {it.get("pic_role", "")}{ch} area={it["area"]:.0%} {pos}')
            elif rk == 'grp':
                md.append(f'{ind}- #{it["id"]} GRP{ch} {pos}')
            elif rk == 'frame':
                extra = f' {len(it.get("table", []))} строк' if it.get('table') else ''
                md.append(f'{ind}- #{it["id"]} {it["frame"].upper()}{extra} {pos}')
            else:
                if it['area'] > 0.01 and it.get('fill') != 'none':
                    cont = f' содержит: {", ".join("#" + c for c in it["contains"])}' if it.get('contains') else ''
                    md.append(f'{ind}- #{it["id"]} SHAPE {it.get("geom", "")} fill={it.get("fill") or "?"}{ch}{cont} {pos}')
        if small:
            md.append(f'- мелкие иконки/маркеры (могут остаться лишними, если пунктов меньше — удаляй через delete): {" ".join(small)}')
        md.append('')
    md += ['# Макеты (layouts) — шаблоны «с нуля»', '',
           'Используй, когда среди готовых слайдов нет подходящего: {"layout": N, "fill": {"title": ..., "body": [...]}}. '
           'Ключи плейсхолдеров указаны ниже. Образцы: render/layout-NN.png', '']
    for lc in lays:
        used = f', используется слайдами {lc["used_by_slides"]}' if lc['used_by_slides'] else ', в файле не используется'
        md.append(f'## Макет L{lc["index"]} «{lc["name"]}» (декор-фигур: {lc["decor_shapes"]}{used})')
        for ph in lc['placeholders']:
            bb = ph['bbox']
            pos = f'@{bb[0]:.2f},{bb[1]:.2f} {bb[2]:.2f}x{bb[3]:.2f}' if bb else ''
            md.append(f'- "{ph["key"]}" ({ph["type"]}) {ph["size"] or "?"}pt cap≈{ph["capacity"]} {pos}')
        md.append('')
    (out / 'profile.md').write_text('\n'.join(md))
    print(f'profile: {out / "profile.md"} ({len(cat)} templates), renders: {len(pngs)}')
    if missing:
        print('WARN fonts missing in renderer: ' + ', '.join(missing))


def cmd_outline(a):
    out = Path(a.out)
    (out / 'img').mkdir(parents=True, exist_ok=True)
    prs = Presentation(a.src)
    W, H = prs.slide_width, prs.slide_height
    data, md = [], [f'# Содержание исходника: {Path(a.src).name}', '']
    for idx, slide in enumerate(prs.slides, 1):
        blocks, images, notes_flags = [], [], []
        rels = slide.part.rels
        for s in walk(slide._element.cSld.spTree):
            el, bbox = s['el'], s['bbox']
            txBody = el.find(P + 'txBody')
            if txBody is not None:
                paras = []
                for p in txBody.findall(A + 'p'):
                    t = para_text(p).strip()
                    if not t:
                        continue
                    ppr = p.find(A + 'pPr')
                    lvl = int(ppr.get('lvl', 0)) if ppr is not None else 0
                    flag = None
                    for r in p.iter(A + 'r'):
                        c = run_color(r.find(A + 'rPr'))
                        if c and c.startswith('#'):
                            rr, gg, bb = int(c[1:3], 16), int(c[3:5], 16), int(c[5:7], 16)
                            if rr > 190 and gg < 90 and bb < 90:
                                flag = 'red'
                                notes_flags.append(''.join(x.text or '' for x in r.iter(A + 't')))
                    paras.append({'text': t, 'level': lvl, **({'flag': flag} if flag else {})})
                if paras:
                    ph = el.find('.//' + P + 'ph')
                    blocks.append({'id': s['id'], 'bbox': bbox, 'size': font_size_pt(txBody),
                                   'placeholder': ph.get('type', 'body') if ph is not None else None, 'paragraphs': paras})
            if s['kind'] == 'pic':
                blip = el.find('.//' + A + 'blip')
                rid = blip.get(R + 'embed') if blip is not None else None
                if rid and rid in rels:
                    part = rels[rid].target_part
                    ext = Path(part.partname).suffix or '.png'
                    fn = out / 'img' / f's{idx:02d}_{s["id"]}{ext}'
                    fn.write_bytes(part.blob)
                    area = (bbox[2] * bbox[3]) / float(W * H) if bbox else 0
                    images.append({'id': s['id'], 'file': str(fn), 'area': round(area, 3)})
            if s['kind'] == 'frame':
                tbl = el.find('.//' + A + 'tbl')
                if tbl is not None:
                    rows = [[body_text(tc.find(A + 'txBody')).strip() for tc in tr.findall(A + 'tc')] for tr in tbl.findall(A + 'tr')]
                    blocks.append({'id': s['id'], 'bbox': bbox, 'table': rows})
        def key(b):
            bb = b.get('bbox') or (0, 0, 0, 0)
            return (round(bb[1] / H * 10), bb[0])
        blocks.sort(key=key)
        title = None
        cands = [b for b in blocks if b.get('placeholder') in ('title', 'ctrTitle')] or \
                [b for b in blocks if b.get('paragraphs')]
        if cands:
            title = max(cands, key=lambda b: ((b.get('size') or 0) + (100 if b.get('placeholder') in ('title', 'ctrTitle') else 0)))
        notes = ''
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip() if slide.notes_slide.notes_text_frame else ''
        for b in blocks:
            b.pop('bbox', None)
        rec = {'index': idx, 'title': (title['paragraphs'][0]['text'] if title and title.get('paragraphs') else None),
               'blocks': blocks, 'images': images, 'notes': notes, 'flags_red_text': notes_flags}
        data.append(rec)
        md.append(f'## Слайд {idx}: {rec["title"] or "(без заголовка)"}')
        for b in blocks:
            if b.get('table'):
                md.append(f'- [#{b["id"]} таблица]')
                for row in b['table']:
                    md.append('  | ' + ' | '.join(row) + ' |')
                continue
            if title is b and len(b['paragraphs']) == 1:
                continue
            for p in b['paragraphs']:
                mark = ' ⚠(красный текст — похоже на пометку редактора)' if p.get('flag') else ''
                md.append(f'{"  " * p["level"]}- {p["text"]}{mark}')
        for im in images:
            md.append(f'- [картинка #{im["id"]} {im["area"]:.0%} площади: {im["file"]}]')
        if notes:
            md.append(f'- [заметки докладчика]: {notes[:400]}')
        md.append('')
    (out / 'outline.json').write_text(json.dumps(data, ensure_ascii=False, indent=1))
    (out / 'outline.md').write_text('\n'.join(md))
    if not a.no_render:
        pdf, pngs = render(a.src, out / 'render', dpi=60)
        if pngs:
            contact_sheet(pngs, out / 'contact.jpg')
    print(f'outline: {out / "outline.md"} ({len(data)} slides)')


# ----------------------------------------------------------------------------- build

def clone_slide(prs, src_slide):
    new = prs.slides.add_slide(src_slide.slide_layout)
    sld = new._element
    # keep the existing cSld/spTree objects: python-pptx caches slide.shapes bound to this spTree
    src_cSld = copy.deepcopy(src_slide._element.cSld)
    cSld, spTree = sld.cSld, sld.cSld.spTree
    for ch in list(spTree):
        spTree.remove(ch)
    for ch in list(src_cSld.spTree):
        spTree.append(ch)
    for ch in list(cSld):
        if ch is not spTree:
            cSld.remove(ch)
    for ch in list(src_cSld):
        if ch.tag == P + 'spTree':
            continue
        if ch.tag == P + 'bg':
            cSld.insert(0, ch)
        else:
            cSld.append(ch)
    if src_cSld.get('name'):
        cSld.set('name', src_cSld.get('name'))
    for tag in ('p:clrMapOvr', 'p:transition', 'p:timing'):
        old = sld.find(q(tag))
        if old is not None:
            sld.remove(old)
        srcel = src_slide._element.find(q(tag))
        if srcel is not None and tag == 'p:clrMapOvr':
            sld.append(copy.deepcopy(srcel))
    rmap = {}
    for rel in list(src_slide.part.rels.values()):
        if rel.reltype in (RT.SLIDE_LAYOUT, RT.NOTES_SLIDE):
            continue
        if rel.is_external:
            rmap[rel.rId] = new.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        else:
            rmap[rel.rId] = new.part.relate_to(rel.target_part, rel.reltype)
    for el in sld.iter():
        for k, v in list(el.attrib.items()):
            if k.startswith(R) and v in rmap:
                el.set(k, rmap[v])
    return new


def delete_slide(prs, slide):
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        if prs.part.related_part(sldId.rId) is slide.part:
            rid = sldId.rId
            sldIdLst.remove(sldId)
            prs.part.drop_rel(rid)
            return


def _make_runs(text, rpr_tpl):
    """'**bold** text<br>next line' -> list of a:r / a:br elements."""
    out = []
    lines = re.split(r'<br\s*/?>|\u000b', text)
    for li, line in enumerate(lines):
        if li:
            br = etree.Element(A + 'br')
            if rpr_tpl is not None:
                b = copy.deepcopy(rpr_tpl)
                b.tag = A + 'rPr'
                br.append(b)
            out.append(br)
        parts = line.split('**')
        for pi, seg in enumerate(parts):
            if not seg:
                continue
            r = etree.Element(A + 'r')
            if rpr_tpl is not None:
                rp = copy.deepcopy(rpr_tpl)
                rp.tag = A + 'rPr'
                for ch in list(rp):
                    if ch.tag == A + 'extLst':
                        rp.remove(ch)
                if pi % 2 == 1:
                    rp.set('b', '1')
                r.append(rp)
            t = etree.SubElement(r, A + 't')
            t.text = seg
            out.append(r)
    return out


def _para_from_tpl(tpl_p, text, level=None):
    p = copy.deepcopy(tpl_p)
    rpr = None
    for r in tpl_p.iter(A + 'r'):
        rpr = r.find(A + 'rPr')
        if rpr is not None:
            break
    if rpr is None:
        rpr = tpl_p.find(A + 'endParaRPr')
    end = p.find(A + 'endParaRPr')
    for ch in list(p):
        if ch.tag != A + 'pPr':
            p.remove(ch)
    for el in _make_runs(text, rpr):
        p.append(el)
    if para_text(tpl_p).endswith('\n') and text and not text.endswith('<br>'):
        br = etree.SubElement(p, A + 'br')
        if rpr is not None:
            b = copy.deepcopy(rpr)
            b.tag = A + 'rPr'
            br.append(b)
    if end is not None:
        p.append(end)
    if level is not None:
        ppr = p.find(A + 'pPr')
        if ppr is None:
            ppr = etree.Element(A + 'pPr')
            p.insert(0, ppr)
        ppr.set('lvl', str(level))
    return p


def _normalize_items(value):
    """fill value -> (items:list[dict], opts:dict)"""
    opts = {}
    if isinstance(value, dict):
        opts = {k: v for k, v in value.items() if k not in ('paragraphs', 'text')}
        value = value.get('paragraphs', value.get('text', ''))
    if isinstance(value, str):
        value = value.split('\n')
    items = []
    for v in value:
        if isinstance(v, str):
            items.append({'text': v})
        else:
            items.append(dict(v))
    return items, opts


def fill_text(el, value):
    txBody = el.find(P + 'txBody')
    if txBody is None:
        return 'no-text-body'
    items, opts = _normalize_items(value)
    tpl_ps = txBody.findall(A + 'p')
    if not tpl_ps:
        tpl_ps = [etree.SubElement(txBody, A + 'p')]
    nonempty = [i for i, p in enumerate(tpl_ps) if para_text(p).strip()] or [0]
    # lists (3+ paragraphs): an item takes the style of "its" template paragraph only when that style is the
    # dominant list style (or it is the first paragraph); one-off accents of the template (e.g. a red last item)
    # are not inherited by new items
    regular = list(nonempty)
    if len(nonempty) >= 3:
        sigs = Counter(_sig(tpl_ps[i]) for i in nonempty[1:])
        # drop one-off accent paragraphs (unique style) from the sequence used for new items
        regular = [nonempty[0]] + [i for i in nonempty[1:] if sigs[_sig(tpl_ps[i])] > 1] or list(nonempty)
        if len(regular) < 2:
            regular = list(nonempty)
    new_ps = []
    for j, it in enumerate(items):
        if 'style' in it:
            k = nonempty[min(int(it['style']), len(nonempty) - 1)]
        elif j < len(regular):
            k = regular[j]
        elif len(regular) >= 3 and _sig(tpl_ps[regular[-1]]) != _sig(tpl_ps[regular[-2]]):
            # alternating pattern (heading / caption …): continue the alternation
            k = regular[-2 + (j - len(regular)) % 2]
        else:
            k = regular[-1]
        # keep template spacer paragraphs between consecutive styled paragraphs
        if j and j < len(regular) and 'style' not in it and opts.get('spacers', True):
            for si in range(regular[j - 1] + 1, regular[j]):
                if not para_text(tpl_ps[si]).strip():  # only empty spacer paragraphs, never template text
                    new_ps.append(copy.deepcopy(tpl_ps[si]))
        if it.get('text', '') == '' and len(items) > 1:
            new_ps.append(_para_from_tpl(tpl_ps[k], '', it.get('level')))
        else:
            new_ps.append(_para_from_tpl(tpl_ps[k], it.get('text', ''), it.get('level')))
    if len(nonempty) > 1 and opts.get('anchor_top', True):
        bp = txBody.find(A + 'bodyPr')
        if bp is not None and bp.get('anchor') in ('ctr', 'b'):
            bp.set('anchor', 't')
    for p in tpl_ps:
        txBody.remove(p)
    for p in new_ps:
        txBody.append(p)
    if opts.get('shrink'):
        ratio = float(opts['shrink']) if not isinstance(opts['shrink'], bool) else None
        return ('shrink', ratio)
    return None


def shrink_text(el, bbox, ratio=None):
    txBody = el.find(P + 'txBody')
    fs = font_size_pt(txBody, 18)
    if ratio is None:
        cap, _ = capacity(bbox, fs, txBody)
        n = len(body_text(txBody))
        ratio = min(1.0, max(0.7, (cap / n) ** 0.5)) if cap and n else 1.0
    for e in txBody.iter():
        if e.tag in (A + 'rPr', A + 'endParaRPr') and e.get('sz'):
            e.set('sz', str(int(int(e.get('sz')) * ratio)))
    return ratio


def fill_table(el, rows):
    tbl = el.find('.//' + A + 'tbl')
    if tbl is None:
        return 'no-table'
    trs = tbl.findall(A + 'tr')
    for ri, row in enumerate(rows):
        if ri >= len(trs):
            nr = copy.deepcopy(trs[-1])
            tbl.append(nr)
            trs.append(nr)
        tcs = trs[ri].findall(A + 'tc')
        for ci, val in enumerate(row[:len(tcs)]):
            tb = tcs[ci].find(A + 'txBody')
            if tb is not None:
                holder = etree.Element(P + 'sp')
                holder.append(tb)
                fill_text(holder, str(val))
                tcs[ci].insert(0, tb)
    for tr in trs[len(rows):]:
        tbl.remove(tr)
    return None


def replace_image(slide, el, path, mode='cover'):
    from PIL import Image
    blip = el.find('.//' + A + 'blip')
    if blip is None:
        return 'no-blip'
    _, rid = slide.part.get_or_add_image_part(path)
    blip.set(R + 'embed', rid)
    bf = blip.getparent()
    for sr in bf.findall(A + 'srcRect'):
        bf.remove(sr)
    x = _xfrm(el)
    if x and mode == 'cover':
        with Image.open(path) as im:
            iw, ih = im.size
        box_ar, img_ar = x[2] / x[3], iw / ih
        sr = etree.Element(A + 'srcRect')
        if img_ar > box_ar:
            c = int((1 - box_ar / img_ar) / 2 * 100000)
            sr.set('l', str(c)); sr.set('r', str(c))
        elif img_ar < box_ar:
            c = int((1 - img_ar / box_ar) / 2 * 100000)
            sr.set('t', str(c)); sr.set('b', str(c))
        blip.addnext(sr)
    return None


def _set_xfrm(el, x=None, cx=None):
    for tag in ('p:spPr', 'p:grpSpPr'):
        sppr = el.find(q(tag))
        if sppr is not None:
            xf = sppr.find(A + 'xfrm')
            if xf is not None:
                if x is not None:
                    xf.find(A + 'off').set('x', str(int(x)))
                if cx is not None:
                    xf.find(A + 'ext').set('cx', str(int(cx)))
                return True
    return False


def respace(root, cards, info, tpl_shapes):
    """Distribute the given cards (with the shapes they contain) evenly over the original row span.
    Row span = horizontal extent of all template shapes with the same fill and similar y/height."""
    base = [info.get(c) for c in cards]
    if not all(base) or any(b.get('depth') for b in base):
        return 'cards must be top-level SHAPE ids from profile'
    ref = base[0]
    y, h = ref['bbox_emu'][1], ref['bbox_emu'][3]
    row = [t for t in tpl_shapes if t.get('role_kind') == 'shape' and t.get('fill') == ref.get('fill') and t['bbox_emu']
           and abs(t['bbox_emu'][1] - y) < h * 0.2 and abs(t['bbox_emu'][3] - h) < h * 0.2]
    row.sort(key=lambda t: t['bbox_emu'][0])
    if len(row) <= len(cards):
        return 'nothing to respace'
    x0 = row[0]['bbox_emu'][0]
    x1 = row[-1]['bbox_emu'][0] + row[-1]['bbox_emu'][2]
    gap = row[1]['bbox_emu'][0] - (row[0]['bbox_emu'][0] + row[0]['bbox_emu'][2])
    n = len(cards)
    neww = (x1 - x0 - gap * (n - 1)) / n
    for i, b in enumerate(sorted(base, key=lambda t: t['bbox_emu'][0])):
        ox, ow = b['bbox_emu'][0], b['bbox_emu'][2]
        nx = x0 + i * (neww + gap)
        dx, dw = nx - ox, neww - ow
        el = find_shape(root, b['id'])
        _set_xfrm(el, nx, neww)
        for cid in b.get('contains', []):
            c = info.get(cid)
            cel = find_shape(root, cid)
            if not c or cel is None or c.get('depth'):
                continue
            cx, cw = c['bbox_emu'][0], c['bbox_emu'][2]
            if c.get('role_kind') == 'text':
                _set_xfrm(cel, cx + dx, cw + dw)
            else:
                _set_xfrm(cel, cx + dx)
    return None


def remove_el(el):
    par = el.getparent()
    if par is not None:
        par.remove(el)


def _hex(c, default):
    c = (c or default).lstrip('#')
    from pptx.dml.color import RGBColor
    return RGBColor.from_string(c.upper())


def apply_adds(slide, sp, n, slot_bbox, slot_el, main_font, warnings):
    """plan "add": [{"type": "image"|"table", "into": <slot>, ...}] — put new content into a slot's area."""
    from pptx.util import Pt
    for add in sp.get('add', []):
        key = str(add.get('into', ''))
        bbox = slot_bbox(key)
        if not bbox:
            warnings.append(f'slide {n}: add into "{key}": slot not found')
            continue
        x, y, w, h = [int(v) for v in bbox]
        el = slot_el(key)
        if el is not None:
            remove_el(el)
        if add.get('type') == 'image':
            from PIL import Image
            with Image.open(add['file']) as im:
                iw, ih = im.size
            if add.get('mode', 'contain') == 'contain':
                sc = min(w / iw, h / ih)
                nw, nh = int(iw * sc), int(ih * sc)
                slide.shapes.add_picture(add['file'], x + (w - nw) // 2, y + (h - nh) // 2, nw, nh)
            else:
                pic = slide.shapes.add_picture(add['file'], x, y, w, h)
                replace_image(slide, pic._element, add['file'], 'cover')
        elif add.get('type') == 'table':
            rows = add['rows']
            st = add.get('style', {})
            nr, nc = len(rows), max(len(r) for r in rows)
            fs = st.get('font_size', 14)
            rh = min(int(h / nr), int(Pt(fs * 2.4)))
            gf = slide.shapes.add_table(nr, nc, x, y, w, rh * nr)
            tbl = gf.table
            tblPr = tbl._tbl.tblPr
            for attr in ('bandRow', 'firstRow'):
                tblPr.set(attr, '0')
            sid = tblPr.find(A + 'tableStyleId')
            if sid is not None:
                tblPr.remove(sid)
            for r, row in enumerate(rows):
                tbl.rows[r].height = rh
                for c in range(nc):
                    cell = tbl.cell(r, c)
                    cell.text = str(row[c]) if c < len(row) else ''
                    head = r == 0 and st.get('header', True)
                    fill = st.get('header_fill', '#333333') if head else (st.get('band_fill') if r % 2 == 0 and st.get('band_fill') else st.get('fill', '#FFFFFF'))
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = _hex(fill, '#FFFFFF')
                    for para in cell.text_frame.paragraphs:
                        for run_ in para.runs:
                            run_.font.size = Pt(fs)
                            run_.font.bold = bool(head)
                            run_.font.name = st.get('font', main_font)
                            run_.font.color.rgb = _hex(st.get('header_color', '#FFFFFF') if head else st.get('color', '#333333'), '#333333')
        else:
            warnings.append(f'slide {n}: unknown add type {add.get("type")}')


def build_from_layout(prs, sp, n, warnings):
    li = int(sp['layout'])
    if li < 1 or li > len(prs.slide_layouts):
        die(f'slide {n}: layout {li} out of range 1..{len(prs.slide_layouts)}')
    layout = prs.slide_layouts[li - 1]
    new = prs.slides.add_slide(layout)
    phs = [ph for ph in new.placeholders if ph_of(ph._element) and ph_of(ph._element)[0] not in META_PH]
    keys = ph_keys([ph._element for ph in phs])
    fill = sp.get('fill') or {}
    images = sp.get('images') or {}
    by_key = dict(zip(keys, phs))
    into = {str(a_.get('into')) for a_ in sp.get('add', [])}
    apply_adds(new, sp, n, lambda k_: ph_inherited(layout, by_key[k_]._element)[0] if k_ in by_key else None,
               lambda k_: by_key[k_]._element if k_ in by_key else None, MAIN_FONT[0], warnings)
    for key, ph in zip(keys, phs):
        if key in into:
            continue
        el = ph._element
        if key in fill:
            fill_text(el, fill[key])
            bbox, fs = ph_inherited(layout, el)
            tb = el.find(P + 'txBody')
            cap = capacity(bbox, fs or 18, tb)[0] if bbox else None
            if cap and len(body_text(tb)) > cap * 1.05:
                warnings.append(f'slide {n} {key}: {len(body_text(tb))} chars > capacity≈{cap}')
        elif key in images:
            try:
                ph.insert_picture(images[key])
            except Exception as e:
                warnings.append(f'slide {n} {key}: insert_picture failed: {e}')
        else:
            remove_el(el)
    for key in list(fill) + list(images):
        if key not in keys:
            warnings.append(f'slide {n}: layout L{li} has no placeholder "{key}" (есть: {keys})')
    if sp.get('notes'):
        new.notes_slide.notes_text_frame.text = sp['notes']


def cmd_build(a):
    plan = json.loads(Path(a.plan).read_text())
    prs = Presentation(a.ref)
    cat = {s['index']: s for s in shape_catalogue(prs)}
    fonts = palette_and_fonts(a.ref)[1]
    MAIN_FONT[0] = fonts.most_common(1)[0][0] if fonts else None
    originals = list(prs.slides)
    warnings = []
    for n, sp in enumerate(plan['slides'], 1):
        if 'layout' in sp:
            build_from_layout(prs, sp, n, warnings)
            continue
        k = int(sp['template'])
        if k < 1 or k > len(originals):
            die(f'slide {n}: template {k} out of range 1..{len(originals)}')
        new = clone_slide(prs, originals[k - 1])
        root = new._element
        fill = {str(i): v for i, v in (sp.get('fill') or {}).items()}
        tables = {str(i): v for i, v in (sp.get('tables') or {}).items()}
        images = {str(i): v for i, v in (sp.get('images') or {}).items()}
        keep = {str(i) for i in sp.get('keep', [])}
        delete = {str(i) for i in sp.get('delete', [])}
        clear = {str(i) for i in sp.get('clear', [])}
        info = {it['id']: it for it in cat[k]['shapes']}
        for sid in list(fill) + list(tables) + list(images) + list(keep) + list(delete) + list(clear):
            if sid not in info:
                warnings.append(f'slide {n} (tpl {k}): shape #{sid} not found in template')
        # 1) default cleanup: drop unfilled text shapes (non-chrome) so no template text leaks
        for it in cat[k]['shapes']:
            sid = it['id']
            el = find_shape(root, sid)
            if el is None:
                continue
            if sid in delete:
                remove_el(el)
                continue
            if sid in {str(a_.get('into')) for a_ in sp.get('add', [])}:
                continue
            if it.get('role_kind') == 'text' and not it['chrome'] and sid not in fill and sid not in keep:
                if sid in clear:
                    fill_text(el, '')
                else:
                    remove_el(el)
            if it.get('role_kind') == 'frame' and it.get('frame') == 'table' and sid not in tables and sid not in keep:
                remove_el(el)
        # 1b) decorative shapes (cards, pills) whose texts were all dropped: remove them with their icons,
        #     otherwise empty cards/pills stay on the slide
        into = {str(a_.get('into')) for a_ in sp.get('add', [])}
        for it in cat[k]['shapes']:
            if it.get('role_kind') != 'shape' or not it.get('contains') or it['id'] in keep or it.get('chrome'):
                continue
            texts = [c for c in it['contains'] if info.get(c, {}).get('role_kind') == 'text']
            if texts and all(c not in fill and c not in keep and c not in clear and c not in into
                             and not info[c].get('chrome') for c in texts):
                card_el = find_shape(root, it['id'])
                if card_el is None:
                    continue
                for c in [it['id']] + [c for c in it['contains'] if c not in keep and c not in images]:
                    el = find_shape(root, c)
                    if el is not None:
                        remove_el(el)
                warnings.append(f'slide {n}: removed empty card/shape #{it["id"]} (its text slots were not filled)')
        # 2) fill
        for sid, val in fill.items():
            el = find_shape(root, sid)
            if el is None:
                continue
            res = fill_text(el, val)
            it = info.get(sid, {})
            if isinstance(res, tuple) and res[0] == 'shrink':
                r = shrink_text(el, it.get('bbox_emu'), res[1])
                if r < 0.995:
                    warnings.append(f'slide {n} #{sid}: font scaled x{r:.2f}')
            if it.get('markers'):
                items, _ = _normalize_items(val)
                for mid in it['markers'][len(items):]:
                    mel = find_shape(root, mid)
                    if mel is not None:
                        remove_el(mel)
                for itx in items:
                    if len(itx.get('text', '')) > it['line_chars']:
                        warnings.append(f'slide {n} #{sid}: item «{itx["text"][:40]}…» longer than one line '
                                        f'(≤{it["line_chars"]} chars) — markers will misalign')
            txt = body_text(el.find(P + 'txBody'))
            cap = it.get('capacity')
            if cap and not it.get('grows') and len(txt) > cap * 1.05:
                warnings.append(f'slide {n} #{sid}: {len(txt)} chars > capacity≈{cap} — shorten text, use "shrink": true or another template')
        for sid, rows in tables.items():
            el = find_shape(root, sid)
            if el is not None:
                fill_table(el, rows)
        for sid, path in images.items():
            el = find_shape(root, sid)
            if el is None:
                continue
            if isinstance(path, dict):
                res = replace_image(new, el, path['file'], path.get('mode', 'cover'))
            else:
                res = replace_image(new, el, path)
            if res:
                warnings.append(f'slide {n} #{sid}: image not replaced ({res})')
        if sp.get('add'):
            apply_adds(new, sp, n, lambda k_: info[k_]['bbox_emu'] if k_ in info else None,
                       lambda k_: find_shape(root, k_), MAIN_FONT[0], warnings)
        # 3) re-space rows of cards: {"respace": [["4","75"]]} — stretch the remaining cards over the row span
        for row in sp.get('respace', []):
            msg = respace(root, [str(c) for c in row], info, cat[k]['shapes'])
            if msg:
                warnings.append(f'slide {n}: respace {row}: {msg}')
        # 4) drop groups that became empty
        changed = True
        while changed:
            changed = False
            for g in list(root.iter(P + 'grpSp')):
                if not any(c.tag in SHAPE_TAGS for c in g):
                    remove_el(g)
                    changed = True
        if sp.get('notes'):
            new.notes_slide.notes_text_frame.text = sp['notes']
    for s in originals:
        delete_slide(prs, s)
    prs.save(a.out)
    print(f'built {a.out}: {len(plan["slides"])} slides')
    for w in warnings:
        print('WARN ' + w)


# ----------------------------------------------------------------------------- qa

def pdf_words(pdf):
    r = run(['pdftotext', '-bbox', str(pdf), '-'], timeout=300)
    pages, cur = [], None
    for line in r.stdout.splitlines():
        m = re.search(r'<page width="([\d.]+)" height="([\d.]+)"', line)
        if m:
            cur = {'w': float(m.group(1)), 'h': float(m.group(2)), 'words': []}
            pages.append(cur)
            continue
        m = re.search(r'<word xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">(.*?)</word>', line)
        if m and cur is not None:
            cur['words'].append((float(m.group(1)), float(m.group(2)), float(m.group(3)), float(m.group(4)),
                                 norm(re.sub(r'&\w+;', ' ', m.group(5)))))
    return pages


def cmd_qa(a):
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    plan = json.loads(Path(a.plan).read_text())
    ref_dir = Path(a.ref_dir)
    prof = json.loads((ref_dir / 'profile.json').read_text())
    prs = Presentation(a.pptx)
    W, H = prs.slide_width, prs.slide_height
    issues = []
    rw, rh = prof['slide_size_emu']
    if (W, H) != (rw, rh):
        issues.append(f'GLOBAL: slide size {W}x{H} differs from reference {rw}x{rh}')
    if len(prs.slides) != len(plan['slides']):
        issues.append(f'GLOBAL: {len(prs.slides)} slides in file, {len(plan["slides"])} in plan')
    if prof.get('fonts_missing_in_renderer'):
        issues.append('INFO: fonts missing in renderer (renders use substitutes, overflow checks approximate): '
                      + ', '.join(prof['fonts_missing_in_renderer']))
    pdf, pngs = render(a.pptx, out / 'render', dpi=a.dpi)
    pages = pdf_words(pdf) if pdf else []
    tpl_shapes = {s['index']: s['shapes'] for s in prof['slides']}
    outline = {}
    if a.src_dir and (Path(a.src_dir) / 'outline.json').exists():
        for rec in json.loads((Path(a.src_dir) / 'outline.json').read_text()):
            words = []
            for b in rec['blocks']:
                for p in b.get('paragraphs', []):
                    if not p.get('flag'):
                        words += norm(p['text']).split()
                for row in b.get('table', []):
                    words += norm(' '.join(row)).split()
            outline[rec['index']] = {w[:5] for w in words if len(w) >= 5}
    elif not a.src_dir:
        issues.append('INFO: --src-dir not given — content completeness not checked')
    from PIL import Image, ImageDraw
    report = []
    for n, (slide, sp) in enumerate(zip(prs.slides, plan['slides']), 1):
        is_layout = 'layout' in sp
        k = int(sp['layout'] if is_layout else sp['template'])
        sl_issues = []
        fill_texts = ' '.join(norm(json.dumps(v, ensure_ascii=False)) for v in (sp.get('fill') or {}).values())
        cur_texts = []
        shapes = list(walk(slide._element.cSld.spTree))
        for s in shapes:
            tb = s['el'].find(P + 'txBody')
            if tb is not None:
                cur_texts.extend(para_text(p).strip() for p in tb.findall(A + 'p') if para_text(p).strip())
        # completeness: share of the source slide's words (5-letter stems) visible on the result slide (notes excluded)
        src_idx = sp.get('source')
        if outline and src_idx is not None:
            srcs = src_idx if isinstance(src_idx, list) else [src_idx]
            need = set().union(*[outline.get(int(i), set()) for i in srcs])
            have = {w[:5] for w in norm(' '.join(t.text or '' for t in slide._element.iter(A + 't'))).split() if len(w) >= 5}
            if len(need) >= 8:
                cov = len(need & have) / len(need)
                if cov < 0.45:
                    sl_issues.append(f'content: only {cov:.0%} of source slide {src_idx} words are on the slide — '
                                     f'content lost or moved to notes; put the key points on the slide')
        # leftovers: template paragraphs still present but not requested
        tpl_paras = set()
        for it in ([] if is_layout else tpl_shapes.get(k, [])):
            if it.get('text') and not it.get('chrome'):
                for ptxt in it['text'].split('\n'):
                    if len(norm(ptxt)) >= 4:
                        tpl_paras.add(ptxt.strip())
        for ptxt in cur_texts:
            if ptxt in tpl_paras and norm(ptxt) not in fill_texts:
                sl_issues.append(f'leftover template text: «{ptxt[:60]}»')
        # overflow via rendered word positions
        if n <= len(pages):
            pg = pages[n - 1]
            sx, sy = pg['w'] / W, pg['h'] / H
            idx = defaultdict(list)
            for w in pg['words']:
                for tok in w[4].split():
                    idx[tok].append(w)
            if is_layout:
                phels = [x for x in shapes if ph_of(x['el']) and ph_of(x['el'])[0] not in META_PH]
                for x, key in zip(phels, ph_keys([x['el'] for x in phels])):
                    x['id'] = key
                    if not x['bbox']:
                        x['bbox'] = ph_inherited(slide.slide_layout, x['el'])[0]
            for sid in (sp.get('fill') or {}):
                s = next((x for x in shapes if x['id'] == str(sid)), None)
                if not s or not s['bbox']:
                    continue
                tb = s['el'].find(P + 'txBody')
                # count occurrences: a repeated word is "inside" only as many times as it is rendered inside the box
                toks = Counter(t for t in norm(body_text(tb)).split() if len(t) >= 4)
                x0, y0 = s['bbox'][0] * sx, s['bbox'][1] * sy
                x1, y1 = x0 + s['bbox'][2] * sx, y0 + s['bbox'][3] * sy
                tol = 0.015 * pg['h']
                outside, total, below = 0, 0, 0.0
                for t, need in toks.items():
                    cands = idx.get(t)
                    if not cands:
                        continue
                    total += need
                    ins = [w for w in cands if x0 - tol <= (w[0] + w[2]) / 2 <= x1 + tol
                           and y0 - tol <= (w[1] + w[3]) / 2 <= y1 + tol]
                    if len(ins) < need:
                        outside += need - len(ins)
                        col = [((w[1] + w[3]) / 2 - y1) / pg['h'] for w in cands
                               if w not in ins and x0 - tol <= (w[0] + w[2]) / 2 <= x1 + tol]
                        if col:
                            below = max(below, max(col))
                if total and outside and (outside / total > 0.1 or below > 0.03):
                    sl_issues.append(f'#{sid}: text runs outside its box ({outside}/{total} words, '
                                     f'{max(below, 0):.0%} of slide height below) — shorten, shrink or pick a larger slot')
        # side-by-side pair
        ref_png = ref_dir / 'render' / (f'layout-{k:02d}.png' if is_layout else f'slide-{k:02d}.png')
        if not is_layout and not any(t.get('text') and not t.get('chrome') for t in tpl_shapes.get(k, [])):
            # empty template slide (only empty placeholders): compare with the sample render of its layout
            lay = next((lc for lc in prof.get('layouts_catalogue', []) if k in lc['used_by_slides']), None)
            if lay and (ref_dir / 'render' / f'layout-{lay["index"]:02d}.png').exists():
                ref_png = ref_dir / 'render' / f'layout-{lay["index"]:02d}.png'
        if n <= len(pngs) and ref_png.exists():
            l, r = Image.open(ref_png).convert('RGB'), Image.open(pngs[n - 1]).convert('RGB')
            tw = 760
            l = l.resize((tw, int(l.height * tw / l.width)))
            r = r.resize((tw, int(r.height * tw / r.width)))
            pair = Image.new('RGB', (tw * 2 + 30, max(l.height, r.height) + 28), (235, 235, 235))
            d = ImageDraw.Draw(pair)
            d.text((10, 6), f'LAYOUT L{k}' if is_layout else f'TEMPLATE {k}', fill='black')
            d.text((tw + 20, 6), f'RESULT {n}', fill='black')
            pair.paste(l, (10, 22))
            pair.paste(r, (tw + 20, 22))
            pair.save(out / f'pair-{n:02d}.jpg', quality=85)
        report.append({'slide': n, 'template': f'L{k}' if is_layout else k, 'issues': sl_issues})
    if pngs:
        contact_sheet(pngs, out / 'contact.jpg')
    md = [f'# QA: {Path(a.pptx).name}', ''] + [f'- {i}' for i in issues] + ['']
    total = 0
    for r in report:
        mark = 'OK' if not r['issues'] else f'{len(r["issues"])} проблем'
        md.append(f'## Слайд {r["slide"]} (шаблон {r["template"]}): {mark} — pair-{r["slide"]:02d}.jpg')
        md += [f'- {i}' for i in r['issues']]
        total += len(r['issues'])
    md += ['', 'Автопроверки не видят: текст, «запечённый» в картинках шаблона; смысловое соответствие; '
           'визуальный баланс. Обязательно просмотри pair-NN.jpg через read_file.']
    (out / 'report.md').write_text('\n'.join(md))
    (out / 'report.json').write_text(json.dumps({'global': issues, 'slides': report}, ensure_ascii=False, indent=1))
    print('\n'.join(md))
    print(f'\nQA: {total} slide issues; pairs in {out}')


# ----------------------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest='cmd', required=True)
    sub.add_parser('doctor')
    p = sub.add_parser('profile'); p.add_argument('ref'); p.add_argument('--out', required=True); p.add_argument('--dpi', type=int, default=80)
    p = sub.add_parser('outline'); p.add_argument('src'); p.add_argument('--out', required=True); p.add_argument('--no-render', action='store_true')
    p = sub.add_parser('build'); p.add_argument('ref'); p.add_argument('plan'); p.add_argument('--out', required=True)
    p = sub.add_parser('qa'); p.add_argument('pptx'); p.add_argument('--plan', required=True); p.add_argument('--ref-dir', required=True)
    p.add_argument('--src-dir', help='outline dir of the source deck (enables the content-completeness check)')
    p.add_argument('--out', required=True); p.add_argument('--dpi', type=int, default=80)
    a = ap.parse_args()
    {'doctor': cmd_doctor, 'profile': cmd_profile, 'outline': cmd_outline, 'build': cmd_build, 'qa': cmd_qa}[a.cmd](a)


if __name__ == '__main__':
    main()
