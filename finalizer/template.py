"""Make a manuscript follow the INFERENSI Word template.

Values below were read from the official template (Inferensi_Journal_Template_English):
body Palatino Linotype 9 pt justified, sub-headings Helvetica 9 pt bold italic, captions Helvetica 8 pt with a blue
bold label, tables Palatino Linotype 8 pt, references Palatino Linotype 8 pt, A4 page with margins
top/right/bottom 1.70 cm and left 2.49 cm, header/footer 0.76 cm, blue #00629B header and footer bars.
"""
from __future__ import annotations

import copy
import re

from lxml import etree

from .docxtools import (NS, WP, WPPR, WR, WRPR, DocxPackage, body_regions, get_t, local, paras_of, ppr_set,
                        ptext, q)

BLUE = '00629B'
BODY_FONT, HEAD_FONT, AFFIL_FONT = 'Palatino Linotype', 'Helvetica', 'Times New Roman'
SZ_BODY, SZ_SMALL, SZ_CAPTION, SZ_TITLE, SZ_AUTHOR, SZ_AFFIL, SZ_EMAIL = '18', '16', '16', '44', '20', '16', '15'
PAGE = {'w': '11909', 'h': '16834'}
MARGIN = {'top': '964', 'right': '964', 'bottom': '964', 'left': '1412', 'header': '431', 'footer': '431', 'gutter': '0'}
KEEP_FONTS = {'symbol', 'wingdings', 'wingdings 2', 'wingdings 3', 'webdings', 'cambria math', 'mt extra',
              'ms gothic', 'segoe ui symbol', 'courier new', 'consolas'}
MAIN_HEAD_RX = re.compile(r'^(?:[IVX]+\.?\s+)?(INTRODUCTION|LITERATURE REVIEW|METHODOLOGY|METHODS?|RESEARCH METHODS?|'
                          r'RESULTS?(?: AND DISCUSSIONS?)?|DISCUSSIONS?|CONCLUSIONS?(?: AND SUGGESTIONS?)?|'
                          r'ACKNOWLEDGE?MENTS?|ACKNOWLEDGEMET|REFERENCES|APPENDIX.*)$')
CAPTION_RX = re.compile(r'^(Figure|Table|Gambar|Tabel)\s+\d+')
CAPTION_VERB = re.compile(r'\b(shows?|presents?|indicates?|illustrates?|displays?|summari[sz]es|lists|reports|reveals|'
                          r'demonstrates|depicts|contains|is|are|was|were)\b')


# --------------------------------------------------------------------------- run helpers
def _rpr(r):
    rp = r.find(WRPR)
    if rp is None:
        rp = etree.Element(WRPR)
        r.insert(0, rp)
    return rp


ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow',
         'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern',
         'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs',
         'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath']


def _set(rp, name, attrs):
    el = rp.find(q('w:' + name))
    if el is None:
        el = etree.Element(q('w:' + name))
        pos = ORDER.index(name)
        for c in rp:
            cn = local(c)
            if cn in ORDER and ORDER.index(cn) > pos:
                c.addprevious(el)
                break
        else:
            rp.append(el)
    for k in list(el.attrib):
        del el.attrib[k]
    for k, v in attrs.items():
        el.set(q('w:' + k), v)
    return el


def _font(r):
    f = r.find('w:rPr/w:rFonts', NS)
    return (f.get(q('w:ascii')) or f.get(q('w:hAnsi')) or '') if f is not None else ''


def set_font(r, name):
    if r.find('w:sym', NS) is not None or _font(r).lower() in KEEP_FONTS:
        return False
    rp = _rpr(r)
    old = rp.find('w:rFonts', NS)
    east = old.get(q('w:eastAsia')) if old is not None else None
    if old is not None and old.get(q('w:ascii')) == name and old.get(q('w:hAnsi')) == name and \
            not any('Theme' in k or 'theme' in k for k in old.attrib):
        return False
    attrs = {'ascii': name, 'hAnsi': name, 'cs': name}
    if east:
        attrs['eastAsia'] = east
    _set(rp, 'rFonts', attrs)
    return True


def set_size(r, half_points):
    rp = _rpr(r)
    s = rp.find('w:sz', NS)
    if s is not None and s.get(q('w:val')) == half_points:
        return False
    _set(rp, 'sz', {'val': half_points})
    _set(rp, 'szCs', {'val': half_points})
    return True


def set_color(r, val):
    rp = _rpr(r)
    c = rp.find('w:color', NS)
    if c is not None and c.get(q('w:val'), '').upper() == val and len(c.attrib) == 1:
        return False
    _set(rp, 'color', {'val': val})
    return True


def text_runs(p):
    """Runs that belong to this paragraph (inside hyperlinks/smart tags too, never inside math or text boxes)."""
    out = []
    for r in p.iter(WR):
        anc = list(r.iterancestors())
        if any(local(a) in ('oMath', 'txbxContent') for a in anc):
            continue
        if next((a for a in anc if a.tag == WP), None) is not p:
            continue
        if r.find('w:t', NS) is None and r.find('w:sym', NS) is None:
            continue
        out.append(r)
    return out


def in_hyperlink(r):
    if any(local(a) == 'hyperlink' for a in r.iterancestors()):
        return True
    st = r.find('w:rPr/w:rStyle', NS)
    return st is not None and 'hyperlink' in st.get(q('w:val'), '').lower()


# --------------------------------------------------------------------------- paragraph classes
def own_text(p):
    return ''.join(get_t(r) for r in text_runs(p)).strip()


def in_table(p):
    return any(local(a) == 'tbl' for a in p.iterancestors())


def classify(p):
    t = own_text(p)
    ppr = p.find(WPPR)
    style = (ppr.find('w:pStyle', NS).get(q('w:val')) if ppr is not None and ppr.find('w:pStyle', NS) is not None else '').lower()
    if in_table(p):
        return 'table'
    if not t:
        if p.find('.//m:oMath', NS) is not None:
            return 'equation'
        return 'empty'
    if MAIN_HEAD_RX.match(t.upper()) and len(t) < 60 or style.startswith(('heading1', 'h1')) and len(t) < 80:
        return 'head'
    if CAPTION_RX.match(t) and len(t) < 300 and not CAPTION_VERB.search(t[:60]):
        return 'caption'
    if p.find('.//m:oMath', NS) is not None and re.fullmatch(r'\(?\d*\)?', re.sub(r'\s', '', t)):
        return 'equation'
    if p.find('.//w:drawing', NS) is not None or p.find('.//w:pict', NS) is not None:
        if len(t) < 40:
            return 'figure'
    numbered = ppr is not None and ppr.find('w:numPr', NS) is not None
    if (numbered or style.startswith(('heading2', 'heading3', 'h2', 'h3'))) and len(t.split()) <= 14 and \
            not t.endswith(('.', ',', ';', ':')):
        runs = text_runs(p)
        if style.startswith(('heading', 'h2', 'h3')) or (runs and all(
                (r.find('w:rPr/w:b', NS) is not None and r.find('w:rPr/w:b', NS).get(q('w:val')) not in ('0', 'false'))
                or not get_t(r).strip() for r in runs)):
            return 'subhead'
    jc = ppr.find('w:jc', NS).get(q('w:val')) if ppr is not None and ppr.find('w:jc', NS) is not None else ''
    if jc in ('center', 'right', 'end') and len(t) < 120:
        return 'centered'
    return 'body'


# --------------------------------------------------------------------------- fixers
def remove_highlight(root):
    n = 0
    for rp in root.iter(WRPR):
        for el in rp.findall('w:highlight', NS):
            rp.remove(el)
            n += 1
        for el in rp.findall('w:shd', NS):
            fill = (el.get(q('w:fill')) or '').upper()
            if fill not in ('', 'AUTO', 'FFFFFF') or el.get(q('w:val')) not in (None, 'clear', 'nil'):
                rp.remove(el)
                n += 1
    body = root.find('w:body', NS)
    for p in body.iter(WP):
        if in_table(p):
            continue
        sh = p.find('w:pPr/w:shd', NS)
        if sh is not None and (sh.get(q('w:fill')) or '').upper() in ('FFFF00', 'FFFF99', 'FFFFCC', '00FF00', 'FFC000', 'FFD966', 'FFE599'):
            sh.getparent().remove(sh)
            n += 1
    return ['%d sorotan (highlight/shading teks) dihapus' % n] if n else []


def unbox_tables(root):
    """Text boxes, frames and paragraph borders inside table cells -> plain cell text."""
    n_box = n_frame = 0
    body = root.find('w:body', NS)
    for tc in list(body.iter(q('w:tc'))):
        for p in list(tc.iter(WP)):
            if p.getparent() is None or any(local(a) == 'txbxContent' for a in p.iterancestors()):
                continue
            for r in list(p.findall('w:r', NS)):
                holder = r.find('mc:AlternateContent', NS)
                src = holder.find('mc:Choice', NS) if holder is not None else r
                if src is None or src.find('.//w:txbxContent', NS) is None or src.find('.//pic:pic', NS) is not None:
                    continue
                inner = []
                for tx in src.iter(q('w:txbxContent')):
                    inner += [copy.deepcopy(x) for x in tx if x.tag in (WP, q('w:tbl'))]
                    break
                r.getparent().remove(r)
                others = own_text(p)
                anchor = p
                host_ppr = p.find(WPPR)
                for x in inner:
                    if x.tag == WP:
                        old = x.find(WPPR)
                        if old is not None:
                            x.remove(old)
                        if host_ppr is not None:
                            x.insert(0, copy.deepcopy(host_ppr))
                    for fp in list(x.iter(q('w:framePr'))) + list(x.iter(q('w:pBdr'))):
                        fp.getparent().remove(fp)
                    anchor.addnext(x)
                    anchor = x
                if not others and p.find('.//w:drawing', NS) is None and p.find('.//m:oMath', NS) is None:
                    if p.getparent() is not None and len([c for c in tc if c.tag == WP]) > 1:
                        p.getparent().remove(p)
                n_box += 1
        for p in tc.iter(WP):
            ppr = p.find(WPPR)
            if ppr is None:
                continue
            for name in ('framePr', 'pBdr', 'shd'):
                el = ppr.find('w:' + name, NS)
                if el is not None:
                    ppr.remove(el)
                    n_frame += 1
    n_special = _clean_table_specials(body)
    _, main, refs = body_regions(root)
    for p in paras_of(main + refs):
        fp = p.find('w:pPr/w:framePr', NS)
        if fp is not None and not in_table(p):
            fp.getparent().remove(fp)
            n_frame += 1
    log = []
    if n_special:
        log.append('%d elemen khusus di tabel dibersihkan (content control, tabel bersarang, teks tidak membungkus, '
                   'tinggi baris terkunci, fit-text, garis/shading karakter dan sel)' % n_special)
    if n_box:
        log.append('%d kotak teks (text box) di dalam sel tabel dibuka menjadi teks sel biasa' % n_box)
    if n_frame:
        log.append('%d bingkai/garis kotak paragraf (frame/border) di tabel dan isi naskah dihapus' % n_frame)
    return log


def _unwrap(el):
    parent = el.getparent()
    content = el.find('w:sdtContent', NS)
    idx = parent.index(el)
    for c in (list(content) if content is not None else []):
        parent.insert(idx, c)
        idx += 1
    parent.remove(el)


def _clean_table_specials(body):
    """Content controls, nested one-cell tables, no-wrap, exact row heights, fit-text, character borders/shading."""
    n = 0
    for tbl in list(body.iter(q('w:tbl'))):
        # content controls (keep Mendeley/Zotero citations and bibliography)
        for sdt in list(tbl.iter(q('w:sdt'))):
            tag = sdt.find('w:sdtPr/w:tag', NS)
            val = tag.get(q('w:val'), '') if tag is not None else ''
            if val.startswith(('MENDELEY', 'ZOTERO')) or sdt.find('w:sdtPr/w:docPartObj', NS) is not None:
                continue
            if sdt.getparent() is not None:
                _unwrap(sdt)
                n += 1
        # a table nested in a cell whose only purpose is a box: keep its text, drop the table
        for inner in list(tbl.iter(q('w:tbl'))):
            if inner is tbl or inner.getparent() is None or local(inner.getparent()) != 'tc':
                continue
            cells = inner.findall('.//w:tc', NS)
            if len(cells) == 1:
                parent = inner.getparent()
                idx = parent.index(inner)
                for c in [x for x in cells[0] if x.tag in (WP, q('w:tbl'))]:
                    parent.insert(idx, c)
                    idx += 1
                parent.remove(inner)
                n += 1
        for tc in tbl.iter(q('w:tc')):
            pr = tc.find('w:tcPr', NS)
            if pr is not None:
                for name in ('noWrap', 'tcFitText'):
                    el = pr.find('w:' + name, NS)
                    if el is not None:
                        pr.remove(el)
                        n += 1
                sh = pr.find('w:shd', NS)
                if sh is not None and (sh.get(q('w:fill')) or 'auto').upper() not in ('AUTO', 'FFFFFF'):
                    pr.remove(sh)
                    n += 1
        for tr in tbl.iter(q('w:tr')):
            h = tr.find('w:trPr/w:trHeight', NS)
            if h is not None and h.get(q('w:hRule')) == 'exact':
                h.set(q('w:hRule'), 'atLeast')
                n += 1
        for rp in tbl.iter(WRPR):
            for name in ('bdr', 'fitText', 'shd', 'highlight', 'effect', 'emboss', 'imprint', 'outline', 'shadow'):
                for el in rp.findall('w:' + name, NS):
                    rp.remove(el)
                    n += 1
    return n


def unwrap_body_content_controls(root):
    """Non-citation content controls in the article body show as boxes in Word; keep their text only."""
    _, main, refs = body_regions(root)
    n = 0
    for e in main + refs:
        for sdt in ([e] if e.tag == q('w:sdt') else []) + list(e.iter(q('w:sdt'))):
            if sdt.getparent() is None:
                continue
            tag = sdt.find('w:sdtPr/w:tag', NS)
            val = tag.get(q('w:val'), '') if tag is not None else ''
            if val.startswith(('MENDELEY', 'ZOTERO')) or sdt.find('w:sdtPr/w:docPartObj', NS) is not None:
                continue
            _unwrap(sdt)
            n += 1
    return ['%d content control (kotak isian) di isi naskah dilepas, teksnya dipertahankan' % n] if n else []


def format_by_template(root):
    """Fonts, sizes, colours and alignment following the template, by paragraph role."""
    front, main, refs = body_regions(root)
    cnt = {'font': 0, 'size': 0, 'color': 0, 'jc': 0}

    def bump(k, changed):
        if changed:
            cnt[k] += 1

    # main text ---------------------------------------------------------------
    for p in paras_of(main, include_fallback=False):
        if any(local(a) == 'txbxContent' for a in p.iterancestors()) and not in_table(p):
            continue
        role = classify(p)
        runs = text_runs(p)
        if role == 'body':
            for r in runs:
                bump('font', set_font(r, BODY_FONT))
                bump('size', set_size(r, SZ_BODY))
                if not in_hyperlink(r):
                    c = r.find('w:rPr/w:color', NS)
                    if c is not None and (c.get(q('w:val'), '').upper() not in ('000000', 'AUTO') or len(c.attrib) > 1):
                        bump('color', set_color(r, '000000'))
            ppr = p.find(WPPR)
            jc = ppr.find('w:jc', NS) if ppr is not None else None
            if jc is None or jc.get(q('w:val')) in ('left', 'start', 'distribute', 'right', 'end', 'center'):
                if jc is None or jc.get(q('w:val')) != 'both':
                    ppr_set(p, 'jc', {'val': 'both'})
                    cnt['jc'] += 1
        elif role == 'subhead':
            for r in runs:
                bump('font', set_font(r, HEAD_FONT))
                bump('size', set_size(r, SZ_BODY))
                bump('color', set_color(r, '000000'))
        elif role == 'head':
            for r in runs:
                f = r.find('w:rPr/w:rFonts', NS)
                if f is not None and _font(r).lower() not in KEEP_FONTS and _font(r) not in ('', HEAD_FONT, 'Arial'):
                    bump('font', set_font(r, HEAD_FONT))
                c = r.find('w:rPr/w:color', NS)
                if c is not None and c.get(q('w:val'), '').upper() != BLUE:
                    bump('color', set_color(r, BLUE))
        elif role == 'caption':
            label_done = False
            for r in runs:
                bump('font', set_font(r, HEAD_FONT))
                bump('size', set_size(r, SZ_CAPTION))
                t = get_t(r)
                if not label_done and re.match(r'^\s*(Figure|Table|Gambar|Tabel)\b|^\s*\d+\s*$', t):
                    bump('color', set_color(r, BLUE))
                    _set(_rpr(r), 'b', {})
                    if re.search(r'\d', t):
                        label_done = True
                else:
                    label_done = True
                    c = r.find('w:rPr/w:color', NS)
                    if c is not None and c.get(q('w:val'), '').upper() not in ('000000', 'AUTO'):
                        bump('color', set_color(r, '000000'))
        elif role == 'table':
            for r in runs:
                bump('font', set_font(r, BODY_FONT))
                bump('size', set_size(r, SZ_SMALL))
                c = r.find('w:rPr/w:color', NS)
                if c is not None and c.get(q('w:val'), '').upper() not in ('000000', 'AUTO') and not in_hyperlink(r):
                    bump('color', set_color(r, '000000'))
    # references --------------------------------------------------------------
    for p in paras_of(refs[1:] if refs else [], include_fallback=False):
        t = own_text(p)
        if 'Creative Commons' in t:
            break
        if not t:
            continue
        for r in text_runs(p):
            bump('font', set_font(r, BODY_FONT))
            bump('size', set_size(r, SZ_SMALL))
            c = r.find('w:rPr/w:color', NS)
            if c is not None and c.get(q('w:val'), '').upper() not in ('000000', 'AUTO') and not in_hyperlink(r):
                bump('color', set_color(r, '000000'))
        jc = p.find('w:pPr/w:jc', NS)
        if re.match(r'^\[\d+\]', t) and (jc is None or jc.get(q('w:val')) != 'both'):
            ppr_set(p, 'jc', {'val': 'both'})
            cnt['jc'] += 1
    # front matter ------------------------------------------------------------
    fps = [p for p in paras_of(front, include_fallback=False)]
    title_done = False
    for p in fps:
        t = own_text(p)
        if not t or any(local(a) == 'txbxContent' for a in p.iterancestors()):
            continue
        runs = text_runs(p)
        if not title_done:
            title_done = True
            for r in runs:
                if _font(r) and _font(r) not in (HEAD_FONT, 'Arial'):
                    bump('font', set_font(r, HEAD_FONT))
                bump('color', set_color(r, BLUE))
            continue
        low = t.lower()
        if low.startswith('abstract'):
            for r in runs:
                bump('font', set_font(r, BODY_FONT))
                if not re.match(r'^\s*abstract', get_t(r), re.I):
                    c = r.find('w:rPr/w:color', NS)
                    if c is not None and c.get(q('w:val'), '').upper() not in ('000000', 'AUTO'):
                        bump('color', set_color(r, '000000'))
            jc = p.find('w:pPr/w:jc', NS)
            if jc is None or jc.get(q('w:val')) != 'both':
                ppr_set(p, 'jc', {'val': 'both'})
                cnt['jc'] += 1
        elif low.startswith('keywords'):
            for r in runs:
                if not re.match(r'^\s*keywords', get_t(r), re.I):
                    bump('font', set_font(r, BODY_FONT))
            jc = p.find('w:pPr/w:jc', NS)
            if jc is None or jc.get(q('w:val')) != 'both':
                ppr_set(p, 'jc', {'val': 'both'})
                cnt['jc'] += 1
        elif low.startswith('*corresponding') or low.startswith('corresponding'):
            for r in runs:
                bump('font', set_font(r, AFFIL_FONT))
                bump('size', set_size(r, SZ_EMAIL))
        elif re.match(r'^\d?\s*(department|departemen|faculty|fakultas|school|program|study program|universit|institut|'
                      r'politeknik|badan|research|center|centre|laborator)', low):
            for r in runs:
                bump('font', set_font(r, AFFIL_FONT))
                bump('size', set_size(r, SZ_AFFIL))
    log = []
    if cnt['font']:
        log.append('Font disesuaikan template pada %d run (isi: Palatino Linotype; subjudul & caption: Helvetica; '
                   'afiliasi: Times New Roman)' % cnt['font'])
    if cnt['size']:
        log.append('Ukuran huruf disesuaikan template pada %d run (isi 9 pt; tabel, caption & referensi 8 pt)' % cnt['size'])
    if cnt['color']:
        log.append('Warna teks disesuaikan template pada %d run (teks hitam; judul, judul bab & label caption biru #00629B)' % cnt['color'])
    if cnt['jc']:
        log.append('%d paragraf isi, abstrak, dan referensi dibuat rata kanan-kiri (justify)' % cnt['jc'])
    return log


def fix_sections(root):
    """Page size, margins and header/footer distance of every section as in the template."""
    n = 0
    secs = list(root.iter(q('w:sectPr')))
    final = root.find('w:body/w:sectPr', NS)
    refs = [copy.deepcopy(x) for x in final if local(x) in ('headerReference', 'footerReference')] if final is not None else []
    for sp in secs:
        changed = False
        pg = sp.find('w:pgSz', NS)
        if pg is None:
            pg = etree.SubElement(sp, q('w:pgSz'))
        land = pg.get(q('w:orient')) == 'landscape' or int(pg.get(q('w:w'), '0') or 0) > int(pg.get(q('w:h'), '1') or 1)
        want = {'w': PAGE['h'], 'h': PAGE['w']} if land else dict(PAGE)
        for k, v in want.items():
            if pg.get(q('w:' + k)) != v:
                pg.set(q('w:' + k), v)
                changed = True
        mar = sp.find('w:pgMar', NS)
        if mar is None:
            mar = etree.Element(q('w:pgMar'))
            pg.addnext(mar)
        for k, v in MARGIN.items():
            if mar.get(q('w:' + k)) != v:
                mar.set(q('w:' + k), v)
                changed = True
        cols = sp.find('w:cols', NS)
        if cols is not None and cols.get(q('w:num')) not in (None, '1'):
            for k in list(cols.attrib):
                del cols.attrib[k]
            cols.set(q('w:space'), '720')
            for c in list(cols):
                cols.remove(c)
            changed = True
        if sp is not final and refs:
            have = [x for x in sp if local(x) in ('headerReference', 'footerReference')]
            if have and [etree.tostring(x) for x in have] != [etree.tostring(x) for x in refs]:
                for x in have:
                    sp.remove(x)
                for x in reversed(refs):
                    sp.insert(0, copy.deepcopy(x))
                changed = True
        n += changed
    return ['Pengaturan halaman %d section disamakan dengan template (A4; margin atas/kanan/bawah 1,70 cm, kiri 2,49 cm; '
            'header/footer 0,76 cm; satu kolom)' % n] if n else []


def fix_header_footer_style(pkg: DocxPackage):
    """Blue bars and white/blue text exactly as in the template, whatever theme the document carries."""
    n = 0
    for name in pkg.parts(r'word/(header|footer)\d*\.xml'):
        root = pkg.xml(name)
        is_header = 'header' in name
        for sp in root.iter('{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}spPr'):
            fills = [f for f in sp.findall('a:solidFill', NS)]
            ln = sp.find('a:ln', NS)
            if ln is not None:
                fills += ln.findall('a:solidFill', NS)
            for fill in fills:
                kids = list(fill)
                if kids and local(kids[0]) == 'srgbClr' and kids[0].get('val', '').upper() in (BLUE, 'FFFFFF') \
                        and len(kids[0]) == 0:
                    continue
                for k in kids:
                    fill.remove(k)
                etree.SubElement(fill, q('a:srgbClr')).set('val', BLUE)
                n += 1
        for v in root.iter():
            if not isinstance(v.tag, str) or not v.tag.startswith('{urn:schemas-microsoft-com:vml}'):
                continue
            for att in ('fillcolor', 'strokecolor'):
                val = v.get(att)
                if val and val.lower() not in ('#00629b', 'white', '#ffffff', 'none') and not val.lower().startswith('window'):
                    v.set(att, '#00629b')
                    n += 1
            if local(v) == 'fill' and v.get('opacity'):
                del v.attrib['opacity']
                n += 1
        for p in root.iter(WP):
            t = ptext(p)
            in_box = any(local(a) == 'txbxContent' for a in p.iterancestors())
            for r in p.findall('w:r', NS):
                if r.find('w:t', NS) is None or not get_t(r).strip():
                    continue
                if is_header and in_box and re.search(r'INFEREN|Vol|ISSN|Print|Online|\d', t):
                    n += set_font(r, HEAD_FONT) + set_color(r, 'FFFFFF')
                    _set(_rpr(r), 'b', {})
                elif not is_header and in_box and t.strip().startswith('DOI'):
                    n += set_font(r, HEAD_FONT) + set_color(r, 'FFFFFF')
                elif not is_header and 'Department of Statistics' in t:
                    n += set_font(r, HEAD_FONT) + set_color(r, BLUE)
                elif is_header and not in_box and re.fullmatch(r'\s*\d+\s*', get_t(r)):
                    n += set_font(r, HEAD_FONT) + set_color(r, BLUE)
    return ['Warna dan font header/footer diseragamkan dengan template (bar biru #00629B, teks putih/biru Helvetica) — %d atribut' % n] if n else []


def copy_header_footer_from(pkg: DocxPackage, template_bytes: bytes):
    """Replace the default header and footer of the manuscript with those of an uploaded template file."""
    tpl = DocxPackage(template_bytes)
    log = []

    def default_targets(p):
        body = p.xml(p.main).find('w:body', NS)
        sp = body.find('w:sectPr', NS)
        rels = p.text(p.main_rels)
        out = {}
        for kind in ('header', 'footer'):
            ref = next((x for x in sp.findall('w:%sReference' % kind, NS) if x.get(q('w:type')) == 'default'), None) \
                if sp is not None else None
            if ref is None:
                continue
            rid = ref.get(q('r:id'))
            m = re.search(r'<Relationship [^>]*Id="%s"[^>]*Target="([^"]+)"' % rid, rels) or \
                re.search(r'<Relationship [^>]*Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
            if m:
                tgt = m.group(1).lstrip('/')
                out[kind] = tgt if tgt.startswith('word/') else 'word/' + tgt
        return out

    mine, theirs = default_targets(pkg), default_targets(tpl)
    for kind in ('header', 'footer'):
        if kind in mine and kind in theirs and theirs[kind] in tpl.files:
            pkg.trees.pop(mine[kind], None)
            pkg.files[mine[kind]] = tpl.files[theirs[kind]]
            src_rels = theirs[kind].replace('word/', 'word/_rels/', 1) + '.rels'
            dst_rels = mine[kind].replace('word/', 'word/_rels/', 1) + '.rels'
            if src_rels in tpl.files:
                pkg.files[dst_rels] = tpl.files[src_rels]
                if dst_rels not in pkg.order:
                    pkg.order.append(dst_rels)
                for tg in re.findall(rb'Target="([^"]+)"', tpl.files[src_rels]):
                    path = 'word/' + tg.decode().lstrip('/').replace('../', '')
                    if path in tpl.files and path not in pkg.files:
                        pkg.files[path] = tpl.files[path]
                        pkg.order.append(path)
            elif dst_rels in pkg.files:
                pkg.remove(dst_rels)
            log.append('%s disalin dari file template yang diunggah' % ('Header' if kind == 'header' else 'Footer'))
    return log


def check_sections_order(root):
    """Return a message if the main template sections are missing or out of order."""
    want = [('INTRODUCTION', r'INTRODUCTION'), ('LITERATURE REVIEW', r'LITERATURE REVIEW|THEORY|THEORETICAL'),
            ('METHODOLOGY', r'METHODOLOGY|METHODS?\b|RESEARCH METHOD'),
            ('RESULTS AND DISCUSSIONS', r'RESULTS?|DISCUSSION'), ('CONCLUSIONS AND SUGGESTIONS', r'CONCLUSION'),
            ('REFERENCES', r'REFERENCES')]
    found = []
    for e in root.find('w:body', NS):
        if e.tag != WP:
            continue
        t = own_text(e).upper()
        if len(t) < 60 and MAIN_HEAD_RX.match(t):
            found.append(t)
    msgs, pos = [], -1
    for name, rx in want:
        idx = next((i for i, t in enumerate(found) if re.search(rx, t)), None)
        if idx is None:
            msgs.append('Bagian “%s” tidak ditemukan (template: I. INTRODUCTION, II. LITERATURE REVIEW, III. METHODOLOGY, '
                        'IV. RESULTS AND DISCUSSIONS, V. CONCLUSIONS AND SUGGESTIONS, REFERENCES)' % name)
        elif idx < pos:
            msgs.append('Urutan bagian “%s” tidak sesuai template' % name)
        else:
            pos = idx
    return msgs, found
