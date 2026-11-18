"""Automatic, low-risk fixes. Every fixer returns a list of human-readable change messages."""
from __future__ import annotations

import copy
import re

from lxml import etree

from .docxtools import (NS, WP, WPPR, WR, WRPR, WT, body_regions, get_t, inner_paras, is_text_run, local,
                        new_run_like, normalize, paras_of, ppr_set, ptext, q, regex_sub_para, rep, replace_span,
                        segments, set_t)

ISSN = 'ISSN: 0216-308X (Print) 2721-3862 (Online)'
INSTITUTION = 'Department of Statistics, Institut Teknologi Sepuluh Nopember'


def header_text(volume, issue, month, year) -> str:
    return 'INFERENSI, Vol. %s(%s), %s. %s. %s' % (volume, issue, month, year, ISSN)


# --------------------------------------------------------------------------- helpers
def _rewrite_paragraph_text(p, new_text: str, start_rx: str):
    """Put new_text in the first run matching start_rx and blank every later plain-text run
    of the paragraph that is not part of a field result (page numbers stay intact)."""
    in_field = 0
    target = None
    for c in list(p):
        if c.tag != WR:
            continue
        fc = c.find('w:fldChar', NS)
        if fc is not None:
            t = fc.get(q('w:fldCharType'))
            in_field += 1 if t == 'begin' else (-1 if t == 'end' else 0)
            continue
        if c.find(WT) is None:
            continue
        txt = get_t(c)
        if target is None:
            if re.search(start_rx, txt):
                lead = re.match(r'^\s*', txt).group(0)
                set_t(c, lead + new_text)
                target = c
        elif in_field == 0:
            p.remove(c)
    return target is not None


# --------------------------------------------------------------------------- header / footer
def fix_header_footer(pkg, volume, issue, month, year, doi: str | None):
    log = []
    htxt = header_text(volume, issue, month, year)
    for name in pkg.parts(r'word/header\d*\.xml'):
        root = pkg.xml(name)
        normalize(root)
        for p in inner_paras(root):
            t = ptext(p)
            if re.search(r'INFEREN|IPTEK|Journal for Technology and Science|science and technology', t, re.I) and \
                    re.search(r'Vol|ISSN|Journal|science', t, re.I):
                if t.strip() != htxt and _rewrite_paragraph_text(
                        p, htxt, r'INFEREN|IPTEK|Journal for Technology|science and technology'):
                    log.append('Header (%s): “%s” → “%s”' % (name.split('/')[-1], t.strip()[:70], htxt))
    doi_txt = 'DOI: ' + (doi if doi else '10.12962/j27213862.v%si%s.xxxx' % (volume, issue))
    for name in pkg.parts(r'word/footer\d*\.xml'):
        root = pkg.xml(name)
        normalize(root)
        for p in inner_paras(root):
            t = ptext(p).strip()
            if t.startswith('DOI') and t != doi_txt:
                if _rewrite_paragraph_text(p, doi_txt, r'DOI'):
                    log.append('Footer: “%s” → “%s”' % (t, doi_txt))
            elif t.startswith('Department of Statistics') and t != INSTITUTION:
                if _rewrite_paragraph_text(p, INSTITUTION, r'Department'):
                    log.append('Footer: “%s” → “%s”' % (t, INSTITUTION))
    return log


# --------------------------------------------------------------------------- dates
def fix_dates(root, dates: dict):
    """dates: {'Received': '15 June 2026', ...}; values may be empty strings."""
    log, found = [], set()
    for p in inner_paras(root):
        t = ptext(p).strip()
        for lab, val in dates.items():
            if not t.startswith(lab + ':'):
                continue
            found.add(lab)
            if not val:
                continue
            runs = [r for r in p.findall('w:r', NS) if r.find(WT) is not None]
            lab_run = next((r for r in runs if (lab + ':') in get_t(r)), None)
            if lab_run is None:
                continue
            old = t[len(lab) + 1:].strip()
            lt = get_t(lab_run)
            i = lt.index(lab + ':') + len(lab) + 1
            after = runs[runs.index(lab_run) + 1:]
            if lt[i:].strip() or not after:
                set_t(lab_run, lt[:i] + ' ' + val)
                for r in after:
                    p.remove(r)
            else:
                if not lt.endswith(' '):
                    set_t(lab_run, lt + ' ')
                set_t(after[0], val)
                for r in after[1:]:
                    p.remove(r)
            if old != val:
                log.append('%s: “%s” → “%s”' % (lab, old, val))
    if found:
        _widen_dates_box(root)
    # each value exists twice (DrawingML + VML fallback): keep the messages unique
    return list(dict.fromkeys(log)), found


def _widen_dates_box(root, min_cx=1599220):
    """Widen the Received/Revised/Accepted text box so long dates (e.g. “13 November 2026”) stay on one line."""
    for alt in root.iter(q('mc:AlternateContent')):
        if 'Received:' not in ptext(alt):
            continue
        an = alt.find('.//wp:anchor', NS)
        if an is None:
            continue
        ext = an.find('wp:extent', NS)
        cx = int(ext.get('cx'))
        if cx >= min_cx:
            continue
        delta = min_cx - cx
        ext.set('cx', str(min_cx))
        for ae in an.iter(q('a:ext')):
            if ae.get('cx') == str(cx):
                ae.set('cx', str(min_cx))
        off = an.find('wp:positionH/wp:posOffset', NS)
        if off is not None:
            off.text = str(int(off.text) - delta)
        pt = delta / 12700.0
        for v in alt.iter():
            st = v.get('style')
            if st and 'width:' in st and v.tag.startswith('{urn:schemas-microsoft-com:vml}'):
                def adj(m, sign):
                    return '%s:%.1fpt' % (m.group(1), float(m.group(2)) + sign * pt)
                st = re.sub(r'(width):([\d.]+)pt', lambda m: adj(m, 1), st)
                st = re.sub(r'(margin-left):([\d.]+)pt', lambda m: adj(m, -1), st)
                v.set('style', st)


def fix_copyright(root, year):
    log = []
    for p in inner_paras(root):
        for old, new in regex_sub_para(p, r'©\s*(\d{4})(\s+by the authors)', r'© %s\2' % year):
            log.append('Lisensi: “%s” → “%s”' % (old, new))
    return log


# --------------------------------------------------------------------------- text cleanup
ZW = '\u200b\u200c\u200d\ufeff\u2060'


def fix_invisible_and_spaces(root):
    n_zw = n_sp = 0
    for t in root.iter(WT):
        if not t.text:
            continue
        s = t.text
        for ch in ZW:
            if ch in s:
                n_zw += s.count(ch)
                s = s.replace(ch, '')
        if t.text != s:
            t.text = s
    front, main, refs = body_regions(root)
    for p in paras_of(front + main + refs):
        for old, new in regex_sub_para(p, r'(?<=\S) {2,}(?=\S)', ' '):
            n_sp += 1
    log = []
    if n_zw:
        log.append('%d karakter tak terlihat (zero-width space/BOM) dihapus' % n_zw)
    if n_sp:
        log.append('%d spasi ganda di dalam kalimat dijadikan satu' % n_sp)
    return log


def fix_math_spacing(root):
    n = 0
    for p in inner_paras(root):
        for om in p.findall('m:oMath', NS):
            mts = om.xpath('.//m:t/text()', namespaces=NS)
            if not mts:
                continue
            nxt = om.getnext()
            while nxt is not None and nxt.tag in (q('w:bookmarkStart'), q('w:bookmarkEnd')):
                nxt = nxt.getnext()
            if nxt is not None and is_text_run(nxt):
                t = get_t(nxt)
                if t and t[0].isalnum():
                    set_t(nxt, ' ' + t)
                    n += 1
            prv = om.getprevious()
            if prv is not None and is_text_run(prv):
                t = get_t(prv)
                if re.search(r'\b[a-z]{2,}$', t) and not mts[0][:1].isspace():
                    set_t(prv, t + ' ')
                    n += 1
    return ['%d spasi yang hilang di sekitar persamaan inline ditambahkan' % n] if n else []


def fix_decimal_commas(root):
    """1426,06 → 1426.06 in the body (two decimals or ≥4 decimals; 12,345 thousands are left alone)."""
    front, main, _ = body_regions(root)
    ex = []
    for p in paras_of(front + main):
        for old, new in regex_sub_para(p, r'(?<![\d,.])(\d+),(\d{2}|\d{4,})(?![\d,])', r'\1.\2'):
            ex.append(old)
    if not ex:
        return []
    return ['%d koma desimal diganti titik (template: desimal memakai titik), mis. %s' %
            (len(ex), ', '.join(dict.fromkeys(ex[:6])))]


SAFE_LANG = [
    ('data train', 'training data'), ('data test', 'test data'), ('Data Train', 'Training Data'),
    ('Data Test', 'Test Data'), ('data training', 'training data'), ('data testing', 'testing data'),
    ('Data Training', 'Training Data'), ('Data Testing', 'Testing Data'),
    ('residues', 'residuals'), ('residue', 'residual'),
    ('Child3.4', 'Niño3.4'), (' it shown in ', ' is shown in '), (' ad the time', ' at the time'),
    ('three-error metrics', 'three error metrics'), ('two-evaluation metrics', 'two evaluation metrics'),
    ('Indeks', 'Index'), (' indeks ', ' index '),
]


def fix_safe_language(root):
    front, main, _ = body_regions(root)
    log = []
    ps = paras_of(front + main)
    for a, b in SAFE_LANG:
        n = 0
        for p in ps:
            n += rep(p, a, b)
        if n:
            log.append('“%s” → “%s” (%d×)' % (a.strip(), b.strip(), n))
    n = 0
    for p in ps:
        n += len(regex_sub_para(p, r'\bet al(?![.\w])', 'et al.'))
    if n:
        log.append('“et al” → “et al.” (%d×)' % n)
    return log


# --------------------------------------------------------------------------- citations
CIT = re.compile(r'^\[\d+(?:[–-]\d+)?\]$')


def fix_citation_groups(root):
    n = 0
    sdt_tag = q('w:sdt')
    for p in inner_paras(root):
        kids = list(p)
        for i, c in enumerate(kids):
            # Mendeley Cite (content controls)
            if c.tag == sdt_tag and c.getnext() is not None:
                nxt = c.getnext()
                if nxt.tag == sdt_tag:
                    sep = etree.Element(WR)
                    set_t(sep, ', ')
                    c.addnext(sep)
                    n += 1
                elif is_text_run(nxt) and get_t(nxt).strip() == '' and nxt.getnext() is not None \
                        and nxt.getnext().tag == sdt_tag:
                    set_t(nxt, ', ')
                    n += 1
            # Mendeley Desktop / Zotero (complex fields)
            if c.tag == WR and CIT.match(get_t(c) or '') and i + 1 < len(kids):
                e = kids[i + 1]
                fc = e.find('w:fldChar', NS)
                if fc is None or fc.get(q('w:fldCharType')) != 'end':
                    continue
                j = i + 2
                while j < len(kids) and kids[j].tag == WR and kids[j].find(WT) is None:
                    j += 1
                if j < len(kids) and kids[j].tag == WR and CIT.match(get_t(kids[j]) or ''):
                    e.addnext(new_run_like(c, ', '))
                    n += 1
                elif j < len(kids) and is_text_run(kids[j]) and get_t(kids[j]).strip() == '':
                    k = j + 1
                    while k < len(kids) and kids[k].tag == WR and kids[k].find(WT) is None:
                        k += 1
                    if k < len(kids) and CIT.match(get_t(kids[k]) or ''):
                        set_t(kids[j], ', ')
                        n += 1
        # plain-text citations inside one run
        n += len(regex_sub_para(p, r'(\[\d+\])\s*(?=\[\d+\])', r'\1, '))
    return ['%d kelompok sitasi dipisah dengan “, ” sesuai IEEE (mis. [1], [3])' % n] if n else []


# --------------------------------------------------------------------------- equations
def fix_equations(root, layout=True, text_width_twips=None):
    log = []
    # accent placed over a subscripted symbol -> accent over the base only (Ŷ_t)
    n = 0
    for a in list(root.iter(q('m:acc'))):
        e = a.find('m:e', NS)
        kids = [k for k in e if local(k) != 'ctrlPr']
        if len(kids) == 1 and local(kids[0]) == 'sSub':
            ss = kids[0]
            base = ss.find('m:e', NS)
            newacc = copy.deepcopy(a)
            ne = newacc.find('m:e', NS)
            for k in list(ne):
                if local(k) != 'ctrlPr':
                    ne.remove(k)
            for k in list(base):
                ne.append(k)
            for k in list(base):
                base.remove(k)
            base.append(newacc)
            a.getparent().replace(a, ss)
            n += 1
    if n:
        log.append('%d aksen (topi/tilde) dipindah dari simbol bersubskrip ke huruf dasarnya (Ŷₜ, bukan topi di atas Yₜ)' % n)
    n = 0
    for mt in root.iter(q('m:t')):
        if mt.text and '∅' in mt.text:
            mt.text = mt.text.replace('∅', 'φ')
            n += 1
    if n:
        log.append('Simbol himpunan kosong “∅” di persamaan diganti “φ” (%d×)' % n)
    # one-row matrices with empty cells used as a layout trick
    n = 0
    for mm in list(root.iter(q('m:m'))):
        rows = mm.findall('m:mr', NS)
        if len(rows) != 1:
            continue
        cells = rows[0].findall('m:e', NS)
        filled = [c for c in cells if any(local(k) != 'ctrlPr' for k in c)]
        if len(cells) > 1 and len(filled) == 1:
            for k in list(filled[0]):
                if local(k) != 'ctrlPr':
                    mm.addprevious(k)
            mm.getparent().remove(mm)
            n += 1
    if n:
        log.append('%d matriks 1-baris berisi sel kosong (trik tata letak) dibuka menjadi persamaan biasa' % n)
    # numbered equations written as "math + tab + (n)" in a paragraph (not in a table)
    if layout and text_width_twips:
        n = 0
        for p in root.find('w:body', NS).findall('w:p', NS):
            if p.find('m:oMath', NS) is None and p.find('m:oMathPara', NS) is None:
                continue
            m_ = re.search(r'\((\d+)\)\s*$', ptext(p))
            outside = ''.join(get_t(r) for r in p.findall('w:r', NS))
            if not m_ or re.sub(r'\s|\(\d+\)', '', outside):
                continue
            num = m_.group(1)
            runs = p.findall('w:r', NS)
            numr = next((r for r in runs if re.search(r'\(%s\)' % num, get_t(r))), None)
            if numr is None:
                continue
            for r in runs:
                if r is not numr and get_t(r).strip() == '' and r.find('w:drawing', NS) is None:
                    p.remove(r)
            for x in list(numr):
                if x.tag != WRPR:
                    numr.remove(x)
            etree.SubElement(numr, q('w:tab'))
            tt = etree.SubElement(numr, WT)
            tt.text = '(%s)' % num
            ppr = p.find(WPPR)
            for nm in ('tabs', 'ind', 'jc'):
                el = ppr.find(q('w:' + nm)) if ppr is not None else None
                if el is not None:
                    ppr.remove(el)
            tabs = ppr_set(p, 'tabs')
            tb = etree.SubElement(tabs, q('w:tab'))
            tb.set(q('w:val'), 'right')
            tb.set(q('w:pos'), str(text_width_twips))
            ppr_set(p, 'ind', {'left': '0', 'firstLine': '0'})
            ppr_set(p, 'jc', {'val': 'left'})
            n += 1
        if n:
            log.append('%d persamaan bernomor dirapikan: persamaan di margin kiri, nomor di margin kanan' % n)
    return log


# --------------------------------------------------------------------------- layout helpers
def fix_keep_with_next(root):
    body = root.find('w:body', NS)
    n = 0
    for p in body.findall('w:p', NS):
        t = ptext(p).strip()
        is_cap = re.match(r'^Table\s+\d+', t) is not None
        is_head = (len(t) < 110 and not t.endswith('.') and
                   (p.find('w:pPr/w:numPr', NS) is not None and len(t.split()) <= 14 or
                    re.match(r'^(I|II|III|IV|V|VI|VII|VIII)\.\s+[A-Z][A-Z ,&-]+$', t) or
                    t.upper() in ('REFERENCES', 'ACKNOWLEDGEMENT', 'ACKNOWLEDGMENT', 'ACKNOWLEDGEMENTS')))
        if (is_cap or is_head) and t and p.find('w:pPr/w:keepNext', NS) is None:
            ppr_set(p, 'keepNext')
            n += 1
    for tbl in body.iter(q('w:tbl')):
        rows = tbl.findall('w:tr', NS)
        if len(rows) > 25:
            continue
        for tr in rows[:-1]:
            for pp in tr.iter(WP):
                if pp.find('w:pPr/w:keepNext', NS) is None:
                    ppr_set(pp, 'keepNext')
    return ['Judul bagian dan caption tabel diberi “keep with next” agar tidak terpisah dari isinya (%d paragraf)' % n] if n else []


# --------------------------------------------------------------------------- references
ACRONYMS = ['ARIMA', 'SARIMA', 'ARIMAX', 'SARIMAX', 'ARMA', 'GARCH', 'ARCH', 'VAR', 'EWMA', 'MEWMA', 'MEWMV', 'CUSUM',
            'MCUSUM', 'LSTM', 'GRU', 'RNN', 'CNN', 'GAN', 'SVR', 'SVM', 'PCA', 'KPCA', 'PLS', 'ANN', 'MLP', 'XGBOOST',
            'SMOTE', 'CART', 'ADABOOST', 'BMKG', 'ENSO', 'PDAM', 'COVID', 'GDP', 'IOT', 'ISPU', 'LASSO', 'MAPE', 'RMSE',
            'SMAPE', 'MAE', 'AIC', 'BIC', 'SPC', 'ARL', 'BPS', 'DKI', 'NTB', 'NTT', 'IPM', 'UMKM']
SMALL_EN = ['a', 'an', 'the', 'and', 'or', 'of', 'in', 'on', 'at', 'for', 'to', 'with', 'by', 'from', 'via']
SMALL_ID = ['dan', 'di', 'pada', 'untuk', 'dengan', 'yang', 'dari', 'ke', 'dalam', 'terhadap', 'atau']
MONTHS_RX = r'(?:January|February|March|April|May|June|July|August|September|October|November|December|' \
            r'Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\.?'
ABBR_OK = {'al', 'inc', 'ltd', 'co', 'corp', 'univ', 'dept', 'ed', 'eds', 'proc', 'vol', 'no', 'pp', 'jr', 'sr',
           'conf', 'int', 'j', 'trans', 'rev', 'res', 'sci', 'eng', 'technol', 'appl', 'stat', 'math'}


_SPELL = None


def _speller():
    global _SPELL
    if _SPELL is None:
        from spellchecker import SpellChecker
        _SPELL = SpellChecker(distance=1)
    return _SPELL


def _is_full_word(word):
    """True for dictionary words ("Methods"), False for journal abbreviations ("Phys", "Manag")."""
    try:
        return word.lower() in _speller() and len(word) >= 5
    except Exception:
        return False


def reference_entries(root):
    """Reference list entries: dicts {n, paras, text, kind}.

    Handles plain paragraphs (“[1] A. Author …”), two-column tables (“[1]” | entry) and Word numbered lists whose
    number “[1]” comes from the list format rather than from the text.
    """
    _, _, refs = body_regions(root)
    out = []
    seen = set()
    numbered = []
    for p in paras_of(refs[1:] if refs else []):
        t = ptext(p).strip()
        if 'Creative Commons' in t:
            break
        if not t or id(p) in seen:
            continue
        m = re.match(r'^\[(\d+)\]\s*(.*)$', t, re.S)
        tc = next((a for a in p.iterancestors() if local(a) == 'tc'), None)
        if m and not m.group(2).strip() and tc is not None:
            row = tc.getparent()
            cells = row.findall('w:tc', NS)
            k = cells.index(tc)
            if k + 1 < len(cells):
                ps = [x for x in cells[k + 1].iter(WP)]
                for x in ps:
                    seen.add(id(x))
                out.append({'n': int(m.group(1)), 'paras': ps, 'kind': 'row',
                            'text': '[%s] %s' % (m.group(1), ' '.join(ptext(x).strip() for x in ps))})
            continue
        if m:
            out.append({'n': int(m.group(1)), 'paras': [p], 'kind': 'para', 'text': t})
        elif p.find('w:pPr/w:numPr', NS) is not None:
            numbered.append(p)
    if not out and len(numbered) >= 3:
        for i, p in enumerate(numbered, 1):
            out.append({'n': i, 'paras': [p], 'kind': 'list', 'text': '[%d] %s' % (i, ptext(p).strip())})
    return out


def reference_paragraphs(root):
    return [e['paras'][0] for e in reference_entries(root) if e['kind'] == 'para']


def _title_span(text):
    m = re.search(r'[“"](.+?)[”"]', text)
    return (m.start(1), m.end(1)) if m else None


def fix_references(root):
    log = []
    acr_rx = re.compile(r'\b(%s)\b' % '|'.join(a.capitalize() for a in ACRONYMS if a.capitalize() != a))
    for e in reference_entries(root):
        num = e['n']
        for p in e['paras']:
            ch = []
            ch += regex_sub_para(p, r',\s*no\.\s*%s(?:\s+\d{4})?\s*,' % MONTHS_RX, ',')
            ch += regex_sub_para(p, r'\[Online\]\.\s*Available:\s*https?://(?:dx\.)?doi\.org/(\S+?)\.?\s*$', r'doi: \1.')
            ch += regex_sub_para(p, acr_rx, lambda m: m.group(0).upper())

            def dotcomma(m):
                word = m.group(1)
                if word.lower() in ABBR_OK or len(word) <= 3 or word.isupper():
                    return m.group(0)
                return word + ',' if _is_full_word(word) else m.group(0)
            ch += regex_sub_para(p, r'(\w+)\.,', dotcomma)
            # lower-case small words inside the quoted title (not the first word, not after a colon)
            ch += _lower_small_words_in_title(p, SMALL_EN + SMALL_ID)
            for old, new in ch:
                log.append('[%s] “%s” → “%s”' % (num, old.strip(), new.strip()))
    return log


def _lower_small_words_in_title(p, small):
    """Within the quoted title only (character positions), lower-case capitalised small words."""
    changes = []
    for seg in list(segments(p)):
        full = ''.join(get_t(r) for r in seg)
        span = _title_span(full)
        if not span:
            continue
        a, b = span
        title = full[a:b]
        for m in reversed(list(re.finditer(r'(?<=[A-Za-z0-9\)\]] )(%s)(?=\s)' % '|'.join(w.capitalize() for w in small),
                                           title))):
            before = title[:m.start()].rstrip()
            if before.endswith(':') or before.endswith('—') or before.endswith('-'):
                continue
            replace_span(seg, a + m.start(), a + m.end(), m.group(1).lower())
            changes.append((m.group(1), m.group(1).lower()))
    return changes


# --------------------------------------------------------------------------- tracked changes & comments
def accept_tracked_changes(root):
    n = 0
    # a deleted paragraph mark joins the paragraph with the next one (Word semantics)
    for p in list(root.iter(WP)):
        mark = p.find('w:pPr/w:rPr/w:del', NS)
        if mark is None:
            continue
        mark.getparent().remove(mark)
        nxt = p.getnext()
        if nxt is not None and nxt.tag == WP:
            content = [c for c in p if c.tag != WPPR]
            npr = nxt.find(WPPR)
            idx = 0 if npr is None else nxt.index(npr) + 1
            for c in content:
                nxt.insert(idx, c)
                idx += 1
            p.getparent().remove(p)
        n += 1
    for tag in ('w:del', 'w:moveFrom'):
        for el in list(root.iter(q(tag))):
            if el.getparent() is not None:
                el.getparent().remove(el)
                n += 1
    for tag in ('w:ins', 'w:moveTo'):
        for el in list(root.iter(q(tag))):
            parent = el.getparent()
            if parent is None:
                continue
            if local(parent) == 'rPr':           # inserted paragraph mark: just drop the marker
                parent.remove(el)
                n += 1
                continue
            idx = parent.index(el)
            for k in list(el):
                parent.insert(idx, k)
                idx += 1
            parent.remove(el)
            n += 1
    for tag in ('w:rPrChange', 'w:pPrChange', 'w:sectPrChange', 'w:tblPrChange', 'w:trPrChange', 'w:tcPrChange',
                'w:tblGridChange', 'w:numberingChange', 'w:cellIns', 'w:cellDel', 'w:cellMerge',
                'w:moveFromRangeStart', 'w:moveFromRangeEnd', 'w:moveToRangeStart', 'w:moveToRangeEnd'):
        for el in list(root.iter(q(tag))):
            el.getparent().remove(el)
            n += 1
    return ['%d jejak track changes diterima/dibersihkan' % n] if n else []


def remove_comments(pkg, root):
    n = 0
    for tag in ('w:commentRangeStart', 'w:commentRangeEnd'):
        for el in list(root.iter(q(tag))):
            el.getparent().remove(el)
            n += 1
    for el in list(root.iter(q('w:commentReference'))):
        r = el.getparent()
        r.getparent().remove(r)
        n += 1
    removed = []
    for part in ('word/comments.xml', 'word/commentsExtended.xml', 'word/commentsIds.xml',
                 'word/commentsExtensible.xml', 'word/people.xml'):
        if pkg.has(part):
            pkg.remove(part)
            removed.append(part.split('/')[-1])
    if removed and pkg.has(pkg.main_rels):
        rels = pkg.text(pkg.main_rels)
        rels = re.sub(r'<Relationship [^>]*Target="/?(?:word/)?(?:%s)"[^>]*/>' % '|'.join(re.escape(x) for x in removed), '', rels)
        pkg.set_text(pkg.main_rels, rels)
        ct = pkg.text('[Content_Types].xml')
        ct = re.sub(r'<Override [^>]*PartName="/word/(?:%s)"[^>]*/>' % '|'.join(re.escape(x) for x in removed), '', ct)
        pkg.set_text('[Content_Types].xml', ct)
    return ['Komentar Word dihapus (%d penanda)' % n] if (n or removed) else []


# --------------------------------------------------------------------------- metadata
def clean_metadata(pkg, title: str):
    log = []
    if pkg.has('docProps/core.xml'):
        c = pkg.text('docProps/core.xml')
        before = c
        esc = (title or '').replace('&', '&amp;').replace('<', '&lt;')
        if '<dc:title' in c:
            c = re.sub(r'<dc:title>.*?</dc:title>|<dc:title/>', '<dc:title>%s</dc:title>' % esc, c, flags=re.S)
        for tag in ('dc:creator', 'cp:lastModifiedBy', 'cp:keywords', 'dc:description', 'dc:subject', 'cp:category'):
            c = re.sub(r'<%s>.*?</%s>' % (tag, tag), '<%s></%s>' % (tag, tag), c, flags=re.S)
        c = re.sub(r'<cp:revision>.*?</cp:revision>', '<cp:revision>1</cp:revision>', c)
        c = re.sub(r'<cp:lastPrinted>.*?</cp:lastPrinted>', '', c)
        if c != before:
            pkg.set_text('docProps/core.xml', c)
            log.append('Metadata dokumen dibersihkan (pembuat, penyunting terakhir, jumlah revisi, tanggal cetak)')
    if pkg.has('docProps/app.xml'):
        a = pkg.text('docProps/app.xml')
        a = re.sub(r'<TotalTime>\d+</TotalTime>', '<TotalTime>0</TotalTime>', a)
        a = re.sub(r'<Company>.*?</Company>', '<Company></Company>', a, flags=re.S)
        a = re.sub(r'<Manager>.*?</Manager>', '<Manager></Manager>', a, flags=re.S)
        pkg.set_text('docProps/app.xml', a)
    if pkg.has('docProps/custom.xml'):
        cx = pkg.text('docProps/custom.xml')
        props = re.findall(r'<property\b.*?</property>', cx, flags=re.S)
        keep = [x for x in props if re.search(r'name="Mendeley', x)]       # needed by Mendeley to refresh citations
        if len(keep) != len(props):
            if keep:
                head = cx[:cx.index('<property')] if '<property' in cx else cx
                pkg.set_text('docProps/custom.xml', head + ''.join(keep) + '</Properties>')
            else:
                pkg.remove('docProps/custom.xml')
                rels = pkg.text('_rels/.rels')
                pkg.set_text('_rels/.rels', re.sub(r'<Relationship [^>]*Target="/?docProps/custom.xml"[^>]*/>', '', rels))
                ct = pkg.text('[Content_Types].xml')
                pkg.set_text('[Content_Types].xml', re.sub(r'<Override [^>]*PartName="/docProps/custom.xml"[^>]*/>', '', ct))
            log.append('Properti kustom dokumen dibersihkan (%d dihapus; properti Mendeley dipertahankan)' % (len(props) - len(keep)))
    if pkg.has('word/settings.xml'):
        s = pkg.text('word/settings.xml')
        n = len(re.findall(r'<w:docVar w:name="__Grammarly[^>]*/>', s))
        s = re.sub(r'<w:docVar w:name="__Grammarly[^>]*/>', '', s)
        s = re.sub(r'<w:docVars>\s*</w:docVars>', '', s)
        s = re.sub(r'<w:attachedTemplate [^>]*/>', '', s)
        pkg.set_text('word/settings.xml', s)
        if n:
            log.append('Jejak Grammarly di pengaturan dokumen dihapus (%d)' % n)
    return log
