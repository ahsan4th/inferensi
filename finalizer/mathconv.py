"""Turn variables and equations typed as plain text into Microsoft Word equations (OMML).

Handled patterns (body text and tables; references, headers and captions are left alone):
- a letter/Greek symbol followed by runs formatted as subscript/superscript:  x + ₍sub₎1,i,t  ->  x_{1,i,t}
- Unicode sub/superscripts and math-italic letters:  H₀, μ₁, T², Σ⁻¹, 𝑦  ->  equation objects
- LaTeX-like snippets typed in text:  x_{i,t}, y_t, e^{-1}
- single italic letters used as variables:  italic “t”, “n”, “k”  ->  equation objects
- numbered equations typed as a text paragraph, e.g. “δ = √[(μ₁ − μ₀)′ Σ⁻¹ (μ₁ − μ₀)]   (12)”
"""
from __future__ import annotations

import re

from lxml import etree

from .docxtools import NS, WR, WRPR, WT, body_regions, get_t, local, paras_of, q

M_NS = NS['m']
SUB = dict(zip('₀₁₂₃₄₅₆₇₈₉₊₋₌₍₎ₐₑₒₓₕₖₗₘₙₚₛₜᵢⱼᵣᵤᵥ', '0123456789+-=()aeoxhklmnpstijruv'))
SUP = dict(zip('⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾ⁿⁱᵃᵇᶜᵈᵉᶠᵍʰʲᵏˡᵐᵒᵖʳˢᵗᵘᵛʷˣʸᶻ', '0123456789+-=()niabcdefghjklmoprstuvwxyz'))
GREEK = 'αβγδεζηθικλμνξοπρστυφχψωΓΔΘΛΞΠΣΥΦΨΩϵϑϕµ'
SUB_CH, SUP_CH = ''.join(SUB), ''.join(SUP)
# mathematical italic / bold-italic Latin and Greek letters (U+1D400 block) -> plain letters
_MATH_ALPHA = {}
for base, start in (('A', 0x1D434), ('a', 0x1D44E), ('A', 0x1D468), ('a', 0x1D482)):
    for k in range(26):
        _MATH_ALPHA[chr(start + k)] = chr(ord(base) + k)
_MATH_ALPHA['\u210e'] = 'h'                       # Planck constant = italic h
for k, g in enumerate('ΑΒΓΔΕΖΗΘΙΚΛΜΝΞΟΠΡΣΤΥΦΧΨΩ'):
    _MATH_ALPHA[chr(0x1D6E2 + k)] = g
for k, g in enumerate('αβγδεζηθικλμνξοπρςστυφχψω'):
    _MATH_ALPHA[chr(0x1D6FC + k)] = g
MATH_ALPHA_CH = ''.join(_MATH_ALPHA)
UNITS = {'m', 'cm', 'mm', 'km', 'kg', 'g', 'mg', 'µg', 'μg', 's', 'ms', 'h', 'l', 'L', 'ml', 'mL', 'K', 'Pa', 'W', 'J',
         'N', 'Hz', 'mol', 'ha'}


def _m(tag):
    return '{%s}%s' % (M_NS, tag)


# --------------------------------------------------------------------------- OMML builders
def _rpr(size):
    rp = etree.Element(WRPR)
    f = etree.SubElement(rp, q('w:rFonts'))
    f.set(q('w:ascii'), 'Cambria Math')
    f.set(q('w:hAnsi'), 'Cambria Math')
    if size:
        etree.SubElement(rp, q('w:sz')).set(q('w:val'), size)
        etree.SubElement(rp, q('w:szCs')).set(q('w:val'), size)
    return rp


def mrun(text, size=None, upright=False):
    r = etree.Element(_m('r'))
    if upright:
        mp = etree.SubElement(r, _m('rPr'))
        etree.SubElement(mp, _m('sty')).set(_m('val'), 'p')
    r.append(_rpr(size))
    t = etree.SubElement(r, _m('t'))
    t.text = text
    if text != text.strip():
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return r


def _ctrl(parent_pr, size):
    c = etree.SubElement(parent_pr, _m('ctrlPr'))
    c.append(_rpr(size))


def _token_runs(text, size):
    """Split a short expression into math runs: words of ≥2 letters upright (ARL, UCL, log), the rest default."""
    out = []
    for tok in re.findall(r'[A-Za-z]{2,}|\d+(?:[.,]\d+)?|.', text):
        out.append(mrun(tok, size, upright=bool(re.fullmatch(r'[A-Za-z]{2,}', tok))))
    return out


def script(base_text, sub_text, sup_text, size):
    """Base with optional subscript and superscript."""
    if sub_text and sup_text:
        el = etree.Element(_m('sSubSup'))
        _ctrl(etree.SubElement(el, _m('sSubSupPr')), size)
    elif sub_text:
        el = etree.Element(_m('sSub'))
        _ctrl(etree.SubElement(el, _m('sSubPr')), size)
    elif sup_text:
        el = etree.Element(_m('sSup'))
        _ctrl(etree.SubElement(el, _m('sSupPr')), size)
    else:
        return _token_runs(base_text, size)
    e = etree.SubElement(el, _m('e'))
    for r in _token_runs(base_text, size):
        e.append(r)
    if sub_text:
        s = etree.SubElement(el, _m('sub'))
        for r in _token_runs(sub_text, size):
            s.append(r)
    if sup_text:
        s = etree.SubElement(el, _m('sup'))
        for r in _token_runs(sup_text, size):
            s.append(r)
    return [el]


def omath(children):
    om = etree.Element(_m('oMath'))
    for c in children:
        om.append(c)
    return om


def _size_of(run):
    s = run.find('w:rPr/w:sz', NS) if run is not None else None
    return s.get(q('w:val')) if s is not None else None


# --------------------------------------------------------------------------- scope
def _paragraphs(root):
    front, main, _ = body_regions(root)
    out = []
    for p in paras_of(main):
        if any(local(a) in ('txbxContent', 'Fallback') for a in p.iterancestors()):
            continue
        t = ''.join(p.xpath('./w:r/w:t/text()', namespaces=NS)).strip()
        if re.match(r'^(Figure|Table|Gambar|Tabel)\s+\d+', t) and len(t) < 300:
            continue                     # captions keep their typography
        if re.match(r'^(?:[IVX]+\.?\s+)?[A-Z][A-Z ,&-]{3,}$', t):
            continue                     # main headings
        out.append(p)
    return out


def _is_text_run(r):
    return r.tag == WR and r.find(WT) is not None and all(c.tag in (WRPR, WT) for c in r)


def _vert(r):
    v = r.find('w:rPr/w:vertAlign', NS)
    return v.get(q('w:val')) if v is not None else None


def _italic(r):
    i = r.find('w:rPr/w:i', NS)
    return i is not None and i.get(q('w:val')) not in ('0', 'false')


def _split_run(r, idx):
    """Split run r at idx; returns the right-hand run (inserted after r)."""
    import copy
    t = get_t(r)
    right = copy.deepcopy(r)
    for x in r.findall(WT):
        r.remove(x)
    tl = etree.SubElement(r, WT)
    tl.text = t[:idx]
    tl.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    for x in right.findall(WT):
        right.remove(x)
    tr = etree.SubElement(right, WT)
    tr.text = t[idx:]
    tr.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    r.addnext(right)
    return right


# --------------------------------------------------------------------------- 1. sub/superscript-formatted runs
def convert_script_runs(root):
    n = 0
    for p in _paragraphs(root):
        kids = list(p)
        i = 0
        while i < len(kids):
            r = kids[i]
            if not (_is_text_run(r) and _vert(r) is None):
                i += 1
                continue
            j = i + 1
            subs, sups, scr = [], [], []
            while j < len(kids) and _is_text_run(kids[j]) and _vert(kids[j]) in ('subscript', 'superscript'):
                (subs if _vert(kids[j]) == 'subscript' else sups).append(get_t(kids[j]))
                scr.append(kids[j])
                j += 1
            if not scr:
                i += 1
                continue
            sub_t, sup_t = ''.join(subs).strip(), ''.join(sups).strip()
            base_text = get_t(r)
            m = re.search(r'([A-Za-z%s]{1,4})(′|\')?$' % GREEK, base_text)
            ok_scripts = all(re.fullmatch(r"[\w,+\-−'′*()= ]+", x) for x in (sub_t, sup_t) if x)
            footnote_mark = (not sub_t) and bool(re.fullmatch(r'\*+|†+|‡+|[a-z]', sup_t)) and len(m.group(1)) > 1 \
                if m else False
            if not m or not ok_scripts or footnote_mark:
                i = j
                continue
            before = base_text[:m.start()]
            if m.group(1) in UNITS and re.search(r'[\d/]\s?$', before):
                i = j
                continue                 # units such as m², µg/m³ stay as text
            if len(m.group(1)) > 1 and m.group(1).lower() in ('st', 'nd', 'rd', 'th'):
                i = j
                continue
            size = _size_of(r)
            node = omath(script(m.group(1) + (m.group(2) or ''), sub_t.replace('−', '-'), sup_t.replace('−', '-'), size))
            if before:
                right = _split_run(r, m.start())
                p.replace(right, node)
            else:
                p.replace(r, node)
            for x in scr:
                p.remove(x)
            n += 1
            kids = list(p)
            i = kids.index(node) + 1
    return n


# --------------------------------------------------------------------------- 2. Unicode scripts, math italics, LaTeX-like
TOKEN = re.compile(
    r'(?<![\w%s])'
    r'(?:(?P<nb>\d+)(?P<nsup>[%s]+)(?![\w%s])|(?P<b>[A-Za-z%s%s]{1,4})(?P<pr>′)?(?P<sub>[%s]+)?(?P<sup>[%s]+)?(?![\w%s])'
    r'|(?P<lb>[A-Za-z%s])_(?:\{(?P<ls>[^{}]{1,12})\}|(?P<ls1>[A-Za-z0-9]))(?:\^(?:\{(?P<lp>[^{}]{1,12})\}|(?P<lp1>[A-Za-z0-9-])))?)'
    % (GREEK + MATH_ALPHA_CH, re.escape(SUP_CH), GREEK, GREEK, MATH_ALPHA_CH, re.escape(SUB_CH), re.escape(SUP_CH), GREEK,
       GREEK))


def _plain(s):
    return ''.join(_MATH_ALPHA.get(c, c) for c in s)


def _join_split_scripts(p):
    """“z” | “ᵢ” in two runs (different formatting) -> keep the script characters with their base."""
    kids = [c for c in p if _is_text_run(c)]
    for a, b in zip(kids, kids[1:]):
        if a.getnext() is not b:
            continue
        ta, tb = get_t(a), get_t(b)
        m = re.match('[%s%s]+' % (re.escape(SUB_CH), re.escape(SUP_CH)), tb)
        if m and re.search(r'[\w%s]$' % GREEK, ta):
            a.find(WT).text = ta + m.group(0)
            a.find(WT).set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            for w in a.findall(WT)[1:]:
                a.remove(w)
            rest = tb[m.end():]
            if rest:
                b.find(WT).text = rest
                for w in b.findall(WT)[1:]:
                    b.remove(w)
            else:
                p.remove(b)


def convert_unicode_math(root):
    n = 0
    for p in _paragraphs(root):
        for _ in range(6):
            before = len(p)
            _join_split_scripts(p)
            if len(p) == before:
                break
        for r in list(p):
            if not _is_text_run(r) or _vert(r):
                continue
            t = get_t(r)
            if not re.search('[%s%s%s_]' % (re.escape(SUB_CH), re.escape(SUP_CH), MATH_ALPHA_CH), t):
                continue
            pieces, pos = [], 0
            for m in TOKEN.finditer(t):
                if m.group('nb'):
                    pieces.append((m.start(), m.end(), m.group('nb'), '', ''.join(SUP.get(c, c) for c in m.group('nsup'))))
                    continue
                if m.group('lb'):
                    base, sub, sup = m.group('lb'), m.group('ls') or m.group('ls1'), m.group('lp') or m.group('lp1')
                else:
                    base = m.group('b')
                    sub = ''.join(SUB.get(c, c) for c in (m.group('sub') or ''))
                    sup = ''.join(SUP.get(c, c) for c in (m.group('sup') or ''))
                    has_alpha = any(c in _MATH_ALPHA for c in base)
                    if not (sub or sup or has_alpha):
                        continue
                    if base in UNITS and re.search(r'[\d/]\s?$', t[:m.start()]) and not sub:
                        continue
                    if len(base) > 1 and not has_alpha and not (sub or sup):
                        continue
                    base = _plain(base) + ('′' if m.group('pr') else '')
                pieces.append((m.start(), m.end(), base, sub or '', sup or ''))
            if not pieces:
                continue
            size = _size_of(r)
            nodes, pos = [], 0
            import copy
            for a, b, base, sub, sup in pieces:
                if t[pos:a]:
                    x = copy.deepcopy(r)
                    for w in x.findall(WT):
                        x.remove(w)
                    w = etree.SubElement(x, WT)
                    w.text = t[pos:a]
                    w.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                    nodes.append(x)
                nodes.append(omath(script(base, sub, sup, size)))
                n += 1
                pos = b
            if t[pos:]:
                x = copy.deepcopy(r)
                for w in x.findall(WT):
                    x.remove(w)
                w = etree.SubElement(x, WT)
                w.text = t[pos:]
                w.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                nodes.append(x)
            for x in nodes:
                r.addprevious(x)
            p.remove(r)
    return n


# --------------------------------------------------------------------------- 3. italic single letters
def convert_italic_letters(root):
    n = 0
    for p in _paragraphs(root):
        for r in list(p):
            if not _is_text_run(r) or not _italic(r) or _vert(r):
                continue
            t = get_t(r)
            m = re.fullmatch(r'(\s*)([A-Za-z%s])(\s*)' % GREEK, t)
            if not m:
                continue
            prev, nxt = r.getprevious(), r.getnext()
            if prev is not None and _is_text_run(prev) and re.search(r'[A-Za-z]$', get_t(prev)) and not m.group(1):
                continue                 # part of a word (e.g. emphasised letter)
            if nxt is not None and _is_text_run(nxt) and re.match(r'^[A-Za-z]', get_t(nxt)) and not m.group(3):
                continue
            size = _size_of(r)
            import copy
            nodes = []
            for txt in (m.group(1),):
                if txt:
                    x = copy.deepcopy(r)
                    x.find('w:rPr', NS).remove(x.find('w:rPr/w:i', NS))
                    for w in x.findall(WT):
                        x.remove(w)
                    w = etree.SubElement(x, WT)
                    w.text = txt
                    w.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                    nodes.append(x)
            nodes.append(omath([mrun(m.group(2), size)]))
            if m.group(3):
                x = copy.deepcopy(r)
                x.find('w:rPr', NS).remove(x.find('w:rPr/w:i', NS))
                for w in x.findall(WT):
                    x.remove(w)
                w = etree.SubElement(x, WT)
                w.text = m.group(3)
                w.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                nodes.append(x)
            for x in nodes:
                r.addprevious(x)
            p.remove(r)
            n += 1
    return n


# --------------------------------------------------------------------------- 4. text paragraphs that are numbered equations
def _linear_to_omml(expr, size):
    """Small linear-format parser: √[...] / √(...) / √x, Unicode scripts, primes, the rest as math runs."""
    out, i = [], 0
    while i < len(expr):
        c = expr[i]
        if c == '√':
            j = i + 1
            while j < len(expr) and expr[j] == ' ':
                j += 1
            if j < len(expr) and expr[j] in '([{':
                close = {'(': ')', '[': ']', '{': '}'}[expr[j]]
                depth, k = 0, j
                while k < len(expr):
                    if expr[k] == expr[j]:
                        depth += 1
                    elif expr[k] == close:
                        depth -= 1
                        if depth == 0:
                            break
                    k += 1
                inner = expr[j + 1:k] if expr[j] == '[' else expr[j:k + 1]
                i = k + 1
            else:
                m = re.match(r'[\w%s]+' % GREEK, expr[j:])
                inner = m.group(0) if m else ''
                i = j + len(inner)
            rad = etree.Element(_m('rad'))
            rp = etree.SubElement(rad, _m('radPr'))
            etree.SubElement(rp, _m('degHide')).set(_m('val'), '1')
            _ctrl(rp, size)
            etree.SubElement(rad, _m('deg'))
            e = etree.SubElement(rad, _m('e'))
            for x in _linear_to_omml(inner, size):
                e.append(x)
            out.append(rad)
            continue
        m = re.match(r'([A-Za-z%s]{1,4})(′)?([%s]+)?([%s]+)?' % (GREEK, re.escape(SUB_CH), re.escape(SUP_CH)), expr[i:])
        if m and (m.group(3) or m.group(4)):
            sub = ''.join(SUB.get(ch, ch) for ch in (m.group(3) or ''))
            sup = ''.join(SUP.get(ch, ch) for ch in (m.group(4) or ''))
            out += script(m.group(1) + (m.group(2) or ''), sub, sup, size)
            i += m.end()
            continue
        if c in SUP_CH and out:
            k = i
            while k < len(expr) and expr[k] in SUP_CH:
                k += 1
            sup = ''.join(SUP[ch] for ch in expr[i:k])
            last = out.pop()
            el = etree.Element(_m('sSup'))
            _ctrl(etree.SubElement(el, _m('sSupPr')), size)
            etree.SubElement(el, _m('e')).append(last)
            s = etree.SubElement(el, _m('sup'))
            for x in _token_runs(sup, size):
                s.append(x)
            out.append(el)
            i = k
            continue
        if c == ' ':
            i += 1
            continue
        m = re.match(r'[A-Za-z]{2,}|\d+(?:\.\d+)?|.', expr[i:])
        tok = m.group(0)
        out.append(mrun(tok.replace('−', '-') if tok == '−' else tok, size, upright=len(tok) > 1 and tok.isalpha()))
        i += len(tok)
    return out


def convert_text_equations(root):
    n = 0
    body = root.find('w:body', NS)
    for p in body.findall('w:p', NS):
        if p.find('.//m:oMath', NS) is not None or p.find('.//w:drawing', NS) is not None:
            continue
        runs = [r for r in p.findall('w:r', NS)]
        if not runs or not all(_is_text_run(r) or (r.find('w:tab', NS) is not None and r.find(WT) is None) for r in runs):
            continue
        t = ''.join(get_t(r) for r in runs).strip()
        m = re.fullmatch(r'(.+?=.+?)\s*\((\d+)\)', t)
        if not m or len(t) > 160 or re.search(r'\b(the|and|is|are|of|with|where|which)\b', m.group(1)):
            continue
        size = _size_of(runs[0])
        expr = m.group(1).strip()
        nodes = _linear_to_omml(expr, size)
        keep_rpr = runs[0].find(WRPR)
        for r in runs:
            p.remove(r)
        p.append(omath(nodes))
        num = etree.SubElement(p, WR)
        if keep_rpr is not None:
            num.append(etree.fromstring(etree.tostring(keep_rpr)))
        etree.SubElement(num, q('w:tab'))
        tt = etree.SubElement(num, WT)
        tt.text = '(%s)' % m.group(2)
        n += 1
    return n


def convert_all(root):
    log = []
    a = convert_text_equations(root)
    b = convert_script_runs(root)
    c = convert_unicode_math(root)
    d = convert_italic_letters(root)
    if a:
        log.append('%d persamaan bernomor yang diketik sebagai teks diubah menjadi Equation Word' % a)
    if b:
        log.append('%d variabel berformat subscript/superscript diubah menjadi Equation Word (mis. x₁,ᵢ,ₜ)' % b)
    if c:
        log.append('%d simbol/variabel teks (H₀, μ₁, T², huruf miring matematika, x_{i,t}) diubah menjadi Equation Word' % c)
    if d:
        log.append('%d variabel huruf miring tunggal diubah menjadi Equation Word' % d)
    return log


def remaining_plain_math(root):
    """Places where variables still look like plain text (for the review list)."""
    out = []
    for p in _paragraphs(root):
        for r in p.findall('w:r', NS):
            t = get_t(r)
            if not t.strip():
                continue
            if _vert(r) in ('subscript', 'superscript') and re.search(r'[A-Za-z0-9]', t):
                out.append(''.join(p.xpath('.//w:t/text()', namespaces=NS))[:120])
                break
            probe = re.sub(r'(?:\b(?:%s))[²³]' % '|'.join(sorted(UNITS, key=len, reverse=True)), '', t)
            if re.search('[%s%s%s]' % (re.escape(SUB_CH), re.escape(SUP_CH), MATH_ALPHA_CH), probe):
                out.append(t[:120])
                break
    return out
