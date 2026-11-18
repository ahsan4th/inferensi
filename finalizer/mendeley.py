"""Make citations updatable by Mendeley.

Documents whose citations are plain text ("[1]", "[2], [5]") and whose reference list is plain text are converted into
Mendeley Desktop citation fields (``ADDIN CSL_CITATION {…}``) with the bibliographic data embedded as CSL-JSON, plus a
``ADDIN Mendeley Bibliography CSL_BIBLIOGRAPHY`` field around the reference list. Mendeley Desktop refreshes such
fields from the embedded data, and Mendeley Cite offers to convert them on first use.
"""
from __future__ import annotations

import json
import re
import uuid

from lxml import etree

from .docxtools import NS, WP, WR, WRPR, WT, body_regions, get_t, local, paras_of, q, segments, set_t

SCHEMA = 'https://github.com/citation-style-language/schema/raw/master/csl-citation.json'
IEEE_STYLE = 'http://www.zotero.org/styles/ieee'


def has_citation_fields(root) -> bool:
    for it in root.iter(q('w:instrText')):
        if it.text and ('CSL_CITATION' in it.text or 'ZOTERO_ITEM' in it.text):
            return True
    for tag in root.iter(q('w:tag')):
        if tag.get(q('w:val'), '').startswith(('MENDELEY_CITATION', 'ZOTERO_ITEM')):
            return True
    return False


# --------------------------------------------------------------------------- IEEE reference parser
def _italic_text(p):
    out = []
    for r in p.iter(WR):
        rp = r.find(WRPR)
        i = rp.find('w:i', NS) if rp is not None else None
        if i is not None and i.get(q('w:val')) not in ('0', 'false'):
            out.append(get_t(r))
        else:
            out.append('\x00')
    spans = [s.strip(' ,.') for s in ''.join(out).split('\x00') if s.strip(' ,.')]
    return spans


def _names(s):
    s = re.sub(r'\bet al\.?', '', s).strip(' ,')
    parts = [x.strip() for x in re.split(r',\s*and\s+|\s+and\s+|,\s*', s) if x.strip()]
    out = []
    for nm in parts:
        m = re.match(r'^((?:[A-Z][a-z]?\.-?\s*)+)(.+)$', nm)
        if m:
            out.append({'given': m.group(1).strip(), 'family': m.group(2).strip(), 'parse-names': False,
                        'dropping-particle': '', 'non-dropping-particle': '', 'suffix': ''})
        else:
            toks = nm.split()
            if len(toks) > 1:
                out.append({'given': ' '.join(toks[:-1]), 'family': toks[-1], 'parse-names': False,
                            'dropping-particle': '', 'non-dropping-particle': '', 'suffix': ''})
            else:
                out.append({'family': nm, 'given': '', 'parse-names': False,
                            'dropping-particle': '', 'non-dropping-particle': '', 'suffix': ''})
    return out


def parse_ieee(text: str, italics: list) -> dict:
    t = re.sub(r'^\s*\[\d+\]\s*', '', text).strip()
    item = {'type': 'article-journal'}
    rest = t
    m = re.search(r'doi:\s*(10\.\S+?)\.?\s*$', t, re.I) or re.search(r'https?://(?:dx\.)?doi\.org/(10\.\S+?)\.?(?:\s|$)', t)
    if m:
        item['DOI'] = m.group(1)
    m = re.search(r'Available:\s*(https?://\S+?)\.?\s*$', t)
    if m and 'doi.org' not in m.group(1):
        item['URL'] = m.group(1)
    q_ = re.search(r'[“"](.+?)[,.]?[”"]', t)
    if q_:
        item['title'] = q_.group(1).strip().rstrip(',')
        item['author'] = _names(t[:q_.start()])
        rest = t[q_.end():]
        cont = next((s for s in italics if s and s in rest), None)
        if not cont:
            m = re.match(r'\s*,?\s*(?:in\s+)?([^,]{3,}?),\s*(?:vol\.|no\.|pp\.|Art\.|(?:19|20)\d\d)', rest)
            cont = m.group(1).strip() if m else None
        if cont:
            item['container-title'] = cont
        if re.search(r'\bin\b\s*$', t[q_.end():q_.end() + 6]) or re.search(
                r'\bProc\.|\bProceedings\b|\bConference\b|\bConf\.|\bSymposium\b|\bSeminar\b', rest[:160]):
            item['type'] = 'paper-conference'
        elif re.search(r'thesis|dissertation|skripsi|tesis', rest, re.I):
            item['type'] = 'thesis'
        elif re.search(r'arXiv', rest):
            item['type'] = 'article'
    else:
        title = italics[0] if italics else ''
        item['type'] = 'book'
        if title and title in t:
            item['title'] = title
            item['author'] = _names(t[:t.index(title)])
            rest = t[t.index(title) + len(title):]
            m = re.search(r'([A-Z][^.:,]+):\s*([^,]+),\s*((?:19|20)\d\d)', rest)
            if m:
                item['publisher-place'], item['publisher'] = m.group(1).strip(), m.group(2).strip()
            m = re.search(r'(\d+)(?:st|nd|rd|th)\s+ed\.', rest)
            if m:
                item['edition'] = m.group(1)
        else:
            item['title'] = t.split(',')[0][:250]
            rest = t
    for key, rx in (('volume', r'\bvol\.\s*([\w.-]+)'), ('issue', r'\bno\.\s*([\w.-]+)'), ('page', r'\bpp?\.\s*([\w–-]+)')):
        m = re.search(rx, rest)
        if m:
            item[key] = m.group(1).rstrip('.,').replace('–', '-')
    years = re.findall(r'\b((?:19|20)\d\d)\b', t if not item.get('DOI') else t.replace(item['DOI'], ''))
    if years:
        item['issued'] = {'date-parts': [[years[-1]]]}
    return item


# --------------------------------------------------------------------------- field builders
def _run(tag_children, rpr=None):
    r = etree.Element(WR)
    if rpr is not None:
        r.append(etree.fromstring(etree.tostring(rpr)))
    for c in tag_children:
        r.append(c)
    return r


def _text_run(text, rpr):
    r = etree.Element(WR)
    if rpr is not None:
        r.append(etree.fromstring(etree.tostring(rpr)))
    set_t(r, text)
    return r


def _fld(kind):
    el = etree.Element(q('w:fldChar'))
    el.set(q('w:fldCharType'), kind)
    return el


def _instr(text):
    el = etree.Element(q('w:instrText'))
    el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    el.text = text
    return el


def citation_field_runs(numbers, items, display, rpr):
    cit_items = []
    for k, n in enumerate(numbers, 1):
        data = dict(items[n])
        data['id'] = 'ITEM-%d' % k
        cit_items.append({'id': 'ITEM-%d' % k, 'itemData': data, 'uris': []})
    payload = {'citationItems': cit_items,
               'mendeley': {'formattedCitation': display, 'plainTextFormattedCitation': display,
                            'previouslyFormattedCitation': display},
               'properties': {'noteIndex': 0}, 'schema': SCHEMA}
    instr = ' ADDIN CSL_CITATION ' + json.dumps(payload, ensure_ascii=False)
    res = etree.Element(WT)
    res.text = display
    res.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    return [_run([_fld('begin')], rpr), _run([_instr(instr)], rpr), _run([_fld('separate')], rpr),
            _run([res], rpr), _run([_fld('end')], rpr)]


# --------------------------------------------------------------------------- conversion
def convert_plain_citations(root):
    """Return (log messages, issues). Does nothing when the document already has citation fields."""
    from .fixes import reference_entries
    if has_citation_fields(root):
        return [], []
    entries = reference_entries(root)
    if len(entries) < 3:
        return [], ['Daftar pustaka bernomor tidak ditemukan; sitasi tidak dikonversi ke field Mendeley.']
    items = {}
    for e in entries:
        ital = []
        for p in e['paras']:
            ital += _italic_text(p)
        items[e['n']] = parse_ieee(e['text'], ital)
    _, main, _ = body_regions(root)
    n_fields = 0
    rx = re.compile(r'\[(\d+)\]')
    for p in paras_of(main):
        if any(local(a) in ('tbl', 'txbxContent') for a in p.iterancestors()):
            continue
        for seg in list(segments(p)):
            for r in seg:
                t = get_t(r)
                mm = [m for m in rx.finditer(t) if int(m.group(1)) in items]
                if not mm:
                    continue
                rp = r.find(WRPR)
                nodes, pos = [], 0
                for m in mm:
                    if t[pos:m.start()]:
                        nodes.append(_text_run(t[pos:m.start()], rp))
                    nodes += citation_field_runs([int(m.group(1))], items, m.group(0), rp)
                    n_fields += 1
                    pos = m.end()
                if t[pos:]:
                    nodes.append(_text_run(t[pos:], rp))
                for x in nodes:
                    r.addprevious(x)
                r.getparent().remove(r)
    notes = []
    if any(e['kind'] != 'para' for e in entries):
        notes.append('Daftar pustaka berbentuk tabel/daftar bernomor otomatis, sehingga tidak dibungkus field Mendeley '
                     'Bibliography; sitasi tetap menjadi field Mendeley dan dapat di-refresh.')
        return ['%d sitasi teks biasa dikonversi menjadi field Mendeley (CSL_CITATION) dengan data %d referensi tertanam'
                % (n_fields, len(items))], notes
    # bibliography field around the reference list
    first, last = entries[0]['paras'][0], entries[-1]['paras'][-1]
    rpr = next((r.find(WRPR) for r in first.iter(WR) if r.find(WRPR) is not None), None)
    ppr = first.find(q('w:pPr'))
    idx = 0 if ppr is None else first.index(ppr) + 1
    for k, node in enumerate([_run([_fld('begin')], rpr), _run([_instr(' ADDIN Mendeley Bibliography CSL_BIBLIOGRAPHY ')], rpr),
                              _run([_fld('separate')], rpr)]):
        first.insert(idx + k, node)
    last.append(_run([_fld('end')], rpr))
    msgs = ['%d sitasi teks biasa dikonversi menjadi field Mendeley (CSL_CITATION) dengan data %d referensi tertanam; '
            'daftar pustaka dibungkus field Mendeley Bibliography sehingga dapat di-refresh Mendeley' % (n_fields, len(items))]
    return msgs, []


def ensure_mendeley_properties(pkg):
    """Mendeley Desktop reads the citation style from the custom document properties."""
    name = 'docProps/custom.xml'
    ns_cp = 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties'
    ns_vt = 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
    if pkg.has(name):
        xml = pkg.text(name)
        if 'Mendeley Citation Style_1' in xml:
            return []
        pids = [int(x) for x in re.findall(r'pid="(\d+)"', xml)] or [1]
        prop = ('<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="%d" name="Mendeley Citation Style_1">'
                '<vt:lpwstr>%s</vt:lpwstr></property>' % (max(pids) + 1, IEEE_STYLE))
        pkg.set_text(name, xml.replace('</Properties>', prop + '</Properties>'))
        return ['Gaya sitasi Mendeley (IEEE) dicatat di properti dokumen']
    xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<Properties xmlns="%s" xmlns:vt="%s"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" '
           'name="Mendeley Citation Style_1"><vt:lpwstr>%s</vt:lpwstr></property></Properties>' % (ns_cp, ns_vt, IEEE_STYLE))
    pkg.files[name] = xml.encode('utf-8')
    pkg.order.append(name)
    ct = pkg.text('[Content_Types].xml')
    if '/docProps/custom.xml' not in ct:
        ct = ct.replace('</Types>', '<Override PartName="/docProps/custom.xml" '
                        'ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/></Types>')
        pkg.set_text('[Content_Types].xml', ct)
    rels = pkg.text('_rels/.rels')
    if 'docProps/custom.xml' not in rels:
        rid = 'rIdMendeley%s' % uuid.uuid4().hex[:6]
        rels = rels.replace('</Relationships>', '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/'
                            '2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>' % rid)
        pkg.set_text('_rels/.rels', rels)
    return ['Gaya sitasi Mendeley (IEEE) dicatat di properti dokumen']
