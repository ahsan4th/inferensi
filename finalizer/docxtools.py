"""Low-level helpers to read, edit and write .docx packages with lxml.

Everything works on the raw OOXML so that Mendeley fields, equations (OMML),
text boxes and headers/footers survive untouched unless a fixer edits them.
"""
from __future__ import annotations

import copy
import io
import re
import zipfile

from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
NS = {
    'w': W,
    'm': M,
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'v': 'urn:schemas-microsoft-com:vml',
    'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    'dc': 'http://purl.org/dc/elements/1.1/',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
}
XMLSPACE = '{http://www.w3.org/XML/1998/namespace}space'


def q(tag: str) -> str:
    prefix, local = tag.split(':')
    return '{%s}%s' % (NS[prefix], local)


def local(el) -> str:
    return etree.QName(el).localname


WR, WT, WP, WPPR, WRPR = q('w:r'), q('w:t'), q('w:p'), q('w:pPr'), q('w:rPr')
NONBREAK = {q('w:pPr'), q('w:bookmarkStart'), q('w:bookmarkEnd'), q('w:proofErr'),
            q('w:permStart'), q('w:permEnd')}


# --------------------------------------------------------------------------- package
class DocxError(Exception):
    """Input that cannot be processed, with a message meant for the end user."""


MAX_UNCOMPRESSED = 600 * 1024 * 1024
STRICT_NS = b'http://purl.oclc.org/ooxml/wordprocessingml/main'
DOCX_MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'


class DocxPackage:
    """A .docx kept in memory; XML parts are parsed lazily and written back on save()."""

    def __init__(self, data: bytes):
        if not data:
            raise DocxError('File kosong.')
        if data[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':
            raise DocxError('File ini berformat Word 97–2003 (.doc) atau dokumen Word yang dikunci password. '
                            'Buka di Microsoft Word, hapus password bila ada, lalu simpan sebagai “Word Document (*.docx)”.')
        if data[:2] != b'PK':
            raise DocxError('File bukan dokumen .docx yang valid (bukan arsip Office Open XML).')
        try:
            zin = zipfile.ZipFile(io.BytesIO(data))
        except zipfile.BadZipFile as exc:
            raise DocxError('File .docx rusak atau tidak lengkap (%s).' % exc)
        infos = [i for i in zin.infolist() if not i.is_dir() and not i.filename.endswith('/')]
        if sum(i.file_size for i in infos) > MAX_UNCOMPRESSED:
            raise DocxError('Isi file terlalu besar setelah diekstrak (> 600 MB).')
        self.order = [i.filename for i in infos]
        try:
            self.files = {n: zin.read(n) for n in self.order}
        except Exception as exc:
            raise DocxError('Sebagian isi file .docx tidak dapat dibaca (%s).' % exc)
        self.trees: dict[str, etree._ElementTree] = {}
        self.main = self._find_main()
        if STRICT_NS in self.files[self.main][:5000]:
            raise DocxError('Dokumen disimpan dalam format “Strict Open XML”. Buka di Word lalu “Save As” → '
                            '“Word Document (*.docx)” (bukan Strict), kemudian unggah ulang.')
        folder, base = self.main.rsplit('/', 1) if '/' in self.main else ('', self.main)
        self.main_rels = (folder + '/' if folder else '') + '_rels/' + base + '.rels'
        try:
            self.xml(self.main)
        except etree.XMLSyntaxError as exc:
            raise DocxError('XML dokumen utama rusak dan tidak dapat dibaca (%s).' % exc)

    def _find_main(self) -> str:
        rels = self.files.get('_rels/.rels')
        if rels:
            m = re.search(rb'Type="[^"]*/officeDocument"[^>]*Target="/?([^"]+)"', rels) or \
                re.search(rb'Target="/?([^"]+)"[^>]*Type="[^"]*/officeDocument"', rels)
            if m and m.group(1).decode() in self.files:
                return m.group(1).decode()
        if 'word/document.xml' in self.files:
            return 'word/document.xml'
        cand = [n for n in self.order if re.fullmatch(r'word/document\d*\.xml', n)]
        if cand:
            return cand[0]
        raise DocxError('Bagian utama dokumen (word/document.xml) tidak ditemukan; file bukan dokumen Word.')

    def has(self, name: str) -> bool:
        return name in self.files

    def xml(self, name: str):
        if name not in self.trees:
            parser = etree.XMLParser(remove_blank_text=False, huge_tree=True, resolve_entities=False, no_network=True)
            self.trees[name] = etree.ElementTree(etree.fromstring(self.files[name], parser))
        return self.trees[name].getroot()

    def parts(self, pattern: str):
        rx = re.compile(pattern)
        return [n for n in self.order if rx.fullmatch(n) and n in self.files]

    def remove(self, name: str):
        self.files.pop(name, None)
        self.trees.pop(name, None)
        if name in self.order:
            self.order.remove(name)

    def text(self, name: str) -> str:
        if name in self.trees:
            self.files[name] = etree.tostring(self.trees.pop(name), xml_declaration=True, encoding='UTF-8',
                                              standalone=True)
        b = self.files[name]
        if b[:2] in (b'\xff\xfe', b'\xfe\xff'):
            return b.decode('utf-16')
        return b.decode('utf-8-sig', errors='replace')

    def set_text(self, name: str, s: str):
        self.trees.pop(name, None)
        s = re.sub(r'^<\?xml[^>]*encoding="[^"]*"[^>]*\?>',
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', s, count=1)
        self.files[name] = s.encode('utf-8')

    def save(self) -> bytes:
        for name, tree in self.trees.items():
            self.files[name] = etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone=True)
        out = io.BytesIO()
        names = [n for n in self.order if n in self.files]
        if '[Content_Types].xml' in names:
            names.remove('[Content_Types].xml')
            names.insert(0, '[Content_Types].xml')
        with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
            for n in names:      # file entries only: Word may refuse archives with folder entries
                z.writestr(n, self.files[n])
        return out.getvalue()


def sanitize_package(pkg: DocxPackage):
    """Turn macro-enabled documents or templates into a plain .docx; returns log messages."""
    log = []
    ct = pkg.text('[Content_Types].xml')
    new = re.sub(r'application/vnd\.ms-word\.(document|template)\.macroEnabled(\.12|\.main\+xml)', DOCX_MAIN, ct)
    new = new.replace('application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml', DOCX_MAIN)
    if new != ct:
        log.append('Dokumen makro/template (.docm/.dotx/.dotm) diubah menjadi dokumen Word biasa (.docx)')
    vba = [n for n in pkg.order if n.lower().endswith(('vbaproject.bin', 'vbadata.xml'))]
    for n in vba:
        pkg.remove(n)
    if vba:
        new = re.sub(r'<Override [^>]*PartName="/[^"]*(?:vbaProject\.bin|vbaData\.xml)"[^>]*/>', '', new, flags=re.I)
        new = re.sub(r'<Default [^>]*Extension="bin"[^>]*ContentType="application/vnd\.ms-office\.vbaProject"[^>]*/>', '', new)
        if pkg.has(pkg.main_rels):
            rels = pkg.text(pkg.main_rels)
            pkg.set_text(pkg.main_rels, re.sub(r'<Relationship [^>]*Target="[^"]*vba[^"]*"[^>]*/>', '', rels, flags=re.I))
        log.append('Makro VBA dihapus dari dokumen')
    # every file extension in the package must have a content type (LibreOffice exports sometimes miss one)
    declared = set(e.lower() for e in re.findall(r'<Default [^>]*Extension="([^"]+)"', new))
    overrides = set(o.lstrip('/') for o in re.findall(r'<Override [^>]*PartName="([^"]+)"', new))
    added = []
    for n in pkg.order:
        if n == '[Content_Types].xml' or n in overrides or '.' not in n.rsplit('/', 1)[-1]:
            continue
        ext = n.rsplit('.', 1)[-1].lower()
        if ext not in declared and ext in CONTENT_TYPES:
            new = new.replace('</Types>', '<Default Extension="%s" ContentType="%s"/></Types>' % (ext, CONTENT_TYPES[ext]))
            declared.add(ext)
            added.append(ext)
    if added:
        log.append('Deklarasi tipe file yang hilang ditambahkan ke paket Word: %s' % ', '.join(added))
    if new != ct:
        pkg.set_text('[Content_Types].xml', new)
    return log


CONTENT_TYPES = {
    'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg', 'gif': 'image/gif', 'bmp': 'image/bmp',
    'tif': 'image/tiff', 'tiff': 'image/tiff', 'emf': 'image/x-emf', 'wmf': 'image/x-wmf', 'svg': 'image/svg+xml',
    'webp': 'image/webp', 'ico': 'image/x-icon', 'wdp': 'image/vnd.ms-photo',
    'bin': 'application/vnd.openxmlformats-officedocument.oleObject',
    'xml': 'application/xml', 'rels': 'application/vnd.openxmlformats-package.relationships+xml',
    'vml': 'application/vnd.openxmlformats-officedocument.vmlDrawing',
    'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    'odttf': 'application/vnd.openxmlformats-officedocument.obfuscatedFont', 'ttf': 'application/x-font-ttf',
    'fntdata': 'application/x-fontdata', 'pdf': 'application/pdf', 'txt': 'text/plain',
}


def validate_package(data: bytes, main: str):
    """Re-open a produced .docx: every XML part must parse and the main part must contain a body."""
    problems = []
    try:
        z = zipfile.ZipFile(io.BytesIO(data))
    except Exception as exc:
        return ['arsip tidak valid: %s' % exc]
    names = set(z.namelist())
    if '[Content_Types].xml' not in names:
        problems.append('[Content_Types].xml hilang')
    for n in names:
        if n.endswith('.xml') or n.endswith('.rels'):
            try:
                etree.fromstring(z.read(n), etree.XMLParser(huge_tree=True, resolve_entities=False))
            except Exception as exc:
                problems.append('%s tidak dapat diparse: %s' % (n, exc))
    for n in [x for x in names if x.endswith('.rels')]:
        base = n.replace('_rels/', '').rsplit('.rels', 1)[0]
        folder = base.rsplit('/', 1)[0] + '/' if '/' in base else ''
        for m in re.finditer(rb'<Relationship [^>]*>', z.read(n)):
            tag = m.group(0)
            if b'TargetMode="External"' in tag:
                continue
            t = re.search(rb'Target="([^"]+)"', tag)
            if not t:
                continue
            tgt = t.group(1).decode()
            path = tgt.lstrip('/') if tgt.startswith('/') else _join(folder, tgt)
            if path not in names:
                problems.append('relasi %s → %s menunjuk bagian yang tidak ada' % (n, tgt))
    try:
        root = etree.fromstring(z.read(main), etree.XMLParser(huge_tree=True))
        if root.find('w:body', NS) is None:
            problems.append('body dokumen hilang')
    except Exception as exc:
        problems.append('dokumen utama rusak: %s' % exc)
    return problems


def _join(folder, rel):
    parts = (folder + rel).split('/')
    out = []
    for p in parts:
        if p == '..':
            if out:
                out.pop()
        elif p and p != '.':
            out.append(p)
    return '/'.join(out)


def visible_text_length(root) -> int:
    return sum(len(t.text or '') for t in root.iter(q('w:t')))


# --------------------------------------------------------------------------- runs & text
def get_t(r) -> str:
    return ''.join(t.text or '' for t in r.findall(WT))


def set_t(r, text: str):
    ts = r.findall(WT)
    if not ts:
        t = etree.SubElement(r, WT)
        ts = [t]
    ts[0].text = text
    ts[0].set(XMLSPACE, 'preserve')
    for t in ts[1:]:
        r.remove(t)


def is_text_run(r) -> bool:
    if r.tag != WR:
        return False
    return all(c.tag in (WRPR, WT) for c in r)


def segments(p):
    """Yield lists of consecutive plain text runs of a paragraph (math, fields, drawings break segments)."""
    seg = []
    for c in p:
        if is_text_run(c):
            seg.append(c)
        elif c.tag in NONBREAK:
            continue
        else:
            if seg:
                yield seg
            seg = []
    if seg:
        yield seg


def ptext(el) -> str:
    return ''.join(el.xpath('.//w:t/text()', namespaces=NS))


def inner_paras(root, include_fallback: bool = True):
    """Paragraphs that do not themselves contain paragraphs (i.e. text-box content is reached)."""
    out = []
    for p in root.iter(WP):
        if p.find('.//w:p', NS) is not None:
            continue
        if not include_fallback and p.xpath('ancestor::mc:Fallback', namespaces=NS):
            continue
        out.append(p)
    return out


def replace_span(seg, start: int, end: int, new: str):
    """Replace characters [start, end) of the concatenated segment text; formatting of the first run is kept."""
    texts = [get_t(r) for r in seg]
    pos = 0
    first = None
    emptied = []
    for k, r in enumerate(seg):
        a, b = pos, pos + len(texts[k])
        pos = b
        if a == b or b <= start or a >= end:
            if not (a == b == start and first is None and start == end):
                continue
        if first is None:
            first = k
            set_t(r, texts[k][:start - a] + new + (texts[k][end - a:] if end <= b else ''))
        else:
            set_t(r, texts[k][end - a:] if end <= b else '')
            emptied.append(r)
    for r in emptied:
        if get_t(r) == '' and r.getparent() is not None:
            r.getparent().remove(r)
    return first is not None


def rep(root, old: str, new: str, count: int | None = None) -> int:
    """Literal replacement across runs inside paragraphs; returns the number of replacements."""
    n = 0
    for p in list(root.iter(WP)):
        guard = 0
        while guard < 500:
            guard += 1
            hit = False
            for seg in segments(p):
                full = ''.join(get_t(r) for r in seg)
                i = full.find(old)
                if i >= 0:
                    replace_span(seg, i, i + len(old), new)
                    hit = True
                    n += 1
                    break
            if not hit or old in new:
                break
    if count is not None and n != count:
        raise ValueError('rep %r: expected %d, got %d' % (old[:50], count, n))
    return n


def regex_sub_para(p, pattern, repl, flags=0):
    """Regex substitution inside each text segment of a paragraph; returns list of (old, new)."""
    rx = re.compile(pattern, flags) if isinstance(pattern, str) else pattern
    changes = []
    for seg in list(segments(p)):
        full = ''.join(get_t(r) for r in seg)
        matches = list(rx.finditer(full))
        for m in reversed(matches):
            new = m.expand(repl) if isinstance(repl, str) else repl(m)
            if new != m.group(0):
                replace_span(seg, m.start(), m.end(), new)
                changes.append((m.group(0), new))
    return list(reversed(changes))


def normalize(root):
    """Remove render hints, split runs mixing text and field codes, merge identical adjacent runs."""
    for tag in ('w:lastRenderedPageBreak', 'w:proofErr'):
        for el in root.xpath('//' + tag, namespaces=NS):
            el.getparent().remove(el)
    for r in list(root.iter(WR)):
        kids = [x for x in r if x.tag != WRPR]
        if len(kids) <= 1 or all(x.tag == WT for x in kids):
            continue
        rpr = r.find(WRPR)
        anchor = r
        for x in kids:
            nr = etree.Element(WR)
            if rpr is not None:
                nr.append(copy.deepcopy(rpr))
            nr.append(x)
            anchor.addnext(nr)
            anchor = nr
        r.getparent().remove(r)
    for p in root.iter(WP):
        prev = None
        for c in list(p):
            if is_text_run(c):
                if prev is not None and _same_rpr(prev, c):
                    set_t(prev, get_t(prev) + get_t(c))
                    p.remove(c)
                    continue
                prev = c
            elif c.tag in NONBREAK:
                continue
            else:
                prev = None


def _same_rpr(a, b) -> bool:
    ra, rb = a.find(WRPR), b.find(WRPR)
    sa = etree.tostring(ra) if ra is not None else b''
    sb = etree.tostring(rb) if rb is not None else b''
    return sa == sb


def new_run_like(r, text: str):
    nr = copy.deepcopy(r)
    for x in list(nr):
        if x.tag != WRPR:
            nr.remove(x)
    set_t(nr, text)
    return nr


PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
             'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
             'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
             'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc',
             'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
             'sectPr', 'pPrChange']


def ppr_set(p, name: str, attrs: dict | None = None):
    ppr = p.find(WPPR)
    if ppr is None:
        ppr = etree.Element(WPPR)
        p.insert(0, ppr)
    old = ppr.find(q('w:' + name))
    if old is not None:
        ppr.remove(old)
    el = etree.Element(q('w:' + name))
    for k, v in (attrs or {}).items():
        el.set(q('w:' + k), v)
    order = PPR_ORDER.index(name)
    for c in ppr:
        cn = local(c)
        if cn in PPR_ORDER and PPR_ORDER.index(cn) > order:
            c.addprevious(el)
            return el
    ppr.append(el)
    return el


def body_regions(root):
    """Return (front, main, refs) lists of top-level body elements.

    front = everything before the INTRODUCTION heading, refs = the reference heading and after.
    """
    body = root.find('w:body', NS)
    kids = [k for k in body if local(k) != 'sectPr']
    intro = ref = None
    for i, k in enumerate(kids):
        t = ptext(k).strip()
        if intro is None and re.match(r'^(I\.?\s*)?(INTRODUCTION|PENDAHULUAN)$', t, re.I):
            intro = i
        if re.match(r'^(REFERENCES|DAFTAR PUSTAKA|BIBLIOGRAPHY)$', t, re.I):
            ref = i
    intro = intro or 0
    ref = ref if ref is not None else len(kids)
    return kids[:intro], kids[intro:ref], kids[ref:]


def paras_of(elements, include_fallback=False):
    """All paragraphs inside the given top-level elements (text-box paragraphs included)."""
    out = []
    for e in elements:
        cand = [e] if e.tag == WP else []
        cand += [p for p in e.iter(WP) if p is not e]
        for p in cand:
            if include_fallback or not p.xpath('ancestor::mc:Fallback', namespaces=NS):
                out.append(p)
    return out
