"""Letter of Acceptance (LOA) and payment receipt for an accepted INFERENSI article.

Both documents are produced from the journal's own Word templates (templates/LOA.docx and templates/Receipt.docx):
dates and identities are replaced in place, the rest of the layout (letterhead, signature, logos) is kept, and the
result is converted to PDF with LibreOffice when it is available.
"""
from __future__ import annotations

import datetime as dt
import os
import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass, field

from .docxtools import NS, WP, DocxPackage, get_t, inner_paras, normalize, ptext, q, regex_sub_para, replace_span, segments, set_t
from .ojs import MONTHS, fmt_date

MONTH_RX = '(?:%s)' % '|'.join(MONTHS)
EMAIL_RX = r'[\w.+-]+@[\w-]+(?:\.[\w-]+)+'
DEFAULT_AMOUNT = 'IDR 1.500.000'


def us_date(d: dt.date) -> str:
    return '%s %d, %d' % (MONTHS[d.month - 1], d.day, d.year)


def today_wib() -> dt.date:
    try:
        from zoneinfo import ZoneInfo
        return dt.datetime.now(ZoneInfo('Asia/Jakarta')).date()
    except Exception:
        return dt.date.today()


@dataclass
class LetterInfo:
    name: str = ''
    email: str = ''
    title: str = ''
    article_id: str = ''
    volume: str = ''
    issue: str = ''
    year: str = ''
    loa_date: dt.date = field(default_factory=today_wib)
    payment_date: dt.date = field(default_factory=today_wib)
    receipt_date: dt.date = field(default_factory=today_wib)
    amount: str = DEFAULT_AMOUNT
    notes: list = field(default_factory=list)


# --------------------------------------------------------------------------- defaults from the article and OJS
def letter_defaults(front: dict, front_text: str, ojs, volume, issue, year) -> LetterInfo:
    info = LetterInfo(volume=str(volume), issue=str(issue), year=str(year))
    info.title = (front or {}).get('title') or (ojs.title if ojs else '')
    m = re.search(r'corresponding author\s*:?\s*(%s)' % EMAIL_RX, front_text or '', re.I)
    email = m.group(1) if m else ''
    contributors = list(ojs.contributors) if ojs else []
    if email:
        info.email = email
        hit = next((n for n, e in contributors if e.lower() == email.lower()), None)
        if hit:
            info.name = hit
        else:
            info.notes.append('Email penulis korespondensi tidak ada di daftar kontributor OJS; periksa nama penerima.')
    if not info.name and ojs:
        submitter = next((e[1] for e in sorted(ojs.events, key=lambda e: e[0])
                          if 'submission completed' in e[2].lower()), None)
        hit = next(((n, e) for n, e in contributors if submitter and n.lower() == submitter.lower()), None)
        if hit is None and contributors:
            hit = contributors[0]
            info.notes.append('Penerima LOA/receipt diambil dari kontributor pertama di OJS; periksa kembali.')
        if hit:
            info.name = hit[0]
            info.email = info.email or hit[1]
    if ojs:
        info.article_id = ojs.submission_id or ''
        if ojs.accepted:
            info.loa_date = ojs.accepted
        else:
            info.notes.append('Tanggal keputusan Accept tidak ditemukan; tanggal LOA memakai hari ini.')
        authors = {n.lower() for n, _ in contributors}
        uploads = [e[0] for e in ojs.events if ojs.accepted and e[0] >= ojs.accepted and 'upload' in e[2].lower()
                   and e[1].lower() in authors]
        if uploads:
            info.payment_date = max(uploads)
        else:
            info.notes.append('Tanggal pembayaran tidak terdeteksi dari riwayat OJS (unggahan bukti bayar oleh penulis '
                              'setelah Accept); memakai hari ini — sesuaikan bila perlu.')
    if not info.article_id:
        info.notes.append('Article ID (nomor submission OJS) tidak ditemukan; isi manual.')
    return info


# --------------------------------------------------------------------------- filling
def _rewrite(p, text):
    runs = [r for r in p.findall('w:r', NS) if r.find('w:t', NS) is not None]
    if not runs:
        return False
    set_t(runs[0], text)
    for r in runs[1:]:
        p.remove(r)
    return True


def _fix_space_padding(pkg):
    """Text pushed to the right with dozens of spaces moves when fonts differ; use a right tab stop instead."""
    from .docxtools import ppr_set
    root = pkg.xml(pkg.main)
    sp = root.find('w:body/w:sectPr', NS)
    try:
        width = int(sp.find('w:pgSz', NS).get(q('w:w'))) - int(sp.find('w:pgMar', NS).get(q('w:left'))) - \
            int(sp.find('w:pgMar', NS).get(q('w:right')))
    except Exception:
        width = 9000
    for name in [pkg.main] + pkg.parts(r'word/(header|footer)\d*\.xml'):
        part = pkg.xml(name)
        for p in part.iter(WP):
            runs = p.findall('w:r', NS)
            for i, r in enumerate(runs):
                t = get_t(r)
                if len(t) >= 15 and not t.strip() and any(get_t(x).strip() for x in runs[i + 1:]):
                    p.remove(r)
                    if not any(x.find('w:tab', NS) is not None for x in runs[:i]):
                        tab = r.makeelement(q('w:r'))
                        tab.append(tab.makeelement(q('w:tab')))
                        runs[i + 1].addprevious(tab)
                    tabs = ppr_set(p, 'tabs')
                    for pos in sorted(set(_style_tabs(pkg, p)) - {str(width)}):
                        c_ = tabs.makeelement(q('w:tab'))
                        c_.set(q('w:val'), 'clear')
                        c_.set(q('w:pos'), pos)
                        tabs.append(c_)
                    t_ = tabs.makeelement(q('w:tab'))
                    t_.set(q('w:val'), 'right')
                    t_.set(q('w:pos'), str(width))
                    tabs.append(t_)
                    break


def _style_tabs(pkg, p):
    """Tab stops inherited from the paragraph style (e.g. the centre tab of the “Header” style)."""
    st = p.find('w:pPr/w:pStyle', NS)
    if st is None or not pkg.has('word/styles.xml'):
        return ['4513', '4680']
    sid = st.get(q('w:val'))
    styles = pkg.xml('word/styles.xml')
    out, guard = [], 0
    while sid and guard < 10:
        guard += 1
        el = next((x for x in styles.findall('w:style', NS) if x.get(q('w:styleId')) == sid), None)
        if el is None:
            break
        out += [t.get(q('w:pos')) for t in el.findall('w:pPr/w:tabs/w:tab', NS) if t.get(q('w:val')) != 'clear']
        b = el.find('w:basedOn', NS)
        sid = b.get(q('w:val')) if b is not None else None
    return out or ['4513', '4680']


def _clean(pkg, title):
    _fix_space_padding(pkg)
    from .fixes import clean_metadata
    clean_metadata(pkg, title)


def fill_loa(template: bytes, info: LetterInfo):
    pkg = DocxPackage(template)
    root = pkg.xml(pkg.main)
    normalize(root)
    done = set()
    for p in inner_paras(root):
        t = ptext(p).strip()
        if re.fullmatch(r'%s \d{1,2}, \d{4}' % MONTH_RX, t) or re.fullmatch(r'\d{1,2} %s \d{4}' % MONTH_RX, t):
            _rewrite(p, us_date(info.loa_date))
            done.add('date')
        elif t.startswith('Dear '):
            _rewrite(p, 'Dear %s,' % info.name)
            done.add('name')
        elif re.search(r'manuscript\s+titled', t):
            for seg in list(segments(p)):
                full = ''.join(get_t(r) for r in seg)
                m = re.search(r'(?<=[“"])(\s*)([^“”"]+?)(\s*)(?=[”"])', full)
                if m:
                    replace_span(seg, m.start(3), m.end(3), '')
                    replace_span(seg, m.start(2), m.end(2), info.title)
                    replace_span(seg, m.start(1), m.end(1), '')
                    done.add('title')
                    break
            if re.search(r'Volume\s+\d+\s+Number\s+\d+\s+\d{4}', t):
                regex_sub_para(p, r'Volume\s+\d+\s+Number\s+\d+\s+\d{4}',
                               'Volume %s Number %s %s' % (info.volume, info.issue, info.year))
                done.add('volume')
    _clean(pkg, 'Letter of Acceptance')
    missing = {'date', 'name', 'title', 'volume'} - done
    return pkg.save(), ['Bagian LOA tidak ditemukan di template: %s' % ', '.join(sorted(missing))] if missing else []


def fill_receipt(template: bytes, info: LetterInfo):
    pkg = DocxPackage(template)
    root = pkg.xml(pkg.main)
    normalize(root)
    done = set()
    for p in inner_paras(root):
        t = ptext(p).strip()
        if t.startswith('Surabaya,'):
            if re.search(r'\d{1,2}\s+%s\s+\d{4}' % MONTH_RX, t):
                regex_sub_para(p, r'\d{1,2}\s+%s\s+\d{4}' % MONTH_RX, fmt_date(info.receipt_date))
                done.add('receipt date')
        elif re.fullmatch(EMAIL_RX, t):
            _rewrite(p, info.email)
            prev = p.getprevious()
            while prev is not None and (prev.tag != WP or not ptext(prev).strip()):
                prev = prev.getprevious()
            if prev is not None:
                _rewrite(prev, info.name)
            done.add('name')
        elif t.startswith('Article ID'):
            regex_sub_para(p, r'(?<=Article ID:)\s*\S*\s*$', ' %s' % info.article_id) or \
                regex_sub_para(p, r'\s*\d+\s*$', ' %s' % info.article_id)
            done.add('article id')
        elif t.startswith('Manuscript Title'):
            for seg in list(segments(p)):
                full = ''.join(get_t(r) for r in seg)
                m = re.search(r'Manuscript Title:\s*(.*)$', full, re.S)
                if m:
                    replace_span(seg, m.start(1), m.end(1), info.title)
                    done.add('title')
                    break
            _autofit_textbox(p)
        elif t.startswith('Amount Received'):
            if re.search(r'(?:IDR|Rp\.?)\s*[\d.,]+', t):
                regex_sub_para(p, r'(?:IDR|Rp\.?)\s*[\d.,]+', info.amount)
                done.add('amount')
        elif t.startswith('Date') and 'Received' in t:
            if re.search(r'\d{1,2}\s+%s\s+\d{4}' % MONTH_RX, t):
                regex_sub_para(p, r'\d{1,2}\s+%s\s+\d{4}' % MONTH_RX, fmt_date(info.payment_date))
                done.add('payment date')
    _clean(pkg, 'Payment Receipt')
    missing = {'receipt date', 'name', 'article id', 'title', 'amount', 'payment date'} - done
    return pkg.save(), ['Bagian receipt tidak ditemukan di template: %s' % ', '.join(sorted(missing))] if missing else []


def _autofit_textbox(p):
    """Let the text box that holds a long manuscript title grow instead of clipping it."""
    for a in p.iterancestors():
        if a.tag.endswith('}wsp'):
            bp = a.find('{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}bodyPr')
            if bp is not None:
                for c in list(bp):
                    if c.tag.endswith(('}noAutofit', '}normAutofit', '}spAutoFit')):
                        bp.remove(c)
                bp.append(bp.makeelement('{http://schemas.openxmlformats.org/drawingml/2006/main}spAutoFit'))
            break
        if a.tag == '{urn:schemas-microsoft-com:vml}textbox':
            st = a.get('style') or ''
            if 'mso-fit-shape-to-text' not in st:
                a.set('style', (st + ';' if st else '') + 'mso-fit-shape-to-text:t')
            break


# --------------------------------------------------------------------------- PDF
def soffice_path():
    for name in ('soffice', 'libreoffice'):
        p = shutil.which(name)
        if p:
            return p
    for p in ('/usr/bin/soffice', '/usr/lib/libreoffice/program/soffice', '/opt/libreoffice/program/soffice',
              '/Applications/LibreOffice.app/Contents/MacOS/soffice', r'C:\Program Files\LibreOffice\program\soffice.exe'):
        if os.path.exists(p):
            return p
    return None


def docx_to_pdf(data: bytes, timeout: int = 180):
    exe = soffice_path()
    if not exe:
        return None
    with tempfile.TemporaryDirectory() as d:
        src = os.path.join(d, 'letter.docx')
        with open(src, 'wb') as f:
            f.write(data)
        env = dict(os.environ, HOME=d)
        cmd = [exe, '-env:UserInstallation=file://%s/profile' % d.replace('\\', '/'), '--headless', '--norestore',
               '--nolockcheck', '--convert-to', 'pdf', '--outdir', d, src]
        try:
            subprocess.run(cmd, timeout=timeout, env=env, capture_output=True)
        except Exception:
            return None
        out = os.path.join(d, 'letter.pdf')
        if os.path.exists(out) and os.path.getsize(out) > 0:
            with open(out, 'rb') as f:
                return f.read()
    return None


def make_letters(info: LetterInfo, loa_template: bytes | None, receipt_template: bytes | None, pdf: bool = True):
    """Return (files {name: bytes}, messages). PDF when LibreOffice is available, otherwise DOCX."""
    files, msgs = {}, []
    for kind, tpl, fn in (('LOA', loa_template, fill_loa), ('Receipt', receipt_template, fill_receipt)):
        if not tpl:
            msgs.append('Template %s tidak tersedia; dokumen %s tidak dibuat.' % (kind, kind))
            continue
        try:
            docx, notes = fn(tpl, info)
        except Exception as exc:
            msgs.append('%s gagal dibuat: %s' % (kind, exc))
            continue
        msgs += notes
        files['%s.docx' % kind] = docx
        if pdf:
            p = docx_to_pdf(docx)
            if p:
                files['%s.pdf' % kind] = p
            else:
                msgs.append('LibreOffice tidak tersedia, sehingga %s hanya tersedia dalam format Word (.docx).' % kind)
    return files, msgs


def load_default_templates(base_dir: str):
    out = {}
    for kind in ('LOA', 'Receipt'):
        path = os.path.join(base_dir, 'templates', '%s.docx' % kind)
        if os.path.exists(path):
            with open(path, 'rb') as f:
                out[kind] = f.read()
    return out


__all__ = ['LetterInfo', 'letter_defaults', 'make_letters', 'fill_loa', 'fill_receipt', 'docx_to_pdf', 'us_date',
           'load_default_templates', 'q']
