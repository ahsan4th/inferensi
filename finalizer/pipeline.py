"""End-to-end finalisation of an accepted INFERENSI manuscript."""
from __future__ import annotations

import base64
import datetime as dt
import json
import re
from dataclasses import dataclass, field

from . import checks, fixes, letters, mathconv, mendeley, template
from .docxtools import (NS, DocxError, DocxPackage, normalize, q, sanitize_package, validate_package,  # noqa: F401
                        visible_text_length)
from .ojs import OJSInfo, parse_history


@dataclass
class Options:
    decimal_comma: bool = True
    safe_language: bool = True
    accept_changes: bool = True
    remove_comments: bool = True
    equation_layout: bool = True
    reference_format: bool = True
    spell_check: bool = True
    template_format: bool = True       # fonts, sizes, colours, justify, page sections, header/footer style
    remove_highlight: bool = True
    unbox_tables: bool = True
    mendeley_fields: bool = True       # plain-text citations -> Mendeley CSL fields
    math_convert: bool = True          # plain-text variables/equations -> Word equations (OMML)


@dataclass
class Result:
    docx: bytes = b''
    header: str = ''
    doi: str = ''
    dates: dict = field(default_factory=dict)
    changes: list = field(default_factory=list)      # (category, message)
    issues: list = field(default_factory=list)       # checks.Issue
    equations: list = field(default_factory=list)
    front: dict = field(default_factory=dict)
    ojs: OJSInfo | None = None
    ris: str = ''
    n_ris: int = 0
    report_md: str = ''
    safe_mode: bool = False
    failed_steps: list = field(default_factory=list)
    letter_info: object = None


def normalize_doi(doi: str | None) -> str:
    if not doi:
        return ''
    d = doi.strip()
    d = re.sub(r'^(https?://)?(dx\.)?doi\.org/', '', d, flags=re.I)
    d = re.sub(r'^doi:\s*', '', d, flags=re.I)
    return d


def run(docx_bytes: bytes, volume, issue, month: str, year, history_html: bytes | None = None,
        doi: str | None = None, options: Options | None = None, template_bytes: bytes | None = None) -> Result:
    """Full pipeline; falls back to a minimal, safe pass if the full output fails validation."""
    res = _run(docx_bytes, volume, issue, month, year, history_html, doi, options or Options(), safe=False,
               template_bytes=template_bytes)
    problems = _check_output(docx_bytes, res)
    if problems:
        safe = _run(docx_bytes, volume, issue, month, year, history_html, doi, options or Options(), safe=True,
                    template_bytes=template_bytes)
        safe.safe_mode = True
        safe.issues.insert(0, checks.Issue(checks.KRITIS, 'Sistem',
                                           'Hasil pemrosesan lengkap tidak lolos validasi (%s). Dokumen dibuat ulang dalam '
                                           'mode aman: hanya header, footer, tanggal, lisensi, dan metadata yang diubah.'
                                           % '; '.join(problems[:3])))
        p2 = _check_output(docx_bytes, safe)
        if p2:
            safe.docx = docx_bytes
            safe.issues.insert(0, checks.Issue(checks.KRITIS, 'Sistem', 'Mode aman juga gagal (%s); file asli dikembalikan '
                                               'tanpa perubahan.' % '; '.join(p2[:3])))
        safe.report_md = build_report(safe)
        return safe
    return res


def _check_output(original: bytes, res: Result):
    pkg_out = DocxPackage(res.docx)          # raises if unreadable
    problems = validate_package(res.docx, pkg_out.main)
    before = visible_text_length(DocxPackage(original).xml(DocxPackage(original).main))
    after = visible_text_length(pkg_out.xml(pkg_out.main))
    if before and before - after > max(400, 0.07 * before):
        problems.append('teks berkurang %.0f%%' % (100 * (1 - after / before)))
    try:
        import docx as python_docx
        python_docx.Document(__import__('io').BytesIO(res.docx))
    except ImportError:
        pass
    except Exception as exc:
        problems.append('tidak dapat dibuka python-docx: %s' % exc)
    return problems


def _run(docx_bytes, volume, issue, month, year, history_html, doi, opt: Options, safe: bool,
         template_bytes: bytes | None = None) -> Result:
    res = Result()
    pkg = DocxPackage(docx_bytes)
    root = pkg.xml(pkg.main)

    def add(cat, msgs):
        for m in msgs or []:
            if (cat, m) not in res.changes:
                res.changes.append((cat, m))

    def step(name, fn, *a, **kw):
        try:
            return fn(*a, **kw)
        except Exception as exc:      # one failing step must not stop the others
            res.failed_steps.append(name)
            res.issues.append(checks.Issue(checks.CEK, 'Sistem', 'Langkah “%s” dilewati karena error: %s: %s'
                                           % (name, type(exc).__name__, exc)))
            return None

    res.doi = normalize_doi(doi)
    res.header = fixes.header_text(volume, issue, month, year)

    add('Dokumen', step('sanitasi paket', sanitize_package, pkg))
    if opt.accept_changes and not safe:
        add('Track changes', step('track changes', fixes.accept_tracked_changes, root))
    if opt.remove_comments and not safe:
        add('Komentar', step('komentar', fixes.remove_comments, pkg, root))
    step('normalisasi run', normalize, root)

    if template_bytes:
        add('Header & footer', step('salin header/footer template', template.copy_header_footer_from, pkg, template_bytes))
    add('Header & footer', step('header/footer', fixes.fix_header_footer, pkg, volume, issue, month, year, res.doi or None))
    if opt.template_format:
        add('Header & footer', step('warna header/footer', template.fix_header_footer_style, pkg))
    if not res.doi:
        res.issues.append(checks.Issue(checks.CEK, 'DOI', 'DOI belum diisi; footer memakai placeholder '
                                       '10.12962/j27213862.v%si%s.xxxx' % (volume, issue)))

    if history_html:
        info = step('riwayat OJS', parse_history, history_html)
        if info is not None:
            res.ojs = info
            res.dates = info.dates_text()
            out = step('tanggal penting', fixes.fix_dates, root, res.dates)
            msgs, found = out if out else ([], set())
            add('Tanggal penting', msgs)
            for lab in ('Received', 'Revised', 'Accepted'):
                if lab not in found:
                    res.issues.append(checks.Issue(checks.KRITIS, 'Tanggal', 'Kotak “%s:” tidak ditemukan di halaman pertama '
                                                   '(isi manual: %s)' % (lab, res.dates.get(lab) or '—')))
                elif not res.dates.get(lab):
                    res.issues.append(checks.Issue(checks.KRITIS, 'Tanggal', 'Tanggal %s tidak dapat ditentukan dari riwayat OJS' % lab))
    else:
        res.issues.append(checks.Issue(checks.CEK, 'Tanggal', 'Riwayat OJS tidak diunggah; tanggal Received/Revised/Accepted tidak diubah'))

    add('Lisensi', step('lisensi', fixes.fix_copyright, root, year))
    if not safe:
        if opt.unbox_tables:
            add('Tabel', step('kotak di tabel', template.unbox_tables, root))
            add('Tabel', step('content control', template.unwrap_body_content_controls, root))
        if opt.math_convert:
            add('Persamaan', step('variabel ke Equation Word', mathconv.convert_all, root))
        add('Teks', step('karakter & spasi', fixes.fix_invisible_and_spaces, root))
        add('Teks', step('spasi persamaan', fixes.fix_math_spacing, root))
        if opt.decimal_comma:
            add('Angka', step('koma desimal', fixes.fix_decimal_commas, root))
        if opt.safe_language:
            add('Bahasa', step('istilah', fixes.fix_safe_language, root))
        add('Sitasi', step('sitasi', fixes.fix_citation_groups, root))
        add('Persamaan', step('persamaan', fixes.fix_equations, root, layout=opt.equation_layout,
                              text_width_twips=_text_width(root)))
        if opt.reference_format:
            add('Referensi', step('referensi', fixes.fix_references, root))
        if opt.remove_highlight:
            add('Template', step('sorotan', template.remove_highlight, root))
        if opt.template_format:
            add('Template', step('font/warna/perataan', template.format_by_template, root))
            add('Template', step('section halaman', template.fix_sections, root))
        if opt.mendeley_fields:
            out = step('field Mendeley', mendeley.convert_plain_citations, root)
            if out:
                msgs, notes = out
                add('Sitasi', msgs)
                for nt in notes:
                    res.issues.append(checks.Issue(checks.CEK, 'Sitasi', nt))
                if msgs:
                    add('Sitasi', step('properti Mendeley', mendeley.ensure_mendeley_properties, pkg))
        add('Tata letak', step('keep with next', fixes.fix_keep_with_next, root))

    # checks ------------------------------------------------------------------
    out = step('cek halaman depan', checks.check_front, root, res.ojs)
    if out:
        issues, fm = out
        res.front = fm
        res.issues += issues
    res.equations = step('daftar persamaan', checks.equations, root) or []
    res.issues += step('cek persamaan', checks.check_equations, root, res.equations) or []
    res.issues += step('cek bahasa', checks.check_language, root, use_spell=opt.spell_check) or []
    res.issues += step('cek referensi', checks.check_references, root) or []
    res.issues += step('cek gambar/tabel', checks.check_figures_tables, root) or []
    left = step('cek variabel teks', mathconv.remaining_plain_math, root) or []
    for t in left[:20]:
        res.issues.append(checks.Issue(checks.CEK, 'Persamaan', 'Variabel/simbol masih berupa teks biasa, belum Equation Word', t))
    out = step('cek urutan bagian', template.check_sections_order, root)
    if out:
        for m in out[0]:
            res.issues.append(checks.Issue(checks.CEK, 'Struktur', m))
    if not mendeley.has_citation_fields(root):
        res.issues.append(checks.Issue(checks.CEK, 'Sitasi', 'Sitasi belum berupa field Mendeley; Mendeley tidak dapat '
                                       'memperbarui sitasi dokumen ini'))

    title = res.front.get('title') or (res.ojs.title if res.ojs else '')
    add('Metadata', step('metadata', fixes.clean_metadata, pkg, title))
    from .docxtools import body_regions, ptext as _pt
    front_text = ' '.join(_pt(e) for e in body_regions(root)[0])
    res.letter_info = step('data LOA/receipt', letters.letter_defaults, res.front, front_text, res.ojs, volume, issue, year)

    out = step('ekspor RIS', export_ris, root)
    res.ris, res.n_ris = out if out else ('', 0)
    res.docx = pkg.save()
    res.report_md = build_report(res)
    return res


def _text_width(root):
    sp = root.find('w:body/w:sectPr', NS)
    try:
        pg = sp.find('w:pgSz', NS)
        mar = sp.find('w:pgMar', NS)
        return int(pg.get(q('w:w'))) - int(mar.get(q('w:left'))) - int(mar.get(q('w:right')))
    except Exception:
        return 9533


# --------------------------------------------------------------------------- RIS
def _csl_items(root):
    items, seen = [], set()

    def take(d):
        for ci in d.get('citationItems', []):
            it = ci.get('itemData') or {}
            key = (it.get('DOI') or it.get('title') or json.dumps(it, sort_keys=True)).lower()
            if key not in seen:
                seen.add(key)
                items.append(it)

    # Mendeley Cite (content controls) and Mendeley Desktop / Zotero (fields), in document order
    stack = []
    for el in root.iter(q('w:tag'), q('w:fldChar'), q('w:instrText')):
        tag = el.tag
        if tag == q('w:tag'):
            v = el.get(q('w:val'), '')
            if v.startswith('MENDELEY_CITATION_v3_'):
                b = v[len('MENDELEY_CITATION_v3_'):]
                try:
                    take(json.loads(base64.b64decode(b + '=' * (-len(b) % 4)).decode('utf-8')))
                except Exception:
                    pass
        elif tag == q('w:fldChar'):
            t = el.get(q('w:fldCharType'))
            if t == 'begin':
                stack.append({'buf': '', 'done': False})
            elif t in ('separate', 'end') and stack:
                top = stack[-1] if t == 'separate' else stack.pop()
                if not top['done']:
                    top['done'] = True
                    s = top['buf']
                    if 'CSL_CITATION' in s and '{' in s:
                        try:
                            take(json.loads(s[s.index('{'):s.rindex('}') + 1]))
                        except Exception:
                            pass
        elif stack and not stack[-1]['done']:
            stack[-1]['buf'] += el.text or ''
    return items


def export_ris(root):
    items = _csl_items(root)
    if not items:
        return '', 0
    ty = {'book': 'BOOK', 'article-journal': 'JOUR', 'paper-conference': 'CONF', 'chapter': 'CHAP',
          'thesis': 'THES', 'report': 'RPRT', 'webpage': 'ELEC', 'article': 'GEN'}
    out = []
    for n, it in enumerate(items, 1):
        rows = [('TY', ty.get(it.get('type'), 'GEN'))]
        for a in it.get('author', []):
            fam = a.get('family') or a.get('literal') or ''
            giv = a.get('given', '')
            rows.append(('AU', fam + (', ' + giv if giv else '')))
        rows.append(('TI', it.get('title', '')))
        if it.get('container-title'):
            rows.append(('T2' if rows[0][1] in ('CONF', 'CHAP') else 'JO', it['container-title']))
        y = (it.get('issued', {}).get('date-parts') or [[None]])[0][0]
        if y:
            rows.append(('PY', str(y)))
        for k, t in (('volume', 'VL'), ('issue', 'IS'), ('edition', 'ET'), ('publisher', 'PB'),
                     ('publisher-place', 'CY'), ('ISSN', 'SN'), ('ISBN', 'SN')):
            if it.get(k):
                rows.append((t, str(it[k])))
        pg = str(it.get('page', '')).replace('–', '-')
        if pg:
            if '-' in pg:
                a_, b_ = pg.split('-', 1)
                rows += [('SP', a_), ('EP', b_)]
            else:
                rows.append(('SP', pg))
        if it.get('DOI'):
            rows += [('DO', it['DOI'].strip()), ('UR', 'https://doi.org/' + it['DOI'].strip())]
        elif it.get('URL'):
            rows.append(('UR', it['URL']))
        rows.append(('ER', ''))
        out.append('\r\n'.join('%s  - %s' % r for r in rows))
    return '\r\n\r\n'.join(out) + '\r\n', len(items)


# --------------------------------------------------------------------------- report
def build_report(res: Result) -> str:
    L = ['# Laporan Finalisasi Artikel INFERENSI', '',
         ('> **Mode aman aktif** — hanya perubahan minimal yang diterapkan.\n' if res.safe_mode else ''),
         'Dibuat: %s' % dt.datetime.now().strftime('%d %B %Y %H:%M'), '',
         '**Header:** %s  ' % res.header,
         '**DOI:** %s  ' % (res.doi or '(belum diisi)')]
    if res.dates:
        L.append('**Tanggal:** Received %s · Revised %s · Accepted %s' %
                 (res.dates.get('Received') or '—', res.dates.get('Revised') or '—', res.dates.get('Accepted') or '—'))
    L += ['', '## Perbaikan otomatis (%d)' % len(res.changes), '']
    cur = None
    for cat, msg in res.changes:
        if cat != cur:
            L += ['', '### ' + cat]
            cur = cat
        L.append('- ' + msg)
    L += ['', '## Perlu ditinjau editor (%d)' % len(res.issues), '']
    for sev in (checks.KRITIS, checks.CEK, checks.INFO):
        items = [i for i in res.issues if i.severity == sev]
        if not items:
            continue
        L += ['', '### %s (%d)' % (sev, len(items))]
        for i in items:
            L.append('- **[%s]** %s%s' % (i.category, i.message, ('  \n  _%s_' % i.context) if i.context else ''))
    L += ['', '## Daftar persamaan (%d)' % len(res.equations), '', '| No. | Persamaan (teks linear) |', '|---|---|']
    for e in res.equations:
        L.append('| %s | `%s` |' % (('(%s)' % e['no']) if e['no'] else '—', e['text'].replace('|', '\\|')))
    if res.ojs:
        L += ['', '## Ringkasan OJS', '', '- Judul: %s' % res.ojs.title,
              '- Kontributor: %s' % ', '.join(n for n, _ in res.ojs.contributors),
              '- Kata kunci: %s' % (', '.join(res.ojs.keywords) or '(kosong)'),
              '- ID submission: %s' % (res.ojs.submission_id or '—')]
    L += ['', '## Sitasi Mendeley', '', ('%d referensi diekspor ke .ris.' % res.n_ris) if res.n_ris else
          'Tidak ada field Mendeley/Zotero di naskah; file .ris tidak dibuat.']
    return '\n'.join(L) + '\n'
