"""Read-only checks. Each returns a list of Issue objects for the editor to review."""
from __future__ import annotations

import difflib
import re
from collections import Counter
from dataclasses import dataclass

from .docxtools import NS, WP, body_regions, get_t, inner_paras, local, paras_of, ptext, q  # noqa: F401
from .fixes import reference_entries

KRITIS, CEK, INFO = 'Kritis', 'Perlu dicek', 'Info'


@dataclass
class Issue:
    severity: str
    category: str
    message: str
    context: str = ''


def _ctx(text, i, j, w=55):
    return ('…' if i > w else '') + text[max(0, i - w):j + w].replace('\n', ' ') + ('…' if j + w < len(text) else '')


def _visible_text(p):
    """Paragraph text with inline math replaced by a placeholder."""
    out = []
    for c in p.iter(q('w:t'), q('m:oMath')):
        if next(c.iterancestors(WP), None) is not p:
            continue          # belongs to a nested text-box paragraph
        if c.tag == q('w:t') and not any(local(a) == 'oMath' for a in c.iterancestors()):
            out.append(c.text or '')
        elif c.tag == q('m:oMath') and not any(local(a) == 'oMath' for a in c.iterancestors()):
            out.append(' ⟨math⟩ ')
    return ''.join(out)


# --------------------------------------------------------------------------- OMML → linear text
def omml_linear(e) -> str:
    tag = local(e)
    kids = [k for k in e if not local(k).endswith('Pr')]
    if tag == 'r':
        return ''.join(t.text or '' for t in e.iter(q('m:t')))
    if tag in ('sSub', 'sSup', 'sSubSup', 'sPre'):
        base = omml_linear(e.find('m:e', NS)) if e.find('m:e', NS) is not None else ''
        sub = e.find('m:sub', NS)
        sup = e.find('m:sup', NS)
        s = ('_{%s}' % omml_linear(sub)) if sub is not None else ''
        p = ('^{%s}' % omml_linear(sup)) if sup is not None else ''
        return (s + p + base) if tag == 'sPre' else base + s + p
    if tag == 'f':
        return '(%s)/(%s)' % (omml_linear(e.find('m:num', NS)), omml_linear(e.find('m:den', NS)))
    if tag == 'rad':
        deg = e.find('m:deg', NS)
        dg = omml_linear(deg) if deg is not None else ''
        return ('root_{%s}' % dg if dg else 'sqrt') + '(%s)' % omml_linear(e.find('m:e', NS))
    if tag == 'nary':
        pr = e.find('m:naryPr', NS)
        ch = pr.find('m:chr', NS).get(q('m:val')) if pr is not None and pr.find('m:chr', NS) is not None else '∫'
        sub, sup = e.find('m:sub', NS), e.find('m:sup', NS)
        s = '_{%s}' % omml_linear(sub) if sub is not None and omml_linear(sub).strip('\u200b') else ''
        p = '^{%s}' % omml_linear(sup) if sup is not None and omml_linear(sup).strip('\u200b') else ''
        return '%s%s%s %s' % (ch, s, p, omml_linear(e.find('m:e', NS)))
    if tag == 'd':
        pr = e.find('m:dPr', NS)

        def val(n, default):
            el = pr.find('m:' + n, NS) if pr is not None else None
            return el.get(q('m:val'), '') if el is not None else default
        es = [omml_linear(x) for x in e.findall('m:e', NS)]
        return val('begChr', '(') + val('sepChr', '|').join(es) + val('endChr', ')')
    if tag == 'acc':
        pr = e.find('m:accPr', NS)
        ch = pr.find('m:chr', NS).get(q('m:val')) if pr is not None and pr.find('m:chr', NS) is not None else '\u0302'
        return omml_linear(e.find('m:e', NS)) + ch
    if tag == 'bar':
        return 'bar(%s)' % omml_linear(e.find('m:e', NS))
    if tag == 'func':
        return omml_linear(e.find('m:fName', NS)) + omml_linear(e.find('m:e', NS))
    if tag == 'm':
        return '[' + '; '.join(', '.join(omml_linear(x) for x in mr.findall('m:e', NS)) for mr in e.findall('m:mr', NS)) + ']'
    if tag == 'eqArr':
        return ' ; '.join(omml_linear(x) for x in e.findall('m:e', NS))
    return ''.join(omml_linear(k) for k in kids)


def equations(root):
    """List display equations: dicts {no, text, where}."""
    out = []
    seen = set()
    for p in inner_paras(root, include_fallback=False):
        oms = p.findall('.//m:oMath', NS)
        oms = [o for o in oms if not any(local(a) == 'oMath' for a in o.iterancestors())]
        if not oms:
            continue
        outside = ''.join(get_t(r) for r in p.findall('w:r', NS)).strip()
        num = None
        m = re.fullmatch(r'\(?(\d+)\)?', outside)
        if m:
            num = m.group(1)
        elif outside:
            continue   # inline math inside prose
        if num is None:
            tr = next((a for a in p.iterancestors() if local(a) == 'tr'), None)
            if tr is not None:
                for tc in tr.findall('w:tc', NS):
                    mm = re.fullmatch(r'\s*\((\d+)\)\s*', ptext(tc))
                    if mm:
                        num = mm.group(1)
        txt = '  '.join(omml_linear(o) for o in oms)
        key = (num, txt)
        if key in seen:
            continue
        seen.add(key)
        out.append({'no': num, 'text': txt})
    return out


def check_equations(root, eqs):
    issues = []
    nums = [int(e['no']) for e in eqs if e['no'] and e['no'].isdigit()]
    nums_display = list(nums)
    # numbered equation typed as plain text
    body = root.find('w:body', NS)
    for p in body.iter(WP):
        if p.find('.//m:oMath', NS) is not None:
            continue
        t = ptext(p).strip()
        mm = re.search(r'=.*\((\d+)\)$', t)
        if mm and len(t) < 160:
            nums.append(int(mm.group(1)))
            issues.append(Issue(CEK, 'Persamaan', 'Persamaan (%s) diketik sebagai teks biasa (bukan Equation Editor)' % mm.group(1), t))
    for n, c in Counter(nums).items():
        if c > 1:
            issues.append(Issue(KRITIS, 'Persamaan', 'Nomor persamaan (%d) dipakai lebih dari sekali' % n))
    if nums:
        expected = list(range(1, max(nums) + 1))
        missing = [n for n in expected if n not in nums]
        if missing:
            issues.append(Issue(CEK, 'Persamaan', 'Nomor persamaan tidak lengkap; tidak ditemukan: %s' %
                                ', '.join('(%d)' % n for n in missing[:15])))
        if nums_display != sorted(nums_display):
            issues.append(Issue(CEK, 'Persamaan', 'Urutan nomor persamaan di dokumen tidak naik: %s' %
                                ' → '.join(str(n) for n in nums_display[:30])))
    for e in eqs:
        t = e['text']
        tag = '(%s)' % e['no'] if e['no'] else 'tanpa nomor'
        plain = re.sub(r'[\[\]]', '', t)
        if plain.count('(') != plain.count(')'):
            issues.append(Issue(CEK, 'Persamaan', 'Persamaan %s: jumlah kurung buka/tutup tidak seimbang' % tag, t))
        for tok, c in Counter(re.findall(r'\b([WUV])_\{(\w+)\}', t)).items():
            if c >= 2:
                issues.append(Issue(CEK, 'Persamaan', 'Persamaan %s: bobot %s_%s muncul %d× — pastikan bukan dua matriks bobot yang '
                                    'seharusnya berbeda (mis. untuk input dan hidden state)' % (tag, tok[0], tok[1], c), t))
        lhs, _, rhs = t.partition('=')
        if '\u0302' in lhs and re.search(r'\+\s*[eε]_\{t\}\s*$', t):
            issues.append(Issue(KRITIS, 'Persamaan', 'Persamaan %s: nilai ramalan (bertopi) ikut menjumlahkan galat acak di ruas '
                                'kanan — ramalan seharusnya tidak memuat eₜ' % tag, t))
        if '∅' in t:
            issues.append(Issue(CEK, 'Persamaan', 'Persamaan %s memakai ∅ (himpunan kosong); maksudnya mungkin φ' % tag, t))
        if re.search(r'\{\s*\}', t):
            issues.append(Issue(CEK, 'Persamaan', 'Persamaan %s memiliki placeholder kosong' % tag, t))
    # references to equations that do not exist
    _, main, _ = body_regions(root)
    text = ' '.join(_visible_text(p) for p in paras_of(main))
    for m in re.finditer(r'\b(?:Equations?|Eq\.|Persamaan)\s*\(?(\d+)\)?', text):
        if nums and int(m.group(1)) not in nums:
            issues.append(Issue(KRITIS, 'Persamaan', 'Teks merujuk ke persamaan (%s) yang tidak ada' % m.group(1),
                                _ctx(text, m.start(), m.end())))
    # math characters in body text that should be in equations
    return issues


# --------------------------------------------------------------------------- language
DOMAIN_WORDS = set('''
hyperparameter hyperparameters autoregressive autocorrelation autocorrelated nonstationary stationarity univariate
multivariate nonlinear nonlinearity heteroscedasticity homoscedasticity datasets dataset preprocessing denormalize
denormalized normalization overfitting underfitting quantile quantiles decorrelation generalizability nonsingular
covariates covariate seasonality seasonal deseasonalized backpropagation sigmoid softmax relu tanh dropout minmax
minmaxscaler eigenvalue eigenvalues eigenvector sarimax arimax sarima arima mewma ewma cusum lstm gru rnn svr
precipitation hydrometeorological interannual multifactor geophysics bogor jakarta surabaya indonesia palu citeko
residuals residual timesteps timestep subgroup subgroups bootstrap bootstrapping metaheuristic heuristic boxplot
boxplots scatterplot scatterplots histogram histograms ljung kolmogorov smirnov wilcoxon kruskal anova manova
multicollinearity parameterization reparameterization estimator estimators unbiasedness consistency asymptotic
asymptotically stochastic stochastically leptokurtic platykurtic kurtosis skewness monte carlo bayesian frequentist
'''.split())

MT_PATTERNS = [
    (r'(?<!\ba )\bvariet(y|ies)\b(?! of)', 'Kemungkinan terjemahan “ragam”. Dalam statistik gunakan “variance”.'),
    (r'\bat the time of\s*-', 'Konstruksi “at the time of-t” — gunakan “at time t”.'),
    (r'(?<!time )\bof-\s*⟨math⟩', 'Konstruksi “… of-k” — gunakan mis. “the k-th …”.'),
    (r'^By\s+⟨math⟩\s*being\b', 'Kalimat penjelasan simbol “By … being …” — gunakan “where … is …”.'),
    (r'\bdifferentiat(ed|ion)\b', 'Untuk deret waktu istilahnya “differencing/differenced”.'),
    (r'\bthe rest of the forecasting\b', 'Maksudnya “residuals”.'),
    (r'\bdata (train|test)\b', 'Gunakan “training/test data”.'),
    (r'\bresidue', 'Untuk model statistik gunakan “residual”.'),
    (r'\b\d+(st|nd|rd|th)-\d\b', 'Sepertinya hasil terjemahan notasi ilmiah (mis. 1e-4 → 10⁻⁴).'),
    (r'\bRMSE[^.%]{0,60}?\d+[.,]\d+\s*%', 'RMSE bukan persentase — cek satuan.'),
    (r'\bit shown\b', '“it shown” → “is shown”.'),
    (r'\bChild\s?3\.4\b', 'Kemungkinan “Niño3.4”.'),
    (r'\b(?:the|a)\s+the\b', 'Artikel ganda.'),
]

PLACEHOLDERS = [
    (r'\[seed\]|\(tolerance\)|\[TBD\]|\bTBD\b|\bXXX+\b|\?\?+|\[citation needed\]|Lorem ipsum', 'Placeholder draf belum diisi'),
    (r'Error! (Reference|Bookmark) source not found|Error! Not a valid', 'Cross-reference Word rusak'),
    (r'\b00 (January|February|March|April|May|June|July|August|September|October|November|December) 20\d\d\b',
     'Tanggal placeholder'),
    (r'\bxx\(x\)|Xxx\. 20xx|vxix', 'Placeholder template belum diganti'),
]


def check_language(root, use_spell=True):
    issues = []
    front, main, _ = body_regions(root)
    ps = paras_of(front + main)
    texts = [_visible_text(p) for p in ps]
    full = '\n'.join(texts)
    for rx, why in MT_PATTERNS:
        for t in texts:
            ms = list(re.finditer(rx, t, re.I | re.M))
            if ms:
                extra = ' (%d× di paragraf ini)' % len(ms) if len(ms) > 1 else ''
                issues.append(Issue(CEK, 'Bahasa', why + extra, _ctx(t, ms[0].start(), ms[0].end())))
    for rx, why in PLACEHOLDERS:
        for t in texts:
            for m in re.finditer(rx, t):
                issues.append(Issue(KRITIS, 'Placeholder', why, _ctx(t, m.start(), m.end())))
    for t in texts:
        for m in re.finditer(r'\b([A-Za-z]{2,})\s+\1\b', t, re.I):
            if m.group(1).lower() not in ('that', 'had', 'is'):
                issues.append(Issue(CEK, 'Ejaan', 'Kata ganda “%s %s”' % (m.group(1), m.group(1)), _ctx(t, m.start(), m.end())))
    if re.search(r'\bSMAPE\b', full) and re.search(r'(?<!S)\bMAPE\b', full):
        for t in texts:
            for m in re.finditer(r'(?<!S)\bMAPE\b', t):
                issues.append(Issue(CEK, 'Konsistensi', 'Naskah memakai SMAPE dan MAPE — pastikan yang dimaksud', _ctx(t, m.start(), m.end())))
    for t in texts:
        for m in re.finditer(r'\b[A-Z][A-Za-z-]+(?: et al\.?| and [A-Z][A-Za-z-]+)? \((?:19|20)\d\d\)', t):
            issues.append(Issue(CEK, 'Sitasi', 'Rujukan gaya author-year di naskah IEEE; gunakan nomor, mis. “Nama et al. [n]”',
                                _ctx(t, m.start(), m.end())))
    if use_spell:
        try:
            from .fixes import _speller
            sp = _speller()
            words = re.findall(r"(?<![\w@/.-])[a-z][a-z]{3,}(?![\w@/-])", re.sub(r'https?://\S+|\S+@\S+', ' ', full))
            cand = [w for w in words if w not in DOMAIN_WORDS]
            unknown = sp.unknown(set(cand))
            cnt = Counter(w for w in cand if w in unknown)
            for k, (w, c) in enumerate(cnt.most_common(60)):
                sug = sp.correction(w) if k < 30 else None
                i = full.find(w)
                issues.append(Issue(CEK, 'Ejaan', 'Kata tidak dikenal “%s” (%d×)%s' % (w, c, ', saran: “%s”' % sug if sug and sug != w else ''),
                                    _ctx(full, i, i + len(w)) if i >= 0 else ''))
        except Exception as exc:   # spell checker is optional
            issues.append(Issue(INFO, 'Ejaan', 'Pemeriksa ejaan tidak tersedia: %s' % exc))
    return issues


# --------------------------------------------------------------------------- citations & references
def _citation_texts(main_elements):
    """Text of citation fields (Mendeley/Zotero) if present, otherwise prose paragraphs outside tables."""
    out = []
    for e in main_elements:
        for sdt in e.iter(q('w:sdt')):
            tag = sdt.find('w:sdtPr/w:tag', NS)
            if tag is not None and tag.get(q('w:val'), '').startswith('MENDELEY_CITATION'):
                out.append(ptext(sdt))
        for p in e.iter(q('w:p')):
            depth, buf, is_cit, instr = 0, '', False, ''
            for r in p.iter(q('w:r')):
                fc = r.find('w:fldChar', NS)
                if fc is not None:
                    t = fc.get(q('w:fldCharType'))
                    if t == 'begin':
                        depth += 1
                        instr, buf = '', ''
                    elif t == 'separate':
                        is_cit = 'CSL_CITATION' in instr or 'ZOTERO_ITEM' in instr
                    elif t == 'end':
                        if is_cit and buf:
                            out.append(buf)
                        depth -= 1
                        is_cit, buf = False, ''
                    continue
                it = r.find('w:instrText', NS)
                if it is not None:
                    instr += it.text or ''
                elif is_cit:
                    buf += get_t(r)
    if out:
        return out
    texts = []
    for p in paras_of(main_elements):
        if any(local(a) == 'tbl' for a in p.iterancestors()):
            continue
        texts.append(ptext(p))
    return texts


def _cited_numbers(main_elements):
    nums = []
    for t in _citation_texts(main_elements):
        for m in re.finditer(r'\[(\d+(?:\s*[–,-]\s*\d+)*)\]', t):
            parts = re.split(r'\s*,\s*', m.group(1))
            if any(x.strip() == '0' for x in parts):
                continue
            for part in parts:
                if re.search(r'[–-]', part):
                    a, b = [int(x) for x in re.split(r'\s*[–-]\s*', part)]
                    nums.extend(range(a, b + 1))
                else:
                    nums.append(int(part))
    return nums


def check_references(root):
    issues = []
    front, main, refs = body_regions(root)
    if not refs:
        return [Issue(KRITIS, 'Referensi', 'Judul bagian REFERENCES tidak ditemukan')]
    entries = {}
    for e in reference_entries(root):
        entries[e['n']] = re.sub(r'^\[\d+\]\s*', '', e['text'])
    if not entries:
        return [Issue(KRITIS, 'Referensi', 'Daftar pustaka tidak bernomor [n] (format IEEE)')]
    keys = sorted(entries)
    if keys != list(range(1, len(keys) + 1)):
        issues.append(Issue(KRITIS, 'Referensi', 'Nomor daftar pustaka tidak berurutan: %s' % keys))
    cited = _cited_numbers(main)
    order = []
    for n in cited:
        if n not in order:
            order.append(n)
    for n in sorted(set(cited) - set(entries)):
        issues.append(Issue(KRITIS, 'Sitasi', 'Sitasi [%d] tidak ada di daftar pustaka' % n))
    for n in sorted(set(entries) - set(cited)):
        issues.append(Issue(CEK, 'Sitasi', 'Referensi [%d] tidak pernah disitasi di naskah' % n, entries[n][:120]))
    for i, n in enumerate(order):
        if n != i + 1 and n in entries:
            issues.append(Issue(CEK, 'Sitasi', 'Urutan kemunculan pertama sitasi tidak sesuai IEEE: [%d] muncul pada posisi ke-%d '
                                '(referensi harus dinomori sesuai urutan pertama kali disitasi)' % (n, i + 1)))
            break
    if len(entries) < 15:
        issues.append(Issue(CEK, 'Referensi', 'Jumlah referensi %d (< 15 sesuai template)' % len(entries)))
    for n, e in entries.items():
        title = re.search(r'[“"](.+?)[”"]', e)
        if title:
            tt = title.group(1)
            letters = [c for c in tt if c.isalpha()]
            if letters and sum(c.isupper() for c in letters) / len(letters) > 0.8 and len(letters) > 12:
                issues.append(Issue(CEK, 'Referensi', '[%d] Judul ditulis huruf kapital semua' % n, e[:140]))
            after = e[title.end():title.end() + 12]
            if re.match(r'\s*,?”?\s*(vol\.|pp\.|no\.)', after) or re.match(r'\s*,?”?\s*(vol\.|pp\.|no\.)', e[title.end():].lstrip('”," ')):
                issues.append(Issue(CEK, 'Referensi', '[%d] Nama jurnal/prosiding tampaknya tidak ada' % n, e[:140]))
        if not re.search(r'\b(19|20)\d\d\b', e):
            issues.append(Issue(CEK, 'Referensi', '[%d] Tahun terbit tidak ditemukan' % n, e[:140]))
        m = re.search(r'doi:\s*(10\.\S+?)\.?$', e)
        if m and re.search(r'\d[a-z]{2,}$', m.group(1)):
            issues.append(Issue(KRITIS, 'Referensi', '[%d] DOI diakhiri huruf yang mencurigakan: %s' % (n, m.group(1)), e[-120:]))
        if re.search(r'\bno\.\s*\d{5,}', e):
            issues.append(Issue(CEK, 'Referensi', '[%d] Nomor artikel ditulis sebagai “no.”; IEEE: “p. <nomor artikel>”' % n, e[:140]))
        first_author = re.split(r',| and ', e, maxsplit=1)[0]
        if not re.search(r'\b[A-Z]\.', first_author) and len(first_author.split()) >= 2 and '“' not in first_author:
            issues.append(Issue(CEK, 'Referensi', '[%d] Nama penulis pertama bukan format inisial IEEE (mis. “A. B. Surname”)' % n,
                                first_author))
        if re.search(r'\b(\w+)[- ]\1\b', e) and not re.search(r'\d', re.search(r'\b(\w+)[- ]\1\b', e).group(0)):
            issues.append(Issue(CEK, 'Referensi', '[%d] Ada kata berulang' % n, e[:140]))
        if 'Online' in e and 'Available' in e and 'doi' not in e.lower():
            pass
    return issues


# --------------------------------------------------------------------------- figures, tables, layout
def check_figures_tables(root):
    issues = []
    body = root.find('w:body', NS)
    caps = {'Figure': [], 'Table': []}
    for p in body.iter(WP):
        if p.xpath('ancestor::mc:Fallback', namespaces=NS):
            continue
        t = ''.join(p.xpath('./w:r/w:t/text()|./w:hyperlink/w:r/w:t/text()', namespaces=NS)).strip()
        m = re.match(r'^(Figure|Table)\s+(\d+)\b', t)
        if m and len(t) < 300 and not re.search(r'\b(shows?|presents?|indicates?|illustrates?|displays?|summari[sz]es|'
                                                 r'lists|reports|reveals|demonstrates|depicts|contains|is|are)\b', t[:70]):
            caps[m.group(1)].append(int(m.group(2)))
    for kind, nums in caps.items():
        if nums and sorted(set(nums)) != list(range(1, max(nums) + 1)):
            issues.append(Issue(CEK, 'Gambar/Tabel', 'Nomor caption %s tidak berurutan: %s' % (kind, nums)))
        for n, c in Counter(nums).items():
            if c > 1:
                issues.append(Issue(KRITIS, 'Gambar/Tabel', 'Caption %s %d muncul %d×' % (kind, n, c)))
    _, main, _ = body_regions(root)
    text = ' '.join(_visible_text(p) for p in paras_of(main))
    for kind in caps:
        mx = max(caps[kind]) if caps[kind] else 0
        for m in re.finditer(r'\b%ss?\s+(\d+)' % kind, text):
            if int(m.group(1)) > mx:
                issues.append(Issue(KRITIS, 'Gambar/Tabel', 'Teks merujuk %s %s yang tidak ada' % (kind, m.group(1)),
                                    _ctx(text, m.start(), m.end())))
    n_float = 0
    for an in body.iter(q('wp:anchor')):
        if an.find('.//pic:pic', NS) is not None:
            n_float += 1
    if n_float:
        issues.append(Issue(CEK, 'Tata letak', '%d gambar berupa objek mengambang (floating). Periksa di Word agar tidak menimpa '
                            'tabel/teks; idealnya gambar inline dengan caption di bawahnya.' % n_float))
    for tbl in body.iter(q('w:tbl')):
        for t in tbl.iter(q('w:t')):
            if t.text and re.search(r'\d\.\d{8,}', t.text):
                issues.append(Issue(CEK, 'Gambar/Tabel', 'Angka dengan digit berlebihan di tabel (template: hindari digit tak perlu)',
                                    t.text))
    return issues


# --------------------------------------------------------------------------- front matter & OJS
def front_matter(root):
    front, _, _ = body_regions(root)
    ps = paras_of(front)
    texts = [ptext(p).strip() for p in ps]
    info = {'title': '', 'abstract': '', 'keywords': [], 'authors_line': '', 'has_corresponding': False}
    for t in texts:
        if t and not info['title'] and not t.startswith(('Received', 'Revised', 'Accepted')):
            info['title'] = t
        if t.upper().startswith('ABSTRACT'):
            info['abstract'] = re.sub(r'^ABSTRACT\s*[⎯–—:\-]*\s*', '', t, flags=re.I)
        if t.lower().startswith('keywords'):
            kw = re.sub(r'^keywords\s*[⎯–—:\-]*\s*', '', t, flags=re.I)
            info['keywords'] = [k.strip(' .') for k in re.split(r'[,;]', kw) if k.strip(' .')]
        if 'corresponding author' in t.lower():
            info['has_corresponding'] = True
    return info


def check_front(root, ojs=None):
    issues = []
    fm = front_matter(root)
    nt = len(fm['title'].split())
    if nt > 16:
        issues.append(Issue(CEK, 'Format', 'Judul %d kata (maksimal 16 kata sesuai template)' % nt, fm['title']))
    na = len(fm['abstract'].split())
    if na > 200:
        issues.append(Issue(CEK, 'Format', 'Abstrak %d kata (maksimal 200 kata sesuai template)' % na))
    if not fm['abstract']:
        issues.append(Issue(KRITIS, 'Format', 'Paragraf ABSTRACT tidak ditemukan'))
    kws = fm['keywords']
    if len(kws) > 5:
        issues.append(Issue(CEK, 'Format', 'Kata kunci %d (maksimal 5)' % len(kws), ', '.join(kws)))
    if kws and [k.lower() for k in kws] != sorted(k.lower() for k in kws):
        issues.append(Issue(CEK, 'Format', 'Kata kunci tidak urut abjad', ', '.join(kws)))
    if not fm['has_corresponding']:
        issues.append(Issue(KRITIS, 'Format', 'Baris “*Corresponding author” tidak ditemukan — naskah mungkin masih versi anonim '
                            '(blok penulis/afiliasi belum ada)'))
    if ojs is not None:
        if ojs.title:
            r = difflib.SequenceMatcher(None, ojs.title.lower(), fm['title'].lower()).ratio()
            if r < 0.95:
                issues.append(Issue(KRITIS if r < 0.7 else CEK, 'OJS', 'Judul naskah berbeda dari metadata OJS (kemiripan %.0f%%)' % (100 * r),
                                    'Naskah: %s | OJS: %s' % (fm['title'], ojs.title)))
        front_text = ' '.join(ptext(e) for e in body_regions(root)[0])
        norm = lambda s: ' '.join(re.sub(r'[^a-z]+', ' ', s.lower()).split())
        for name, email in ojs.contributors:
            if norm(name) not in norm(front_text):
                issues.append(Issue(KRITIS, 'OJS', 'Nama kontributor OJS tidak ditemukan di blok penulis: %s (%s)' % (name, email)))
        if ojs.keywords and kws:
            a, b = {k.lower() for k in ojs.keywords}, {k.lower() for k in kws}
            if a != b:
                issues.append(Issue(CEK, 'OJS', 'Kata kunci berbeda dengan OJS', 'Naskah: %s | OJS: %s' % (', '.join(kws), ', '.join(ojs.keywords))))
        elif not ojs.keywords:
            issues.append(Issue(INFO, 'OJS', 'Kata kunci di metadata OJS kosong'))
        if ojs.abstract and fm['abstract']:
            r = difflib.SequenceMatcher(None, ojs.abstract, fm['abstract']).ratio()
            if r < 0.97:
                issues.append(Issue(CEK, 'OJS', 'Abstrak di metadata OJS berbeda dari naskah final (kemiripan %.0f%%); '
                                    'samakan metadata OJS sebelum terbit' % (100 * r)))
        for n in ojs.notes:
            issues.append(Issue(CEK, 'OJS', n))
    return issues, fm
