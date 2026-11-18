"""Parse a saved OJS 3 submission page (File > Save Page As .html).

Extracts the activity log (history) and the metadata needed for consistency
checks: title, abstract, keywords and contributors.
"""
from __future__ import annotations

import datetime as dt
import html as htmlmod
import re
from dataclasses import dataclass, field

MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August',
          'September', 'October', 'November', 'December']


def fmt_date(d: dt.date | None) -> str:
    return '' if d is None else '%d %s %d' % (d.day, MONTHS[d.month - 1], d.year)


@dataclass
class OJSInfo:
    title: str = ''
    abstract: str = ''
    keywords: list = field(default_factory=list)
    contributors: list = field(default_factory=list)   # (name, email)
    events: list = field(default_factory=list)         # (date, user, text)
    received: dt.date | None = None
    revised: dt.date | None = None
    accepted: dt.date | None = None
    submission_id: str = ''
    notes: list = field(default_factory=list)

    def dates_text(self):
        return {'Received': fmt_date(self.received), 'Revised': fmt_date(self.revised),
                'Accepted': fmt_date(self.accepted)}


def _lines(raw: str):
    raw = re.sub(r'<script.*?</script>', '', raw, flags=re.S | re.I)
    raw = re.sub(r'<style.*?</style>', '', raw, flags=re.S | re.I)
    raw = re.sub(r'<[^>]+>', '\n', raw)
    raw = htmlmod.unescape(raw)
    return [ln.strip() for ln in raw.split('\n') if ln.strip()]


def parse_history(data: bytes | str) -> OJSInfo:
    raw = data.decode('utf-8', errors='replace') if isinstance(data, bytes) else data
    info = OJSInfo()

    m = re.search(r'<title>(.*?)</title>', raw, re.S | re.I)
    if m:
        parts = [p.strip() for p in htmlmod.unescape(m.group(1)).split('|')]
        if len(parts) >= 3:
            info.title = ' | '.join(parts[1:-1])
        elif parts:
            info.title = parts[0]

    L = _lines(raw)

    # contributors -------------------------------------------------------------
    if 'List of Contributors' in L:
        i = L.index('List of Contributors')
        for k in range(i, min(i + 200, len(L) - 1)):
            if L[k] in ('No Items', 'Done') and info.contributors:
                break
            if re.fullmatch(r'[^@\s]+@[^@\s]+\.[^@\s]+', L[k]) and k > 0:
                name = L[k - 1]
                if name not in ('Settings', 'E-mail', 'Primary Contact') and not name.startswith('Remove'):
                    info.contributors.append((name, L[k]))

    # keywords (first "Selected:" block after an English "Keywords" heading) ----
    for i, ln in enumerate(L):
        if ln == 'Keywords' and i + 3 < len(L):
            for k in range(i, min(i + 6, len(L))):
                if L[k] == 'Selected:':
                    kws = []
                    for x in L[k + 1:k + 60]:
                        if x.startswith('Remove '):
                            continue
                        if 'languages completed' in x or x in ('Bahasa Indonesia', 'English', 'Save'):
                            break
                        kws.append(x)
                    info.keywords = kws
                    break
            if info.keywords:
                break

    # abstract: first long line after an "Abstract" label ----------------------
    for i, ln in enumerate(L):
        if ln == 'Abstract':
            for x in L[i + 1:i + 8]:
                if len(x) > 300:
                    info.abstract = x
                    break
        if info.abstract:
            break

    # history ------------------------------------------------------------------
    start = 0
    for i, ln in enumerate(L):
        if ln == 'History' and i + 1 < len(L) and L[i + 1] in ('Notes', 'Date'):
            start = i
    for k in range(start, len(L) - 2):
        if re.fullmatch(r'\d{4}-\d{2}-\d{2}', L[k]):
            try:
                d = dt.date.fromisoformat(L[k])
            except ValueError:
                continue
            info.events.append((d, L[k + 1], L[k + 2]))
    m = re.search(r'for article (\d+)', ' '.join(e[2] for e in info.events))
    if m:
        info.submission_id = m.group(1)

    _infer_dates(info)
    return info


def _infer_dates(info: OJSInfo):
    ev = info.events
    init = [e for e in ev if 'submission completed' in e[2].lower()]
    if init:
        info.received = min(e[0] for e in init)
        submitter = min(init, key=lambda e: e[0])[1]
    else:
        info.received = min((e[0] for e in ev), default=None)
        submitter = None
        if ev:
            info.notes.append('“Initial submission completed” tidak ditemukan; Received memakai tanggal log paling awal.')
    acc = [e for e in ev if 'accept submission' in e[2].lower()]
    if acc:
        info.accepted = max(e[0] for e in acc)
    else:
        info.notes.append('Keputusan “Accept Submission” tidak ditemukan di riwayat.')
    limit = info.accepted or dt.date.max
    rv = [e for e in ev if 'revised version uploaded' in e[2].lower() and e[0] <= limit]
    if not rv and submitter:
        rv = [e for e in ev if e[1] == submitter and 'revision' in e[2].lower() and 'uploaded' in e[2].lower()
              and info.received and info.received < e[0] <= limit]
    if rv:
        info.revised = max(e[0] for e in rv)
    else:
        info.notes.append('Unggahan revisi oleh penulis tidak ditemukan; tanggal Revised perlu diisi manual.')
