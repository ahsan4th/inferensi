"""Streamlit app: finalisasi artikel INFERENSI siap terbit."""
import dataclasses
import datetime as dt
import inspect
import io
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

from finalizer import __version__
from finalizer.checks import CEK, INFO, KRITIS
from finalizer.ojs import MONTHS
from finalizer.pipeline import Options, run

try:
    from finalizer.docxtools import DocxError
except ImportError:                      # older finalizer package
    class DocxError(Exception):
        pass

try:
    from finalizer.letters import DEFAULT_AMOUNT, LetterInfo, load_default_templates, make_letters
    HAS_LETTERS = True
except Exception:                        # letters.py missing or incompatible
    HAS_LETTERS = False
    DEFAULT_AMOUNT = 'IDR 1.500.000'

BASE = Path(__file__).resolve().parent
OPTION_FIELDS = {f.name for f in dataclasses.fields(Options)}
RUN_PARAMS = set(inspect.signature(run).parameters)
NEEDED_OPTIONS = {'template_format', 'remove_highlight', 'unbox_tables', 'mendeley_fields', 'math_convert'}
OUTDATED = sorted(NEEDED_OPTIONS - OPTION_FIELDS) or ([] if 'template_bytes' in RUN_PARAMS else ['template_bytes'])


def make_options(**kw):
    """Options accepted by the installed finalizer package (unknown ones are ignored, so the app never crashes)."""
    return Options(**{k: v for k, v in kw.items() if k in OPTION_FIELDS})


def run_pipeline(*args, **kw):
    """Call finalizer.pipeline.run with only the keyword arguments it supports."""
    return run(*args, **{k: v for k, v in kw.items() if k in RUN_PARAMS})


def show_df(df, **kw):
    """st.dataframe at full width on both older and newer Streamlit versions."""
    try:
        st.dataframe(df, width='stretch', **kw)
    except TypeError:
        st.dataframe(df, use_container_width=True, **kw)

st.set_page_config(page_title='INFERENSI Finalizer', page_icon='📰', layout='wide')

st.title('📰 INFERENSI Finalizer')
st.caption('Finalisasi artikel terakhir sebelum terbit: header, tanggal penting, typo, format template, persamaan, '
           'sitasi & referensi IEEE, metadata dokumen. Versi %s' % __version__)
if OUTDATED or not HAS_LETTERS:
    st.warning('Folder **finalizer/** di server belum versi terbaru (terpasang: %s). Aplikasi tetap berjalan, tetapi fitur '
               'berikut nonaktif: %s. Unggah ulang seluruh folder **finalizer/** dari paket terbaru (v1.3.0) agar semua '
               'fitur aktif.' % (__version__, ', '.join(
                   [{'template_format': 'format template', 'remove_highlight': 'hapus highlight',
                     'unbox_tables': 'hapus kotak di tabel', 'mendeley_fields': 'field Mendeley',
                     'math_convert': 'variabel ke Equation Word',
                     'template_bytes': 'salin header/footer template'}.get(x, x) for x in OUTDATED] +
                   ([] if HAS_LETTERS else ['LOA & Receipt']))))

with st.sidebar:
    st.header('1. Terbitan')
    c1, c2 = st.columns(2)
    volume = c1.number_input('Volume', min_value=1, max_value=99, value=9, step=1)
    issue = c2.number_input('Nomor (issue)', min_value=1, max_value=12, value=3, step=1)
    month = st.selectbox('Bulan', MONTHS, index=10)
    year = st.number_input('Tahun', min_value=2000, max_value=2100, value=dt.date.today().year, step=1)
    st.info('Header: **INFERENSI, Vol. %d(%d), %s. %d.** ISSN: 0216-308X (Print) 2721-3862 (Online)'
            % (volume, issue, month, year))

    st.header('2. DOI (opsional)')
    doi = st.text_input('DOI artikel', placeholder='10.12962/j27213862.v9i3.12345')

    st.header('3. Opsi')
    opt = make_options(
        decimal_comma=st.checkbox('Ubah koma desimal menjadi titik (1426,06 → 1426.06)', value=True),
        safe_language=st.checkbox('Perbaikan istilah yang aman (data train → training data, residue → residual, …)', value=True),
        reference_format=st.checkbox('Rapikan format referensi IEEE (no. Bulan, [Online] DOI, akronim)', value=True),
        equation_layout=st.checkbox('Rapikan posisi nomor persamaan', value=True),
        accept_changes=st.checkbox('Terima semua track changes', value=True),
        remove_comments=st.checkbox('Hapus komentar Word', value=True),
        spell_check=st.checkbox('Pemeriksaan ejaan (Inggris)', value=True),
        template_format=st.checkbox('Format sesuai template (font, ukuran, warna, justify, section halaman, header/footer)', value=True),
        remove_highlight=st.checkbox('Hapus highlight/sorotan teks', value=True),
        unbox_tables=st.checkbox('Hapus kotak teks/bingkai di dalam tabel', value=True),
        mendeley_fields=st.checkbox('Jadikan sitasi teks biasa sebagai field Mendeley (dapat di-refresh)', value=True),
        math_convert=st.checkbox('Ubah variabel & persamaan teks menjadi Equation Word', value=True),
    )
    st.header('4. Template (opsional)')
    up_tpl = st.file_uploader('Template INFERENSI (.docx)', type=['docx'], key='tpl',
                              help='Bila diunggah, header dan footer naskah disalin persis dari template ini.')

    st.header('5. LOA & Receipt')
    default_letters = load_default_templates(str(BASE)) if HAS_LETTERS else {}
    st.caption('Template bawaan: LOA %s · Receipt %s' % ('✅' if 'LOA' in default_letters else '—',
                                                        '✅' if 'Receipt' in default_letters else '—'))
    up_loa = st.file_uploader('Template LOA (.docx, opsional)', type=['docx'], key='loa')
    up_rcpt = st.file_uploader('Template Receipt (.docx, opsional)', type=['docx'], key='rcpt')
    amount = st.text_input('Jumlah pembayaran', value=DEFAULT_AMOUNT)

col_a, col_b = st.columns(2)
with col_a:
    up_docx = st.file_uploader('Artikel (MS Word .docx)', type=['docx', 'docm', 'dotx'])
    st.caption('File .doc (Word 97–2003) atau yang dikunci password harus disimpan ulang sebagai .docx terlebih dahulu.')
with col_b:
    up_html = st.file_uploader('Riwayat OJS (halaman submission disimpan sebagai .html)', type=['html', 'htm'])
    st.caption('Di OJS buka halaman submission → Activity Log/History, lalu *Save Page As* → HTML.')

go = st.button('🚀 Proses artikel', type='primary', disabled=up_docx is None)

if go and up_docx is not None:
    with st.spinner('Memproses…'):
        try:
            res = run_pipeline(up_docx.getvalue(), volume, issue, month, year,
                      history_html=up_html.getvalue() if up_html else None,
                      doi=doi or None, options=opt,
                      template_bytes=up_tpl.getvalue() if up_tpl else None)
        except DocxError as exc:
            st.error('File Word tidak dapat diproses: %s' % exc)
            st.stop()
        except Exception as exc:   # show the error instead of a blank page
            st.error('Terjadi kesalahan yang tidak terduga. Silakan laporkan beserta file uji (tanpa data sensitif).')
            st.exception(exc)
            st.stop()
    st.session_state['res'] = res
    st.session_state['name'] = Path(up_docx.name).stem
    tpl_loa = up_loa.getvalue() if up_loa else default_letters.get('LOA')
    tpl_rcpt = up_rcpt.getvalue() if up_rcpt else default_letters.get('Receipt')
    st.session_state['letter_tpl'] = (tpl_loa, tpl_rcpt)
    if HAS_LETTERS:
        li = getattr(res, 'letter_info', None) or LetterInfo(volume=str(volume), issue=str(issue), year=str(year))
        li.amount = amount or DEFAULT_AMOUNT
        st.session_state['linfo'] = li
        with st.spinner('Membuat LOA & receipt…'):
            try:
                st.session_state['letters'] = make_letters(li, tpl_loa, tpl_rcpt)
            except Exception as exc:
                st.session_state['letters'] = ({}, ['LOA/Receipt gagal dibuat: %s' % exc])
    else:
        st.session_state['linfo'] = None
        st.session_state['letters'] = ({}, ['Modul LOA/Receipt (finalizer/letters.py) belum terpasang di server.'])

res = st.session_state.get('res')
if res:
    stem = st.session_state.get('name', 'artikel')
    n_k = sum(i.severity == KRITIS for i in res.issues)
    n_c = sum(i.severity == CEK for i in res.issues)
    if res.safe_mode:
        st.warning('Mode aman aktif: pemrosesan lengkap tidak lolos validasi, sehingga hanya header, footer, tanggal, '
                   'lisensi, dan metadata yang diubah. Lihat tab *Perlu ditinjau*.')
    m1, m2, m3, m4 = st.columns(4)
    m1.metric('Perbaikan otomatis', len(res.changes))
    m2.metric('Kritis', n_k)
    m3.metric('Perlu dicek', n_c)
    m4.metric('Persamaan terdeteksi', len(res.equations))

    final_name = '%s_final.docx' % stem
    letter_files, letter_msgs = st.session_state.get('letters', ({}, []))
    d1, d2, d3, d4 = st.columns(4)
    d1.download_button('⬇️ Word siap terbit', res.docx, file_name=final_name,
                       mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document', type='primary')
    d2.download_button('⬇️ Laporan (.md)', res.report_md.encode('utf-8'), file_name='%s_laporan.md' % stem)
    if res.ris:
        d3.download_button('⬇️ Referensi (.ris)', res.ris.encode('utf-8'), file_name='%s_references.ris' % stem)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr(final_name, res.docx)
        z.writestr('%s_laporan.md' % stem, res.report_md)
        if res.ris:
            z.writestr('%s_references.ris' % stem, res.ris)
        for kind in ('LOA', 'Receipt'):
            ext = 'pdf' if '%s.pdf' % kind in letter_files else 'docx'
            if '%s.%s' % (kind, ext) in letter_files:
                z.writestr('%s_%s.%s' % (stem, kind, ext), letter_files['%s.%s' % (kind, ext)])
    d4.download_button('⬇️ Semua (.zip)', buf.getvalue(), file_name='%s_finalisasi.zip' % stem)
    l1, l2, l3, l4 = st.columns(4)
    for col, kind in ((l1, 'LOA'), (l2, 'Receipt')):
        if '%s.pdf' % kind in letter_files:
            col.download_button('⬇️ %s (PDF)' % kind, letter_files['%s.pdf' % kind], file_name='%s_%s.pdf' % (stem, kind),
                                mime='application/pdf')
        elif '%s.docx' % kind in letter_files:
            col.download_button('⬇️ %s (Word)' % kind, letter_files['%s.docx' % kind], file_name='%s_%s.docx' % (stem, kind))

    t1, t2, t3, t4, t5, t6 = st.tabs(['Ringkasan', 'Perbaikan otomatis', 'Perlu ditinjau', 'Persamaan', 'OJS',
                                      'LOA & Receipt'])
    with t1:
        st.write('**Header:** ' + res.header)
        st.write('**DOI di footer:** ' + (res.doi or '_placeholder — isi DOI bila sudah terbit_'))
        if res.dates:
            st.write('**Tanggal:** Received %s · Revised %s · Accepted %s' % (
                res.dates.get('Received') or '—', res.dates.get('Revised') or '—', res.dates.get('Accepted') or '—'))
        fm = res.front
        st.write('**Judul:** %s (%d kata)' % (fm.get('title', ''), len(fm.get('title', '').split())))
        st.write('**Abstrak:** %d kata · **Kata kunci:** %s' % (len(fm.get('abstract', '').split()),
                                                           ', '.join(fm.get('keywords', [])) or '—'))
        crit = [i for i in res.issues if i.severity == KRITIS]
        if crit:
            st.error('Ada %d temuan kritis yang harus diselesaikan sebelum terbit.' % len(crit))
            for i in crit:
                st.markdown('- **[%s]** %s' % (i.category, i.message) + (' — _%s_' % i.context if i.context else ''))
        else:
            st.success('Tidak ada temuan kritis.')
    with t2:
        if res.changes:
            show_df(pd.DataFrame(res.changes, columns=['Kategori', 'Perubahan']), hide_index=True)
        else:
            st.write('Tidak ada perubahan otomatis.')
    with t3:
        sev = st.multiselect('Tingkat', [KRITIS, CEK, INFO], default=[KRITIS, CEK])
        rows = [(i.severity, i.category, i.message, i.context) for i in res.issues if i.severity in sev]
        cats = sorted({r[1] for r in rows})
        pick = st.multiselect('Kategori', cats, default=cats)
        rows = [r for r in rows if r[1] in pick]
        show_df(pd.DataFrame(rows, columns=['Tingkat', 'Kategori', 'Temuan', 'Konteks']), hide_index=True, height=520)
    with t4:
        st.caption('Semua persamaan display ditampilkan sebagai teks linear agar mudah diperiksa satu per satu.')
        show_df(pd.DataFrame([(('(%s)' % e['no']) if e['no'] else '—', e['text']) for e in res.equations],
                             columns=['No.', 'Persamaan']), hide_index=True)
    with t5:
        if res.ojs:
            o = res.ojs
            st.write('**Judul OJS:** ' + o.title)
            st.write('**Kontributor:** ' + ', '.join('%s (%s)' % c for c in o.contributors))
            st.write('**Kata kunci OJS:** ' + (', '.join(o.keywords) or '(kosong)'))
            st.write('**ID submission:** ' + (o.submission_id or '—'))
            ev = pd.DataFrame([(e[0].isoformat(), e[1], e[2]) for e in o.events], columns=['Tanggal', 'Pengguna', 'Aktivitas'])
            show_df(ev, hide_index=True, height=360)
        else:
            st.write('Riwayat OJS tidak diunggah.')
    with t6:
        li = st.session_state.get('linfo')
        tpl_loa, tpl_rcpt = st.session_state.get('letter_tpl', (None, None))
        for m in (li.notes if li else []) + letter_msgs:
            st.warning(m)
        if li is not None:
            st.caption('Data diisi otomatis dari naskah dan riwayat OJS. Ubah bila perlu, lalu klik **Perbarui**.')
            with st.form('letters'):
                c1, c2 = st.columns(2)
                name = c1.text_input('Nama penerima (penulis korespondensi)', li.name)
                email = c2.text_input('Email', li.email)
                title = st.text_area('Judul naskah', li.title, height=80)
                c3, c4, c5 = st.columns(3)
                article_id = c3.text_input('Article ID (OJS)', li.article_id)
                amt = c4.text_input('Jumlah pembayaran', li.amount)
                vol = c5.text_input('Volume / Nomor / Tahun', '%s / %s / %s' % (li.volume, li.issue, li.year))
                c6, c7, c8 = st.columns(3)
                loa_date = c6.date_input('Tanggal LOA (Accept)', li.loa_date)
                pay_date = c7.date_input('Tanggal pembayaran diterima', li.payment_date)
                rc_date = c8.date_input('Tanggal receipt', li.receipt_date)
                if st.form_submit_button('🔄 Perbarui LOA & Receipt'):
                    parts = [x.strip() for x in vol.split('/')] + ['', '', '']
                    new = LetterInfo(name=name, email=email, title=title.strip(), article_id=article_id, amount=amt,
                                     volume=parts[0], issue=parts[1], year=parts[2], loa_date=loa_date,
                                     payment_date=pay_date, receipt_date=rc_date)
                    st.session_state['linfo'] = new
                    with st.spinner('Membuat ulang LOA & receipt…'):
                        try:
                            st.session_state['letters'] = make_letters(new, tpl_loa, tpl_rcpt)
                        except Exception as exc:
                            st.session_state['letters'] = ({}, ['LOA/Receipt gagal dibuat: %s' % exc])
                    st.rerun()
else:
    st.markdown('''
**Cara pakai**
1. Isi volume, nomor, bulan, dan tahun terbitan di sidebar (DOI opsional).
2. Unggah naskah final (.docx, sudah memakai template INFERENSI) dan halaman riwayat OJS (.html).
3. Klik **Proses artikel**, unduh Word siap terbit, lalu selesaikan temuan di tab *Perlu ditinjau*.
''')
