"""Compatibility test: python tests/test_compat.py <folder-with-docx> [ojs.html]

Every .docx in the folder is processed; the output must open in python-docx and every XML part must parse.
Files that cannot be processed must fail with a friendly DocxError (never a crash).
"""
import glob
import io
import os
import sys
import time
import traceback

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from finalizer.docxtools import DocxError, DocxPackage, validate_package  # noqa: E402
from finalizer.pipeline import run  # noqa: E402

folder = sys.argv[1]
html = open(sys.argv[2], 'rb').read() if len(sys.argv) > 2 else None
os.makedirs(os.path.join(folder, 'out'), exist_ok=True)
ok = True
for f in sorted(glob.glob(os.path.join(folder, '*.docx'))):
    name = os.path.basename(f)
    t0 = time.time()
    try:
        res = run(open(f, 'rb').read(), 9, 3, 'November', 2026, html, None)
        out = os.path.join(folder, 'out', name)
        open(out, 'wb').write(res.docx)
        probs = validate_package(res.docx, DocxPackage(res.docx).main)
        import docx
        docx.Document(io.BytesIO(res.docx))
        status = 'OK' if not probs else 'INVALID ' + '; '.join(probs[:2])
        if probs:
            ok = False
        print('%-26s %-8s %5.1fs changes=%-3d issues=%-3d safe=%s failed=%s' % (
            name, status, time.time() - t0, len(res.changes), len(res.issues), res.safe_mode, res.failed_steps))
    except DocxError as exc:
        print('%-26s REJECTED (friendly): %s' % (name, str(exc)[:90]))
    except Exception:
        ok = False
        print('%-26s CRASH' % name)
        traceback.print_exc()
print('ALL GOOD' if ok else 'PROBLEMS FOUND')
