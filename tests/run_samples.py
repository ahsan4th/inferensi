"""Local smoke test: python tests/run_samples.py <docx> <html> [doi]"""
import sys, time
sys.path.insert(0, '.')
from finalizer.pipeline import run
docx, html = sys.argv[1], sys.argv[2]
t0 = time.time()
res = run(open(docx, 'rb').read(), 9, 3, 'November', 2026, open(html, 'rb').read(), sys.argv[3] if len(sys.argv) > 3 else None)
out = sys.argv[4] if len(sys.argv) > 4 else '/tmp/out.docx'
open(out, 'wb').write(res.docx)
open(out.replace('.docx', '.md'), 'w').write(res.report_md)
print('time %.1fs | changes %d | issues %d | eqs %d | ris %d' % (time.time() - t0, len(res.changes), len(res.issues), len(res.equations), res.n_ris))
print(res.dates)
