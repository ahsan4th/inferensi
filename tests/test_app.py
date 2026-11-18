"""UI smoke test with Streamlit's AppTest: python tests/test_app.py <docx> [ojs.html]"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from streamlit.testing.v1 import AppTest  # noqa: E402

from finalizer.pipeline import run  # noqa: E402

APP = os.path.join(os.path.dirname(__file__), '..', 'app.py')
at = AppTest.from_file(APP, default_timeout=60).run()
assert not at.exception, at.exception
print('initial render OK; sidebar widgets:', len(at.sidebar.number_input), 'number inputs')
if len(sys.argv) > 1:
    res = run(open(sys.argv[1], 'rb').read(), 9, 3, 'November', 2026,
              open(sys.argv[2], 'rb').read() if len(sys.argv) > 2 else None)
    at.session_state['res'] = res
    at.session_state['name'] = 'uji'
    at.run()
    assert not at.exception, at.exception
    print('result render OK; metrics:', [(m.label, m.value) for m in at.metric])
    at.multiselect[0].set_value(['Kritis']).run()
    assert not at.exception, at.exception
    print('filter interaction OK')
