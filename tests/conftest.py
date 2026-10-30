import os
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

# In-process tests check API shapes only; real model calls are covered by tests/test_final_e2e_live.py.
os.environ.setdefault("GEMINI_OFFICE_BACKEND", "offline")


def pytest_configure(config):
    config.addinivalue_line("markers", "word: needs desktop Microsoft Word (hidden DispatchEx instance)")
