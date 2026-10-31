"""Reference: open and re-save in Microsoft Word. Shows what normalisation Word itself applies."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from word import Word  # noqa: E402

with Word() as w:
    w.resave(sys.argv[1], sys.argv[2])
