"""Control: open and save with python-docx (no edits). Expected to be near-lossless."""
import sys

from docx import Document

Document(sys.argv[1]).save(sys.argv[2])
