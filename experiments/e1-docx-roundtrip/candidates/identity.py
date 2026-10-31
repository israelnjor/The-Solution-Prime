"""Control: byte-for-byte copy. Must score zero findings, or the checker itself is wrong."""
import shutil
import sys

shutil.copyfile(sys.argv[1], sys.argv[2])
