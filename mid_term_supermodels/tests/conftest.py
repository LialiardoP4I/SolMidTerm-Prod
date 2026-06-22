import sys
from pathlib import Path

# Permette ai test di importare i moduli del package (matching, montecarlo, ...)
PKG = Path(__file__).resolve().parent.parent
if str(PKG) not in sys.path:
    sys.path.insert(0, str(PKG))
