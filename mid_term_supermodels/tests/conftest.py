import sys
from pathlib import Path

# Permette ai test di importare i moduli del package (matching, montecarlo, ...)
PKG = Path(__file__).resolve().parent.parent
if str(PKG) not in sys.path:
    sys.path.insert(0, str(PKG))

# Il package usa import assoluti `from mid_term...`, quindi anche la cartella
# genitore (SolMidTerm - Prod) finisce su sys.path e contiene un suo pacchetto
# `tests/` che maschererebbe questo. Forziamo la risoluzione del pacchetto
# `tests` locale (quello di mid_term_supermodels) ribindandolo esplicitamente.
_LOCAL_TESTS_INIT = Path(__file__).resolve().parent / "__init__.py"
_t = sys.modules.get("tests")
if _t is None or getattr(_t, "__file__", None) != str(_LOCAL_TESTS_INIT):
    import importlib.util
    _spec = importlib.util.spec_from_file_location(
        "tests", str(_LOCAL_TESTS_INIT),
        submodule_search_locations=[str(_LOCAL_TESTS_INIT.parent)],
    )
    _mod = importlib.util.module_from_spec(_spec)
    sys.modules["tests"] = _mod
    _spec.loader.exec_module(_mod)
