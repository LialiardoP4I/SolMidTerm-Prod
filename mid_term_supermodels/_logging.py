# -*- coding: utf-8 -*-
"""Log uniforme per tutti i moduli.

In parole semplici: fornisce un modo unico per scrivere i messaggi di
avanzamento e diagnostica (i log), cosi' che ogni passo della pipeline stampi le
informazioni nello stesso formato.
"""
import logging


def get_logger(name: str) -> logging.Logger:
    """Ritorna logger mid_term.<name>."""
    return logging.getLogger(f"mid_term.{name}")


def configure_default_logging(
    level: int = logging.INFO,
    format_str: str = "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
) -> None:
    """Configura handler stdout sul logger root mid_term.

    Il consumer puo' sovrascrivere o ignorare questa configurazione
    fornendo la propria via logging.basicConfig o handler dedicati.
    """
    handler = logging.StreamHandler()
    handler.setFormatter(logging.Formatter(format_str))
    root = logging.getLogger("mid_term")
    root.handlers.clear()
    root.addHandler(handler)
    root.setLevel(level)
    root.propagate = False
