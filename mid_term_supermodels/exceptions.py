# -*- coding: utf-8 -*-
"""Errori specifici della pipeline.

In parole semplici: definisce i tipi di errore (configurazione mancante, file di
input assente, cambio valuta mancante, ecc.) cosi' che, quando qualcosa non va,
il programma si fermi con un messaggio chiaro invece di un errore generico.
"""


class MidTermError(Exception):
    """Base di tutte le eccezioni del package."""


class ConfigurationError(MidTermError):
    """Configurazione invalida: path mancanti, parametri fuori range."""


class InputFileError(MidTermError):
    """File di input atteso ma assente o malformato."""


class TRParseError(MidTermError):
    """Errore parsing Take Rate / Mappatura (step 0)."""


class BOMParseError(MidTermError):
    """Errore parsing BOM o espansione dipendenze (step 1)."""


class SimulationError(MidTermError):
    """Errore Monte Carlo: mix non valido, Dirichlet con alpha NaN, ecc. (step 2)."""


class MatchingError(MidTermError):
    """Errore matching SKU o calcolo safety stock (step 3-4)."""
