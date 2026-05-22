# Revisione Codice Python — Setup Claude Code + Superpowers

## VINCOLO ASSOLUTO
Non modificare MAI:
- La logica di business e gli algoritmi
- I calcoli e le formule
- I flussi condizionali (if/else/match/try/except)
- Le strutture dati e i valori restituiti
- Qualsiasi comportamento osservabile dall'esterno del modulo

Il comportamento a runtime deve essere identico prima e dopo ogni intervento.
In caso di dubbio su un blocco: segnalalo con un commento, non toccarlo.

## Regola operativa
NON applicare modifiche autonomamente.
Produci prima il report completo, poi attendi conferma esplicita.
Ogni modifica proposta deve mostrare PRIMA e DOPO affiancati.

---

## INSTALLAZIONE PLUGIN (una tantum)

Lancia Claude Code dalla cartella degli script, poi esegui:

    /plugin install superpowers@claude-plugins-official
    /plugin install security-guidance@claude-plugins-official

Verifica con:

    /plugins

---

## AVVIO SESSIONE

    cd C:\percorso\tuoi-script-python
    claude

Claude Code legge questo file automaticamente. Nessun altro setup necessario.

---

## PROMPT — REVISIONE SINGOLO FILE

Incolla questo prompt in Claude Code sostituendo [nome_file.py]:

------------------------------------------------------------
Leggi il file [nome_file.py] nella directory corrente.
Applica il protocollo di revisione definito in CLAUDE.md.
Non modificare nessun file. Solo analisi e report.

### FASE 1 — COMPRENSIONE
Descrivi in massimo 10 righe:
- Cosa fa il codice
- Qual è la logica centrale
- Quali sono input e output principali
- Dipendenze esterne rilevanti

### FASE 2 — BUG E CORRETTEZZA
Analizza riga per riga e identifica:
- Errori certi: eccezioni non gestite, logica errata, off-by-one,
  variabili non inizializzate, type mismatch
- Rischi potenziali: divisioni per zero, None non controllati,
  race condition, overflow, accessi fuori indice
- Edge case non gestiti

Per ogni problema:
| Riga | Descrizione | Gravità (CRITICO / MEDIO / BASSO) |

### FASE 3 — OTTIMIZZAZIONE (senza alterare la logica)
Identifica esclusivamente interventi che non cambiano il comportamento:
- Codice morto: variabili mai usate, blocchi irraggiungibili
- Import non utilizzati
- Operazioni ridondanti o ripetute inutilmente
- Loop migliorabili con list comprehension o built-in equivalenti
- Strutture dati non ottimali per l'uso specifico

Per ogni punto mostra:
  PRIMA:  [codice attuale]
  DOPO:   [codice proposto]
  MOTIVO: [spiegazione in una riga]

### FASE 4 — QUALITÀ E LEGGIBILITÀ
Segnala (senza correggere autonomamente):
- Type hints mancanti su funzioni pubbliche
- Docstring assenti su funzioni/classi pubbliche
- Nomi di variabili non descrittivi
- Violazioni PEP8 sostanziali (non stilistiche minori)

### FASE 5 — CASI DI TEST LOGICI
Proponi 4-5 scenari di test (NON scrivere il codice di test):
- 1 caso normale atteso
- 1+ edge case
- 1 caso di errore/eccezione attesa
Formato: input → output atteso → motivazione

### FASE 6 — RIEPILOGO FINALE
Tabella riepilogativa:
| Categoria          | N. problemi trovati | Priorità intervento |
|--------------------|---------------------|---------------------|
| Bug certi          |                     |                     |
| Rischi potenziali  |                     |                     |
| Ottimizzazioni     |                     |                     |
| Qualità/leggibilità|                     |                     |

Lista interventi ordinata per priorità decrescente.

REGOLA FINALE: tutto ciò che proponi in Fase 3 e 4 deve produrre
un comportamento a runtime IDENTICO all'originale.
Se non sei certo al 100%, scrivi "DUBBIO:" e spiega, senza intervenire.
------------------------------------------------------------

---

## PROMPT — REVISIONE BATCH (più file)

Incolla questo per analizzare tutti gli script della cartella:

------------------------------------------------------------
Elenca tutti i file .py nella directory corrente.
Analizzali uno per uno seguendo il protocollo in CLAUDE.md.
Per ogni file completa il report prima di passare al successivo.
Non modificare nessun file durante l'analisi.

Al termine produci una tabella riepilogativa cross-file:
| File | Bug critici | Bug medi/bassi | Ottimizzazioni | Priorità |
|------|-------------|----------------|----------------|----------|

Ordina per priorità decrescente (i file più critici prima).
------------------------------------------------------------

---

## APPLICARE LE MODIFICHE (dopo aver letto il report)

Per applicare interventi specifici dopo aver valutato il report:

    "Applica solo le modifiche di Fase 3, punti 1 e 3 del file [nome].
     Mostrami il diff completo prima di scrivere qualsiasi file."

Per bloccare tutto tranne un intervento preciso:

    "Applica solo la rimozione degli import inutilizzati nel file [nome].
     Non toccare nient'altro."

---

## HOOK AUTOMATICO post-modifica (opzionale)

Dopo ogni modifica approvata, esegui automaticamente mypy e flake8.
Aggiungi in %USERPROFILE%\.claude\settings.json:

    {
      "hooks": {
        "PostToolUse": [
          { "command": "python -m mypy {file} --ignore-missing-imports" },
          { "command": "python -m flake8 {file} --max-line-length=120" }
        ]
      }
    }

Installa i tool se non presenti:

    pip install mypy flake8
