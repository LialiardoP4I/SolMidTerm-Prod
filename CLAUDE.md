# Python Code Review — Istruzioni Claude Code

## VINCOLO ASSOLUTO
Non modificare MAI la logica di business, gli algoritmi, i calcoli,
i flussi condizionali, le strutture dati e i valori di output.
Ogni intervento deve produrre un comportamento identico all'originale.
Se hai dubbi su un blocco, segnalalo senza toccarlo.

## Obiettivo
Revisione approfondita in sola lettura + proposta di ottimizzazione.
NON applicare modifiche autonomamente: presenta sempre un diff commentato
e attendi conferma esplicita prima di scrivere qualsiasi file.

---

## PLUGIN CONSIGLIATI
Installa questi plugin prima di iniziare (una tantum):

```powershell
/plugin install code-review@claude-plugins-official
/plugin install security-guidance@claude-plugins-official
/plugin install superpowers@claude-plugins-community
```

---

## PROMPT — REVISIONE SINGOLO FILE

Copia e incolla questo prompt in Claude Code, sostituendo `[nome_file.py]`:

```
Esegui una revisione approfondita del file [nome_file.py] seguendo
esattamente queste fasi nell'ordine indicato. Non modificare nulla
senza la mia approvazione esplicita.

## FASE 1 — COMPRENSIONE
Descrivi in 5-10 righe cosa fa il codice, qual è la logica centrale
e quali sono gli input/output principali.

## FASE 2 — BUG E CORRETTEZZA
Identifica:
- Errori certi (eccezioni non gestite, logica errata, off-by-one, ecc.)
- Rischi potenziali (race condition, divisioni per zero, None non controllati)
- Comportamenti inattesi in edge case
Per ogni problema: riga, descrizione, gravità (CRITICO / MEDIO / BASSO).

## FASE 3 — OTTIMIZZAZIONE (senza alterare la logica)
Individua:
- Ridondanze e codice morto (variabili inutilizzate, blocchi mai raggiunti)
- Inefficienze computazionali (loop annidati migliorabili, operazioni ripetute)
- Import inutilizzati
- Strutture dati non ottimali per l'uso che ne viene fatto
Per ogni punto: mostra il PRIMA e il DOPO proposto.

## FASE 4 — QUALITÀ E MANUTENIBILITÀ
Verifica:
- Type hints mancanti o errati
- Docstring assenti su funzioni/classi pubbliche
- Nomi di variabili/funzioni non descrittivi
- Violazioni PEP8 sostanziali (non stilistiche minori)

## FASE 5 — VERIFICA FUNZIONALE
Proponi 3-5 casi di test (input → output atteso) che coprano:
- Il caso normale
- Almeno un edge case
- Almeno un caso di errore atteso
Non scrivere il codice di test, solo i casi logici.

## FASE 6 — RIEPILOGO
Tabella finale con:
| Categoria        | N. problemi | Priorità intervento |
|------------------|-------------|---------------------|
| Bug certi        |             |                     |
| Rischi potenziali|             |                     |
| Ottimizzazioni   |             |                     |
| Qualità codice   |             |                     |

Lista ordinata per priorità degli interventi suggeriti.

## REGOLA FINALE
Tutto ciò che proponi nella Fase 3 e 4 deve produrre
un comportamento a runtime IDENTICO all'originale.
Se non sei certo al 100%, esplicita il dubbio invece di procedere.
```

---

## PROMPT — REVISIONE BATCH (più file)

```
Analizza tutti i file .py nella directory corrente uno per uno,
seguendo il protocollo di revisione definito in CLAUDE.md.
Per ogni file, produci il report completo prima di passare al successivo.
Al termine, aggrega i risultati in una tabella riepilogativa cross-file:

| File | Bug critici | Bug medi | Ottimizzazioni | Priorità complessiva |
|------|-------------|----------|----------------|----------------------|

Ordina la tabella per priorità decrescente.
Non modificare nessun file. Solo analisi e report.
```

---

## FLUSSO DI LAVORO CONSIGLIATO

```powershell
# 1. Posizionati nella cartella degli script
cd C:\tuoi-script-python

# 2. Avvia Claude Code
claude

# 3. Attiva il plugin di review
/review

# 4. Incolla il prompt operativo (singolo file o batch)

# 5. Valuta il report prodotto

# 6. Approva solo le modifiche che vuoi applicare, es:
#    "Applica solo le modifiche della Fase 3, punto 2"
```

---

## HOOK AUTOMATICO (opzionale)
Per eseguire mypy + flake8 automaticamente dopo ogni modifica approvata,
aggiungi questo blocco in `~/.claude/settings.json`:

```json
{
  "hooks": {
    "PostToolUse": [
      { "command": "python -m mypy {file} --ignore-missing-imports" },
      { "command": "python -m flake8 {file} --max-line-length=120" }
    ]
  }
}
```

Assicurati di avere mypy e flake8 installati:
```powershell
pip install mypy flake8
```

---

## NOTE
- Questo file (CLAUDE.md) viene letto automaticamente da Claude Code
  all'avvio di ogni sessione nella cartella in cui si trova.
- Mettilo nella root della cartella dei tuoi script Python.
- Il vincolo di non modificare la logica è enforced dal prompt,
  non da meccanismi tecnici: verifica sempre il diff prima di approvare.
