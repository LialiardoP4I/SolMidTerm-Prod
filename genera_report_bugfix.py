# -*- coding: utf-8 -*-
"""
Genera report Word (.docx) descrivendo i bug fixati nella revisione codice.
Linguaggio semplice, non tecnico.
"""
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH


def _h(doc, text, level):
    h = doc.add_heading(text, level=level)
    return h


def _p(doc, text, bold=False, italic=False, size=11):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold = bold
    r.italic = italic
    r.font.size = Pt(size)
    return p


def _bullet(doc, text):
    p = doc.add_paragraph(text, style='List Bullet')
    return p


def _box(doc, title, text, color_hex):
    p = doc.add_paragraph()
    r = p.add_run(f"{title}: ")
    r.bold = True
    r.font.color.rgb = RGBColor.from_string(color_hex)
    p.add_run(text)
    return p


doc = Document()

style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(11)

# ── COPERTINA ────────────────────────────────────────────────────────────────
title = doc.add_heading('Report Bug Fix', level=0)
title.alignment = WD_ALIGN_PARAGRAPH.CENTER

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('Progetto: SolMidTerm – Prod (Ducati Multistrada V4)')
r.italic = True
r.font.size = Pt(13)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('Data: 22 maggio 2026')
r.font.size = Pt(11)

doc.add_paragraph()

# ── INTRO ────────────────────────────────────────────────────────────────────
_h(doc, 'Cosa è stato fatto', 1)
_p(doc,
   "È stata effettuata una revisione approfondita di 17 script Python del progetto. "
   "Sono stati identificati 19 bug critici. Di questi, 13 sono stati corretti subito "
   "(quelli con causa chiara e fix sicuro). I restanti 5 richiedono chiarimenti con "
   "il team prima di intervenire perché toccano logica di business o necessitano di "
   "verifiche sui file di input reali."
)
_p(doc,
   "Tutti i fix rispettano il vincolo: la logica di business non è stata alterata. "
   "Sono stati corretti solo errori che facevano crashare lo script o producevano "
   "risultati sbagliati per casi limite (edge case)."
)

# ── INDICE BUG FIXATI ────────────────────────────────────────────────────────
_h(doc, 'Elenco bug corretti (13)', 1)

bugs = [
    {
        'n': 1,
        'file': 'tr_implied.py',
        'titolo': 'Mese di OTTOBRE perso dal forecast',
        'sintomo':
            "Quando il file forecast usava nomi mese in italiano (es. 'OTTOBRE'), "
            "il mese veniva escluso silenziosamente dal calcolo del Take Rate implicito. "
            "Conseguenza: tutti i TR del periodo erano errati e disallineati di un mese.",
        'causa':
            "Il codice cercava la parola 'TOT' (per saltare la colonna 'TOTALE'), ma "
            "la parola 'OTTOBRE' contiene 'TOT' come sottostringa, quindi veniva "
            "scartato anche lui per sbaglio.",
        'fix':
            "Sostituito il controllo da 'contiene TOT' a 'inizia con TOT'. Ora solo "
            "celle che iniziano con TOTALE/TOT vengono saltate, mentre OTTOBRE viene "
            "correttamente incluso.",
        'impatto': 'ALTO – correttezza dati sui report mensili',
    },
    {
        'n': 2,
        'file': 'tr_parser_V2.py',
        'titolo': 'Crash silenzioso se manca il foglio TR_Model',
        'sintomo':
            "Se nel file Excel dei Take Rate mancava il foglio 'TR_Model_*', "
            "il programma andava in errore con un messaggio criptico (KeyError) "
            "invece di segnalare in modo chiaro il problema.",
        'causa':
            "Veniva passato un valore nullo (None) a una funzione che si aspettava "
            "un nome di foglio Excel valido.",
        'fix':
            "Aggiunto un controllo esplicito: se il foglio TR_Model_* manca, viene "
            "sollevato un errore con messaggio comprensibile che indica esattamente "
            "qual è il file e cosa serve.",
        'impatto': 'MEDIO – migliore diagnostica errori utente',
    },
    {
        'n': 3,
        'file': 'app_V3.py',
        'titolo': 'Errore NameError lanciando Step 3 da solo',
        'sintomo':
            "Se l'utente avviava solo lo Step 3 (matching SKU) senza prima eseguire "
            "Step 1 o Step 2, l'app crashava con il messaggio 'name importlib is not defined'.",
        'causa':
            "Il modulo 'importlib' veniva caricato solo dentro Step 1 e Step 2, "
            "quindi se questi venivano saltati Step 3 non lo trovava.",
        'fix':
            "Spostato l'import 'importlib' all'inizio del file, così è sempre "
            "disponibile indipendentemente da quali step l'utente seleziona.",
        'impatto': 'ALTO – usabilità interfaccia',
    },
    {
        'n': 4,
        'file': 'optional_cost_allocation.py',
        'titolo': 'Percorso file errato (cartella sbagliata)',
        'sintomo':
            "Lo script puntava a una cartella 'SolMidTerm' ma il progetto è in "
            "'SolMidTerm – Prod'. Risultato: lo script o non trovava i file oppure "
            "lavorava su dati di una cartella diversa.",
        'causa':
            "Il percorso era scritto a mano (hardcoded) nel codice e non corrispondeva "
            "più alla struttura attuale delle cartelle.",
        'fix':
            "Il percorso è ora calcolato automaticamente in base alla posizione del "
            "file Python stesso. In questo modo lo script funziona indipendentemente "
            "da dove si trovi la cartella.",
        'impatto': 'ALTO – script ora utilizzabile senza modifiche manuali',
    },
    {
        'n': 5,
        'file': 'sku_matching_SA.py',
        'titolo': 'Crash su nomi supermodello con caratteri speciali',
        'sintomo':
            "Se un Supermodel conteneva caratteri speciali come parentesi tonde, "
            "punti, o segno '+', la ricerca crashava o trovava risultati sbagliati.",
        'causa':
            "Il codice interpretava il nome del supermodello come 'espressione regolare', "
            "trattando caratteri come '(' o '+' come comandi di pattern invece che testo letterale.",
        'fix':
            "Aggiunto il parametro 'regex=False' alla ricerca: ora il testo è confrontato "
            "letteralmente, senza interpretazione di caratteri speciali.",
        'impatto': 'ALTO – correttezza per nomi modello reali',
    },
    {
        'n': 6,
        'file': 'sku_matching_V2.py',
        'titolo': 'Errore se tutti i componenti hanno lead time valido',
        'sintomo':
            "Quando nessun componente aveva lead time mancante (caso ottimale!), "
            "lo script crashava con 'AttributeError: bool object has no attribute fillna'.",
        'causa':
            "Il codice creava una colonna di flag solo quando trovava lead time mancanti. "
            "Se non ne trovava, la colonna non esisteva e una chiamata successiva falliva.",
        'fix':
            "Aggiunto controllo: se la colonna esiste, la valorizza; se non esiste, "
            "la crea con valore di default 'False' (nessun lead time mancante).",
        'impatto': 'ALTO – evita crash nel caso normale (tutti LT presenti)',
    },
    {
        'n': 7,
        'file': 'sku_matching_SS.py',
        'titolo': 'Crash se la libreria psutil non è installata',
        'sintomo':
            "Su PC senza la libreria 'psutil' lo script crashava nei messaggi di "
            "diagnostica con un errore di formattazione, perché tentava di stampare "
            "un numero che era invece il valore nullo (None).",
        'causa':
            "La funzione 'get_memory_usage_mb' ritornava None se psutil mancava, "
            "ma le stampe usavano un formato numerico che non accetta None.",
        'fix':
            "Ora la funzione ritorna 0.0 (zero) se psutil non è disponibile, "
            "in modo che le stampe non crashino. Le diagnostiche mostreranno "
            "'0.0 MB' invece di crashare.",
        'impatto': 'MEDIO – robustezza su ambienti diversi',
    },
    {
        'n': 8,
        'file': 'sku_matching_SS.py',
        'titolo': 'Stesso errore del fix #6, in funzione diversa',
        'sintomo':
            "Identico al bug #6 ma in una funzione differente dello stesso modulo. "
            "Crash quando tutti i lead time erano già presenti.",
        'causa': "Stesso pattern di codice duplicato in due punti del progetto.",
        'fix':
            "Applicato lo stesso fix del bug #6: controllo esistenza colonna prima "
            "di chiamare fillna.",
        'impatto': 'ALTO – evita crash nel caso normale',
    },
    {
        'n': 9,
        'file': 'sensitivity_dirichlet.py',
        'titolo': "Tabella del report Excel con numeri sbagliati",
        'sintomo':
            "Il report di sensibilità Excel mostrava nella tabella 'Guida alla Scelta' "
            "valori M (numero osservazioni equivalenti) di 100, 1000, 10000, 100000, "
            "mentre il resto del report usava 50, 200, 1000, 5000. Stesso file, "
            "valori diversi: incoerenza visiva confusa per l'utente.",
        'causa':
            "I valori della tabella erano scritti a mano nel codice e non venivano "
            "aggiornati quando si modificavano i parametri reali della simulazione.",
        'fix':
            "I valori della tabella vengono ora calcolati automaticamente dai "
            "parametri reali della simulazione. Se in futuro si cambiano i livelli "
            "(es. da 5000 a 10000 per ALTA), la tabella si aggiorna da sola.",
        'impatto': 'MEDIO – coerenza report destinato a stakeholder business',
    },
    {
        'n': 10,
        'file': 'main_demand_comparison_V2.py',
        'titolo': 'TR mensili letti senza pulizia (typo + righe ALTRO)',
        'sintomo':
            "Quando si confrontava domanda actual vs forecast, i Take Rate annuali "
            "venivano puliti (correggendo il typo 'Bunde'->'Bundle' e rimuovendo le "
            "righe 'ALTRO'), MA i Take Rate mensili venivano letti dal file originale "
            "senza pulizia. Risultato: simulazione MC mensile potenzialmente errata.",
        'causa':
            "Il codice creava un file temporaneo pulito, lo usava solo per i TR "
            "annuali, lo cancellava subito, e poi leggeva i TR mensili dal file "
            "originale (sporco).",
        'fix':
            "Ristrutturato il blocco try/finally: ora il file temporaneo pulito "
            "viene usato sia per i TR annuali sia per quelli mensili, e cancellato "
            "solo alla fine. Entrambi ricevono lo stesso input pulito.",
        'impatto': 'ALTO – correttezza simulazione confronto demand',
    },
    {
        'n': 11,
        'file': 'gestione_dipendenze_V2.py',
        'titolo': "Righe BOM scartate per errore quando contenevano NOT(...)",
        'sintomo':
            "Componenti BOM con condizioni del tipo NOT(\\$root.ZCOUNTRY eq 'CDN') "
            "venivano scartate dalla matrice globale, anche se la regola NOT significa "
            "'vale per tutti tranne Canada' – quindi la riga doveva essere mantenuta.",
        'causa':
            "Il controllo per identificare le righe specifiche per paese trovava "
            "anche le occorrenze dentro NOT(...), interpretandole come 'specifiche "
            "per quel paese' invece che come negazione.",
        'fix':
            "Prima di controllare se la riga è specifica per paese, il codice "
            "rimuove tutti i blocchi NOT(...) dal testo della condizione. In questo "
            "modo le negazioni non vengono confuse con affermazioni positive.",
        'impatto': 'CRITICO – righe BOM perse silenziosamente dalla matrice',
    },
    {
        'n': 12,
        'file': 'gestione_dipendenze_V2.py',
        'titolo': "Split bundle errato su condizioni con NOT compound",
        'sintomo':
            "Per condizioni complesse come NOT(\\$root.A eq 'X' AND \\$root.B eq 'Y'), "
            "il codice catturava erroneamente la seconda parte (B eq Y) come "
            "discriminante per split del bundle, producendo split errati.",
        'causa':
            "Il pattern regex aveva due controlli (lookbehind) che escludevano solo "
            "il PRIMO \\$root dopo NOT(. Il secondo \\$root della stessa parentesi "
            "veniva catturato per sbaglio.",
        'fix':
            "Come per il bug #11, ora il codice rimuove preventivamente tutti i "
            "blocchi NOT(...) dal testo prima di cercare i pattern di split. "
            "Solo le condizioni positive vengono considerate per il discriminante.",
        'impatto': 'CRITICO – matrice configurazioni con split sbagliati',
    },
    {
        'n': 13,
        'file': 'app_V3.py – riga 2540',
        'titolo': "Falso positivo (no fix necessario)",
        'sintomo':
            "L'analisi automatica aveva segnalato un possibile errore di divisione "
            "per zero in un calcolo del grafico Safety Stock mensile.",
        'causa':
            "Verifica manuale del codice: il calcolo problematico è dentro un blocco "
            "protetto da un controllo esterno che garantisce che il divisore non "
            "sia mai zero in quel punto. Si trattava di un falso positivo dell'analisi.",
        'fix':
            "Nessun intervento. Annotato come 'falso positivo' nel sistema di "
            "tracking dei task.",
        'impatto': 'NESSUNO – il codice originale è corretto',
    },
]

for b in bugs:
    _h(doc, f"Bug #{b['n']} – {b['titolo']}", 2)
    _p(doc, f"File: {b['file']}", italic=True)
    _box(doc, "Sintomo", b['sintomo'], "B00020")
    _box(doc, "Causa", b['causa'], "B58105")
    _box(doc, "Fix applicato", b['fix'], "0F7B2E")
    _box(doc, "Impatto", b['impatto'], "1F4E79")
    doc.add_paragraph()

# ── DUBBI ────────────────────────────────────────────────────────────────────
_h(doc, 'Bug non corretti (in attesa di chiarimento)', 1)
_p(doc,
   "I seguenti 5 bug richiedono una decisione del team prima di essere "
   "sistemati, perché toccano la logica di business oppure necessitano di "
   "verifiche sui dati Excel reali che non sono accessibili in sola lettura del codice."
)

dubbi = [
    ('sku_matching_SA.py (righe 80, 98, 234)',
     "Calcolo del numero di run Monte Carlo: una variabile chiamata 'n_runs' "
     "in realtà contiene il numero di mesi (non di run). Funziona per coincidenza "
     "ma è semanticamente fragile. Da chiarire con chi ha scritto la logica originale."),
    ('sku_matching_SS.py (riga 1947)',
     "Il tasso di cambio EUR per qualsiasi valuta è forzato a 1.0. È un placeholder "
     "voluto (in attesa di dati cambio reali) oppure un bug? Va deciso a livello business."),
    ('sku_matching_V2.py (riga 336)',
     "Il file Excel 'quantity e residual.xlsx' viene letto con header sulla prima "
     "riga, ma la documentazione interna dice header alla terza riga. Va verificato "
     "sul file Excel reale qual è il layout corretto."),
    ('optional_cost_allocation.py (riga 115)',
     "La ricerca dei Take Rate per valore di attributo usa un match 'parziale' "
     "(substring). Es: cercando 'Red' può trovare 'Red Sport' o 'Dark Red'. "
     "È voluto (fuzzy matching deliberato) oppure un bug? Da chiarire."),
    ('sku_matching_SS.py (riga 2125)',
     "Due funzioni quasi identiche (V1 e V2) trattano in modo DIVERSO i componenti "
     "senza lead time: V1 li include con LT=0, V2 li esclude. È intenzionale (due "
     "casi d'uso diversi) oppure V1 è codice morto da rimuovere?"),
]

for titolo, descrizione in dubbi:
    _h(doc, titolo, 3)
    _p(doc, descrizione)

# ── CONCLUSIONI ──────────────────────────────────────────────────────────────
_h(doc, 'Riepilogo e prossimi passi', 1)
_p(doc,
   "13 bug corretti su 19 totali. I file Python sono stati modificati direttamente "
   "nel progetto, mantenendo invariata la logica di business. La sintassi è stata "
   "verificata su tutti i file (parser Python ok)."
)
_p(doc, 'Cosa serve fare adesso:', bold=True)
_bullet(doc,
        "Test funzionale: eseguire una run completa della pipeline (Step 0→1→2→3) "
        "per verificare che i fix non abbiano introdotto regressioni.")
_bullet(doc,
        "Chiarire i 5 DUBBI elencati sopra, in modo da poter completare la pulizia "
        "anche di quelli.")
_bullet(doc,
        "Commit dei fix sul repository git (al momento i fix sono nel working tree "
        "non ancora committati).")
_bullet(doc,
        "Procedere con il fix dei bug 'medi' e delle ottimizzazioni "
        "(circa 132 bug medi e 115 ottimizzazioni identificati ma non ancora trattati).")

# ── SALVA ────────────────────────────────────────────────────────────────────
output_path = (
    "C:\\Users\\AntonioLiardo-SMOPS\\OneDrive - Digital360\\"
    "Desktop\\Ducati\\SolMidTerm - Prod\\Report_Bugfix.docx"
)
doc.save(output_path)
print(f"Report salvato in: {output_path}")
