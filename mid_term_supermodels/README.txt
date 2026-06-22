==============================================================================
 CALCOLO SAFETY STOCK - DUCATI (MID TERM)
==============================================================================

Sistema per il calcolo della scorta di sicurezza (safety stock) dei componenti
delle moto Ducati Multistrada, su orizzonte di medio termine.


------------------------------------------------------------------------------
 1. A COSA SERVE (in parole semplici)
------------------------------------------------------------------------------

Una moto e' composta da centinaia di componenti. Alcuni sono sempre presenti,
altri dipendono dalla configurazione scelta dal cliente (colore, pacchetti
optional, ecc.).

La domanda futura di moto e' incerta: non sappiamo in anticipo quante moto
verranno vendute ne' con quali optional. Di conseguenza e' incerta anche la
domanda di ogni singolo componente.

Questo sistema risponde alla domanda:

   "Quanta scorta di sicurezza devo tenere per ogni componente, per non
    rimanere senza pezzi durante il tempo di approvvigionamento, tenendo
    conto dell'incertezza della domanda?"

Lo fa simulando migliaia di scenari di vendita possibili (metodo Monte Carlo)
e misurando, per ogni componente, quanta scorta serve per coprire un livello
di servizio elevato (es. il 99 percentile della domanda).


------------------------------------------------------------------------------
 2. GLOSSARIO (per non tecnici)
------------------------------------------------------------------------------

  Componente / SKU   Singolo pezzo che compone la moto (codice articolo).
  BOM (Distinta Base) Elenco di tutti i componenti che servono per costruire
                     una configurazione di moto.
  Configurazione     Una specifica versione di moto: modello + colore +
                     optional scelti.
  Optional / Pack    Accessori opzionali (es. Touring Pack, Top Case).
  Take Rate (TR)     Percentuale di clienti che sceglie un certo optional/
                     colore. Es: il 30% sceglie il rosso.
  Forecast           Previsione di quante moto verranno vendute, mese per mese.
  Monte Carlo        Tecnica che genera migliaia di scenari casuali (ma
                     realistici) per stimare l'incertezza.
  Lead Time          Tempo che serve per ricevere un componente dopo averlo
                     ordinato (in mesi).
  Safety Stock       Scorta di sicurezza: pezzi tenuti a magazzino in piu' per
                     coprire i picchi di domanda.
  Percentile (P99)   Livello di servizio: P99 = la scorta copre il 99% degli
                     scenari simulati.
  Esclusioni         Combinazioni di optional incompatibili tra loro (non
                     possono coesistere).
  Affidabilita' TR   Quanto ci fidiamo dei take rate dichiarati: piu' alta = la
                     simulazione segue di piu' il TR; piu' bassa = piu'
                     variabilita'.


------------------------------------------------------------------------------
 3. IL FLUSSO COMPLETO IN 5 FASI
------------------------------------------------------------------------------

Il sistema esegue in sequenza 5 step (numerati 0->4). Ogni step usa i
risultati del precedente.

   Step 0  ->  Step 1  ->  Step 2      ->  Step 3     ->  Step 4
   Dati TR     Distinta    Simulazione     Abbinamento    Calcolo
   +forecast   Base (BOM)  Monte Carlo     SKU<->scenari  Safety Stock
   +esclusioni                                            + prezzi


STEP 0 - Caricamento dati (TR, forecast, esclusioni)
  Legge:
    - il file TR (mix modelli + take rate di ogni caratteristica/optional);
    - il forecast di domanda mensile (quante moto al mese);
    - le esclusioni (optional incompatibili);
    - quanti mesi di orizzonte servono (derivati dai lead time piu' lunghi).
  Inoltre applica l'affidabilita' TR: regola quanto la simulazione deve
  "seguire" i take rate dichiarati rispetto al lasciare piu' variabilita'
  casuale (parametro tecnico: moltiplicatore Dirichlet).

STEP 1 - Costruzione della Distinta Base (BOM)
  Per ogni modello moto (scoperti automaticamente in Input/MODEL/):
    1. Pulizia dei file intermedi vecchi (evita risultati "sporchi" da run
       precedenti).
    2. Mappatura unificata: fonde le mappature dei singoli modelli in una sola
       coerente.
    3. Controlli qualita' (prezzi, residual, tracciabilita' caratteristiche).
    4. Costruzione matrice BOM per ogni modello.
    5. Espansione esclusioni: le righe con optional incompatibili vengono
       divise in righe valide (es. "Touring + Radar" incompatibili -> due
       righe separate).
    6. Consolidamento: unisce tutte le configurazioni in un catalogo unico,
       deduplicato.
  Risultato: il catalogo tutte_righe_univoche_V2.xlsx.

STEP 2 - Simulazione Monte Carlo
  Genera migliaia di scenari di vendita (numero impostabile, n_runs), mese per
  mese:
    - per ogni moto simulata estrae una configurazione (colore, optional)
      rispettando i take rate e l'affidabilita' impostata;
    - risolve i conflitti tra optional incompatibili in modo probabilistico
      (l'optional piu' richiesto ha piu' probabilita' di "vincere");
    - deriva i colori dei pack dal colore carena (es. il colore delle Panniers
      segue il colore della moto), inclusa la gestione delle versioni Rally
      (V22E).
  Risultato: una grande tabella che, per ogni configurazione, indica quante
  volte e' uscita in ogni scenario e in ogni mese.

STEP 3 - Abbinamento SKU <-> scenari (Matching)
  Confronta il catalogo componenti (Step 1) con gli scenari simulati (Step 2):
  per ogni componente capisce in quali configurazioni compare e quindi quanta
  domanda riceve in ogni scenario/mese.
  I componenti senza lead time valido (mancante o = 0) non possono essere
  dimensionati e vengono messi da parte in un file dedicato
  (sku_senza_lead_time.xlsx).

STEP 4 - Calcolo Safety Stock + valorizzazione
  Per ogni componente:
    - aggrega la domanda sui mesi di lead time (quanto serve durante l'attesa
      di rifornimento);
    - calcola la safety stock al percentile richiesto (es. P99);
    - produce sia il dato per SKU sia quello per mese;
    - aggancia i prezzi (convertiti in EUR tramite i tassi di cambio) per
      stimare il valore economico della safety stock.
  Risultato: i file finali di safety stock (vedi sezione Output).


------------------------------------------------------------------------------
 4. FILE DI INPUT (cartella mid_term/Input/)
------------------------------------------------------------------------------

  TR TOTALV21E - Copia.xlsx          Take rate: mix modelli e caratteristiche/
                                     optional.
  Total_demand.xlsx                  Forecast domanda moto, mese per mese.
  esclusioni.xlsx                    Coppie di optional incompatibili.
  States.xlsx                        Stati/configurazioni usati nei controlli
                                     qualita'.
  quantity e residual.xlsx           Quantita' per moto + dati residual / lead
                                     time.
  residual_lead_time.xlsx            Lead time per componente.
  MODEL/<MODELLO>/Mappatura_<MODELLO>.xlsx   Mappatura caratteristiche del
                                     modello.
  MODEL/<MODELLO>/BOM_150<MODELLO>.xlsx      Distinta base del modello.
  Dati Logistica/PREZZI.XLSX         Prezzi dei componenti.
  Cambio Valuta.xlsx                 Tassi di cambio verso EUR.
  affidabilita_config.json           Mappa Affidabilita' -> moltiplicatore
                                     (BASSA/MEDIA/ALTA).
  run_config.json                    Parametri della simulazione (vedi sotto).

  NOTA: i modelli (es. V21E, V22E, V24T) vengono scoperti automaticamente:
  basta che esista la sottocartella in Input/MODEL/ con dentro la sua
  Mappatura_<modello>.xlsx.


------------------------------------------------------------------------------
 5. FILE DI OUTPUT (cartella mid_term/Output/)
------------------------------------------------------------------------------

  sku_safety_stock_leadtime_PER_SKU_V2.xlsx   Safety stock per componente
                                     (aggregata sul lead time) + valore
                                     economico.
  sku_safety_stock_PER_MESE_V2.xlsx           Safety stock per componente,
                                     dettaglio mensile.
  sku_senza_lead_time.xlsx           Componenti esclusi perche' senza lead time
                                     valido (con motivo).
  tutte_righe_univoche_V2.xlsx (in Input/)    Catalogo configurazioni
                                     consolidato (Step 1).


------------------------------------------------------------------------------
 6. PARAMETRI DELLA SIMULAZIONE (run_config.json)
------------------------------------------------------------------------------

  n_runs                    Numero di scenari simulati. Piu' alto = piu'
                            preciso ma piu' lento.   [intero > 0, es. 200]
  random_seed               Seme casuale: stesso seme -> stessi risultati
                            (riproducibilita').       [intero]
  start_month               Mese di partenza dell'orizzonte.
                            [nome mese presente nel forecast, es. MAY 2026]
  percentile_safety_stock   Livello di servizio della safety stock.
                            [1-99, es. 99]
  affidabilita_model        Affidabilita' del mix modelli.
                            [NULLA / BASSA / MEDIA / ALTA]
  affidabilita_optional     Affidabilita' degli optional.
                            [NULLA / BASSA / MEDIA / ALTA]

  Il file e' obbligatorio e completo: se manca un parametro o un valore non e'
  valido, il sistema si ferma con un messaggio chiaro (nessun default
  silenzioso).


------------------------------------------------------------------------------
 7. COME SI ESEGUE
------------------------------------------------------------------------------

  Requisiti: Python 3 con pandas, numpy, openpyxl.

      cd "...\SolMidTerm - Prod\mid_term"
      python run_pipeline.py

  Il launcher imposta automaticamente la cartella di lavoro su mid_term/,
  quindi tutti i percorsi Input/ e Output/ vengono risolti li' dentro. Al
  termine stampa un riepilogo (numero righe per SKU, percentile, affidabilita',
  file prodotti).


------------------------------------------------------------------------------
 8. STRUTTURA DEL CODICE (cartella mid_term/)
------------------------------------------------------------------------------

  run_pipeline.py   Punto di avvio: legge i parametri e lancia la pipeline.
  pipeline.py       Orchestratore: esegue gli step 0->4 in sequenza.
  config.py         Configurazione (percorsi file, parametri) con validazione.
  tr.py             Lettura file TR: mix modelli, caratteristiche,
                    affidabilita'.
  bom.py            Costruzione Distinta Base: mappature, controlli qualita',
                    esclusioni, catalogo.
  montecarlo.py     Simulazione Monte Carlo: campionamento, conflitti, colori
                    pack.
  matching.py       Abbinamento componenti <-> scenari e calcolo statistiche/
                    safety stock.
  results.py        Strutture dati dei risultati di ogni step.
  exceptions.py     Errori specifici (es. configurazione/cambio valuta
                    mancanti).
  _logging.py       Log uniforme per tutti i moduli.


------------------------------------------------------------------------------
 9. NOTE TECNICHE
------------------------------------------------------------------------------

  - Esclusioni applicate in due punti: nello Step 1 le righe BOM con optional
    incompatibili vengono divise in alternative valide (deterministico); nello
    Step 2 la simulazione risolve gli stessi conflitti in modo probabilistico
    durante il campionamento. I due meccanismi sono complementari.
  - Affidabilita' TR: regola gli alpha del campionamento Dirichlet.
    Affidabilita' piu' alta concentra le estrazioni sui take rate dichiarati;
    piu' bassa lascia piu' dispersione.
  - Lead time decimale: l'aggregazione mensile della domanda supporta lead time
    non interi (es. 1,5 mesi) tramite interpolazione lineare.
  - Gestione memoria: per matrici di simulazione molto grandi (> 4 GB) il
    calcolo usa un file temporaneo su disco (memmap) invece della RAM, in modo
    trasparente.
  - Robustezza prezzi/cambio: se mancano valute nei tassi di cambio il sistema
    si ferma con errore esplicito (evita valorizzazioni sbagliate).
  - Riproducibilita': con lo stesso random_seed e gli stessi input i risultati
    sono identici.

==============================================================================
