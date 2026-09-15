# PROJECT_AI_NOTES — vetronaviglio_gestione_energia

> Decision log e stato sessione. Data ultimo aggiornamento: 2026-09-11.

## Obiettivo
Rendicontazione produzione/consumo/energia Vetronaviglio 2026: unire dati IcoPower (kWh macchine), ProdWare (ODL/quantità), ore lavorate per turno e bollette ENEL in un excel finale di analisi per macchina.

## Input in `input/` (caricati 11/09/2026)

| File | Fonte | Struttura verificata |
|---|---|---|
| `CSV Vetronaviglio_2026.csv` | IcoPower | `;`-sep, 72.851 righe, timestamp univoci ogni **5 min**, da 2026-01-01 a 2026-09-11 11:15. Header `Timestamp` + 27 colonne `Vetronaviglio.<MACCHINA>#Energia attiva nel carico#kWh`. **Contatori cumulativi** → consumo intervallo = diff letture consecutive. Numeri con virgola (`4691,07`). Macchine: F01-F11, ISOLA1, M1030, M1050, MG5-MG8, MMAN, MP1-MP4, Sycma, TMPM2_TMPMI, TMPMA1, TROV. |
| `Dettaglio Analisi.xlsx` | ProdWare (ODL) | Sheet `Dettaglio Analisi`, 711 righe / **357 ODL**. Col: Macchina, Odl, Art., Qta lancio, Qta residua, *Data VEFI (VUOTA)*, *Qta VEFI (VUOTA)*, Addetto, Ore Prep.Cons., Ore Lav. cons., Ore Fermo, Causa Fermo, Qta scarti, TLavSCons (tutti 0), TPrepPrev. Ogni ODL appare su 2+ righe (fasi/macchine diverse). Senza date → analisi su totale periodo. |
| `Dettaglio ore lavorate per turno 2026.xlsx` | Ore presenza | 38 righe macchina × date (gen→set 2026) × turno 1/2/3. Header 3 livelli (riga1 mese, riga2 date, riga3 n° turno). Colonne A/B = gruppo+macchina. Valori = ore per turno/giorno/macchina. Colonna finale "Totale complessivo". |
| `Qta prodotta 2026.xlsx` | ProdWare | Sheet `Qta prodotta`, **una sola riga** = totali periodo per macchina (26 macchine, ordine sparso). Nessuna dimensione mese. |
| 8 PDF bollette | ENEL | Periodi **DIC 2025 → LUG 2026** (manca AGO 2026). Totale bolletta = importo ripetuto 7× nel PDF (verificato incrociando solleciti "non pagata"). |

## Bollette ENEL 2026 (verificate)
| Periodo | kWh | Totale € | €/kWh |
|---|---|---|---|
| DIC 2025 | 51.372 | 15.283,15 | 0,2975 |
| GEN 2026 | 54.755 | 18.174,09 | 0,3319 |
| FEB 2026 | 54.897 | 16.321,89 | 0,2973 |
| MAR 2026 | 47.216 | 15.093,24 | 0,3197 |
| APR 2026 | 36.199 | 9.817,06 | 0,2712 |
| MAG 2026 | 25.994 | 7.523,72 | 0,2894 |
| GIU 2026 | 47.377 | 13.902,72 | 0,2934 |
| LUG 2026 | 78.390 | 26.859,91 | 0,3426 |
| **Tot** | **396.200** | **122.975,78** | — |

## Analisi ODL per macchina (Dettaglio Analisi, totale periodo)
- Tot Ore Lav. cons = **12.186 h**, Ore Prep = 2.286 h, Ore Fermo = **15.560 h** (fermo/lavoro = 1,28).
- Cause fermo (in ore, top due per macchina): Mancanza operatore, Pausa pranzo/cena, problemi qualità, Mancanza Attrezzista, rifacimento retino serigrafico, intervento tecnico, Fine turno, Avviamento produzione, Mancanza materiale.
- % scarti su lancio: TMPMA2 5,40%, TROV 3,44%, M1030 2,38%.
- Macchine più caricate (OreLav+Prep): TROV, MG5, F04, MG8, M1030.

## Decisioni confermate da Stefano (2026-09-11)
1. **N/A**: escludere dall'analisi energia — 351 ODL / 4,7M pezzi lanciati con 0 ore = fase "lancio ordine" non assegnata a macchina.
2. **Data VEFI / Qta VEFI vuote**: accettato, si lavora su **totale periodo** (gen→11/09/2026). Se un export futuro include la data, si potrà disaggregare per mese.
3. **ISOLA2 / ISOLA3**: escludere dal confronto consumi (non misurate nel CSV IcoPower), valgono solo per ore/quantità.

## Sessione 2026-09-14 — Analisi efficienza e report Direttore

### FATTO
- Creato foglio "Analisi efficienza" (foglio 14 del rendiconto) da dati ICO-F500 contatore generale:
  - Profilo notte/giorno 2025 vs 2026 con tabelle mensili
  - Impatto 3 turni: extra notte = 17.785 kWh ≈ €5.300 (baseline 70 kWh/notte vs 388 durante 3-turni)
  - Idle analysis per macchina (GEN-SET 2026): F04 = 25% idle (1.150 kWh), TROV = 32% idle (593 kWh)
  - Raccomandazioni prioritarie per Direttore di Produzione
- Corretto bug computazione idle: ora replica esattamente `_analisi_efficienza.py` (kWh = diff cumulativi, potenza = kWh/dt_h in kW, soglie su p99)
- Creato `report_direttore_produzione.md`: sintesi esecutiva con quadro energetico, tabella idle, 5 azioni (idle → aria compressa → sotto-contatori → FV → HVAC) con risparmio/CAPEX/payback
- Aggiornato foglio "Note di lettura" nel rendiconto: aggiunta riga foglio 14, fonte ICO-F500, definizioni regime turni/idle, note su copertura contatori e periodo 3-turni
- Creato AGENTS.md del progetto (Current Focus, struttura file, note tecniche)
- Aggiornato PROJECT_AI_NOTES.md con documentazione sessione

### DECISIONE
- L'analisi efficienza usa **solo** il contatore generale ICO-F500 (~82-85% dei kWh ENEL). I dati notte/giorno e idle NON sono disaggregati per macchina (manca/copertura insufficiente nei contatori per-macchina).

### TODO
- Commitare la pipeline consolidata, il report e l'AGENTS.md (tutti attualmente untracked)
- Cancellare file temporanei: `_scratch_2025.py`, `_analisi_efficienza.py`, `_analisi_ico_f500.py`
- Attendere bolletta AGO 2026 per completare l'anno

## Prossimo step (lunedì)
Costruire pipeline di consolizione 2026 → excel finale per macchina con:
`consumo_kwh` (diff contatori IcoPower, totale periodo) ⨝ `ore_lavo` + `ore_prep` + `ore_fermo_causa` (ODL) ⨝ `pezzi_prodotti` (Qta prodotta) ⨝ `ore_turni` (Dettaglio ore)
→ indicatori: kWh/pezzo, kWh/ora_lavoro, pezzi/ora_lavoro, % scarti, ore fermo per causa.
Footer mensile: consumo e costo ENEL da bollette (tabella sopra); bolletta AGO 2026 in attesa.

## Accuratezza dati
- CSV IcoPower: macchine spente nel periodo (changes=0): F08, F11, MP1, MP3, MP4; quasi spenta F10 (43 changes).
- `F04` = macchina più energivora (max lettura 113.804,70 kWh).