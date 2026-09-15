# AGENTS.md — vetronaviglio_gestione_energia

> Source of truth per regole e note operative del progetto. Ultimo aggiornamento: 2026-09-15.

## Descrizione
Pipeline Python (pandas/openpyxl) che unisce dati IcoPower (kWh per macchina), ProdWare (ODL/quantità), ore lavorate per turno e bollette ENEL in un excel consolidato di rendicontazione energetica e produzione.

## Current Focus

> **Data:** 2026-09-15

### Completato
- **Pipeline completa**: `consolida_2025_2026.py` → `rendiconto_2025_2026.xlsx` (14 fogli)
  - Fogli 1-9: tabelle dati 2026 e 2025, confronti e note di lettura
  - Foglio 10: confronti mensili 2026 vs 2025
  - Foglio 11: confronti annuali 2025 vs 2026
  - Foglio 12: tabella oraria statistiche
  - Foglio 13: tabelle orarie per mese
  - Foglio 14: **Analisi efficienza** (NUOVO)
- **Analisi efficienza (foglio 14)** contenente:
  - Profilo notte/giorno 2025 vs 2026 (tabella mensile da ICO-F500)
  - Impatto 3 turni: extra notte ≈ +17.785 kWh / ~€5.300
  - Idle analysis per macchina (GEN-SET 2026): F04 25% idle, TROV 32% idle (top 2)
  - Raccomandazioni prioritarie per Direttore di Produzione
- **Mancata trasmissione contatori 2026**: MG5/MG8 (fermi dal 14/05/2026) e MG6 (dall'11/05/2026) risultano in produzione per censimento ma kWh SOTTOSTIMATI (contatore non aggiorna). Marcati "dato incompleto: contatore fermo dal 14/05/2026 (macchina in produzione)" in Confronto/Sintesi 2026 e bullet 12 nelle Note di lettura. Report sezione 5 "Rilevazioni contatori da verificare" con punto di misura IcoPower esatto e ultima lettura valida per ciascuna — MG6 priorità (133,5 h a settembre, macchina più attiva dell'area).
- **Profilo giorno vs notte** (sezione 2 report): contatore generale 2026 = 313.144 kWh; NOTTE 22-06 = 36.952 kWh (11,8%), GIORNO 06-22 = 276.193 (88,2%). Decomposizione: baseline servizi sempre-on ~5,4 kW ≈ 10,3%; macchine con contatore ≈ 23,2%; residuo diurno ≈ 66,6% = produzione presso contatore generale senza sotto-misura. Benchmark ACEEE/LBNL (aria compressa 10-11%, nonprocess 15-30%) → Vetronaviglio NELLA NORMA: il "82% servizi" era un malinteso, il grosso del residuo è produzione diurna non attribuita.
- **Report Direttore di Produzione**: `report_direttore_produzione.md` (+ HTML) — sintesi esecutiva con quadro energetico, profilo giorno/notte (sez. 2), tabella idle (4), rilevazioni contatori (5), 5 azioni prioritarie con risparmio/CAPEX/payback (6)
- **Bug idle corretto**: computazione ora replica `_analisi_efficienza.py` (kWh = diff contatori cumulativi, potenza kW = kWh/dt_h, soglie su p99 kW, load factor corretto)
- **Testo rivisto dopo analisi**: rimosso il "82% servizi" fuorviante; sotto-contatori consigliati su macchine senza misura dedicata ad alto carico (residuo ~67% contatore generale)

### In corso / Da chiudere
- **Commit `0980f75` pushato su master** (rendiconto, report md+html, AGENTS.md)
- Email a Federico NON inviata (decisione utente: attendere richiesta di riverifica prima di evitare figureccia)
- Manca bolletta AGO 2026

### Prossimo step
- Attendere bolletta AGO 2026 per completare l'anno 2026
- Sessionare con Federico la riverifica dei contatori bloccati (MG5/MG8/MG6) se richiesto
- Aggiornare `readme.md` (descrive solo il flusso legacy, non la pipeline 2026)

## Struttura file

```
vetronaviglio_gestione_energia/
├── consolida_2025_2026.py   ← script principale (14 fogli output)
├── consolida_2026.py        ← helper importati (config, mapping, stili)
├── rendiconto_2025_2026.xlsx← OUTPUT principale (14 fogli)
├── rendiconto_2026.xlsx     ← output singolo 2026 (legacy)
├── report_direttore_produzione.md ← report sintetico per Direttore
├── report_direttore_produzione.html ← versione HTML print-friendly del report
├── PROJECT_AI_NOTES.md      ← decision log e note IA
├── input/
│   ├── CSV Vetronaviglio_2026.csv   ← IcoPower per-macchina (cumulativi 5min)
│   ├── ICO-F500-2025-2026.csv       ← contatore generale (cumulativi 5min)
│   ├── ico-f500.csv                 ← copia contatore generale
│   ├── Dati consumi e costi energetici al 30_11_2025.xlsx ← fonte 2025
│   ├── Dettaglio Analisi 2025.xlsx  ← ODL ProdWare (totale periodo)
│   ├── Dettaglio ore lavorate per turno 2025.xlsx ← ore presenza
│   └── PDF bollette ENEL 2024-2026
└── venv/                           ← Python venv (pandas, openpyxl) [in realtà .venv]
```

## Note tecniche

### CSV IcoPower (cumulativi)
- Diff di letture consecutive = kWh per intervallo 5 min
- Potenza kW = kWh / (5/60) = kWh × 12
- Soglie idle: 5-25% di p99 (potenza kW)
- Periodo 3-turni: 29/06/2026 – ~fine agosto 2026
- Notte: 22:00 → 06:00

### Copertura contatori
- ICO-F500 (contatore generale): ~82-85% dei kWh ENEL totali
- Contatori per-macchina: ~18% del totale (il resto è aria compressa, HVAC, illuminazione, servizi)

### Fonte dati 2025
- `Dati consumi e costi energetici al 30_11_2025.xlsx` → foglio "Tabelle"
- Header macchine riga 3; blocchi: consumi (righe 4-17), ore (20-33), quantità (36-49)
- ENEL kWh righe 111-134
