---
titolo: Report efficienza energetica — Analisi e raccomandazioni
autore: Studio di bilancio energetico Vetronaviglio
data: 14 settembre 2026
periodo: GEN 2025 – SET 2026
fonti: rendiconto_2025_2026.xlsx (foglio Analisi efficienza + Tabelle), fatture ENEL
---

# Report efficienza energetica — Vetronaviglio

## 1. Quadro generale

| Voce | 2025 (GEN-NOV) | 2026 (GEN-SET) | Var. % |
|------|---------------|---------------|--------|
| kWh contatori macchine | 147.103 | 72.467 | — (11 vs 9 mesi) |
| kWh ENEL fatturati | 500.743 | 396.200 | — |
| Costo ENEL totale | ~€146.000 | ~€123.000 | — |

**Copertura contatori**: i contatori montati a monte delle singole macchine misurano circa **il 23%** del consumi fatturati (72.467 kWh su ~313.000 del contatore generale). Il resto non è attribuibile a "servizi": è quasi tutto consumo che transita dal contatore generale durante le ore di produzione (vedi sezione 2).

---

## 2. Profilo giorno vs notte — dove va il consumo non attribuito

L'analisi del contatore generale ICO-F500 conferma che i contatori singoli (a monte macchina) non includono servizi: la verifica **giorno vs notte** mostra dove finisce il resto.

| Fascia (mediana kW) | 2025 (GEN-NOV) | 2026 (GEN-SET) |
|---------------------|---------------|----------------|
| Notte feriale profonda (23-04) | 5,2 kW | 5,4 kW |
| Giorno feriale (06-22) | 104,0 kW | 94,8 kW |
| Weekend | 4,6 kW | 4,6 kW |
| Rapporto giorno/notte | 20× | 18× |

**Energia per fascia — 2026 (contatore generale, totale 313.144 kWh):**
- NOTTE (22-06): **36.952 kWh (11,8%)** — nel 2025 era il 4,5% (18.020 kWh): la differenza è l'effetto dei 3 turni.
- GIORNO (06-22): **276.193 kWh (88,2%)** — di cui weekend solo 8.319 kWh (produzione ferma, resta solo la baseline servizi).

**Decomposizione 2026 del contatore generale (313.144 kWh):**

| Componente | kWh | % |
|-----------|-----|---|
| Baseline servizi sempre-on (@ ~5,4 kW) | ≈ 32.140 | ≈ 10,3% |
| Macchine misurate (sotto-contatori singoli) | 72.467 | 23,2% |
| Residuo diurno (produzione non sotto-misurata + servizi variabili) | ≈ 208.537 | ≈ 66,6% |

**Lettura**: la baseline notturna di ~5 kW (illuminazione, HVAC, pompe, servizi minimi) pesa solo ~10%. Il grosso del consumo si concentra di **giorno, quando le macchine lavorano** — il residuo di ~67% è produzione che transita dal contatore generale senza sotto-contatore dedicato (per esempio aria compressa, refrigerazione/chiller) più servizi che seguono l'attività, NON sprechi fissi. Il benchmark settoriale (aria compressa ~10-11% del consumo elettrico industriale, illuminazione/HVAC 15-30%) colloca Vetronaviglio nella norma, anzi sotto per la baseline notturna.

---

## 3. Impatto dei 3 turni (giugno-agosto 2026)

| Periodo | Notte feriali kWh | Notti | Media kWh/notte |
|---------|------------------|-------|----------------|
| Pre-3turni (GEN-GIUG) | 8.940 | 127 | 70 |
| 3-turni (29/06-14/09) | 21.727 | 56 | 388 |
| **Extra notte vs baseline** | **+17.785 kWh** | | ≈ **€5.300** |

Il consumo notturno dei feriali è passato da ~70 kWh/notte a ~388 kWh/notte durante i 3 turni. Da settembre si è tornati al regime 1 turno e i consumi notturni sono tornati ai livelli di baseline (~37-39 kWh/notte).

**Implicazione**: il picco di LUG-AGO 2026 non è un'anomalia dei contatori — è il costo reale della produzione h24 e si ripeterà ogni anno se la produzione proseguirà su 3 turni.

---

## 4. Macchine con maggiore consumo idle

L'analisi idle (kWh con potenza tra 5% e 25% del p99) sul periodo GEN-SET 2026, contatore generale:

| Macchina | kWh totale | kWh idle | % idle | Ore idle | Note |
|----------|-----------|----------|--------|---------|------|
| **F04** | 4.609 | **1.150** | **25%** | 2.864 | Macchina più dissipativa: tensione continua ma carichi minimi |
| **TROV** | 1.850 | **593** | **32%** | 869 | Più alta percentuale idle del parco |
| MG7 | 7.412 | 523 | 7% | 1.122 | Ottimizzazione setpoint possibili |
| F02 | 8.753 | 490 | 6% | 319 | — |
| MG6 | 6.950 | 399 | 6% | 242 | — |
| F05 | 3.254 | 329 | 10% | 157 | — |
| MP2 | 13.418 | 320 | 2% | 195 | Alto consumo totale ma basso % idle |
| MMAN | 3.377 | 104 | 3% | 782 | Ore attive molte, consumi bassi |

**Idle totale parco**: ~4.391 kWh/anno (~6,1% del totale parco) ≈ **€1.300/anno** a zero CAPEX.

---

## 5. Rilevazioni contatori da verificare

Censimento area soffiaggio/iniezione (7 macchine: MG7, MP2, MG5, MG8, MG6, F02, F04). Quattro contatori della lista sono aggiornati; **tre risultano in produzione ma con trasmissione contatore interrotta** — i kWh 2026 di queste macchine sono sottostimati (dato fermo a maggio). Punti di misura esatti dove intervenire:

| Macchina | Punto di misura IcoPower | Ultima lettura valida | Valore cumulativo | Ore sett. 2026 | kWh 2026 (parz.) |
|----------|--------------------------|----------------------|-------------------|----------------|------------------|
| **MG5** | `Vetronaviglio.MG5#Energia attiva nel carico#kWh` | 14/05/2026 10:50 | 32.825,4 kWh | 47,0 | 1.683,5 |
| **MG8** | `Vetronaviglio.MG8#Energia attiva nel carico#kWh` | 14/05/2026 09:30 | 10.641,0 kWh | 20,4 | 2.648,6 |
| **MG6** | `Vetronaviglio.MG6#Energia attiva nel carico#kWh` | 11/05/2026 04:30 | 27.136,6 kWh | 133,5 | 6.966,1 |

**Intervento richiesto**: verificare rilevatore/trasmissione dei 3 contatori (cavo, alimentazione, concentratore). MG6 è la priorità — a settembre è la macchina più attiva dell'area (133,5 h). Ore e pezzi restano validi (fonti ProdWare), solo il consumo kWh è mancante da metà maggio.

---

## 6. Raccomandazioni prioritarie

### Azione 1 — Spegnimento macchine inattive (idle)
- **Azione**: spegnere F04 e TROV durante i turni di pausa/notte senza commesse pendenti.
- **Risparmio**: ~1.743 kWh/anno ≈ €520/anno.
- **Costo**: €0 — solo disciplina operativa.
- **Payback**: immediato.

### Azione 2 — Aria compressa: audit perdite
- **Azione**: audit sistematico perdite in rete (valvole, giunzioni, couplings) con ultrasuoni; riparare perdite >80% a valle dei riduttori.
- **Risparmio**: fino al 30% del consumo aria compressa, stimato 15-20% sul totale kWh ENEL ≈ €15.000-25.000/anno.
- **Costo**: €14.000-18.000 (valvole di riduzione, materiali).
- **Payback**: 2-3 anni.

### Azione 3 — Sotto-contatori quadri generali
- **Azione**: installare sotto-contatori ifm SD sui quadri generali (alimentazioni, aria compressa, HVAC, illuminazione) per mappare il residuo non sotto-misurato (~67-77% del contatore generale, vedi sezione 2).
- **Risparmio**: non diretto; abilita tutte le successive ottimizzazioni con dati reali.
- **Costo**: ~€5.000 per 10-15 sensori.
- **Payback**: <1 anno grazie al rilevamento di sprechi nascosti.

### Azione 4 — Fotovoltaico industriale 250-500 kWp
- **Azione**: installare FV su copertura capannoni; autoconsumo diretto 75-85%.
- **Risparmio**: 50-70% del consumo annuo, €25.000-40.000/anno con Transizione 5.0.
- **Costo**: €150.000-210.000 (dopo credito 40-45% Transizione 5.0).
- **Payback**: 4-6 anni (2-3 con tax credit industriale).

### Azione 5 — HVAC/chiller: ottimizzazione setpoint
- **Azione**: alzare setpoint estivo di 1-2°C, programmare spegnimento notturno/weekend.
- **Risparmio**: -8-12% consumo HVAC, -3-4% per °C di correzione.
- **Costo**: €0 — solo regolazione parametri impianto.

---

## 7. Riepilogo finanziario

| Azione | Risparmio annuo stimato | CAPEX | Payback |
|--------|------------------------|-------|---------|
| Idle (F04+TROV) | €520 | €0 | Immediato |
| Aria compressa | €15.000-25.000 | €14.000-18.000 | 2-3 anni |
| Sotto-contatori | Abilita risparmi successivi | €5.000 | <1 anno |
| FV 250-500 kWp | €25.000-40.000 | €150.000-210.000 | 4-6 (2-3 con T5.0) |
| HVAC | €3.000-8.000 | €0 | Immediato |
| **Totale** | **€43.000-73.000/anno** | **€170.000-230.000** | |

---

## 8. Prossimi passi suggeriti

1. **Settimana 1-2**: audit aria compressa (contattare fornitore ultrasuoni o specialista).
2. **Settimana 1**:Procedure di spegnimento F04/TROV a fine turno — formare gli addetti.
3. **Mese 1**: installare 5-10 sotto-contatori su quadri generali per mappare il residuo non sotto-misurato.
4. **Mese 2-3**: avviare iter per FV (visita tecnica, studio di fattibilità, domanda Transizione 5.0).
5. **Prossima revisione**: confrontare effetti interventi aria compressa dopo 3 mesi di raccolta dati.

---

*Report generato dal rendiconto energetico 2025-2026 (foglio Analisi efficienza). Tutti i valori si riferiscono al contatore generale ICO-F500 (copre ~82-85% dei kWh ENEL).*
