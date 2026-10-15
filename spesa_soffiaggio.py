# -*- coding: utf-8 -*-
"""
Spesa energetica macchine di soffiaggio (MP2, MG5, MG6, MG7, MG8).

Riusa i loader di consolida_2025_2026 (nessun ricalcolo parallelo):
  - 2025: load_consumi_2025()  -> kWh/mese dal foglio Tabelle (GEN-NOV)
  - 2026: load_icopower()      -> kWh/mese da contatori IcoPower (GEN-SET, al 11/09)
Costo = kWh macchina x Euro/kWh della bolletta ENEL del mese
(stesso metodo del foglio bollette del rendiconto; mesi senza bolletta stimati
al prezzo medio dei mesi disponibili).

OUTPUT: spesa_soffiaggio.xlsx
  1. Sintesi annua   - kWh e € per macchina, 2025 vs 2026 + variazione
  2. Mensile 2025    - kWh e € per macchina e per mese (GEN-NOV)
  3. Mensile 2026    - kWh e € per macchina e per mese (GEN-SET)
  4. Note            - metodo, fonti, limiti (contatori fermi, bollette stimate)
"""

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from consolida_2025_2026 import (
    BOLLETTE_2025,
    BOLLETTE_2026,
    CONTATORI_INCOMPLETI,
    load_consumi_2025,
    load_icopower,
)

OUTPUT = Path("spesa_soffiaggio.xlsx")

MACCHINE = ["MP2", "MG5", "MG6", "MG7", "MG8"]
MACCHINE_SOFFIAGGIO_LABEL = "Macchine di soffiaggio: MP2, MG5, MG6, MG7, MG8"

MESI_IT = {"GEN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAG": 5, "GIU": 6,
           "LUG": 7, "AGO": 8, "SET": 9, "OTT": 10, "NOV": 11, "DIC": 12}
NOME_MESE = {1: "Gennaio", 2: "Febbraio", 3: "Marzo", 4: "Aprile", 5: "Maggio",
             6: "Giugno", 7: "Luglio", 8: "Agosto", 9: "Settembre",
             10: "Ottobre", 11: "Novembre", 12: "Dicembre"}

# --- Stili ---
THIN = Side(style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(bold=True, color="FFFFFF", size=10)
SUB_FILL = PatternFill("solid", fgColor="DDEBF7")
TOT_FILL = PatternFill("solid", fgColor="FFF2CC")
WARN_FILL = PatternFill("solid", fgColor="FCE4D6")
TITLE_FONT = Font(bold=True, size=13, color="1F4E78")
NOTE_FONT = Font(italic=True, size=9, color="808080")
FMT_KWH = "#,##0"
FMT_EUR = "#,##0.00"


def tariffe_2025() -> dict:
    """mese 'YYYY-MM' -> (Euro/kWh, stimato: bool) per il 2025."""
    prezzi = [(k, e, kwh) for k, e, kwh in
              ((p, e, kwh) for p, kwh, e in BOLLETTE_2025) if e is not None]
    avg = sum(e for _, e, _ in prezzi) / sum(kwh for _, _, kwh in prezzi)
    out = {}
    for periodo, kwh, euro in BOLLETTE_2025:
        anno, mese = periodo.split()[1], MESI_IT[periodo.split()[0]]
        key = f"{anno}-{mese:02d}"
        if euro is None:
            out[key] = (round(kwh * avg, 2) / kwh, True)
        else:
            out[key] = (euro / kwh, False)
    return out


def tariffe_2026() -> dict:
    """mese 'YYYY-MM' -> (Euro/kWh, stimato: bool) per il 2026 (GEN-SET)."""
    disp = [(p, e, kwh) for p, kwh, e in BOLLETTE_2026 if e is not None
            and p != "DIC 2025"]
    avg = sum(e for _, e, _ in disp) / sum(kwh for _, _, kwh in disp)
    out = {}
    for periodo, kwh, euro in BOLLETTE_2026:
        if periodo == "DIC 2025":
            continue
        anno, mese = periodo.split()[1], MESI_IT[periodo.split()[0]]
        key = f"{anno}-{mese:02d}"
        if euro is None:
            out[key] = (avg, True)
        else:
            out[key] = (euro / kwh, False)
    for mese in (8, 9):  # AGO/SET 2026: bolletta in attesa -> stima prezzo medio
        out[f"2026-{mese:02d}"] = (avg, True)
    return out


def _kwh_2025() -> pd.DataFrame:
    cons = load_consumi_2025()
    mancanti = [m for m in MACCHINE if m not in cons.columns]
    if mancanti:
        raise SystemExit(f"Macchine assenti nel foglio Tabelle 2025: {mancanti}")
    return cons[MACCHINE].round(2)


def _kwh_2026() -> pd.DataFrame:
    cons = load_icopower()
    mancanti = [m for m in MACCHINE if m not in cons.columns]
    if mancanti:
        raise SystemExit(f"Macchine assenti nel CSV IcoPower 2026: {mancanti}")
    mesi = [i for i in cons.index if i <= "2026-09"]
    return cons.loc[mesi, MACCHINE].round(2)


def _costi(kwh: pd.DataFrame, tariffe: dict) -> pd.DataFrame:
    """Stessa forma di kwh, con i costi € per mese."""
    out = pd.DataFrame(index=kwh.index, columns=kwh.columns, dtype=float)
    for mese in kwh.index:
        prezzo, _ = tariffe[mese]
        out.loc[mese] = kwh.loc[mese] * prezzo
    return out.round(2)


def _stima_mese(tariffe: dict, mese: str) -> bool:
    return tariffe[mese][1]


# --- Scrittura Excel ---

def _title(ws, testo, sottotitolo=None):
    ws["A1"] = testo
    ws["A1"].font = TITLE_FONT
    riga = 2
    if sottotitolo:
        ws["A2"] = sottotitolo
        ws["A2"].font = NOTE_FONT
        riga = 3
    return riga + 1


def _sheet_sintesi(wb, k25, e25, k26, e26):
    ws = wb.create_sheet("Sintesi annua")
    r = _title(ws, "Spesa energetica annua - macchine di soffiaggio",
               f"{MACCHINE_SOFFIAGGIO_LABEL} | 2025 GEN-NOV (11 mesi) vs 2026 GEN-SET (9 mesi, al 11/09) | "
               "Confronto sulle MEDIE MENSILI: i totali coprono periodi diversi e non sono confrontabili")
    headers = ["Macchina", "kWh 2025\n(GEN-NOV)", "Euro 2025\n(GEN-NOV)",
               "kWh 2026\n(GEN-SET)", "Euro 2026\n(GEN-SET)",
               "Euro/mese\nmedia 2025", "Euro/mese\nmedia 2026",
               "Var. %\n(su medie\nmensili)", "Nota"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(r, c, h)
        cell.fill, cell.font, cell.border = HDR_FILL, HDR_FONT, BORDER
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
    ws.row_dimensions[r].height = 42
    ws.freeze_panes = ws.cell(r + 1, 1)

    mesi25, mesi26 = len(k25.index), len(k26.index)  # 11 e 9
    riga = r + 1
    for m in MACCHINE:
        kwh25, eur25 = k25[m].sum(), e25[m].sum()
        kwh26, eur26 = k26[m].sum(), e26[m].sum()
        med25, med26 = eur25 / mesi25, eur26 / mesi26
        var = (med26 - med25) / med25 * 100 if med25 else None
        nota = CONTATORI_INCOMPLETI.get(m, "").replace(
            "dato incompleto: ", "2026 sottostimato: ")
        vals = [m, round(kwh25, 2), round(eur25, 2), round(kwh26, 2),
                round(eur26, 2), round(med25, 2), round(med26, 2),
                round(var, 1) if var is not None else None, nota]
        for c, v in enumerate(vals, 1):
            cell = ws.cell(riga, c, v)
            cell.border = BORDER
            if c in (2, 4):
                cell.number_format = FMT_KWH
            elif c in (3, 5, 6, 7):
                cell.number_format = FMT_EUR
            elif c == 8:
                cell.number_format = "0.0"
            elif c == 9:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
            if nota and c == 9:
                cell.fill = WARN_FILL
        riga += 1

    tot = ["TOTALE AREA",
           round(k25.sum().sum(), 2), round(e25.sum().sum(), 2),
           round(k26.sum().sum(), 2), round(e26.sum().sum(), 2),
           round(e25.sum().sum() / mesi25, 2), round(e26.sum().sum() / mesi26, 2),
           None,
           "Variazione non calcolata sui totali: MG5/MG6/MG8 sottostimate da metà maggio 2026"]
    for c, v in enumerate(tot, 1):
        cell = ws.cell(riga, c, v)
        cell.border, cell.fill, cell.font = BORDER, TOT_FILL, Font(bold=True)
        if c in (2, 4):
            cell.number_format = FMT_KWH
        elif c in (3, 5, 6, 7):
            cell.number_format = FMT_EUR
        elif c == 9:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[riga].height = 30

    for c, w in zip(range(1, 10), (12, 12, 13, 12, 13, 12, 12, 12, 55)):
        ws.column_dimensions[get_column_letter(c)].width = w


def _sheet_mensile(wb, nome, anno, kwh, euro, tariffe):
    ws = wb.create_sheet(nome)
    r = _title(ws, f"Spesa mensile - macchine di soffiaggio {anno}",
               f"{MACCHINE_SOFFIAGGIO_LABEL} | Euro/kWh da bolletta ENEL del mese")

    # header a 2 livelli: riga r = macchina (merge), riga r+1 = kWh / Euro
    ws.cell(r, 1, "Mese")
    ws.merge_cells(start_row=r, start_column=1, end_row=r + 1, end_column=1)
    col = 2
    for m in MACCHINE:
        titolo = m + (" *" if m in CONTATORI_INCOMPLETI and anno == "2026" else "")
        ws.cell(r, col, titolo)
        ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col + 1)
        ws.cell(r + 1, col, "kWh")
        ws.cell(r + 1, col + 1, "Euro")
        col += 2
    ws.cell(r, col, "TOTALE")
    ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col + 1)
    ws.cell(r + 1, col, "kWh")
    ws.cell(r + 1, col + 1, "Euro")
    ws.cell(r, col + 2, "Euro/kWh")
    ws.cell(r, col + 3, "Nota")
    ws.merge_cells(start_row=r, start_column=col + 3, end_row=r + 1, end_column=col + 3)

    for rr in (r, r + 1):
        for c in range(1, col + 4):
            cell = ws.cell(rr, c)
            cell.fill, cell.font, cell.border = HDR_FILL, HDR_FONT, BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = ws.cell(r + 2, 2)

    riga = r + 2
    for mese in kwh.index:
        prezzo, stimato = tariffe[mese]
        vals = [NOME_MESE[int(mese[5:7])]]
        for m in MACCHINE:
            vals += [kwh.loc[mese, m], euro.loc[mese, m]]
        vals += [kwh.loc[mese].sum(), euro.loc[mese].sum(), round(prezzo, 4)]
        nota = ""
        if stimato:
            nota = "prezzo stimato (bolletta non disponibile)"
        if anno == "2026" and int(mese[5:7]) >= 5:
            offerti = [m for m in MACCHINE if m in CONTATORI_INCOMPLETI]
            nota = (nota + "; " if nota else "") + \
                f"contatore {'/'.join(offerti)} fermo da metà maggio: kWh sottostimati"
        if anno == "2026" and mese == "2026-09":
            nota += "; settembre al 11/09"
        vals.append(nota)
        for c, v in enumerate(vals, 1):
            cell = ws.cell(riga, c, v)
            cell.border = BORDER
            if c >= 2 and c < len(vals):
                cell.number_format = FMT_KWH if c % 2 == 0 else FMT_EUR
            if c == len(vals) - 1:  # colonna Euro/kWh
                cell.number_format = "0.0000"
            if c == len(vals) and nota:
                cell.fill = WARN_FILL
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        riga += 1

    # riga totale
    tot = ["TOTALE"]
    for m in MACCHINE:
        tot += [kwh[m].sum(), euro[m].sum()]
    tot += [kwh.sum().sum(), euro.sum().sum(), None, ""]
    for c, v in enumerate(tot, 1):
        cell = ws.cell(riga, c, round(v, 2) if isinstance(v, float) else v)
        cell.border, cell.fill, cell.font = BORDER, TOT_FILL, Font(bold=True)
        if c >= 2 and c < len(tot):
            cell.number_format = FMT_KWH if c % 2 == 0 else FMT_EUR

    ws.column_dimensions["A"].width = 12
    for i in range(2, col):
        ws.column_dimensions[get_column_letter(i)].width = 11
    ws.column_dimensions[get_column_letter(col + 2)].width = 10
    ws.column_dimensions[get_column_letter(col + 3)].width = 60


def _sheet_note(wb, tar25, tar26):
    ws = wb.create_sheet("Note")
    r = _title(ws, "Note di metodo - spesa macchine di soffiaggio")
    righe = [
        ("Oggetto", "Stima della spesa energetica attribuibile alle macchine di "
                    "soffiaggio MP2, MG5, MG6, MG7, MG8, su base mensile e annua."),
        ("Metodo", "Costo = kWh consumati dalla macchina (contatore IcoPower) x "
                   "Euro/kWh della bolletta ENEL del corrispondente mese "
                   "(stesso metodo del foglio bollette del rendiconto)."),
        ("Fonte kWh 2025", "Dati consumi e costi energetici al 30_11_2025.xlsx, foglio "
                           "Tabelle (consolidato aziendale, GEN-NOV 2025)."),
        ("Fonte kWh 2026", "CSV Vetronaviglio_2026.csv (contatori cumulativi IcoPower "
                           "ogni 5 min, GEN-SET 2026, dato al 11/09/2026)."),
        ("Tariffe 2025", "Bollette ENEL 2025: € reali per GEN, APR, MAG, LUG, AGO, SET, "
                         "OTT, NOV; FEB, MAR, GIU STIMATI al prezzo medio dei mesi "
                         "disponibili (come nel rendiconto)."),
        ("Tariffe 2026", "Bollette ENEL GEN-LUG 2026 (€ reali); AGO e SET 2026 STIMATI "
                         "al prezzo medio GEN-LUG 2026 (bollette in attesa)."),
        ("Limite - contatori fermi", "MG5 e MG8: contatore fermo dal 14/05/2026; MG6: "
                                     "dall'11/05/2026. Le macchine sono in produzione ma "
                                     "i kWh (e quindi il costo) dal 15/05/2026 sono "
                                     "SOTTOSTIMATI. La spesa reale di queste 3 macchine "
                                     "e' piu' alta di quanto indicato."),
        ("Limite - periodo 2026", "Totali 2026 su 9 mesi (settembre solo al 11/09): non "
                                  "confrontabili 1:1 con i 11 mesi del 2025. Per questo "
                                  "il confronto nella Sintesi annua usa le MEDIE MENSILI "
                                  "di euro, non i totali."),
        ("Limite - perimetro", "Il costo riguarda solo l'energia elettrica misurata dai "
                               "contatori macchina (circa 18% dei kWh dell'impresa). "
                               "Aree di soffiaggio: aria compressa, HVAC e servizi non "
                               "sono attribuiti alle macchine."),
        ("Periodi tariffe", "Il mese della bolletta e' associato al mese di consumo "
                            "(GEN 2026 -> gennaio 2026), come nel rendiconto."),
    ]
    for c, h in enumerate(["Voce", "Descrizione"], 1):
        cell = ws.cell(r, c, h)
        cell.fill, cell.font, cell.border = HDR_FILL, HDR_FONT, BORDER
    for i, (voce, desc) in enumerate(righe, 1):
        ws.cell(r + i, 1, voce).font = Font(bold=True, size=10)
        ws.cell(r + i, 1).border = BORDER
        cell = ws.cell(r + i, 2, desc)
        cell.border, cell.alignment = BORDER, Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r + i].height = 45
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 110


def main():
    k25, k26 = _kwh_2025(), _kwh_2026()
    t25, t26 = tariffe_2025(), tariffe_2026()
    e25, e26 = _costi(k25, t25), _costi(k26, t26)

    wb = Workbook()
    wb.remove(wb.active)
    _sheet_sintesi(wb, k25, e25, k26, e26)
    _sheet_mensile(wb, "Mensile 2025", "2025", k25, e25, t25)
    _sheet_mensile(wb, "Mensile 2026", "2026", k26, e26, t26)
    _sheet_note(wb, t25, t26)
    wb.save(OUTPUT)

    print(f"Scritto {OUTPUT}")
    print("\nSintesi (Euro):")
    df = pd.DataFrame({
        "2025 EUR": e25.sum(), "2026 EUR": e26.sum(),
        "2025 kWh": k25.sum(), "2026 kWh": k26.sum(),
    }).round(2)
    df.loc["TOTALE"] = df.sum().round(2)
    print(df)


if __name__ == "__main__":
    main()
