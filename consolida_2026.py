# -*- coding: utf-8 -*-
"""
Pipeline rendiconto energia 2026 - Vetronaviglio.
Unisce le 4 sorgenti in un unico excel finale per macchina, pronto per il
Direttore di Produzione / Responsabile Qualità (autodocumentato).

INPUT (cartella input/):
  - CSV Vetronaviglio_2026.csv          # IcoPower: contatori cumulativi ogni 5 min, 27 macchine
  - Dettaglio Analisi.xlsx              # ProdWare ODL: ore lav/prep/fermo, cause, scarti, qta lancio
  - Qta prodotta 2026.xlsx              # ProdWare: pezzi finiti per macchina (totale periodo)
  - Dettaglio ore lavorate per turno 2026.xlsx  # ore per macchina/mese/turno (disagg mensile)

OUTPUT: rendiconto_2026.xlsx (6 fogli):
  1. Note di lettura         - come leggere il file, fonti, definizioni, limiti, cross-check
  2. Sintesi per macchina    - indicatori consolidati per macchina (+ TOTALE)
  3. Consumi mensili CSV     - kWh per macchina x mese (diff contatori IcoPower) + TOTALE
  4. Ore lavorate mensili    - ore per macchina x mese (da file turni) + TOTALE
  5. Ore fermo per causa     - dettaglio cause fermo per macchina + TOTALE
  6. Bollette ENEL           - consumi/costi bollettati DIC25-LUG26 (+ nota AGO in attesa)

NOTA IMPORTANTE (confermata 2026-09-14): il consumo misurato IcoPower (72.467 kWh)
copre solo ~18% del totale bollettato ENEL (396.200 kWh). I contatori misurano una
parte dei carichi; il resto (aria compressa, luci, uffici, ISOLA2/3/5, MANUA, ...)
non è contabilizzato a macchina. Il rendiconto riporta quindi "kWh misurati" per
macchina + "kWh bollettati ENEL" come dato di bolletta separato.

Mapping macchine (canonico = produttivo):
  CSV 'Sycma'        -> ASS1          # ASS1 = ASS.POMPE SYCMA
  CSV 'TMPMA1'       -> TMPMA2        # naming drift sul codice
  CSV 'TMPM2_TMPMI'  -> TMPM2_TMPMI   # contatore unico; grandezze TMPM2+TMPMI fuse
  MANUA, ISOLA2, ISOLA3, ISOLA5 -> kWh non misurato (vuoto nel rendiconto)
"""

from pathlib import Path

import pandas as pd

INPUT_DIR = Path("input")
OUTPUT = "rendiconto_2026.xlsx"

# --- Mapping contatori IcoPower -> codice canonico (produttivo) ---
MAPPING_COUNTER = {
    "Sycma": "ASS1",
    "TMPMA1": "TMPMA2",
    "TMPM2_TMPMI": "TMPM2_TMPMI",  # già canonico (produzione TMPM2+TMPMI fuse)
}
# Macchine produttive senza contatore: kWh = non misurato
NON_MISURATE = {"MANUA", "ISOLA2", "ISOLA3", "ISOLA5"}
MERGE_PROD = {"TMPM2": "TMPM2_TMPMI", "TMPMI": "TMPM2_TMPMI"}
# Macchine presenti solo sui contatori IcoPower, letture trascurabili (~18 kWh totali)
SOLO_CONTATORE = ["F08", "F10", "F11", "M1050", "MP1", "MP3", "MP4"]

# --- Bollette ENEL DIC25 -> LUG26 (valori verificati su PDF, 2026-09-11) ---
BOLLETTE = [
    ("DIC 2025", 51372, 15283.15),
    ("GEN 2026", 54755, 18174.09),
    ("FEB 2026", 54897, 16321.89),
    ("MAR 2026", 47216, 15093.24),
    ("APR 2026", 36199, 9817.06),
    ("MAG 2026", 25994, 7523.72),
    ("GIU 2026", 47377, 13902.72),
    ("LUG 2026", 78390, 26859.91),
    ("AGO 2026", None, None),  # in attesa
]

# --- Elenco macchine presenti nel file turni (per documentazione) ---


def load_icopower() -> pd.DataFrame:
    """kWh consumati per (macchina, mese) dai contatori cumulativi IcoPower."""
    df = pd.read_csv(INPUT_DIR / "CSV Vetronaviglio_2026.csv", sep=";", decimal=",")
    df.columns = [c.split("#")[0].replace("Vetronaviglio.", "").strip() for c in df.columns]
    df["timestamp"] = pd.to_datetime(df["Timestamp"])
    meters = [c for c in df.columns if c not in ("Timestamp", "timestamp")]
    df[meters] = df[meters].apply(pd.to_numeric, errors="coerce")
    delta = df[meters].diff().clip(lower=0.0)  # consumo per intervallo (resetta conteggi negativi)
    df["mese"] = df["timestamp"].dt.to_period("M")
    monthly = delta.groupby(df["mese"], sort=True).sum()
    monthly.columns = [MAPPING_COUNTER.get(c, c) for c in monthly.columns]
    monthly = monthly.T.groupby(level=0).sum().T  # unisci colonne con stessa canonica
    monthly.index = [str(m) for m in monthly.index]  # "2026-01"
    return monthly


def load_odl() -> pd.DataFrame:
    """Aggregati ODL per macchina: ore, scarti, qta lancio (totale periodo)."""
    df = pd.read_excel(INPUT_DIR / "Dettaglio Analisi.xlsx", sheet_name="Dettaglio Analisi")
    df.columns = [c.strip() for c in df.columns]
    return df.groupby("Macchina", sort=True).agg(
        ore_lav=("Ore Lav. cons.", "sum"),
        ore_prep=("Ore Prep.Cons.", "sum"),
        ore_fermo=("Ore Fermo", "sum"),
        scarti=("Qta scarti", "sum"),
        qta_lancio=("Qta lancio", "sum"),
    ).round(2)


def load_qta() -> pd.Series:
    """Pezzi finiti per macchina (riga unica del file Qta prodotta)."""
    df = pd.read_excel(INPUT_DIR / "Qta prodotta 2026.xlsx", sheet_name="Qta prodotta")
    df.columns = [str(c).strip() for c in df.columns]
    s = df.iloc[0].drop(labels=df.columns[0])  # scarta prima colonna (N/A)
    s = pd.to_numeric(s, errors="coerce").fillna(0.0)
    s.name = "pezzi_prodotti"
    return s


def load_turni() -> pd.DataFrame:
    """Ore lavorate per macchina x mese dal file turni (header a 3 livelli).

    Il file turni è la disaggregazione mensile delle ore lavorate:
    la somma per macchina coincide con la colonna sorgente "Totale complessivo"
    (ordine di grandezza == Ore Lav. cons. ODL, con piccoli scostamenti).
    Nota: le colonne giorno/turno sono sparse => somma per gruppo mese
    (groupby), mai accumulo colonna per colonna (NaN lo inquinerebbe).
    """
    raw = pd.read_excel(INPUT_DIR / "Dettaglio ore lavorate per turno 2026.xlsx",
                        sheet_name="Dettaglio ore lavorate per turn", header=None)
    hdr = pd.Series(pd.to_datetime(raw.iloc[0, 2:], errors="coerce").values)  # mesi in riga 0
    val_cols = hdr.notna().values  # False solo per l'ultima colonna "Totale complessivo"
    rows = (raw[0].notna()
            & ~raw[0].astype(str).str.contains("Totale")
            & raw[1].notna()
            & (raw[1].astype(str) != "Totale complessivo"))
    sub = raw.loc[rows]
    mcodes = sub[1].astype(str).str.split(" - ").str[0].values
    vals = sub.iloc[:, 2:].apply(pd.to_numeric, errors="coerce")
    vals = vals.loc[:, val_cols].reset_index(drop=True)
    months = hdr[val_cols].values  # mesi in riga 0, allineati per posizione a vals
    month_key = sorted(set(months))
    key_idx = {m: i for i, m in enumerate(month_key)}
    vals.columns = [key_idx[m] for m in months]
    out = vals.T.groupby(level=0).sum().T  # colonne => mesi (NaN ignorate)
    out.index = mcodes
    out = out.groupby(level=0).sum().round(2)
    out.columns = [pd.to_datetime(m).strftime("%Y-%m") for m in month_key]  # "2026-01"
    return out


def load_fermo_causa() -> pd.DataFrame:
    """Ore fermo per (macchina, causa), solo righe > 0, ordinate per macchina."""
    df = pd.read_excel(INPUT_DIR / "Dettaglio Analisi.xlsx", sheet_name="Dettaglio Analisi")
    df.columns = [c.strip() for c in df.columns]
    g = (df.groupby(["Macchina", "Causa Fermo"])["Ore Fermo"].sum()
           .reset_index())
    g = g[g["Ore Fermo"] > 0].sort_values(["Macchina", "Ore Fermo"], ascending=[True, False])
    g = g.rename(columns={"Macchina": "macchina", "Causa Fermo": "causa_fermo",
                          "Ore Fermo": "ore_fermo"}).round(2)
    return g


def build_sintesi(consumi: pd.DataFrame, odl: pd.DataFrame, qta: pd.Series) -> pd.DataFrame:
    """Foglio 2: indicatori per macchina (canonico produttivo)."""
    macchine = sorted(set(odl.index) | set(qta.index))
    prod = pd.DataFrame(index=macchine)
    prod["ore_lav"] = odl["ore_lav"].reindex(macchine)
    prod["ore_prep"] = odl["ore_prep"].reindex(macchine)
    prod["ore_fermo"] = odl["ore_fermo"].reindex(macchine)
    prod["scarti"] = odl["scarti"].reindex(macchine)
    prod["qta_lancio"] = odl["qta_lancio"].reindex(macchine)
    prod["pezzi_prodotti"] = qta.reindex(macchine)

    canone = [MERGE_PROD.get(m, m) for m in prod.index]
    prod = prod.groupby(canone).sum()  # fonde TMPM2+TMPMI in TMPM2_TMPMI

    out = prod.copy()
    out.insert(0, "kwh_misurati", consumi.sum(axis=0).reindex(out.index))
    out["perc_scarti"] = out["scarti"] / out["qta_lancio"] * 100
    out["kwh_per_pezzo"] = out["kwh_misurati"] / out["pezzi_prodotti"]
    out["kwh_per_ora_lav"] = out["kwh_misurati"] / out["ore_lav"]
    out["pezzi_per_ora_lav"] = out["pezzi_prodotti"] / out["ore_lav"]
    out["nota"] = ""
    nm = out.index.intersection(NON_MISURATE)
    out.loc[nm, "nota"] = "kWh non misurato"
    out.loc[nm, ["kwh_per_pezzo", "kwh_per_ora_lav"]] = None
    out.loc["TMPM2_TMPMI", "nota"] = "contatore unico: TMPM2+TMPMI fuse"
    out = out.replace([float("inf"), float("-inf")], None).round(2)

    # Ordine e intestazioni leggibili (italiano + unità)
    out = out[["kwh_misurati", "ore_lav", "ore_prep", "ore_fermo", "qta_lancio",
               "scarti", "pezzi_prodotti", "perc_scarti", "kwh_per_pezzo",
               "kwh_per_ora_lav", "pezzi_per_ora_lav", "nota"]]
    out = out.rename(columns={
        "kwh_misurati": "kWh misurati",
        "ore_lav": "Ore lavoro (h)",
        "ore_prep": "Ore prep. (h)",
        "ore_fermo": "Ore fermo (h)",
        "qta_lancio": "Qta lancio (pz)",
        "scarti": "Scarti (pz)",
        "pezzi_prodotti": "Pezzi prodotti (pz)",
        "perc_scarti": "% scarti",
        "kwh_per_pezzo": "kWh/pezzo",
        "kwh_per_ora_lav": "kWh/ora lavoro",
        "pezzi_per_ora_lav": "Pezzi/ora lavoro",
        "nota": "Nota",
    })
    out.index.name = "Macchina"
    return out


# ------------------------------- Excel output -------------------------------

HEADER_ROW = 3  # riga excel della tabella (1 = titolo, 2 = sottotitolo)
TITLE_FONT_COLOR = "1F4E78"


def it_num(x, dec=1) -> str:
    """Formatta un numero in stile italiano: '72.467,3'."""
    if x is None or pd.isna(x):
        return "-"
    s = f"{float(x):,.{dec}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def _style_sheet(ws, ncols: int, nrows: int, fmt_map=None, widths=None,
                 header: int = HEADER_ROW) -> None:
    """Formatta una tabella: header, freeze, autofilter, riga TOTALE, formati, larghezze.

    Allinea formati/larghezze sui nomi reali dell'header del foglio: robusto rispetto
    a index=True/False e all'ordine delle colonne.
    """
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF")
    total_fill = PatternFill("solid", fgColor="DDEBF7")

    last_row = header + nrows

    for col in range(1, ncols + 1):
        c = ws.cell(row=header, column=col)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.freeze_panes = f"A{header + 1}"
    ws.auto_filter.ref = f"A{header}:{get_column_letter(ncols)}{last_row}"

    for col in range(1, ncols + 1):  # riga TOTALE (ultima): sfondo
        ws.cell(row=last_row, column=col).fill = total_fill

    if fmt_map:
        for col in range(1, ncols + 1):
            fmt = fmt_map.get(ws.cell(row=header, column=col).value)
            if fmt is None:
                continue
            for r in range(header + 1, last_row + 1):
                ws.cell(row=r, column=col).number_format = fmt

    if widths:
        for col in range(1, ncols + 1):
            w = widths.get(ws.cell(row=header, column=col).value)
            if w:
                ws.column_dimensions[get_column_letter(col)].width = w


def write_excel(sintesi: pd.DataFrame, odl: pd.DataFrame, consumi: pd.DataFrame,
                ore: pd.DataFrame, fermo: pd.DataFrame, qta: pd.Series) -> None:
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF")
    total_fill = PatternFill("solid", fgColor="DDEBF7")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # --- cross-check per il foglio Note ---
    cross = {
        "kwh_misurati": float(consumi.sum().sum()),
        "kwh_enel": float(sum(b[1] for b in BOLLETTE if b[1])),
        "ore_lav_odl": float(odl["ore_lav"].sum()),
        "ore_lav_turni": float(ore.sum().sum()),
        "ore_fermo": float(odl["ore_fermo"].sum()),
        "pezzi": float(qta.sum()),
        "righe": int(len(sintesi)),
    }

    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
        # ---- 1. Note di lettura -------------------------------------------------
        note_ws = writer.book.create_sheet("Note di lettura", 0)
        _write_note_sheet(note_ws, cross)

        # ---- 2. Sintesi per macchina -------------------------------------------
        disp = sintesi.reset_index()  # colonna "Macchina"
        tot = {}
        for key, col in [("kWh misurati", "kWh misurati"), ("Ore lavoro (h)", "Ore lavoro (h)"),
                         ("Ore prep. (h)", "Ore prep. (h)"), ("Ore fermo (h)", "Ore fermo (h)"),
                         ("Qta lancio (pz)", "Qta lancio (pz)"), ("Scarti (pz)", "Scarti (pz)"),
                         ("Pezzi prodotti (pz)", "Pezzi prodotti (pz)")]:
            tot[key] = disp[col].sum()
        tot["% scarti"] = tot["Scarti (pz)"] / tot["Qta lancio (pz)"] * 100
        tot["kWh/pezzo"] = tot["kWh misurati"] / tot["Pezzi prodotti (pz)"] if tot["Pezzi prodotti (pz)"] else None
        tot["kWh/ora lavoro"] = tot["kWh misurati"] / tot["Ore lavoro (h)"] if tot["Ore lavoro (h)"] else None
        tot["Pezzi/ora lavoro"] = tot["Pezzi prodotti (pz)"] / tot["Ore lavoro (h)"] if tot["Ore lavoro (h)"] else None
        tot["Nota"] = ""
        tot_row = pd.DataFrame([{"Macchina": "TOTALE", **tot}])
        disp = pd.concat([disp, tot_row], ignore_index=True)

        disp.to_excel(writer, sheet_name="Sintesi per macchina", index=False, startrow=HEADER_ROW - 1)
        ws = writer.book["Sintesi per macchina"]
        _write_title(ws, "Sintesi per macchina (kWh / ore / pezzi)",
                     "Periodo GEN-SET 2026 | Fonti: IcoPower (kWh), ProdWare ODL/Qta (ore, pezzi, scarti)")
        _style_sheet(ws, len(disp.columns), len(disp), fmt_map={
            "kWh misurati": "#,##0.0", "Ore lavoro (h)": "#,##0.0", "Ore prep. (h)": "#,##0.0",
            "Ore fermo (h)": "#,##0.0", "Qta lancio (pz)": "#,##0", "Scarti (pz)": "#,##0",
            "Pezzi prodotti (pz)": "#,##0", "% scarti": '0.00"%"', "kWh/pezzo": "0.000",
            "kWh/ora lavoro": "#,##0.0", "Pezzi/ora lavoro": "#,##0", "Nota": None,
        }, widths={"Macchina": 14, "kWh misurati": 13, "Ore lavoro (h)": 13, "Ore prep. (h)": 13,
                   "Ore fermo (h)": 13, "Qta lancio (pz)": 13, "Scarti (pz)": 12,
                   "Pezzi prodotti (pz)": 15, "% scarti": 10, "kWh/pezzo": 10,
                   "kWh/ora lavoro": 13, "Pezzi/ora lavoro": 14, "Nota": 44})
        _tot_row(ws, HEADER_ROW + len(disp), len(disp.columns))

        # ---- 3. Consumi mensili CSV --------------------------------------------
        cons = consumi.round(1).copy()
        cons["TOTALE"] = cons.sum(axis=1)
        cons.loc["TOTALE"] = cons.sum(axis=0)
        cons.index.name = "Mese"
        cons.to_excel(writer, sheet_name="Consumi mensili CSV", startrow=HEADER_ROW - 1)
        ws = writer.book["Consumi mensili CSV"]
        _write_title(ws, "Consumi mensili per macchina (kWh)",
                     "Fonte: contatori IcoPower - differenza letture cumulative | GEN-SET 2026")
        _style_sheet(ws, len(cons.columns) + 1, len(cons),  # +1 colonna indice Mese
                     fmt_map={**{c: "#,##0.0" for c in cons.columns}},
                     widths={"Mese": 11})
        _tot_row(ws, HEADER_ROW + len(cons), len(cons.columns) + 1)

        # ---- 4. Ore lavorate mensili -------------------------------------------
        ore_d = ore.copy()
        ore_d["TOTALE"] = ore_d.sum(axis=1)
        ore_d.loc["TOTALE"] = ore_d.sum(axis=0)
        ore_d.index.name = "Macchina"
        ore_d.to_excel(writer, sheet_name="Ore lavorate mensili", startrow=HEADER_ROW - 1)
        ws = writer.book["Ore lavorate mensili"]
        _write_title(ws, "Ore lavorate mensili per macchina (h)",
                     "Fonte: file turni (disaggregazione mensile ore lavorate) | GEN-SET 2026")
        _style_sheet(ws, len(ore_d.columns) + 1, len(ore_d),  # +1 colonna indice Macchina
                     fmt_map={**{c: "#,##0.0" for c in ore_d.columns}},
                     widths={"Macchina": 14})
        _tot_row(ws, HEADER_ROW + len(ore_d), len(ore_d.columns) + 1)

        # ---- 5. Ore fermo per causa --------------------------------------------
        fer = fermo.copy()
        fer["% su totale"] = fer["ore_fermo"] / fer["ore_fermo"].sum() * 100
        fer = fer.rename(columns={"macchina": "Macchina", "causa_fermo": "Causa fermo",
                                  "ore_fermo": "Ore fermo (h)"})
        ttot = {"Macchina": "TOTALE", "Causa fermo": "",
                "Ore fermo (h)": fer["Ore fermo (h)"].sum(),
                "% su totale": 100.0}
        fer = pd.concat([fer, pd.DataFrame([ttot])], ignore_index=True)
        fer.to_excel(writer, sheet_name="Ore fermo per causa", index=False, startrow=HEADER_ROW - 1)
        ws = writer.book["Ore fermo per causa"]
        _write_title(ws, "Ore fermo per causa (h)",
                     "Fonte: ProdWare ODL (causa fermo dichiarata) | GEN-SET 2026")
        _style_sheet(ws, len(fer.columns), len(fer), fmt_map={
            "Ore fermo (h)": "#,##0.0", "% su totale": '0.00"%"',
        }, widths={"Macchina": 14, "Causa fermo": 50, "Ore fermo (h)": 14, "% su totale": 12})
        _tot_row(ws, HEADER_ROW + len(fer), len(fer.columns))

        # ---- 6. Bollette ENEL ---------------------------------------------------
        bol_df = pd.DataFrame([b for b in BOLLETTE if b[1] is not None],
                              columns=["Periodo", "kWh", "Totale Euro"])
        bol_df["Euro/kWh"] = bol_df["Totale Euro"] / bol_df["kWh"]
        tot_row = {"Periodo": "TOTALE",
                   "kWh": bol_df["kWh"].sum(),
                   "Totale Euro": round(bol_df["Totale Euro"].sum(), 2),
                   "Euro/kWh": round(bol_df["Totale Euro"].sum() / bol_df["kWh"].sum(), 4)}
        bol_df = pd.concat([bol_df, pd.DataFrame([tot_row])], ignore_index=True)
        bol_df.to_excel(writer, sheet_name="Bollette ENEL", index=False, startrow=HEADER_ROW - 1)
        ws = writer.book["Bollette ENEL"]
        _write_title(ws, "Bollette ENEL (consumo e costo energia)",
                     "Fonte: fatture ENEL in input/ | DIC 2025 - LUG 2026 | quota AGO 2026 in attesa")
        _style_sheet(ws, len(bol_df.columns), len(bol_df), fmt_map={
            "kWh": "#,##0", "Totale Euro": "#,##0.00", "Euro/kWh": "#,##0.000",
        }, widths={"Periodo": 12, "kWh": 14, "Totale Euro": 14, "Euro/kWh": 12})
        _tot_row(ws, HEADER_ROW + len(bol_df), len(bol_df.columns))

        # bordo su tutte le celle dati di ogni tabella
        for sheet in writer.book.worksheets:
            if sheet.title == "Note di lettura":
                continue
            for row in sheet.iter_rows(min_row=HEADER_ROW, max_col=sheet.max_column,
                                       max_row=sheet.max_row):
                for cell in row:
                    cell.border = border


def _write_title(ws, title: str, subtitle: str) -> None:
    """Titolo e sottotitolo sopra una tabella (righe 1-2)."""
    from openpyxl.styles import Font

    ncols = max(ws.max_column, 1)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(row=1, column=1, value=title)
    c.font = Font(bold=True, size=14, color=TITLE_FONT_COLOR)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
    s = ws.cell(row=2, column=1, value=subtitle)
    s.font = Font(italic=True, size=10, color="595959")


def _tot_row(ws, row: int, ncols: int) -> None:
    """Stile riga TOTALE: grassetto (sfondo già applicato da _style_sheet)."""
    from openpyxl.styles import Font

    for col in range(1, ncols + 1):
        ws.cell(row=row, column=col).font = Font(bold=True)


def _write_note_sheet(ws, cross: dict) -> None:
    """Foglio 1: documentazione + istruzioni di lettura (autocontenuto)."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 60
    ws.column_dimensions["C"].width = 60
    ws.column_dimensions["D"].width = 60

    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    h_fill = PatternFill("solid", fgColor="1F4E78")
    sec_font = Font(bold=True, size=12, color="1F4E78")
    wrap = Alignment(wrap_text=True, vertical="top")

    r = [1]

    def put(row, col, val, font=None, fill=None, w=None, b=False):
        c = ws.cell(row=row, column=col, value=val)
        if font:
            c.font = font
        if fill:
            c.fill = fill
        if w:
            c.alignment = w
        if b or fill:
            c.border = border
        return c

    def blank(n=1):
        r[0] += n

    def section(text):
        blank(1)
        c = ws.cell(row=r[0], column=1, value=text)
        c.font = sec_font
        r[0] += 1

    def para(text, width=None):
        ws.merge_cells(start_row=r[0], start_column=1,
                       end_row=r[0], end_column=4 if not width else 2)
        put(r[0], 1, text, w=Alignment(wrap_text=True, vertical="top"))
        r[0] += 1

    def table(headers, rows, widths=(34, 60, 60, 60)):
        for j, h in enumerate(headers, start=1):
            put(r[0], j, h, font=Font(bold=True, color="FFFFFF"), fill=h_fill,
                w=Alignment(wrap_text=True, vertical="center"), b=True)
        r[0] += 1
        for row in rows:
            for j, v in enumerate(row, start=1):
                put(r[0], j, v, w=wrap, b=True)
            r[0] += 1

    # --- intestazione
    ws.merge_cells("A1:D1")
    put(1, 1, "RENDICONTO ENERGIA VETRONAVIGLIO 2026",
        font=Font(bold=True, size=16, color="1F4E78"))
    ws.merge_cells("A2:D2")
    put(2, 1, "Documento consuntivo energia/lavoro per macchina | Destinatario: "
              "Direttore di Produzione / Responsabile Qualità",
        font=Font(italic=True, size=10, color="595959"))
    ws.merge_cells("A3:D3")
    put(3, 1, "Redatto da: ZetaByteNexus | Data elaborazione: 14/09/2026",
        font=Font(italic=True, size=10, color="595959"))
    r[0] = 5

    # --- 1. contenuto file
    section("1. COSA CONTIENE IL FILE")
    table(
        ["Foglio", "Contenuto"],
        [["Note di lettura", "Questo foglio: istruzioni, fonti, definizioni, limiti e consistenza dati."],
         ["Sintesi per macchina", "Indicatori consolidati per macchina: kWh misurati, ore (lavoro/prep./fermo), "
                                   "pezzi (lancio/scarti/prodotti), % scarti e indici di efficienza. Riga TOTALE in fondo."],
         ["Consumi mensili CSV", "kWh per macchina e per mese calcolati dai contatori IcoPower (GEN-SET 2026). Riga e colonna TOTALE."],
         ["Ore lavorate mensili", "Ore lavorate per macchina e per mese, da rilevazione turni (GEN-SET 2026). Riga e colonna TOTALE."],
         ["Ore fermo per causa", "Ore di fermo per macchina con causa dichiarata e incidenza % sul totale."],
         ["Bollette ENEL", "Consumo e costo energia bollettato DIC 2025 - LUG 2026 (AGO 2026 in attesa al momento della stesura)."]],
        widths=(30, 190),
    )

    # --- 2. fonti
    section("2. FONTI DATI")
    table(
        ["Sorgente", "File", "Contenuto", "Periodo"],
        [["Contatori IcoPower", "CSV Vetronaviglio_2026.csv", "Energia attiva per macchina (letture cumulative ogni 5 min). "
                                                               "I kWh consumati sono calcolati come differenza fra letture consecutive.",
          "01/01 - 11/09/2026"],
         ["ProdWare - ODL", "Dettaglio Analisi.xlsx", "Ordini di lavoro: ore lavoro/preparazione/fermo, causa fermo, scarti, "
                                                       "quantità emesse in lancio.",
          "Periodo complessivo (senza data singola)"],
         ["ProdWare - Qta prodotta", "Qta prodotta 2026.xlsx", "Pezzi finiti per macchina.", "01/01 - 11/09/2026"],
         ["Rilevazione turni", "Dettaglio ore lavorate per turno 2026.xlsx", "Ore lavorate per macchina, mese e turno: "
                                                                              "disaggregazione mensile delle ore lavorate.",
          "GEN - SET 2026"],
         ["Fatture ENEL", "8 PDF in input/", "Consumo e costo energia fatturato (bimestre).", "DIC 2025 - LUG 2026"]],
        widths=(30, 50, 90, 30),
    )

    # --- 3. definizioni
    section("3. DEFINIZIONI E NOTE DI CALCOLO")
    table(
        ["Campo", "Definizione"],
        [["kWh misurati", "Differenza (senza valori negativi) fra letture cumulative consecutive dei contatori IcoPower, aggregata per mese."],
         ["Ore lavoro / Ore prep.", "Ore di lavoro effettivo e ore di preparazione (attrezzaggio) per macchina, da ODL."],
         ["Ore fermo", "Ore di fermo macchina con causa dichiarata, da ODL."],
         ["Qta lancio / Scarti / Pezzi prodotti", "Quantità avviata all'ordine di lavoro, pezzi scartati e pezzi finiti contati."],
         ["% scarti", "Scarti / Qta lancio x 100."],
         ["kWh/pezzo", "kWh misurati / pezzi prodotti (solo macchine con contatore)."],
         ["kWh/ora lavoro", "kWh misurati / ore lavoro (solo macchine con contatore)."],
         ["Pezzi/ora lavoro", "Pezzi prodotti / ore lavoro."]],
        widths=(30, 190),
    )

    # --- 4. limiti
    section("4. LIMITI E NOTE CRITICHE (DA LEGGERE)")
    for t in [
        "1. Misura parziale dei consumi. I contatori IcoPower coprono circa 72.500 kWh nel periodo = ~18% dei 396.200 kWh "
        "bollettati ENEL. Il resto è carico NON misurato a macchina (aria compressa, illuminazione, uffici, ISOLA2/3/5, MANUA e "
        "macchine senza presa). I consumi per macchina vanno usati per confronti RELATIVI fra macchine, non sommati alla bolletta.",
        "2. Macchine senza contatore. MANUA, ISOLA2, ISOLA3, ISOLA5 hanno kWh = \"non misurato\": la riga riporta ore e pezzi, "
        "ma nessun consumo (celle vuote negli indici kWh).",
        "3. Contatore unico TMPM2_TMPMI. Le grandezze di TMPM2 e TMPMI (ore, pezzi, scarti) sono sommate nella riga "
        "TMPM2_TMPMI perché l'impianto di misura è unico.",
        "4. Macchine solo contatore. F08, F10, F11, M1050, MP1, MP3, MP4: letture trascurabili (circa 18 kWh totali in periodo); "
        "presenti nel foglio 'Consumi mensili CSV', escluse dalla sintesi.",
        "5. N/A esclusa. Quantità N/A (fase lancio, ~4,7 M pezzi avviati con 0 ore lavoro) non risulta confrontabile e non "
        "compare nel rendiconto.",
        "6. ODL senza data. Ore e scarti ODL sono aggregati sul periodo; la ripartizione MENSILE delle ore lavorate proviene dal "
        "file turni (totale 12.234,7 h vs 12.186,0 h ODL: scostamento 0,4%, derivante da modalità di registrazione aziendali).",
        "7. Bolletta AGO 2026 non ancora disponibile alla data di elaborazione: il totale ENEL comprende DIC 2025 - LUG 2026.",
    ]:
        para(t)

    # --- 5. consistenza
    section("5. CONSISTENZA DATI (CROSS-CHECK)")
    table(
        ["Indicatore", "Valore", "Fonte"],
        [["kWh misurati IcoPower (GEN-SET)", it_num(cross["kwh_misurati"], 1), "CSV IcoPower"],
         ["kWh bollettati ENEL (DIC 25 - LUG 26)", it_num(cross["kwh_enel"], 0), "Fatture ENEL"],
         ["Ore lavoro ODL", it_num(cross["ore_lav_odl"], 1), "ProdWare ODL"],
         ["Ore lavoro file turni (mensili)", it_num(cross["ore_lav_turni"], 1), "File turni"],
         ["Ore fermo", it_num(cross["ore_fermo"], 1), "ProdWare ODL"],
         ["Pezzi prodotti (esclusa N/A)", it_num(cross["pezzi"], 0), "ProdWare Qta"],
         ["Macchine in sintesi", str(cross["righe"]), "Macchine produttive"]],
        widths=(34, 30, 30),
    )

    # --- 6. firma
    section("6. NOTE FINALI")
    para("Elaborazione a cura di ZetaByteNexus. Dati estratti dai sistemi aziendali (IcoPower, ProdWare) e dalle fatture "
         "ENEL allegate in input/. Per aggiornamenti o chiarimenti rivolgersi al referente tecnico.")
    blank(1)
    para("Versione 1.0 | 14/09/2026", width=2)


def main():
    print("1/6 Leggo IcoPower (kWh/macchina/mese)...")
    consumi = load_icopower()
    print("2/6 Leggo ODL (ore/scarti/qta lancio)...")
    odl = load_odl()
    print("3/6 Leggo Qta prodotta (pezzi)...")
    qta = load_qta()
    print("4/6 Leggo turni (ore/macchina/mese)...")
    ore = load_turni()
    print("5/6 Assemblo sintesi...")
    sintesi = build_sintesi(consumi, odl, qta)
    print("6/6 Scrivo excel (6 fogli)...")
    fermo = load_fermo_causa()
    write_excel(sintesi, odl, consumi.round(1), ore, fermo, qta)

    print("\n=== CROSS-CHECK ===")
    print(f"kWh misurati CSV totale periodo : {consumi.sum().sum():>10,.1f}")
    print(f"kWh bollette ENEL (8 mesi)      : {sum(b[1] for b in BOLLETTE if b[1]):>10,.0f}")
    print(f"Ore lav ODL totali              : {odl['ore_lav'].sum():>10,.1f}")
    print(f"Ore lav turni (mensili)         : {ore.sum().sum():>10,.1f}")
    print(f"Ore fermo ODL totali            : {odl['ore_fermo'].sum():>10,.1f}")
    print(f"Pezzi Qta totali (escl. N/A)    : {int(qta.sum()):>10,}")
    print(f"Righe sintesi macchine          : {len(sintesi)}")
    print(f"Fatto. Apri: {OUTPUT}")


if __name__ == "__main__":
    main()