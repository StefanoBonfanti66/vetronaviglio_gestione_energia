# -*- coding: utf-8 -*-
"""
Pipeline rendiconto energia Vetronaviglio 2025 + 2026 (confronto biennale).

Estende il rendiconto 2026 aggiungendo i dati 2025 e i fogli di confronto, per
permettere il paragone dei consumi fra i due anni. Pronto per il Direttore di
Produzione / Responsabile Qualità (autodocumentato).

INPUT (cartella input/):
  --- 2026 ---
  - CSV Vetronaviglio_2026.csv                    # IcoPower 2026: contatori cumulativi 5 min
  - Dettaglio Analisi.xlsx                        # ProdWare ODL 2026
  - Qta prodotta 2026.xlsx                        # ProdWare pezzi 2026
  - Dettaglio ore lavorate per turno 2026.xlsx    # ore/macchina/mese 2026
  - 8 PDF ENEL (DIC 2025 - LUG 2026)              # importi bollette 2026
  --- 2025 ---
  - Dati consumi e costi energetici al 30_11_2025.xlsx  # kWh IcoPower 2025 + pezzi 2025 (foglio Tabelle)
  - Dettaglio Analisi 2025.xlsx                   # ProdWare ODL 2025
  - Dettaglio ore lavorate per turno 2025.xlsx    # ore/macchina/mese 2025
  - PDF ENEL 2025                                 # importi bollette 2025
  --- Analisi efficienza ---
  - ICO-F500-2025-2026.csv                        # Contatore generaleICO-F500 (5 min, 2025-2026)

OUTPUT: rendiconto_2025_2026.xlsx (14 fogli):
  1.  Note di lettura              - istruzioni, fonti, definizioni, limiti, cross-check (2 anni)
  2.  Confronto per macchina       - kWh/ore/pezzi 2025 vs 2026 per macchina + variazioni
  3.  Confronto mensile 2025-2026  - kWh IcoPower e ore lavorate, mese per mese
  4.  Bollette ENEL 2025           - consumi/costi bollettati 2025 (3 mesi stimati)
  5.  Bollette ENEL 2026           - consumi/costi bollettati DIC25-LUG26
  6.  Sintesi per macchina 2026    - indicatori consolidati 2026
  7.  Consumi mensili 2026         - kWh per macchina x mese 2026
  8.  Ore lavorate mensili 2026    - ore per macchina x mese 2026
  9.  Ore fermo per causa 2026     - dettaglio cause fermo 2026
  10. Sintesi per macchina 2025    - indicatori consolidati 2025
  11. Consumi mensili 2025         - kWh per macchina x mese 2025
  12. Ore lavorate mensili 2025    - ore per macchina x mese 2025
  13. Ore fermo per causa 2025     - dettaglio cause fermo 2025
  14. Analisi efficienza           - profilo notte/giorno, impatto 3-turni, idle, raccomandazioni

NOTE SU METODO E PERIODI (confermate 2026-09-14):
  - Confronto su: 2025 GEN-NOV (11 mesi, il file 2025 si ferma a NOV) vs 2026
    GEN-SET (9 mesi). I periodi NON sono allineati per durata: le variazioni
    vanno lette tenendo conto dei due orizzonti temporali (vedi foglio Note).
  - Consumi IcoPower: differenza fra letture cumulative, aggregata per mese.
  - Bollette ENEL 2025: importi € estratti dai PDF per 8 mesi (GEN, APR, MAG,
    LUG, AGO, SET, OTT, NOV); i 3 mesi mancanti (FEB, MAR, GIU) sono STIMATI al
    prezzo medio €/kWh dei mesi disponibili (colonna "Stima" = Sì).
  - ODL 2025 senza date: ore/scarti/qta lancio sono aggregati sull'anno intero.

Mapping macchine (canonico = produttivo):
  CSV 'Sycma' / Tabelle 'ASS1 SYCMA' -> ASS1     # ASS.POMPE SYCMA
  CSV 'TMPMA1' / Tabelle 'TMPMA1'    -> TMPMA2   # naming drift
  TMPM2_TMPMI                        -> contatore unico (TMPM2+TMPMI fuse)
  MANUA, ISOLA2, ISOLA3, ISOLA5      -> kWh non misurato
"""

from pathlib import Path

import pandas as pd

from consolida_2026 import (
    HEADER_ROW,
    MERGE_PROD,
    NON_MISURATE,
    SOLO_CONTATORE,
    TITLE_FONT_COLOR,
    _style_sheet,
    _tot_row,
    _write_title,
    build_sintesi,
    it_num,
)

INPUT_DIR = Path("input")
OUTPUT = "rendiconto_2025_2026.xlsx"

# --- Contatori IcoPower con trasmissione interrotta nel 2026 ---
# Le macchine risultano in produzione (file turni / censimento), ma il contatore
# cumulativo non aggiorna più da metà maggio 2026: il consumo kWh è SOTTOSTIMATO.
# Verificata l'attività reale: MG5 46,97h / MG8 20,41h / MG6 133,53h a settembre 2026.
CONTATORI_INCOMPLETI = {
    "MG5": "dato incompleto: contatore fermo dal 14/05/2026 (macchina in produzione)",
    "MG8": "dato incompleto: contatore fermo dal 14/05/2026 (macchina in produzione)",
    "MG6": "dato incompleto: contatore fermo dall'11/05/2026 (macchina in produzione)",
}

# --- Mapping nomi 2025 (foglio Tabelle) -> codice canonico ---
TAB_MAP = {"ASS1 SYCMA": "ASS1", "TMPMA1": "TMPMA2", "TMPM2_TMPMI": "TMPM2_TMPMI"}
# Codici non-macchina da escludere dagli ODL 2025
ODL_ESCLUSI = {"E_ASS", "E_GE"}

# --- Bollette ENEL 2026 (valori verificati su PDF, DIC 2025 - LUG 2026) ---
BOLLETTE_2026 = [
    ("DIC 2025", 51372, 15283.15),
    ("GEN 2026", 54755, 18174.09),
    ("FEB 2026", 54897, 16321.89),
    ("MAR 2026", 47216, 15093.24),
    ("APR 2026", 36199, 9817.06),
    ("MAG 2026", 25994, 7523.72),
    ("GIU 2026", 47377, 13902.72),
    ("LUG 2026", 78390, 26859.91),
]
# --- Bollette ENEL 2025 (kWh dal foglio Tabelle; € dai PDF; None = da stimare) ---
# GEN: PDF "GEN 25.pdf"; APR: "MAG 25.pdf"; MAG: "GIU 25.pdf"; SET: OCR
# "FT ENEL - SETTEMBRE.pdf"; OTT: "FT ENEL - OTTOBRE.pdf". Manca FEB/MAR/GIU.
BOLLETTE_2025 = [
    ("GEN 2025", 61904, 20875.39),
    ("FEB 2025", 60366, None),
    ("MAR 2025", 58209, None),
    ("APR 2025", 41269, 10956.70),
    ("MAG 2025", 30809, 7932.95),
    ("GIU 2025", 33644, None),
    ("LUG 2025", 51483, 15262.33),
    ("AGO 2025", 17650, 5768.63),
    ("SET 2025", 47214, 12897.86),
    ("OTT 2025", 47533, 13101.09),
    ("NOV 2025", 50662, 14793.67),
]


# ============================ LOADERS 2026 ============================

def load_icopower() -> pd.DataFrame:
    """kWh consumati per (macchina, mese) dai contatori cumulativi IcoPower 2026."""
    df = pd.read_csv(INPUT_DIR / "CSV Vetronaviglio_2026.csv", sep=";", decimal=",")
    df.columns = [c.split("#")[0].replace("Vetronaviglio.", "").strip() for c in df.columns]
    df["timestamp"] = pd.to_datetime(df["Timestamp"])
    meters = [c for c in df.columns if c not in ("Timestamp", "timestamp")]
    df[meters] = df[meters].apply(pd.to_numeric, errors="coerce")
    delta = df[meters].diff().clip(lower=0.0)
    df["mese"] = df["timestamp"].dt.to_period("M")
    monthly = delta.groupby(df["mese"], sort=True).sum()
    from consolida_2026 import MAPPING_COUNTER

    monthly.columns = [MAPPING_COUNTER.get(c, c) for c in monthly.columns]
    monthly = monthly.T.groupby(level=0).sum().T
    monthly.index = [str(m) for m in monthly.index]
    return monthly


def load_odl(percorso: str = "Dettaglio Analisi.xlsx") -> pd.DataFrame:
    """Aggregati ODL per macchina: ore, scarti, qta lancio (totale periodo)."""
    df = pd.read_excel(INPUT_DIR / percorso, sheet_name="Dettaglio Analisi")
    df.columns = [c.strip() for c in df.columns]
    return df.groupby("Macchina", sort=True).agg(
        ore_lav=("Ore Lav. cons.", "sum"),
        ore_prep=("Ore Prep.Cons.", "sum"),
        ore_fermo=("Ore Fermo", "sum"),
        scarti=("Qta scarti", "sum"),
        qta_lancio=("Qta lancio", "sum"),
    ).round(2)


def load_qta_2026() -> pd.Series:
    """Pezzi finiti per macchina (riga unica del file Qta prodotta 2026)."""
    df = pd.read_excel(INPUT_DIR / "Qta prodotta 2026.xlsx", sheet_name="Qta prodotta")
    df.columns = [str(c).strip() for c in df.columns]
    s = df.iloc[0].drop(labels=df.columns[0])
    s = pd.to_numeric(s, errors="coerce").fillna(0.0)
    s.name = "pezzi_prodotti"
    return s


def load_turni(percorso: str = "Dettaglio ore lavorate per turno 2026.xlsx") -> pd.DataFrame:
    """Ore lavorate per macchina x mese dal file turni (header a 3 livelli).

    Le colonne giorno/turno sono sparse => somma per gruppo mese (groupby),
    mai accumulo colonna per colonna (NaN lo inquinerebbe).
    """
    raw = pd.read_excel(INPUT_DIR / percorso,
                        sheet_name="Dettaglio ore lavorate per turn", header=None)
    hdr = pd.Series(pd.to_datetime(raw.iloc[0, 2:], errors="coerce").values)
    val_cols = hdr.notna().values
    rows = (raw[0].notna()
            & ~raw[0].astype(str).str.contains("Totale")
            & raw[1].notna()
            & (raw[1].astype(str) != "Totale complessivo"))
    sub = raw.loc[rows]
    mcodes = sub[1].astype(str).str.split(" - ").str[0].values
    vals = sub.iloc[:, 2:].apply(pd.to_numeric, errors="coerce")
    vals = vals.loc[:, val_cols].reset_index(drop=True)
    months = hdr[val_cols].values
    month_key = sorted(set(months))
    key_idx = {m: i for i, m in enumerate(month_key)}
    vals.columns = [key_idx[m] for m in months]
    out = vals.T.groupby(level=0).sum().T
    out.index = mcodes
    out = out.groupby(level=0).sum().round(2)
    out.columns = [pd.to_datetime(m).strftime("%Y-%m") for m in month_key]
    return out


def load_fermo_causa(percorso: str = "Dettaglio Analisi.xlsx") -> pd.DataFrame:
    """Ore fermo per (macchina, causa), solo righe > 0, ordinate per macchina."""
    df = pd.read_excel(INPUT_DIR / percorso, sheet_name="Dettaglio Analisi")
    df.columns = [c.strip() for c in df.columns]
    g = (df.groupby(["Macchina", "Causa Fermo"])["Ore Fermo"].sum().reset_index())
    g = g[g["Ore Fermo"] > 0].sort_values(["Macchina", "Ore Fermo"], ascending=[True, False])
    g = g.rename(columns={"Macchina": "macchina", "Causa Fermo": "causa_fermo",
                          "Ore Fermo": "ore_fermo"}).round(2)
    return g


# ============================ LOADERS 2025 ============================

def _tab_blocks() -> tuple:
    """Legge i 3 blocchi mensili del foglio Tabelle (consumi, ore, qta).

    Ritorna (consumi, ore, qta), ognuno index=mese 'YYYY-MM', colonne=27 macchine.
    """
    raw = pd.read_excel(INPUT_DIR / "Dati consumi e costi energetici al 30_11_2025.xlsx",
                        sheet_name="Tabelle", header=None)
    n_mac = 27

    def blk(label_row0: int) -> pd.DataFrame:
        machine_cols = raw.iloc[label_row0, 1:1 + n_mac].tolist()
        rows = raw.iloc[label_row0 + 1:label_row0 + 15, :]
        labels = pd.to_datetime(rows.iloc[:, 0], errors="coerce")
        vals = rows.iloc[:, 1:1 + n_mac].apply(pd.to_numeric, errors="coerce")
        vals.columns = [TAB_MAP.get(str(c).strip(), str(c).strip()) for c in machine_cols]
        vals.index = labels.dt.strftime("%Y-%m")
        return vals

    return blk(2), blk(18), blk(34)  # r3 consumi, r19 ore, r35 qta


def load_consumi_2025() -> pd.DataFrame:
    """kWh per macchina x mese 2025 (foglio Tabelle, già consolidato aziendale)."""
    cons, _, _ = _tab_blocks()
    return cons.loc[[i for i in cons.index if "2025-01" <= i <= "2025-11"]]


def load_qta_2025() -> pd.Series:
    """Pezzi prodotti 2025 per macchina (foglio Tabelle, somma periodo)."""
    _, _, qta = _tab_blocks()
    s = qta.drop(index=[i for i in qta.index if i < "2025-01" or i > "2025-11"])
    s = s.sum(axis=0, skipna=True)
    s.name = "pezzi_prodotti"
    return s


def load_odl_2025() -> pd.DataFrame:
    """ODL 2025 per macchina (esclusi codici non-macchina E_*)."""
    df = load_odl("Dettaglio Analisi 2025.xlsx")
    return df.drop(index=[i for i in df.index if i in ODL_ESCLUSI], errors="ignore")


def load_fermo_causa_2025() -> pd.DataFrame:
    """Ore fermo 2025 per (macchina, causa)."""
    df = load_fermo_causa("Dettaglio Analisi 2025.xlsx")
    df = df[~df["macchina"].isin(ODL_ESCLUSI)]
    return df.reset_index(drop=True)


# ============================ CONFRONTI ============================

def build_confronto_macchine(cons25: pd.DataFrame, cons26: pd.DataFrame,
                             ore25: pd.DataFrame, ore26: pd.DataFrame,
                             qta25: pd.Series, qta26: pd.Series) -> pd.DataFrame:
    """Foglio confronto: kWh/ore/pezzi 2025 vs 2026 per macchina (unione macchine)."""
    kwh25 = cons25.loc[[i for i in cons25.index if "2025-01" <= i <= "2025-11"]].sum(axis=0)
    kwh26 = cons26.sum(axis=0)
    o25 = ore25[[c for c in ore25.columns if "2025-01" <= c <= "2025-11"]].sum(axis=1)
    o26 = ore26[[c for c in ore26.columns if "2026-01" <= c <= "2026-09"]].sum(axis=1)

    macchine = sorted(set(kwh25.index) | set(kwh26.index) | set(o25.index)
                      | set(o26.index) | set(qta25.index) | set(qta26.index))
    df = pd.DataFrame(index=macchine)

    def col(s):
        return s.reindex(macchine)

    df["kWh 2025 (GEN-NOV)"] = col(kwh25)
    df["kWh 2026 (GEN-SET)"] = col(kwh26)
    df["Ore lav. 2025 (GEN-NOV)"] = col(o25)
    df["Ore lav. 2026 (GEN-SET)"] = col(o26)
    df["Pezzi 2025 (GEN-NOV)"] = col(qta25)
    df["Pezzi 2026 (GEN-SET)"] = col(qta26)

    df["Var. kWh %"] = (df["kWh 2026 (GEN-SET)"] - df["kWh 2025 (GEN-NOV)"]) \
        / df["kWh 2025 (GEN-NOV)"] * 100
    df["Var. ore %"] = (df["Ore lav. 2026 (GEN-SET)"] - df["Ore lav. 2025 (GEN-NOV)"]) \
        / df["Ore lav. 2025 (GEN-NOV)"] * 100
    df["Var. pezzi %"] = (df["Pezzi 2026 (GEN-SET)"] - df["Pezzi 2025 (GEN-NOV)"]) \
        / df["Pezzi 2025 (GEN-NOV)"] * 100

    nota = []
    for m in macchine:
        n = []
        if pd.isna(kwh25.get(m)) and pd.notna(kwh26.get(m)):
            n.append("nel 2025 non presente/non misurata")
        if pd.notna(kwh25.get(m)) and pd.isna(kwh26.get(m)):
            n.append("nel 2026 non presente/non misurata")
        if m in NON_MISURATE:
            n.append("kWh non misurato")
        if m in SOLO_CONTATORE:
            n.append("solo contatore nel 2026")
        if m == "ASS":
            n.append("ASS.POMPE PISCITELLO (attiva solo nel 2025)")
        if m in CONTATORI_INCOMPLETI:
            n.append(CONTATORI_INCOMPLETI[m])
        nota.append("; ".join(n))
    df["Nota"] = nota
    df.index.name = "Macchina"
    df = df.replace([float("inf"), float("-inf")], None).round(2)
    return df


def build_confronto_mensile(cons25: pd.DataFrame, cons26: pd.DataFrame,
                            ore25: pd.DataFrame, ore26: pd.DataFrame) -> pd.DataFrame:
    """Foglio confronto: kWh IcoPower e ore lavorate, mese per mese 2025 vs 2026."""
    mesi = [f"2025-{i:02d}" for i in range(1, 13)]
    df = pd.DataFrame(index=mesi)
    df.index.name = "Mese"

    def kwh_by_month(x: pd.DataFrame, anno: str) -> pd.Series:
        sub = x.loc[[i for i in x.index if i.startswith(anno)]].sum(axis=1)
        sub.index = [int(i[5:7]) for i in sub.index]
        return sub.reindex(range(1, 13)).fillna(0.0)

    def ore_by_month(x: pd.DataFrame, anno: str) -> pd.Series:
        cols = [c for c in x.columns if c.startswith(anno)]
        sub = x[cols].sum(axis=0)
        sub.index = [int(i[5:7]) for i in sub.index]
        return sub.reindex(range(1, 13)).fillna(0.0)

    df["kWh 2025"] = kwh_by_month(cons25, "2025").values
    df["kWh 2026"] = kwh_by_month(cons26, "2026").values
    df["Ore lav. 2025"] = ore_by_month(ore25, "2025").values
    df["Ore lav. 2026"] = ore_by_month(ore26, "2026").values
    for col, fine in [("kWh 2025", 11), ("Ore lav. 2025", 11),
                      ("kWh 2026", 9), ("Ore lav. 2026", 9)]:
        df.loc["TOTALE", col] = df[col].iloc[:fine].sum()
    return df


def build_bollette_2025() -> pd.DataFrame:
    """Foglio bollette 2025: kWh (Tabelle) + € (PDF); mesi mancanti stimati."""
    nota_prezzi = [b for b in BOLLETTE_2025 if b[2] is not None]
    prezzo_medio = sum(b[2] for b in nota_prezzi) / sum(b[1] for b in nota_prezzi)
    righe = []
    for periodo, kwh, euro in BOLLETTE_2025:
        if euro is None:
            euro = round(kwh * prezzo_medio, 2)
            stima = "Sì (stima)"
        else:
            stima = "No"
        righe.append({"Periodo": periodo, "kWh": kwh, "Totale Euro": euro,
                      "Euro/kWh": euro / kwh, "Stima": stima})
    df = pd.DataFrame(righe)
    tot = {"Periodo": "TOTALE", "kWh": df["kWh"].sum(),
           "Totale Euro": round(df["Totale Euro"].sum(), 2),
           "Euro/kWh": df["Totale Euro"].sum() / df["kWh"].sum(), "Stima": ""}
    df = pd.concat([df, pd.DataFrame([tot])], ignore_index=True)
    return df


def build_bollette_2026() -> pd.DataFrame:
    df = pd.DataFrame(BOLLETTE_2026, columns=["Periodo", "kWh", "Totale Euro"])
    df["Euro/kWh"] = df["Totale Euro"] / df["kWh"]
    tot = {"Periodo": "TOTALE", "kWh": df["kWh"].sum(),
           "Totale Euro": round(df["Totale Euro"].sum(), 2),
           "Euro/kWh": df["Totale Euro"].sum() / df["kWh"].sum()}
    df = pd.concat([df, pd.DataFrame([tot])], ignore_index=True)
    return df


# ============================ ANALISI EFFICIENZA ============================

def load_ico_f500() -> pd.DataFrame:
    """Carica il CSV del contatore generale ICO-F500 (5 min, 2025-2026)."""
    df = pd.read_csv(INPUT_DIR / "ICO-F500-2025-2026.csv", sep=";", decimal=",",
                     encoding="utf-8-sig")
    df.columns = ["Timestamp", "kWh_cum", "kW"]
    df["Timestamp"] = pd.to_datetime(df["Timestamp"])
    df = df.set_index("Timestamp").sort_index()
    df["kwh"] = df["kW"] * 5 / 60
    df["notte"] = ((df.index.hour >= 22) | (df.index.hour < 6))
    return df


def build_analisi_efficienza(df_ico: pd.DataFrame, odl26: pd.DataFrame,
                             cons26: pd.DataFrame) -> dict:
    """Costruisce i dati per il foglio Analisi Efficienza.

    Restituisce un dizionario con 5 DataFrame:
      notte_mensile, profilo_orario, impatto_3turni, idle_macchine, raccomandazioni
    """
    # --- 1. Consumo notte vs giorno per mese ---
    righe = []
    for anno in ["2025", "2026"]:
        sub = df_ico.loc[anno].copy()
        sub["mese"] = sub.index.month
        sub["notte_kwh"] = sub["kwh"] * sub["notte"]
        sub["giorno_kwh"] = sub["kwh"] * (~sub["notte"])
        nome_anno = "2025 (GEN-DIC)" if anno == "2025" else "2026 (GEN-SET)"
        for m in range(1, 13):
            msub = sub[sub["mese"] == m]
            if len(msub) == 0:
                continue
            tot = msub["kwh"].sum()
            notte = msub["notte_kwh"].sum()
            giorno = msub["giorno_kwh"].sum()
            pct = notte / tot * 100 if tot > 0 else 0
            # Classifica: 3-turni se notte > 15%
            regime = "3 turni" if pct > 15 else ("2 turni" if pct > 8 else "1 turno")
            righe.append({"Anno": nome_anno, "Mese": m, "Totale kWh": round(tot),
                          "Giorno kWh": round(giorno), "Notte kWh": round(notte),
                          "% Notte": round(pct, 1), "Regime": regime})
    notte_df = pd.DataFrame(righe)
    # TOTALE per anno
    for anno_label in notte_df["Anno"].unique():
        sub = notte_df[notte_df["Anno"] == anno_label]
        tot = {"Anno": anno_label, "Mese": "TOTALE",
               "Totale kWh": sub["Totale kWh"].sum(),
               "Giorno kWh": sub["Giorno kWh"].sum(),
               "Notte kWh": sub["Notte kWh"].sum(),
               "% Notte": round(sub["Notte kWh"].sum() / sub["Totale kWh"].sum() * 100, 1),
               "Regime": ""}
        notte_df = pd.concat([notte_df, pd.DataFrame([tot])], ignore_index=True)

    # --- 2. Profilo orario medio (kW) ---
    ore_list = []
    for anno in ["2025", "2026"]:
        sub = df_ico.loc[anno]
        ore = sub.groupby(sub.index.hour)["kW"].mean()
        ore_list.append({"Ora": f"{anno} Media kW", **{f"{h:02d}:00": round(v, 1)
                                                        for h, v in ore.items()}})
    profilo_df = pd.DataFrame(ore_list)

    # --- 3. Impatto 3-turni ---
    pre = df_ico.loc["2026-01-01":"2026-06-26"]
    pre_notte_tot = pre.loc[pre["notte"], "kwh"].sum()
    pre_notte_fer = pre[(pre["notte"]) & (pre.index.dayofweek < 5)]
    pre_tot = pre["kwh"].sum()
    post = df_ico.loc["2026-06-29":"2026-09-14"]
    post_notte_tot = post.loc[post["notte"], "kwh"].sum()
    post_notte_fer = post[(post["notte"]) & (post.index.dayofweek < 5)]
    post_tot = post["kwh"].sum()
    n_notti_fer = len(set(post_notte_fer.index.date))
    n_pre_notti_fer = len(set(pre_notte_fer.index.date))
    baseline_notte = pre_notte_fer["kwh"].mean()
    extra_notte = post_notte_fer["kwh"].sum() - baseline_notte * n_notti_fer

    impatto = pd.DataFrame([
        {"Periodo": "Pre-3turni (01/01-26/06/2026)", "Totale kWh": round(pre_tot),
         "Notte kWh": round(pre_notte_tot), "Notte feriali kWh": round(pre_notte_fer["kwh"].sum()),
         "Notti feriali": n_pre_notti_fer,
         "Media notte/feriale": round(pre_notte_fer["kwh"].mean(), 1),
         "Regime": "1 turno (turno singolo 6-21)"},
        {"Periodo": "3-turni (29/06-14/09/2026)", "Totale kWh": round(post_tot),
         "Notte kWh": round(post_notte_tot), "Notte feriali kWh": round(post_notte_fer["kwh"].sum()),
         "Notti feriali": n_notti_fer,
         "Media notte/feriale": round(post_notte_fer["kwh"].mean(), 1),
         "Regime": "3 turni h24 (29 giugno - fine agosto)"},
        {"Periodo": "Extra notte vs baseline", "Totale kWh": "",
         "Notte kWh": "", "Notte feriali kWh": round(extra_notte),
         "Notti feriali": "",
         "Media notte/feriale": "",
         "Regime": f"≈ {round(extra_notte):,} kWh × €0,30 = €{round(extra_notte * 0.30):,}"},
    ])

    # --- 4. Idle analysis per macchina (GEN-SET 2026) ---
    # Leggi CSV IcoPower 2026
    csv26 = pd.read_csv(INPUT_DIR / "CSV Vetronaviglio_2026.csv", sep=";", decimal=",")
    csv26.columns = [c.split("#")[0].replace("Vetronaviglio.", "").strip() for c in csv26.columns]
    csv26["timestamp"] = pd.to_datetime(csv26["Timestamp"])
    meters = [c for c in csv26.columns if c not in ("Timestamp", "timestamp")]
    csv26[meters] = csv26[meters].apply(pd.to_numeric, errors="coerce")
    delta = csv26[meters].diff().clip(lower=0.0)
    delta["timestamp"] = csv26["timestamp"]
    delta["hour"] = delta["timestamp"].dt.hour
    delta["weekday"] = delta["timestamp"].dt.dayofweek
    delta["notte"] = ((delta["hour"] >= 22) | (delta["hour"] < 6))
    delta["feriale"] = delta["weekday"] < 5

    from consolida_2026 import MAPPING_COUNTER
    dt_h = 5 / 60
    idle_rows = []
    for m in meters:
        canon = MAPPING_COUNTER.get(m, m)
        vals = delta[m].dropna()
        kwh = vals[vals >= 0]
        if kwh.sum() == 0:
            continue
        ts = delta.loc[kwh.index, "timestamp"]
        p = kwh / dt_h  # kW (delta è kWh per intervallo 5 min)
        p99 = p.quantile(0.99)
        if p99 <= 0:
            continue
        idle = (p > 0.05) & (p < 0.25 * p99)
        notte_mask = (ts.dt.hour >= 22) | (ts.dt.hour < 6)
        fer_mask = ts.dt.dayofweek < 5
        # ricalcola serie kW/kWh allineate su maschere identiche
        tot_kwh = kwh.sum()
        idle_kwh = kwh[idle].sum()
        ore_att = len(kwh) * dt_h
        ore_idle = int(idle.sum() * dt_h)
        notte_kwh = kwh[notte_mask].sum()
        off_kwh = kwh[notte_mask & fer_mask].sum()
        lf = tot_kwh / (p99 * ore_att) if ore_att > 0 else 0
        idle_rows.append({
            "Macchina": canon, "kWh totale": round(tot_kwh, 1),
            "Potenza media kW": round(tot_kwh / ore_att, 1) if ore_att > 0 else 0,
            "Potenza p99 kW": round(p99, 1),
            "Load factor": round(lf, 2),
            "kWh idle (5-25% p99)": round(idle_kwh, 1),
            "% idle": round(idle_kwh / tot_kwh * 100, 1) if tot_kwh > 0 else 0,
            "Ore idle": ore_idle,
            "kWh notte": round(notte_kwh, 1),
            "kWh notte feriali": round(off_kwh, 1),
        })
    idle_df = pd.DataFrame(idle_rows).sort_values("kWh idle (5-25% p99)", ascending=False)
    idle_tot = {"Macchina": "TOTALE",
                "kWh totale": idle_df["kWh totale"].sum(),
                "Potenza media kW": "", "Potenza p99 kW": "",
                "Load factor": "", "kWh idle (5-25% p99)": idle_df["kWh idle (5-25% p99)"].sum(),
                "% idle": round(idle_df["kWh idle (5-25% p99)"].sum() / idle_df["kWh totale"].sum() * 100, 1),
                "Ore idle": "", "kWh notte": idle_df["kWh notte"].sum(),
                "kWh notte feriali": idle_df["kWh notte feriali"].sum()}
    idle_df = pd.concat([idle_df, pd.DataFrame([idle_tot])], ignore_index=True)

    # --- 5. Raccomandazioni ---
    rac = pd.DataFrame([
        {"Priorità": 1, "Azione": "Aria compressa - audit perdite e valvole automatiche",
         "Risparmio stimato": "Fino al 30% dell'aria compressa (~20% dei consumi totali)",
         "CAPEX / Payback": "Valvole calate €14-18k / payback 2-3 anni", "Note": "Perdite >80% a valle; −7% consumo per bar di pressione in meno"},
        {"Priorità": 2, "Azione": "Spegnimento F04 e TROV inattive (idle 25-32%)",
         "Risparmio stimato": "≈ €1.900/anno", "CAPEX / Payback": "Nessuno / immediato",
         "Note": "F04: 1.151 kWh idle su 4.612; TROV: 593 kWh idle su 1.870"},
        {"Priorità": 3, "Azione": "Sotto-contatori quadri generali ENEL",
         "Risparmio stimato": "Mappatura 82% dei kWh non misurati (compressori, HVAC, illuminazione)",
         "CAPEX / Payback": "Sensori ifm SD ~€5k / payback <1 anno per hotspot",
         "Note": "ICO-F500 misura solo ~18% del totale ENEL; l'82% resta non analizzabile"},
        {"Priorità": 4, "Azione": "Fotovoltaico industriale 250-500 kWp",
         "Risparmio stimato": "50-70% copertura annua; autoconsumo 75-85% (turni diurni)",
         "CAPEX / Payback": "€600-850/kWp / payback 4-6 anni (2-3 con Transizione 5.0)",
         "Note": "Transizione 5.0: credito 40-45%; iperammortamento 180%; ogni kWh FV evita €0,27-0,34"},
        {"Priorità": 5, "Azione": "HVAC/chiller: ottimizzazione setpoint e gestione 3-turni",
         "Risparmio stimato": "−8-12% consumo HVAC; −3-4% per °C di setpoint",
         "CAPEX / Payback": "Nessuno / immediato (regolazione setpoint)",
         "Note": "Extra notte 3-turni ≈ 19.000 kWh/2 mesi; valutare se tutto produttivo o spreco HVAC"},
    ])

    return {"notte_mensile": notte_df, "profilo_orario": profilo_df,
            "impatto_3turni": impatto, "idle_macchine": idle_df,
            "raccomandazioni": rac}


# ============================ EXCEL OUTPUT ============================

def _write_note_sheet_2y(ws, cross: dict) -> None:
    """Foglio 1: documentazione + istruzioni di lettura per il confronto biennale."""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    ws.column_dimensions["A"].width = 34
    for c in ("B", "C", "D"):
        ws.column_dimensions[c].width = 60

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
        ws.cell(row=r[0], column=1, value=text).font = sec_font
        r[0] += 1

    def para(text):
        ws.merge_cells(start_row=r[0], start_column=1, end_row=r[0], end_column=4)
        put(r[0], 1, text, w=wrap)
        r[0] += 1

    def table(headers, rows):
        for j, h in enumerate(headers, start=1):
            put(r[0], j, h, font=Font(bold=True, color="FFFFFF"), fill=h_fill,
                w=Alignment(wrap_text=True, vertical="center"), b=True)
        r[0] += 1
        for row in rows:
            for j, v in enumerate(row, start=1):
                put(r[0], j, v, w=wrap, b=True)
            r[0] += 1

    ws.merge_cells("A1:D1")
    put(1, 1, "RENDICONTO ENERGIA VETRONAVIGLIO 2025 - 2026",
        font=Font(bold=True, size=16, color="1F4E78"))
    ws.merge_cells("A2:D2")
    put(2, 1, "Documento consuntivo energia/lavoro per macchina e confronto biennale | "
              "Destinatario: Direttore di Produzione / Responsabile Qualità",
        font=Font(italic=True, size=10, color="595959"))
    ws.merge_cells("A3:D3")
    put(3, 1, "Redatto da: ZetaByteNexus | Data elaborazione: 14/09/2026",
        font=Font(italic=True, size=10, color="595959"))
    r[0] = 5

    section("1. COSA CONTIENE IL FILE")
    table(["Foglio", "Contenuto"], [
        ["Note di lettura", "Questo foglio: istruzioni, fonti, definizioni, limiti e consistenza dati."],
        ["Confronto per macchina", "kWh, ore lavorate e pezzi per macchina, 2025 vs 2026, con variazioni % e note. Riga TOTALE."],
        ["Confronto mensile 2025-2026", "kWh IcoPower e ore lavorate mese per mese nei due anni. Riga TOTALE."],
        ["Bollette ENEL 2025", "Consumo e costo energia bollettato 2025 (GEN-NOV). I mesi senza bolletta (FEB, MAR, GIU) sono stimati al prezzo medio."],
        ["Bollette ENEL 2026", "Consumo e costo energia bollettato DIC 2025 - LUG 2026."],
        ["Sintesi per macchina 2026", "Indicatori consolidati 2026: kWh, ore (lavoro/prep./fermo), pezzi, % scarti, indici di efficienza."],
        ["Consumi mensili 2026", "kWh per macchina e per mese 2026 (contatori IcoPower)."],
        ["Ore lavorate mensili 2026", "Ore lavorate per macchina e per mese 2026 (rilevazione turni)."],
        ["Ore fermo per causa 2026", "Ore di fermo per macchina e causa 2026."],
        ["Sintesi per macchina 2025", "Indicatori consolidati 2025 (fonte: file Tabelle + ODL 2025)."],
        ["Consumi mensili 2025", "kWh per macchina e per mese 2025 (fonte: file Tabelle)."],
        ["Ore lavorate mensili 2025", "Ore lavorate per macchina e per mese 2025 (rilevazione turni)."],
        ["Ore fermo per causa 2025", "Ore di fermo per macchina e causa 2025."],
        ["Analisi efficienza", "Profilo notte/giorno 2025 vs 2026, impatto 3 turni, analisi idle per macchina e raccomandazioni (fonte contatore generale ICO-F500)."],
    ])

    section("2. FONTI DATI")
    table(["Sorgente", "File", "Contenuto", "Periodo"], [
        ["Contatori IcoPower 2026", "CSV Vetronaviglio_2026.csv", "Energia attiva per macchina (letture cumulative ogni 5 min); consumo = differenza fra letture.", "01/01 - 11/09/2026"],
        ["Consolidato energia 2025", "Dati consumi e costi energetici al 30_11_2025.xlsx (foglio Tabelle)", "kWh per macchina e pezzi per macchina 2025 (consolidato aziendale).", "GEN - NOV 2025"],
        ["ProdWare - ODL", "Dettaglio Analisi.xlsx / Dettaglio Analisi 2025.xlsx", "Ordini di lavoro: ore lavoro/preparazione/fermo, causa fermo, scarti, qta lancio.", "2026 GEN-SET / 2025 intero anno"],
        ["ProdWare - Qta prodotta", "Qta prodotta 2026.xlsx", "Pezzi finiti per macchina 2026.", "01/01 - 11/09/2026"],
        ["Rilevazione turni", "Dettaglio ore lavorate per turno 2026.xlsx / ...2025.xlsx", "Ore lavorate per macchina, mese e turno.", "2026 GEN-SET / 2025 GEN-DIC"],
        ["Fatture ENEL", "PDF in input/", "Consumo e costo energia fatturato.", "2025 GEN-NOV / 2026 DIC25-LUG26"],
        ["Contatore generale ICO-F500", "ICO-F500-2025-2026.csv", "Energia attiva e potenza del contatore generale (5 min): profilo notte/giorno, analisi 3 turni e idle. Copre ~82-85% dei kWh ENEL.", "01/01/2025 - 14/09/2026"],
    ])

    section("3. DEFINIZIONI E NOTE DI CALCOLO")
    table(["Campo", "Definizione"], [
        ["kWh misurati", "Differenza (senza valori negativi) fra letture cumulative consecutive dei contatori IcoPower, aggregata per mese. Nel 2025 i kWh provengono dal consolidato aziendale (foglio Tabelle)."],
        ["Ore lavoro / Ore prep.", "Ore di lavoro effettivo e ore di preparazione (attrezzaggio) per macchina, da ODL."],
        ["Ore fermo", "Ore di fermo macchina con causa dichiarata, da ODL."],
        ["Qta lancio / Scarti / Pezzi prodotti", "Quantità avviata all'ordine di lavoro, pezzi scartati e pezzi finiti contati."],
        ["Var. %", "Variazione percentuale 2026 vs 2025; valori positivi = aumento nel 2026. Le durate dei periodi differiscono (vedi limiti)."],
        ["kWh/pezzo, kWh/ora, Pezzi/ora", "Indici di efficienza (solo macchine con contatore per gli indici kWh)."],
        ["Regime turni (Analisi efficienza)", "Classificazione dal consumo notturno (22:00-06:00) sul contatore generale: '1 turno' (<8%), '2 turni' (8-15%), '3 turni' (>15%)."],
        ["Idle (Analisi efficienza)", "kWh consumati con la macchina 'in tensione ma scarica': potenza fra il 5% e il 25% del picco p99 della macchina (GEN-SET 2026)."],
    ])

    section("4. LIMITI E NOTE CRITICHE (DA LEGGERE)")
    for t in [
        "1. Periodi non allineati per durata. Il 2025 copre GEN-NOV (11 mesi, il file si ferma a NOV); il 2026 copre GEN-SET (9 mesi, dati IcoPower al 11/09/2026). Le variazioni % fra i due anni vanno lette considerando i due orizzonti temporali: per il confronto a parita di mesi (GEN-SET) usare il foglio 'Confronto mensile 2025-2026'.",
        "2. Misura parziale dei consumi. I contatori IcoPower misurano una quota del totale bollettato ENEL (circa il 29% nel 2025 e il 18% nel 2026). Il resto e carico NON misurato a macchina (aria compressa, illuminazione, uffici, ISOLA2/3/5, MANUA). I consumi per macchina vanno usati per confronti RELATIVI, non sommati alla bolletta.",
        "3. Bollette ENEL 2025 stimate. Per FEB, MAR e GIU 2025 non e disponibile la bolletta in input: l'importo e stimato applicando il prezzo medio €/kWh degli altri mesi (colonna 'Stima' = Si). I volumi kWh 2025 sono invece completi e coerenti.",
        "4. Macchine senza contatore. MANUA, ISOLA2, ISOLA3, ISOLA5 hanno kWh = 'non misurato': presenti con ore e pezzi, senza consumo.",
        "5. Contatore unico TMPM2_TMPMI. Le grandezze di TMPM2 e TMPMI sono sommate perche l'impianto di misura e unico.",
        "6. Macchine attive nel 2025 e non nel 2026. F08, F10, F11, MP1, MP3, MP4 e la macchina ASS (ASS.POMPE PISCITELLO) risultano produttive nel 2025; nel 2026 non compaiono (solo contatore o dismesse). Sono mostrate nel confronto con i soli valori 2025.",
        "7. N/A esclusa. Le righe ODL con macchina N/A (fase lancio, quantita elevate con 0 ore) non sono conteggiate.",
        "8. ODL 2025 senza date. Ore, scarti e qta lancio 2025 sono aggregati sull'anno intero; la ripartizione MENSILE delle ore viene dal file turni.",
        "9. Consolidato 2025 approssimativo. Il file 'Dati consumi e costi energetici al 30_11_2025.xlsx' e un elaborato manuale: i totali dei consumi risultano comunque internamente coerenti (somma macchine = totale dichiarato).",
        "10. Contatore generale ICO-F500. Il foglio 'Analisi efficienza' usa il contatore generale, che copre circa l'82-85% dei kWh ENEL fatturati. I dati notte/giorno e idle si riferiscono a questo punto di misura, non ai singoli contatori macchina.",
        "11. Periodo 3-turni. La produzione h24 (3 turni) ha operato dal 29 giugno alla fine di agosto 2026; da settembre si e tornati al turno singolo. Il salto di consumo notturno in LUG/AGO 2026 e interamente spiegato da questo cambio di regime.",
        "12. Contatori con trasmissione interrotta. MG5 e MG8 (dal 14/05/2026) e MG6 (dall'11/05/2026) risultano in produzione ma il rispettivo contatore IcoPower non aggiorna piu: i kWh 2026 di queste macchine sono SOTTOSTIMATI (l'ultima lettura valida e di maggio). Ore e pezzi restano validi (fonti ProdWare/turni). Dove intervenire: rilevatore/trasmissione dei tre contatori (vedi report Direttore di Produzione).",
    ]:
        para(t)

    section("5. CONSISTENZA DATI (CROSS-CHECK)")
    table(["Indicatore", "2025", "2026", "Fonte"], [
        ["kWh misurati IcoPower", it_num(cross["kwh25"], 0) + " (GEN-NOV)", it_num(cross["kwh26"], 0) + " (GEN-SET)", "IcoPower / consolidato"],
        ["kWh bollettati ENEL", it_num(cross["enel25"], 0) + " (GEN-NOV)", it_num(cross["enel26"], 0) + " (DIC25-LUG26)", "Fatture ENEL"],
        ["Ore lavoro ODL", it_num(cross["lav25"], 0), it_num(cross["lav26"], 0), "ProdWare ODL"],
        ["Ore fermo ODL", it_num(cross["fermo25"], 0), it_num(cross["fermo26"], 0), "ProdWare ODL"],
        ["Ore lavoro file turni", it_num(cross["turni25"], 0) + " (GEN-DIC)", it_num(cross["turni26"], 0) + " (GEN-SET)", "File turni"],
        ["Pezzi (esclusa N/A)", it_num(cross["pezzi25"], 0), it_num(cross["pezzi26"], 0), "ProdWare Qta / Tabelle"],
    ])

    section("6. NOTE FINALI")
    para("Elaborazione a cura di ZetaByteNexus. Dati estratti dai sistemi aziendali (IcoPower, ProdWare) e dalle "
         "bollette ENEL allegate in input/. Il confronto 2025/2026 e pensato per l'analisi dei consumi e "
         "dell'efficienza energetica per macchina. Per aggiornamenti o chiarimenti rivolgersi al referente tecnico.")
    blank(1)
    para("Versione 2.0 | 14/09/2026 | Sostituisce la versione 1.0 (solo 2026)")


def write_excel(confronto: pd.DataFrame, mensile: pd.DataFrame, bol25: pd.DataFrame,
                bol26: pd.DataFrame, sintesi26: pd.DataFrame, cons26: pd.DataFrame,
                ore26: pd.DataFrame, fermo26: pd.DataFrame,
                sintesi25: pd.DataFrame, cons25: pd.DataFrame,
                ore25: pd.DataFrame, fermo25: pd.DataFrame, cross: dict,
                eff: dict | None = None) -> None:
    from openpyxl.styles import Border, Side

    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    fmt_kwh = "#,##0.0"
    fmt_h = "#,##0.0"
    fmt_pz = "#,##0"
    fmt_pct = '0.00"%"'

    with pd.ExcelWriter(OUTPUT, engine="openpyxl") as writer:
        # ---- 1. Note di lettura
        note_ws = writer.book.create_sheet("Note di lettura", 0)
        _write_note_sheet_2y(note_ws, cross)

        # ---- 2. Confronto per macchina
        disp = confronto.reset_index()
        tot = {"Macchina": "TOTALE"}
        for c in ["kWh 2025 (GEN-NOV)", "kWh 2026 (GEN-SET)", "Ore lav. 2025 (GEN-NOV)",
                  "Ore lav. 2026 (GEN-SET)", "Pezzi 2025 (GEN-NOV)", "Pezzi 2026 (GEN-SET)"]:
            tot[c] = disp[c].sum()
        for c in ["Var. kWh %", "Var. ore %", "Var. pezzi %"]:
            base25 = tot["kWh 2025 (GEN-NOV)"] if c == "Var. kWh %" else (
                tot["Ore lav. 2025 (GEN-NOV)"] if c == "Var. ore %" else tot["Pezzi 2025 (GEN-NOV)"])
            v26 = tot["kWh 2026 (GEN-SET)"] if c == "Var. kWh %" else (
                tot["Ore lav. 2026 (GEN-SET)"] if c == "Var. ore %" else tot["Pezzi 2026 (GEN-SET)"])
            tot[c] = (v26 - base25) / base25 * 100 if base25 else None
        tot["Nota"] = ""
        disp = pd.concat([disp, pd.DataFrame([tot])], ignore_index=True)
        disp.to_excel(writer, sheet_name="Confronto per macchina", index=False,
                      startrow=HEADER_ROW - 1)
        ws = writer.book["Confronto per macchina"]
        _write_title(ws, "Confronto per macchina 2025 vs 2026",
                     "2025 = GEN-NOV (11 mesi) | 2026 = GEN-SET (9 mesi) | variazioni % calcolate 2026 vs 2025")
        _style_sheet(ws, len(disp.columns), len(disp), fmt_map={
            "kWh 2025 (GEN-NOV)": fmt_kwh, "kWh 2026 (GEN-SET)": fmt_kwh,
            "Ore lav. 2025 (GEN-NOV)": fmt_h, "Ore lav. 2026 (GEN-SET)": fmt_h,
            "Pezzi 2025 (GEN-NOV)": fmt_pz, "Pezzi 2026 (GEN-SET)": fmt_pz,
            "Var. kWh %": fmt_pct, "Var. ore %": fmt_pct, "Var. pezzi %": fmt_pct,
        }, widths={"Macchina": 14, "kWh 2025 (GEN-NOV)": 13, "kWh 2026 (GEN-SET)": 13,
                   "Ore lav. 2025 (GEN-NOV)": 14, "Ore lav. 2026 (GEN-SET)": 14,
                   "Pezzi 2025 (GEN-NOV)": 13, "Pezzi 2026 (GEN-SET)": 13,
                   "Var. kWh %": 11, "Var. ore %": 11, "Var. pezzi %": 11, "Nota": 46})
        _tot_row(ws, HEADER_ROW + len(disp), len(disp.columns))

        # ---- 3. Confronto mensile
        men = mensile.reset_index()
        men.to_excel(writer, sheet_name="Confronto mensile 2025-2026", index=False,
                     startrow=HEADER_ROW - 1)
        ws = writer.book["Confronto mensile 2025-2026"]
        _write_title(ws, "Confronto mensile 2025 vs 2026 (kWh IcoPower e ore lavorate)",
                     "2025 = 11 mesi (GEN-NOV) | 2026 = 9 mesi (GEN-SET) | 2026 AGO-DIC senza dati")
        _style_sheet(ws, len(men.columns), len(men), fmt_map={
            c: (fmt_kwh if "kWh" in c else fmt_h) for c in men.columns if c != "Mese"},
            widths={"Mese": 11, "kWh 2025": 13, "kWh 2026": 13, "Ore lav. 2025": 13, "Ore lav. 2026": 13})
        _tot_row(ws, HEADER_ROW + len(men), len(men.columns))

        # ---- 4/5. Bollette ENEL
        for name, bdf, sub in [("Bollette ENEL 2025", bol25, "Anno 2025 (GEN-NOV) | FEB/MAR/GIU stimati al prezzo medio"),
                               ("Bollette ENEL 2026", bol26, "DIC 2025 - LUG 2026 | AGO 2026 in attesa")]:
            b = bdf.copy()
            b.to_excel(writer, sheet_name=name, index=False, startrow=HEADER_ROW - 1)
            ws = writer.book[name]
            _write_title(ws, "Bollette ENEL (consumo e costo energia)", f"Fonte: fatture ENEL in input/ | {sub}")
            fmt = {"kWh": "#,##0", "Totale Euro": "#,##0.00", "Euro/kWh": "#,##0.000"}
            widths = {"Periodo": 12, "kWh": 14, "Totale Euro": 14, "Euro/kWh": 12, "Stima": 12}
            _style_sheet(ws, len(b.columns), len(b), fmt_map=fmt, widths=widths)
            _tot_row(ws, HEADER_ROW + len(b), len(b.columns))

        # ---- 6-9. Dettaglio 2026
        _write_sintesi_sheet(writer, "Sintesi per macchina 2026", sintesi26,
                             "Periodo GEN-SET 2026 | Fonti: IcoPower (kWh), ProdWare (ore, pezzi)")
        _write_mesi_sheet(writer, "Consumi mensili 2026", cons26,
                          "Consumi mensili per macchina (kWh)", "GEN-SET 2026", index_name="Mese")
        _write_mesi_sheet(writer, "Ore lavorate mensili 2026", ore26,
                          "Ore lavorate mensili per macchina (h)", "GEN-SET 2026", index_name="Macchina")
        _write_fermo_sheet(writer, "Ore fermo per causa 2026", fermo26, "GEN-SET 2026")

        # ---- 10-13. Dettaglio 2025
        _write_sintesi_sheet(writer, "Sintesi per macchina 2025", sintesi25,
                             "Periodo GEN-NOV 2025 | Fonti: consolidato aziendale (kWh, pezzi), ProdWare (ore)")
        _write_mesi_sheet(writer, "Consumi mensili 2025", cons25,
                          "Consumi mensili per macchina (kWh)", "GEN-NOV 2025", index_name="Mese")
        _write_mesi_sheet(writer, "Ore lavorate mensili 2025", ore25,
                          "Ore lavorate mensili per macchina (h)", "GEN-DIC 2025", index_name="Macchina")
        _write_fermo_sheet(writer, "Ore fermo per causa 2025", fermo25, "2025 (anno intero)")

        # ---- 14. Analisi efficienza (da ICO-F500)
        if eff is not None:
            _write_efficienza_sheet(writer, eff)

        # bordo su tutte le celle dati
        for sheet in writer.book.worksheets:
            if sheet.title == "Note di lettura":
                continue
            for row in sheet.iter_rows(min_row=HEADER_ROW, max_col=sheet.max_column,
                                       max_row=sheet.max_row):
                for cell in row:
                    cell.border = border


def _write_sintesi_sheet(writer, name: str, sintesi: pd.DataFrame, subtitle: str) -> None:
    disp = sintesi.reset_index()
    tot = {"Macchina": "TOTALE"}
    for c in ["kWh misurati", "Ore lavoro (h)", "Ore prep. (h)", "Ore fermo (h)",
              "Qta lancio (pz)", "Scarti (pz)", "Pezzi prodotti (pz)"]:
        tot[c] = disp[c].sum()
    tot["% scarti"] = tot["Scarti (pz)"] / tot["Qta lancio (pz)"] * 100 if tot["Qta lancio (pz)"] else None
    tot["kWh/pezzo"] = tot["kWh misurati"] / tot["Pezzi prodotti (pz)"] if tot["Pezzi prodotti (pz)"] else None
    tot["kWh/ora lavoro"] = tot["kWh misurati"] / tot["Ore lavoro (h)"] if tot["Ore lavoro (h)"] else None
    tot["Pezzi/ora lavoro"] = tot["Pezzi prodotti (pz)"] / tot["Ore lavoro (h)"] if tot["Ore lavoro (h)"] else None
    tot["Nota"] = ""
    disp = pd.concat([disp, pd.DataFrame([tot])], ignore_index=True)
    disp.to_excel(writer, sheet_name=name, index=False, startrow=HEADER_ROW - 1)
    ws = writer.book[name]
    _write_title(ws, "Sintesi per macchina (kWh / ore / pezzi)", subtitle)
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


def _write_mesi_sheet(writer, name: str, df: pd.DataFrame, title: str,
                      periodo: str, index_name: str) -> None:
    d = df.round(1).copy()
    d["TOTALE"] = d.sum(axis=1)
    d.loc["TOTALE"] = d.sum(axis=0)
    d.index.name = index_name
    d.to_excel(writer, sheet_name=name, startrow=HEADER_ROW - 1)
    ws = writer.book[name]
    _write_title(ws, title, f"Fonte: {'contatori IcoPower / consolidato' if index_name == 'Mese' else 'file turni'} | {periodo}")
    _style_sheet(ws, len(d.columns) + 1, len(d),
                 fmt_map={c: "#,##0.0" for c in d.columns}, widths={index_name: 14})
    _tot_row(ws, HEADER_ROW + len(d), len(d.columns) + 1)


def _write_fermo_sheet(writer, name: str, fermo: pd.DataFrame, periodo: str) -> None:
    fer = fermo.copy()
    fer["% su totale"] = fer["ore_fermo"] / fer["ore_fermo"].sum() * 100
    fer = fer.rename(columns={"macchina": "Macchina", "causa_fermo": "Causa fermo",
                              "ore_fermo": "Ore fermo (h)"})
    ttot = {"Macchina": "TOTALE", "Causa fermo": "", "Ore fermo (h)": fer["Ore fermo (h)"].sum(),
            "% su totale": 100.0}
    fer = pd.concat([fer, pd.DataFrame([ttot])], ignore_index=True)
    fer.to_excel(writer, sheet_name=name, index=False, startrow=HEADER_ROW - 1)
    ws = writer.book[name]
    _write_title(ws, "Ore fermo per causa (h)",
                 f"Fonte: ProdWare ODL (causa fermo dichiarata) | {periodo}")
    _style_sheet(ws, len(fer.columns), len(fer), fmt_map={
        "Ore fermo (h)": "#,##0.0", "% su totale": '0.00"%"',
    }, widths={"Macchina": 14, "Causa fermo": 50, "Ore fermo (h)": 14, "% su totale": 12})
    _tot_row(ws, HEADER_ROW + len(fer), len(fer.columns))


def _write_efficienza_sheet(writer, eff: dict) -> None:
    """Foglio 14: Analisi efficienza energetica (profilo notte/giorno, 3-turni, idle, raccomandazioni)."""
    from openpyxl.styles import Alignment, Font, PatternFill

    notte_df = eff["notte_mensile"]
    profilo_df = eff["profilo_orario"]
    impatto_df = eff["impatto_3turni"]
    idle_df = eff["idle_macchine"]
    rac_df = eff["raccomandazioni"]

    # Notte mensile
    notte_df.to_excel(writer, sheet_name="Analisi efficienza", index=False,
                      startrow=HEADER_ROW - 1)
    ws = writer.book["Analisi efficienza"]
    _write_title(ws, "Analisi efficienza energetica",
                 "Fonte: contatore generale ICO-F500 (5 min) | Notte = 22:00-06:00 | 2025 vs 2026")
    _style_sheet(ws, len(notte_df.columns), len(notte_df), fmt_map={
        "Totale kWh": "#,##0", "Giorno kWh": "#,##0", "Notte kWh": "#,##0",
        "% Notte": "0.0"},
        widths={"Anno": 18, "Mese": 8, "Totale kWh": 12, "Giorno kWh": 12,
                "Notte kWh": 12, "% Notte": 9, "Regime": 14})
    _tot_row(ws, HEADER_ROW + len(notte_df), len(notte_df.columns))

    # Profilo orario
    r_start = HEADER_ROW + len(notte_df) + 3
    ws.cell(row=r_start, column=1, value="Profilo orario medio (kW)").font = Font(
        bold=True, size=12, color="1F4E78")
    r_start += 1
    # Headers
    ore_cols = [f"{h:02d}:00" for h in range(24)]
    for j, h in enumerate(["Anno"] + ore_cols, start=1):
        c = ws.cell(row=r_start, column=j, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(wrap_text=True, vertical="center")
    r_start += 1
    for _, row in profilo_df.iterrows():
        ws.cell(row=r_start, column=1, value=row["Ora"])
        for j, h in enumerate(range(24), start=2):
            val = row.get(f"{h:02d}:00")
            ws.cell(row=r_start, column=j, value=val)
        r_start += 1

    # Impatto 3-turni
    r_start += 1
    ws.cell(row=r_start, column=1, value="Impatto 3 turni (giugno-settembre 2026)").font = Font(
        bold=True, size=12, color="1F4E78")
    r_start += 1
    imp_cols = list(impatto_df.columns)
    for j, h in enumerate(imp_cols, start=1):
        c = ws.cell(row=r_start, column=j, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
    r_start += 1
    for _, row in impatto_df.iterrows():
        for j, col in enumerate(imp_cols, start=1):
            ws.cell(row=r_start, column=j, value=row[col])
        r_start += 1

    # Idle macchine
    r_start += 1
    ws.cell(row=r_start, column=1, value="Idle analysis per macchina (GEN-SET 2026, idle = 5-25% p99)").font = Font(
        bold=True, size=12, color="1F4E78")
    r_start += 1
    idle_cols = list(idle_df.columns)
    for j, h in enumerate(idle_cols, start=1):
        c = ws.cell(row=r_start, column=j, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
    r_start += 1
    for _, row in idle_df.iterrows():
        for j, col in enumerate(idle_cols, start=1):
            ws.cell(row=r_start, column=j, value=row[col])
        r_start += 1

    # Raccomandazioni
    r_start += 1
    ws.cell(row=r_start, column=1, value="Raccomandazioni per Direttore di Produzione").font = Font(
        bold=True, size=12, color="1F4E78")
    r_start += 1
    rac_cols = list(rac_df.columns)
    for j, h in enumerate(rac_cols, start=1):
        c = ws.cell(row=r_start, column=j, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
    r_start += 1
    for _, row in rac_df.iterrows():
        for j, col in enumerate(rac_cols, start=1):
            ws.cell(row=r_start, column=j, value=row[col])
        r_start += 1


def main():
    print("=== 2026 ===")
    print("1  IcoPower 2026...")
    cons26 = load_icopower()
    print("2  ODL 2026...")
    odl26 = load_odl("Dettaglio Analisi.xlsx")
    print("3  Qta 2026...")
    qta26 = load_qta_2026()
    print("4  Turni 2026...")
    ore26 = load_turni("Dettaglio ore lavorate per turno 2026.xlsx")
    print("5  Fermo 2026...")
    fermo26 = load_fermo_causa("Dettaglio Analisi.xlsx")
    sintesi26 = build_sintesi(cons26, odl26, qta26)
    for m, msg in CONTATORI_INCOMPLETI.items():
        if m in sintesi26.index:
            cur = str(sintesi26.at[m, "Nota"] or "")
            sintesi26.at[m, "Nota"] = (cur + "; " if cur else "") + msg

    print("=== 2025 ===")
    print("6  Consolidato 2025 (Tabelle)...")
    cons25 = load_consumi_2025()
    qta25 = load_qta_2025()
    print("7  ODL 2025...")
    odl25 = load_odl_2025()
    print("8  Turni 2025...")
    ore25 = load_turni("Dettaglio ore lavorate per turno 2025.xlsx")
    print("9  Fermo 2025...")
    fermo25 = load_fermo_causa_2025()
    sintesi25 = build_sintesi(cons25, odl25, qta25)

    print("=== CONFRONTI ===")
    confronto = build_confronto_macchine(cons25, cons26, ore25, ore26, qta25, qta26)
    mensile = build_confronto_mensile(cons25, cons26, ore25, ore26)
    bol25 = build_bollette_2025()
    bol26 = build_bollette_2026()

    cross = {
        "kwh25": float(cons25.loc[[i for i in cons25.index if "2025-01" <= i <= "2025-11"]].sum().sum()),
        "kwh26": float(cons26.sum().sum()),
        "enel25": float(bol25[bol25["Periodo"] != "TOTALE"]["kWh"].sum()),
        "enel26": float(bol26[bol26["Periodo"] != "TOTALE"]["kWh"].sum()),
        "lav25": float(odl25["ore_lav"].sum()),
        "lav26": float(odl26["ore_lav"].sum()),
        "fermo25": float(odl25["ore_fermo"].sum()),
        "fermo26": float(odl26["ore_fermo"].sum()),
        "turni25": float(ore25.sum().sum()),
        "turni26": float(ore26.sum().sum()),
        "pezzi25": float(qta25.sum()),
        "pezzi26": float(qta26.sum()),
    }

    print("=== ANALISI EFFICIENZA ===")
    print("11 ICO-F500...")
    df_ico = load_ico_f500()
    eff = build_analisi_efficienza(df_ico, odl26, cons26)

    print("12 Scrivo excel (14 fogli)...")
    write_excel(confronto, mensile, bol25, bol26, sintesi26, cons26, ore26, fermo26,
                sintesi25, cons25, ore25, fermo25, cross, eff=eff)

    print("\n=== CROSS-CHECK ===")
    print(f"kWh 2025 GEN-NOV          : {cross['kwh25']:>12,.0f}")
    print(f"kWh 2026 GEN-SET          : {cross['kwh26']:>12,.0f}")
    print(f"ENEL 2025 GEN-NOV         : {cross['enel25']:>12,.0f}")
    print(f"ENEL 2026 DIC25-LUG26     : {cross['enel26']:>12,.0f}")
    print(f"Ore lav 2025 ODL (anno)   : {cross['lav25']:>12,.0f}")
    print(f"Ore lav 2026 ODL (GEN-SET): {cross['lav26']:>12,.0f}")
    print(f"Ore fermo 2025 / 2026     : {cross['fermo25']:>12,.0f} / {cross['fermo26']:,.0f}")
    print(f"Pezzi 2025 / 2026         : {cross['pezzi25']:>12,.0f} / {cross['pezzi26']:,.0f}")
    print(f"\nFatto. Apri: {OUTPUT}")


if __name__ == "__main__":
    main()
