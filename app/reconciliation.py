import io
import re
from pathlib import Path
from typing import Optional

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# Source of truth: Réconciliation 2(1).ipynb
SIGNES = {
    "1_RECEIPT": 1,
    "12_DESPATCHED": -1,
    "13_DESPATCHED": -1,
    "16_DESPATCHED": 0,
    "17_CANCELED": 1,
    "17_DESPATCHED": -1,
    "2_RECEIPT": 0,
    "23_DESPATCHED": -1,
    "24_CANCELED": 1,
    "24_RECEIPT": 1,
    "28_UM V": 0,
    "30_RECEIPT": 1,
    "4_UM V": 0,
    "41_DESPATCHED": -1,
    "5_CANCELED": 1,
    "5_RECEIPT": 1,
    "1_CANCELED": 1,
    "16_CANCELED": 0,
    "23_": 0,
    "23_CANCELED": 1,
    "4_": 0,
    "41_CANCELED": 1,
    "13_CANCELED": 1,
    "37_DESPATCHED": -1,
    "42_UM V": 0,
    "36_UM V": -1,
    "20_DESPATCHED": 0,
    "2_": 0,
    "71_RECEIPT": 0,
    "72_DESPATCHED": 0,
    "21_RECEIPT": 0,
    "6_RECEIPT": 0,
    "2_CANCELED": 0,
}


def clean_text_series(s: pd.Series) -> pd.Series:
    return (
        s.fillna("")
        .astype(str)
        .str.replace("\x00", "", regex=False)
        .str.replace("\xa0", " ", regex=False)
        .str.replace("\ufeff", "", regex=False)
        .str.strip()
        .replace({"nan": "", "NaN": "", "None": ""})
    )


def read_bytes_without_nulls(data: bytes) -> bytes:
    return data.replace(b"\x00", b"")


def detect_separator(data: bytes, sample_lines: int = 30) -> str:
    text = data.decode("latin1", errors="ignore")
    lines = text.splitlines()[:sample_lines]
    candidates = [",", ";", "\t", "|"]
    scores = {sep: sum(line.count(sep) for line in lines) for sep in candidates}
    return max(scores, key=scores.get)


def read_csv_bytes(data: bytes, skiprows: int) -> pd.DataFrame:
    data = read_bytes_without_nulls(data)
    sep = detect_separator(data)
    return pd.read_csv(
        io.BytesIO(data),
        sep=sep,
        skiprows=skiprows,
        encoding="latin1",
        dtype=str,
        low_memory=False,
    )


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    # Work column-by-column to tolerate duplicate column names.
    new_cols = []
    for col in df.columns:
        name = str(col).replace("\xa0", " ").replace("\ufeff", "").strip()
        new_cols.append(name)
    df = df.copy()
    df.columns = new_cols
    for i in range(df.shape[1]):
        df.iloc[:, i] = clean_text_series(df.iloc[:, i])
    return df


def process_stock(data: bytes) -> pd.DataFrame:
    df = clean_dataframe(read_csv_bytes(data, skiprows=9))
    if df.shape[1] <= 9:
        raise ValueError("Fichier stock invalide : au moins 10 colonnes sont nécessaires.")

    stock = pd.to_numeric(
        clean_text_series(df.iloc[:, 6]).str.replace(",", ".", regex=False),
        errors="coerce",
    )
    value = pd.to_numeric(
        clean_text_series(df.iloc[:, 9]).str.replace(",", ".", regex=False),
        errors="coerce",
    )
    df["PU"] = (value / stock).replace([float("inf"), -float("inf")], 0).fillna(0)
    df["_PartNumber"] = clean_text_series(df.iloc[:, 0])
    df["_Stock"] = stock.fillna(0)
    return df


def process_movement(data: bytes, file_number: int):
    df = clean_dataframe(read_csv_bytes(data, skiprows=12))
    if df.shape[1] <= 11:
        raise ValueError(f"FMALS032 #{file_number} invalide : au moins 12 colonnes sont nécessaires.")

    # Notebook: G_Type = column G + '_' + column L, then movement logic.
    col_g = df.columns[6]
    col_l = df.columns[11]
    df.insert(12, "G_Type", clean_text_series(df[col_g]) + "_" + clean_text_series(df[col_l]))

    required = ["PartNr.", "BookType", "Quantity", "Type"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"FMALS032 #{file_number} : colonnes manquantes : {missing}")

    m = df[["PartNr.", "BookType", "Quantity", "Type", "G_Type"]].copy()
    for col in ["PartNr.", "BookType", "Type", "G_Type"]:
        m[col] = clean_text_series(m[col])
    m["Quantity"] = pd.to_numeric(
        clean_text_series(m["Quantity"])
        .str.replace("\xa0", "", regex=False)
        .str.replace(",", ".", regex=False),
        errors="coerce",
    ).fillna(0)

    # Critical for file #2: 01 -> 1, 02 -> 2, etc.
    m["BookType_normalise"] = m["BookType"].str.replace(r"^0+(\d+)$", r"\1", regex=True)
    m["Mouv"] = m["BookType_normalise"] + "_" + m["Type"]
    m["Signe"] = m["Mouv"].map(SIGNES).fillna(0)
    m["Quantity_signed"] = m["Quantity"] * m["Signe"]

    unknown = sorted(set(m.loc[(~m["Mouv"].isin(SIGNES)) & (m["Mouv"] != ""), "Mouv"]))
    tcd = pd.pivot_table(
        m,
        index="PartNr.",
        columns="Mouv",
        values="Quantity",
        aggfunc="sum",
        fill_value=0,
    ).reset_index().rename(columns={"PartNr.": "PartNumber"})

    movement_cols = [c for c in tcd.columns if c != "PartNumber"]
    for col in movement_cols:
        tcd[col] = pd.to_numeric(tcd[col], errors="coerce").fillna(0) * SIGNES.get(str(col).strip(), 0)
    tcd["Total"] = tcd[movement_cols].sum(axis=1)
    return tcd, len(unknown), unknown


def build_reconciliation(si: pd.DataFrame, sf: pd.DataFrame, tcd1: pd.DataFrame, tcd2: pd.DataFrame) -> pd.DataFrame:
    parts_si = si.iloc[:, 0].astype(str).str.strip()
    parts_sf = sf.iloc[:, 0].astype(str).str.strip()
    parts_si = parts_si[parts_si != ""]
    parts_sf = parts_sf[parts_sf != ""]
    parts = sorted(set(parts_si) | set(parts_sf))
    recon = pd.DataFrame({"PartNumber": parts})

    si_par = si.groupby("_PartNumber")["_Stock"].sum()
    sf_par = sf.groupby("_PartNumber")["_Stock"].sum()
    m1 = tcd1.set_index("PartNumber")["Total"]
    m2 = tcd2.set_index("PartNumber")["Total"]

    recon["SI"] = recon["PartNumber"].map(si_par).fillna(0)
    recon["SF"] = recon["PartNumber"].map(sf_par).fillna(0)
    recon["Mouvement 1"] = recon["PartNumber"].map(m1).fillna(0)
    recon["Mouvement 2"] = recon["PartNumber"].map(m2).fillna(0)
    recon["Mouvement"] = recon["Mouvement 1"] + recon["Mouvement 2"]

    # Notebook uses Type from SF first, then SI; PU from SF first, then SI.
    type_si = si.drop_duplicates("_PartNumber").set_index("_PartNumber").iloc[:, 3]
    type_sf = sf.drop_duplicates("_PartNumber").set_index("_PartNumber").iloc[:, 3]
    recon["Type"] = recon["PartNumber"].map(type_sf).fillna(recon["PartNumber"].map(type_si)).fillna("")

    pu_si = si.drop_duplicates("_PartNumber").set_index("_PartNumber")["PU"]
    pu_sf = sf.drop_duplicates("_PartNumber").set_index("_PartNumber")["PU"]
    recon["PU"] = recon["PartNumber"].map(pu_sf).fillna(recon["PartNumber"].map(pu_si)).fillna(0)
    recon["PU"] = pd.to_numeric(recon["PU"], errors="coerce").fillna(0)

    recon["Ecart"] = recon["SF"] - recon["SI"] - recon["Mouvement"]
    recon["Ecart €"] = recon["Ecart"] * recon["PU"]

    return recon[["PartNumber", "Type", "PU", "SI", "Mouvement 1", "Mouvement 2", "Mouvement", "SF", "Ecart", "Ecart €"]]


def export_excel(recon: pd.DataFrame, tcd1: pd.DataFrame, tcd2: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        recon.to_excel(writer, sheet_name="Réconcil", index=False)
        tcd1.to_excel(writer, sheet_name="TCD1", index=False)
        tcd2.to_excel(writer, sheet_name="TCD2", index=False)
    out.seek(0)

    # Light professional formatting without changing the data structure.
    wb = load_workbook(out)
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="002857")
            cell.alignment = Alignment(horizontal="center")
        for col_cells in ws.columns:
            letter = get_column_letter(col_cells[0].column)
            max_len = min(max(len(str(c.value or "")) for c in col_cells) + 2, 45)
            ws.column_dimensions[letter].width = max(10, max_len)
        if ws.title == "Réconcil":
            for col in [3, 4, 5, 6, 7, 8, 9, 10]:
                for cell in ws[get_column_letter(col)][1:]:
                    cell.number_format = '#,##0.00'
    final = io.BytesIO()
    wb.save(final)
    final.seek(0)
    return final.getvalue()


def reconcile(si_bytes: bytes, sf_bytes: bytes, movement1_bytes: bytes, movement2_bytes: Optional[bytes] = None):
    si = process_stock(si_bytes)
    sf = process_stock(sf_bytes)
    tcd1, unknown1_count, unknown1 = process_movement(movement1_bytes, 1)
    if movement2_bytes:
        tcd2, unknown2_count, unknown2 = process_movement(movement2_bytes, 2)
    else:
        tcd2 = pd.DataFrame(columns=["PartNumber", "Total"])
        unknown2_count, unknown2 = 0, []

    recon = build_reconciliation(si, sf, tcd1, tcd2)
    excel = export_excel(recon, tcd1, tcd2)
    summary = {
        "partnumbers": len(recon),
        "tcd1_partnumbers": len(tcd1),
        "tcd2_partnumbers": len(tcd2),
        "si_total": float(recon["SI"].sum()),
        "sf_total": float(recon["SF"].sum()),
        "movement1_total": float(recon["Mouvement 1"].sum()),
        "movement2_total": float(recon["Mouvement 2"].sum()),
        "movement_total": float(recon["Mouvement"].sum()),
        "ecart_total": float(recon["Ecart"].sum()),
        "ecart_eur_total": float(recon["Ecart €"].sum()),
        "unknown_movements_file_1": unknown1_count,
        "unknown_movements_file_2": unknown2_count,
        "unknown_movement_codes_file_1": unknown1,
        "unknown_movement_codes_file_2": unknown2,
    }
    preview = recon.head(100).where(pd.notnull(recon.head(100)), None).to_dict(orient="records")
    return excel, summary, preview
