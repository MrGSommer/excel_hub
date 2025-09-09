# main.py
import io
import os
import re
import json
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, List, Dict, Any

import pandas as pd
import streamlit as st

# Optional: YAML fuer Regeln
try:
    import yaml
except Exception:
    yaml = None

# Optional: HTTP-Download fuer Regeln (RAW-URL)
try:
    import requests
except Exception:
    requests = None

# Eigene Utilities (mussen vorhanden sein)
from excel_utils import clean_columns_values, rename_columns_to_standard, convert_quantity_columns


# ========= Text-Normalisierung (Diakritik & Schweizer 'ss') =========
def _fold_text(t: Any) -> str:
    """Trim, lower, Diakritik entfernen, ß->ss (Schweizer Schreibweise)."""
    if t is None:
        return ""
    t = str(t).strip().lower().replace("ß", "ss")
    nfkd = unicodedata.normalize("NFKD", t)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def _norm_series(s: pd.Series) -> pd.Series:
    return s.astype(str).map(_fold_text)


def _value_labels_with_counts_multi(series_list: List[pd.Series]) -> Dict[str, str]:
    """Zaehlwerte ueber mehrere Serien; leer ausschliessen. Rueckgabe: val -> 'val (Anzahl)'."""
    vals = []
    for s in series_list:
        if s is not None:
            vals.append(s.astype(str).str.strip())
    if not vals:
        return {}
    s = pd.concat(vals, ignore_index=True)
    s = s[~s.isna()]
    s = s[s != ""]
    counts = s.value_counts(dropna=True)
    return {val: f"{val} ({counts[val]})" for val in counts.index}


# ========= Materialisierungs-Regeln =========
@dataclass
class _Cond:
    col: str
    op: str
    value: Any
    case_insensitive: bool = True  # wird durch _fold_text ohnehin abgedeckt


def _apply_single_condition(df: pd.DataFrame, cond: _Cond) -> pd.Series:
    col = cond.col
    if not col or col not in df.columns:
        return pd.Series(False, index=df.index)

    s_fold = _norm_series(df[col])
    op = str(cond.op or "equals").lower()
    val = cond.value
    val_folded = _fold_text(val) if isinstance(val, str) else val

    if op in ("eq", "equals"):
        if isinstance(val, str) and val_folded == "":
            return s_fold.eq("") | df[col].isna()
        return s_fold.eq(val_folded)

    if op in ("neq", "not_equals"):
        if isinstance(val, str) and val_folded == "":
            return ~(s_fold.eq("") | df[col].isna())
        return ~s_fold.eq(val_folded)

    if op in ("contains", "icontains"):
        if val is None:
            return pd.Series(False, index=df.index)
        return s_fold.str.contains(re.escape(val_folded), na=False)

    if op in ("in",):
        vals = [(_fold_text(v) if isinstance(v, str) else v) for v in (val or [])]
        return s_fold.isin(vals)

    if op in ("regex", "matches"):
        try:
            return s_fold.str.contains(val, flags=re.IGNORECASE, regex=True, na=False)
        except Exception:
            return pd.Series(False, index=df.index)

    return pd.Series(False, index=df.index)


def apply_materialization_rules(df: pd.DataFrame, rules: list, first_match_wins: bool = False) -> pd.DataFrame:
    """Regeln (UND-verknuepft) anwenden; '__KEEP__' laesst Spalten unveraendert."""
    if not rules:
        return df

    out = df.copy()
    matched_col = "__matched__"

    for rule in rules:
        conds = rule.get("when", []) or []
        then_raw = rule.get("then", {}) or {}
        set_map: dict = (then_raw.get("set") or {})

        if not conds or not isinstance(set_map, dict):
            continue

        masks = []
        for c in conds:
            cond = _Cond(
                col=c.get("col"),
                op=c.get("op", "equals"),
                value=c.get("value"),
                case_insensitive=bool(c.get("case_insensitive", True)),
            )
            masks.append(_apply_single_condition(out, cond))
        mask = masks[0]
        for m in masks[1:]:
            mask &= m

        idx = out.index[mask]
        if len(idx) == 0:
            continue

        for k, v in set_map.items():
            if k not in out.columns or v == "__KEEP__":
                continue
            out.loc[idx, k] = v

        if first_match_wins:
            if matched_col not in out.columns:
                out[matched_col] = False
            out.loc[idx, matched_col] = True

    if matched_col in out.columns:
        out.drop(columns=matched_col, inplace=True, errors="ignore")

    return out


# ========= Regeln laden =========
def parse_rules_text(text: str) -> list:
    if not text or not str(text).strip():
        return []
    t = str(text).strip()
    try:
        data = json.loads(t)
        return data if isinstance(data, list) else []
    except Exception:
        pass
    if yaml is not None:
        try:
            data = yaml.safe_load(t)
            return data if isinstance(data, list) else []
        except Exception:
            return []
    return []


def load_rules_from_repo(filename: str = "rules.json") -> list:
    """Laedt rules.json aus dem Deploy-Repo (gleicher Ordner wie main.py)."""
    try:
        base_dir = Path(__file__).parent
        path = base_dir / filename
        if path.exists():
            return parse_rules_text(path.read_text(encoding="utf-8")) or []
    except Exception:
        pass
    return []


# ========= Kernverarbeitung (vektorisiert) =========
def _process_df(
    df: pd.DataFrame,
    drop_sub_values: Optional[List[str]] = None,  # eBKP-H exakte Werte: nur Sub-Zeilen droppen
) -> pd.DataFrame:
    """
    Ablauf:
    - Master-Kontext an Subs vererben.
    - eBKP-H der Mutter an alle Nicht-Muetter vererben, wenn Zeilen-eBKP-H undefiniert (""/NaN/Platzhalter).
    - Werte uebernehmen: '... Sub' bevorzugen, sonst Mutter (nur wenn Ziel leer).
    - Hauptspalten vor 'Einzelteile' ohne Pendant '... Sub' (ausser GUID) an Nicht-Muetter vererben (wenn Ziel leer).
    - Subs promoten; wenn mind. 1 Sub bleibt -> Mutter droppen.
    - 'GUID Sub' bleibt erhalten; keine Deduplizierung.
    - Standardisieren & Werte bereinigen.
    """
    drop_set = {str(v).strip().lower() for v in (drop_sub_values or []) if str(v).strip()}
    cols = pd.Index(df.columns)

    # Master-Kontext & Sub-Paare
    master_cols = [c for c in ["Teilprojekt", "Gebaeude", "Baufeld", "Geschoss", "Umbaustatus", "Unter Terrain", "Typ"] if c in cols]
    sub_pairs = sorted({c[:-4] for c in cols if c.endswith(" Sub") and c[:-4] in cols and c[:-4] != "GUID"})

    # Mutter/Sub-Flags (robust)
    if master_cols:
        any_master = df[master_cols].notna().any(axis=1)
        all_master_na = df[master_cols].isna().all(axis=1)
        is_mother = any_master
        is_sub = all_master_na
    else:
        is_mother = pd.Series(False, index=df.index)
        is_sub = pd.Series(False, index=df.index)

    # Gruppen-IDs (laufende Nummern ueber Muetter)
    grp_id = is_mother.cumsum()
    grp_id = grp_id.mask(grp_id.eq(0))  # vor erster Mutter -> NA

    # Meta-Spalten
    df["Mehrschichtiges Element"] = is_sub
    df["Promoted"] = False
    if "GUID Gruppe" not in df.columns:
        df["GUID Gruppe"] = pd.NA

    # Mutter-GUID je Gruppe
    if "GUID" in cols:
        mother_guid_map = df.loc[is_mother, ["GUID"]].assign(grp=grp_id[is_mother]).set_index("grp")["GUID"]
        df.loc[is_mother, "GUID Gruppe"] = df.loc[is_mother, "GUID"]
    else:
        mother_guid_map = pd.Series(dtype=object)

    # ---------- (1) Master-Kontext + eBKP-H vererben ----------
    if not grp_id.isna().all():
        # Master-Kontext fuer echte Subs (alle Master leer)
        if master_cols:
            mother_ctx = df.loc[is_mother, master_cols].assign(grp=grp_id[is_mother]).set_index("grp")
            for c in master_cols:
                df.loc[is_sub, c] = grp_id[is_sub].map(mother_ctx[c])

        # eBKP-H der Mutter an ALLE Nicht-Muetter, wenn Zeilen-eBKP-H undefiniert
        if "eBKP-H" in cols:
            mother_ebkp_map = df.loc[is_mother, ["eBKP-H"]].assign(grp=grp_id[is_mother]).set_index("grp")["eBKP-H"]

            ebkp_sub = df["eBKP-H Sub"] if "eBKP-H Sub" in cols else pd.Series(pd.NA, index=df.index)
            ebkp_row = ebkp_sub.where(ebkp_sub.astype(str).str.strip().ne(""), df.get("eBKP-H", pd.Series(pd.NA, index=df.index)))

            undef_row = (
                ebkp_row.isna()
                | ebkp_row.astype(str).str.strip().eq("")
                | ebkp_row.astype(str).str.contains(r"(?i)nicht\s*klassifiziert|keine\s*zuordnung|nicht\s+verf[uü]gbar", na=False)
            )

            eligible = (~is_mother) & grp_id.notna() & undef_row
            df.loc[eligible, "eBKP-H"] = grp_id[eligible].map(mother_ebkp_map)

    # ---------- (2) Werte uebernehmen: Sub bevorzugen, sonst Mutter ----------
    tgt = (~is_mother) & grp_id.notna()

    for base in sub_pairs:
        sub_col = f"{base} Sub"

        has_sub_val = df[sub_col].notna() & df[sub_col].astype(str).str.strip().ne("")
        if not grp_id.isna().all():
            mother_base_map = df.loc[is_mother, [base]].assign(grp=grp_id[is_mother]).set_index("grp")[base]
            mother_vals = grp_id.map(mother_base_map)
        else:
            mother_vals = pd.Series(pd.NA, index=df.index)

        need_fill = df[base].isna() | df[base].astype(str).str.strip().eq("")

        # 1) Sub-Wert uebernehmen, wenn vorhanden
        df.loc[tgt & has_sub_val & need_fill, base] = df.loc[tgt & has_sub_val & need_fill, sub_col]
        # 2) sonst Mutterwert
        df.loc[tgt & ~has_sub_val & need_fill, base] = mother_vals[tgt & ~has_sub_val & need_fill]

    # ---------- (2b) Hauptspalten ohne Pendant '... Sub' vor 'Einzelteile' vererben ----------
    cols_list = list(df.columns)
    boundary = cols_list.index("Einzelteile") if "Einzelteile" in cols_list else len(cols_list)
    inherit_cols = [c for c in cols_list[:boundary] if c != "GUID" and f"{c} Sub" not in df.columns]

    if inherit_cols and not grp_id.isna().all():
        for base in inherit_cols:
            mother_map = df.loc[is_mother, [base]].assign(grp=grp_id[is_mother]).set_index("grp")[base]
            need_fill = df[base].isna() | df[base].astype(str).str.strip().eq("")
            df.loc[tgt & need_fill, base] = grp_id[tgt & need_fill].map(mother_map)

    # ---------- (3) Sub-Drop gem. eBKP-H + Promotion ----------
    if "eBKP-H Sub" in cols:
        ebkp_sub_norm = df["eBKP-H Sub"].astype(str).str.strip().str.lower()
    else:
        ebkp_sub_norm = pd.Series("", index=df.index)

    if "eBKP-H" in cols:
        ebkp_norm = df["eBKP-H"].astype(str).str.strip().str.lower()
    else:
        ebkp_norm = pd.Series("", index=df.index)

    if drop_set:
        drop_sub_mask = is_sub & (ebkp_sub_norm.isin(drop_set) | ebkp_norm.isin(drop_set))
    else:
        drop_sub_mask = pd.Series(False, index=df.index)

    keep_sub_mask = is_sub & ~drop_sub_mask & grp_id.notna()
    grps_with_kept_sub = pd.Index(grp_id[keep_sub_mask].unique()).dropna()
    drop_mother_mask = is_mother & grp_id.isin(grps_with_kept_sub)

    promoted = df.loc[keep_sub_mask].copy()
    if "GUID Sub" in cols:
        guid_sub_ok = promoted["GUID Sub"].notna() & promoted["GUID Sub"].astype(str).str.strip().ne("")
        promoted["GUID"] = promoted["GUID"].where(~guid_sub_ok, promoted["GUID Sub"])
    promoted["Mehrschichtiges Element"] = False
    promoted["Promoted"] = True
    if not mother_guid_map.empty:
        promoted["GUID Gruppe"] = grp_id[keep_sub_mask].map(mother_guid_map).values

    to_drop_idx = df.index[drop_mother_mask | drop_sub_mask]
    if len(to_drop_idx) > 0:
        df = df.drop(index=to_drop_idx)
    if not promoted.empty:
        df = pd.concat([df, promoted], ignore_index=True)

    # ---------- (4) Sub-Spalten entfernen, aber 'GUID Sub' behalten ----------
    subs_to_drop = [c for c in df.columns if c.endswith(" Sub") and c != "GUID Sub"]
    df.drop(columns=subs_to_drop, inplace=True, errors="ignore")

    # ---------- (5) Restbereinigung ----------
    if "Unter Terrain" in df.columns:
        df.loc[df["Unter Terrain"] == "oi", "Unter Terrain"] = pd.NA
    if "eBKP-H" in df.columns:
        mask_invalid = df["eBKP-H"].astype(str).str.lower().str.contains(
            r"nicht\s*klassifiziert|keine\s*zuordnung|nicht\s+verf[uü]gbar", na=True
        )
        df = df[~mask_invalid]
    for c in ["Einzelteile", "Farbe"]:
        if c in df.columns:
            df.drop(columns=c, inplace=True)

    df.reset_index(drop=True, inplace=True)

    # ---------- (6) Standardisieren & Werte bereinigen ----------
    df = rename_columns_to_standard(df)
    df = clean_columns_values(df, delete_enabled=True, custom_chars="")

    return df


# ========= Streamlit App =========
def app():
    st.set_page_config(page_title="Vererbung & Mengenuebernahme", layout="wide")
    st.header("Vererbung & Mengenuebernahme")

    # Session-State
    if "df_raw" not in st.session_state:
        st.session_state["df_raw"] = None
    if "df_step1" not in st.session_state:
        st.session_state["df_step1"] = None
    if "df_final" not in st.session_state:
        st.session_state["df_final"] = None

    st.markdown("""
**Ablauf**
1) **Einlesen & Bereinigung** (Vererbung/Promotion, keine Regeln, kein Sub-Drop).  
2) **Subs droppen (eBKP-H exakt)** + **Material-Filter** → danach **Regeln aus rules.json** anwenden.  
Vorschauen: **Raw**, **Schritt 1**, **Finalisiert**.
    """)

    # --- Datei laden ---
    uploaded_file = st.file_uploader("Excel-Datei laden", type=["xlsx", "xls"], key="upl_file")
    if not uploaded_file:
        st.stop()

    try:
        df_raw = pd.read_excel(uploaded_file, engine="openpyxl")
        st.session_state["df_raw"] = df_raw.copy()
    except Exception as e:
        st.error(f"Fehler beim Einlesen: {e}")
        st.stop()

    # ===== Schritt 1: Einlesen & Bereinigung =====
    st.subheader("1) Einlesen & Bereinigung")
    with st.form(key="form_step1_process"):
        btn_step1 = st.form_submit_button("Schritt 1 starten (Bereinigung)")
    if btn_step1:
        with st.spinner("Schritt 1 laeuft ..."):
            df_step1 = _process_df(df_raw.copy(), drop_sub_values=[])   # kein Sub-Drop hier
            df_step1 = convert_quantity_columns(df_step1)
        st.session_state["df_step1"] = df_step1.copy()
        st.session_state["df_final"] = None

    if st.session_state["df_step1"] is None:
        st.info("Bitte zuerst Schritt 1 starten (Bereinigung).")
        st.stop()

    df_step1 = st.session_state["df_step1"]

    # ===== Schritt 2: Sub-Drop + Material-Filter + REGELN (am Ende) =====
    st.markdown("---")
    st.subheader("2) Subs droppen (eBKP-H) + Material-Filter + REGELN")

    # eBKP-H Optionen (aus RAW)
    ebkp_options = sorted(
        pd.Series(
            pd.concat([
                df_raw.get("eBKP-H", pd.Series(dtype=object)),
                df_raw.get("eBKP-H Sub", pd.Series(dtype=object))
            ], ignore_index=True)
        ).dropna().astype(str).str.strip().unique()
    )

    # Material-Optionen (aus RAW: Material + Material Sub)
    mat_series = []
    if "Material" in df_raw.columns:
        mat_series.append(df_raw["Material"])
    if "Material Sub" in df_raw.columns:
        mat_series.append(df_raw["Material Sub"])
    materials_map = _value_labels_with_counts_multi(mat_series)
    materials_labels = [materials_map[k] for k in sorted(materials_map.keys(), key=lambda x: x.lower())]
    materials_inv = {v: k for k, v in materials_map.items()}

    with st.form(key="form_step2"):
        sel_drop_values = st.multiselect(
            "Subs ignorieren (droppen), wenn eBKP-H exakt gleich ist",
            options=ebkp_options,
            default=[],
            help="Exakter Textvergleich; wirkt nur auf Sub-Zeilen waehrend der Promotion."
        )
        sel_material_labels = st.multiselect(
            "Material zum Entfernen (Material & Material Sub zusammengefasst)",
            options=materials_labels,
            default=[],
            help="Zeilen mit diesen effektiven Material-Werten werden entfernt."
        )
        first_match_wins = st.checkbox("Materialisierungs-Regeln: erste Regel gewinnt (Stop nach Match)", value=False)
        btn_step2 = st.form_submit_button("Schritt 2 starten")

    if btn_step2:
        with st.spinner("Schritt 2 laeuft ..."):
            # 2a) Neu aus RAW prozessieren MIT Sub-Drop (deterministisch)
            df_after_subdrop = _process_df(df_raw.copy(), drop_sub_values=sel_drop_values)
            df_after_subdrop = convert_quantity_columns(df_after_subdrop)

            # 2b) Material-Filter VOR Regeln
            selected_materials = [materials_inv[lbl] for lbl in sel_material_labels]
            if selected_materials:
                mat_eff = df_after_subdrop.get("Material", pd.Series("", index=df_after_subdrop.index)).astype(str).str.strip()
                df_after_filter = df_after_subdrop.loc[~mat_eff.isin(set(selected_materials))].reset_index(drop=True)
            else:
                df_after_filter = df_after_subdrop

            # 2c) REGELN GANZ AM ENDE (aus rules.json im Repo)
            rules_all: List[dict] = load_rules_from_repo("rules.json") or []
            df_final = apply_materialization_rules(df_after_filter.copy(), rules_all, first_match_wins=first_match_wins) if rules_all else df_after_filter.copy()

        st.session_state["df_final"] = df_final.copy()
        st.success("Schritt 2 abgeschlossen.")

    # ===== Vorschauen & Downloads =====
    st.markdown("---")
    st.subheader("3) Vorschauen & Downloads")

    st.markdown("**Raw (15 Zeilen)**")
    st.dataframe(st.session_state["df_raw"].head(15), use_container_width=True)

    st.markdown("**Schritt 1 – Bereinigt (15 Zeilen)**")
    st.dataframe(df_step1.head(15), use_container_width=True)
    out_step1 = io.BytesIO()
    with pd.ExcelWriter(out_step1, engine="openpyxl") as writer:
        df_step1.to_excel(writer, index=False, sheet_name="Bereinigt_Step1")
    out_step1.seek(0)
    st.download_button(
        "Download: Schritt 1 (ohne Sub-Drop & ohne Regeln)",
        data=out_step1,
        file_name="export_bereinigt_step1.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_bereinigt_step1"
    )

    if st.session_state.get("df_final") is not None:
        df_final = st.session_state["df_final"]
        st.markdown("**Finalisiert (15 Zeilen)**")
        st.dataframe(df_final.head(15), use_container_width=True)
        out_final = io.BytesIO()
        with pd.ExcelWriter(out_final, engine="openpyxl") as writer:
            df_final.to_excel(writer, index=False, sheet_name="Final_Step2")
        out_final.seek(0)
        st.download_button(
            "Download: Finalisiert (nach Sub-Drop, Material-Filter, REGELN)",
            data=out_final,
            file_name="export_final_step2.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_final_step2"
        )
    else:
        st.info("Schritt 2 noch nicht ausgefuehrt.")
