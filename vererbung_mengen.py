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


# ========= Hilfen =========
def _has_value(x) -> bool:
    return pd.notna(x) and str(x).strip() != ""


def _is_na(x) -> bool:
    try:
        return pd.isna(x)
    except Exception:
        return False


def _value_labels_with_counts(series: pd.Series) -> Dict[str, str]:
    """Mapping raw_value -> 'raw_value (Anzahl)'; leere Werte ausgeschlossen."""
    s = series.astype(str).str.strip()
    s = s[s != ""]
    counts = s.value_counts(dropna=True)
    return {val: f"{val} ({counts[val]})" for val in counts.index}


# ========= Normalisierung (Diakritik & Schweizer 'ss') =========
def _fold_text(t: Any) -> str:
    """Trim, kleinschreiben, Diakritik entfernen, ß->ss, fuer robustes Matching."""
    if t is None:
        return ""
    t = str(t).strip().lower()
    t = t.replace("ß", "ss")
    # NFKD zerlegt Umlaute, dann Combining-Zeichen entfernen -> a, o, u etc.
    nfkd = unicodedata.normalize("NFKD", t)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def _norm_series(s: pd.Series) -> pd.Series:
    """Serie auf gefaltete Vergleichsform bringen (siehe _fold_text)."""
    return s.astype(str).map(_fold_text)


# ========= Materialisierungs-Regeln =========
@dataclass
class _Cond:
    col: str
    op: str
    value: Any
    case_insensitive: bool = True  # wird durch _fold_text ohnehin abgedeckt


def _apply_single_condition(df: pd.DataFrame, cond: _Cond) -> pd.Series:
    """Wendet eine einzelne Bedingung (equals/contains/in/regex...) an."""
    col = cond.col
    if not col or col not in df.columns:
        return pd.Series(False, index=df.index)

    s_fold = _norm_series(df[col])
    op = str(cond.op or "equals").lower()
    val = cond.value
    val_folded = _fold_text(val) if isinstance(val, str) else val

    if op in ("eq", "equals"):
        # equals "" soll auch NaN/None matchen
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
        vals = [( _fold_text(v) if isinstance(v, str) else v ) for v in (val or [])]
        return s_fold.isin(vals)

    if op in ("regex", "matches"):
        # Achtung: Regex sieht gefaltete Strings (keine Umlaute/Diakritik)
        try:
            return s_fold.str.contains(val, flags=re.IGNORECASE, regex=True, na=False)
        except Exception:
            return pd.Series(False, index=df.index)

    # Unbekannter Operator -> kein Treffer
    return pd.Series(False, index=df.index)


def apply_materialization_rules(df: pd.DataFrame, rules: list, first_match_wins: bool = False) -> pd.DataFrame:
    """
    Wendet ein Regel-Array an (UND-Verknuepfung pro Regel).
    Schema Regel:
      {
        "when": [
          {"col":"eBKP-H","op":"contains","value":"C02.01"},
          {"col":"Material","op":"contains","value":"Daemmung"},
          {"col":"Unter Terrain","op":"equals","value":"x"}
        ],
        "then": { "set": {"eBKP-H": "E01.02 ..."} }
      }

    Spezialwert: "__KEEP__" = Spalte nicht aendern.
    """
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

        # Kombinierte UND-Maske
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

        # Setzen
        for k, v in set_map.items():
            if k not in out.columns or v == "__KEEP__":
                continue
            out.loc[idx, k] = v

        # Optional: erste-Regel-gewinnt -> markiere und spaetere Regeln ueberspringen
        if first_match_wins:
            if matched_col not in out.columns:
                out[matched_col] = False
            out.loc[idx, matched_col] = True

    # Hilfsspalte entfernen, falls vorhanden
    if matched_col in out.columns:
        out.drop(columns=matched_col, inplace=True, errors="ignore")

    return out


# ========= Regeln laden (Online-tauglich) =========
def parse_rules_text(text: str) -> list:
    """Parst JSON oder YAML aus Text. Leer -> []."""
    if not text or not str(text).strip():
        return []
    t = str(text).strip()
    # JSON zuerst
    try:
        data = json.loads(t)
        return data if isinstance(data, list) else []
    except Exception:
        pass
    # YAML Fallback
    if yaml is not None:
        try:
            data = yaml.safe_load(t)
            return data if isinstance(data, list) else []
        except Exception:
            return []
    return []


def load_rules_from_upload(file) -> list:
    """Liest Regeln aus einem Upload (JSON/YAML)."""
    if file is None:
        return []
    try:
        content = file.read()
        txt = content.decode("utf-8", errors="ignore")
        return parse_rules_text(txt)
    except Exception:
        return []


def load_rules_from_url(url: str) -> list:
    """Liest Regeln von einer RAW-URL (GitHub raw, S3...)."""
    if not url or not str(url).strip():
        return []
    if requests is None:
        st.warning("Requests nicht verfuegbar. Bitte Regeln einkopieren oder Datei hochladen.")
        return []
    try:
        r = requests.get(url, timeout=10)
        r.raise_for_status()
        return parse_rules_text(r.text)
    except Exception as e:
        st.error(f"Regeln konnten nicht geladen werden: {e}")
        return []


def load_rules_from_secrets() -> list:
    """Laedt Regeln aus Streamlit Secrets (rules_json oder RULES_URL)."""
    try:
        if "rules_json" in st.secrets:
            return parse_rules_text(st.secrets["rules_json"]) or []
        if "RULES_URL" in st.secrets:
            return load_rules_from_url(st.secrets["RULES_URL"]) or []
    except Exception:
        pass
    return []


def load_rules_from_env() -> list:
    """Laedt Regeln aus Umgebungsvariablen (RULES_URL oder RULES_PATH)."""
    url = os.getenv("RULES_URL", "").strip()
    if url:
        return load_rules_from_url(url) or []
    p = os.getenv("RULES_PATH", "").strip()
    if p and os.path.exists(p):
        try:
            with open(p, "r", encoding="utf-8") as f:
                return parse_rules_text(f.read()) or []
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
    - Master-Kontext an Subs vererben; eBKP-H der Mutter vererben, wenn 'eBKP-H Sub' undefiniert.
    - Generisch fuer Basis/'... Sub': 'Sub bevorzugen, sonst Mutter'.
    - Subs promoten; wenn mind. 1 Sub bleibt -> Mutter droppen.
    - 'GUID' bleibt Sub-GUID (falls vorhanden); 'GUID Gruppe' = GUID der Mutter.
    - 'GUID Sub' bleibt als Spalte erhalten; keine Deduplizierung.
    - Danach: Standardisieren & Werte bereinigen.
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
        if master_cols:
            mother_ctx = df.loc[is_mother, master_cols].assign(grp=grp_id[is_mother]).set_index("grp")
            for c in master_cols:
                df.loc[is_sub, c] = grp_id[is_sub].map(mother_ctx[c])

        if "eBKP-H" in cols:
            mother_ebkp_map = df.loc[is_mother, ["eBKP-H"]].assign(grp=grp_id[is_mother]).set_index("grp")["eBKP-H"]
            if "eBKP-H Sub" in cols:
                sub_ebkp = df["eBKP-H Sub"].astype(str).str.strip()
                # undefiniert: leer oder "nicht klassifiziert/keine zuordnung/nicht verfuegbar"
                undef_mask = (
                    df["eBKP-H Sub"].isna()
                    | sub_ebkp.eq("")
                    | sub_ebkp.str.contains(
                        r"(?i)(nicht klassifiziert|keine zuordnung|nicht verfu[e|�]gbar)",
                        na=False
                    )
                )
            else:
                undef_mask = pd.Series(False, index=df.index)
            mask_set = is_sub & undef_mask
            df.loc[mask_set, "eBKP-H"] = grp_id[mask_set].map(mother_ebkp_map)

    # ---------- (2) Sub bevorzugen, sonst Mutter ----------
    for base in sub_pairs:
        sub_col = f"{base} Sub"
        has_sub_val = df[sub_col].notna() & df[sub_col].astype(str).str.strip().ne("")
        if not grp_id.isna().all():
            mother_base_map = df.loc[is_mother, [base]].assign(grp=grp_id[is_mother]).set_index("grp")[base]
            mother_vals = grp_id.map(mother_base_map)
        else:
            mother_vals = pd.Series(pd.NA, index=df.index)
        tgt = is_sub
        df.loc[tgt, base] = df.loc[tgt, sub_col].where(has_sub_val[tgt], other=mother_vals[tgt])

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

    # Promoted-Zeilen erzeugen
    promoted = df.loc[keep_sub_mask].copy()
    if "GUID Sub" in cols:
        guid_sub_ok = promoted["GUID Sub"].notna() & promoted["GUID Sub"].astype(str).str.strip().ne("")
        promoted["GUID"] = promoted["GUID"].where(~guid_sub_ok, promoted["GUID Sub"])
    promoted["Mehrschichtiges Element"] = False
    promoted["Promoted"] = True
    if not mother_guid_map.empty:
        promoted["GUID Gruppe"] = grp_id[keep_sub_mask].map(mother_guid_map).values

    # Original-Muetter + gedroppte Subs entfernen; Promotions anhaengen
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
            "nicht klassifiziert|keine zuordnung|nicht verfu[e|�]gbar", na=True
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
def app(supplement_name: str, delete_enabled: bool, custom_chars: str):
    st.set_page_config(page_title="Vererbung & Mengenuebernahme", layout="wide")
    st.header("Vererbung & Mengenuebernahme")

    # Session-State fuer drei eigenstaendige Versionen
    if "df_raw" not in st.session_state:
        st.session_state["df_raw"] = None
    if "df_processed" not in st.session_state:
        st.session_state["df_processed"] = None
    if "df_final" not in st.session_state:
        st.session_state["df_final"] = None

    st.markdown("""
**Ablauf**
1) Subs anhand eBKP-H **ignorieren** (droppen) → **Verarbeitung starten**.  
2) **Regeln** (Materialisierung) optional laden (Secrets/Env/Repo/URL/Upload/Text).  
3) **Filter (UeBERSCHRIFTEN)**: Spalte waehlen → Werte waehlen → **Finalisieren**.  
4) Drei Vorschauen und Downloads: **Raw**, **Bearbeitet**, **Finalisiert**.
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

    # --- Regeln laden (Online-tauglich) ---
    with st.expander("Regelwerk (optional) – Quellen: Secrets/Env/Repo/URL/Upload/Text"):
        rules_text = st.text_area("Regeln einkopieren (JSON oder YAML)", value="", height=180, key="rules_text_area")
        rules_file = st.file_uploader("Regel-Datei hochladen (JSON/YAML)", type=["json", "yaml", "yml"], key="rules_file")
        rules_url = st.text_input("RAW-URL (z. B. GitHub raw)", value="", key="rules_url")

        rules_all: List[dict] = []
        # 0) Automatische Quellen (ohne UI)
        rules_all += load_rules_from_secrets() or []
        rules_all += load_rules_from_env() or []
        rules_all += load_rules_from_repo("rules.json") or []
        # 1) UI-Quellen (ergaenzend)
        rules_all += parse_rules_text(rules_text) or []
        rules_all += load_rules_from_upload(rules_file) or []
        rules_all += load_rules_from_url(rules_url) or []

        st.caption(f"Geladene Regeln (Summe): {len(rules_all)}")

    # --- Vorschau: Raw ---
    st.subheader("1) Originale Daten (Raw, 15 Zeilen)")
    st.dataframe(df_raw.head(15), use_container_width=True)

    # eBKP-H Optionen fuer Sub-Drop
    ebkp_options = sorted(
        pd.Series(
            pd.concat([
                df_raw.get("eBKP-H", pd.Series(dtype=object)),
                df_raw.get("eBKP-H Sub", pd.Series(dtype=object))
            ])
        ).dropna().astype(str).str.strip().unique()
    )

    # ===== Formular 1: Verarbeitung =====
    with st.form(key="form_process_001"):
        sel_drop_values = st.multiselect(
            "Subs ignorieren (droppen), wenn eBKP-H exakt gleich ist",
            options=ebkp_options,
            default=[],
            help="Exakter Textvergleich; wirkt nur auf Sub-Zeilen waehrend der Promotion."
        )
        first_match_wins = st.checkbox("Materialisierungs-Regeln: erste Regel gewinnt (Stop nach Match)", value=False)
        btn_process = st.form_submit_button("Verarbeitung starten")

    if btn_process:
        with st.spinner("Verarbeitung laeuft ..."):
            df_processed = _process_df(df_raw.copy(), drop_sub_values=sel_drop_values)
            df_processed = convert_quantity_columns(df_processed)
            if rules_all:
                df_processed = apply_materialization_rules(df_processed, rules_all, first_match_wins=first_match_wins)
        st.session_state["df_processed"] = df_processed.copy()
        st.session_state["df_final"] = None  # Reset Finalisierung

    # Verarbeitung muss erfolgt sein
    if st.session_state["df_processed"] is None:
        st.info("Bitte zuerst die Verarbeitung starten.")
        st.stop()

    # ===== Bereich 2: Filter (UeBERSCHRIFTEN) =====
    st.markdown("---")
    st.subheader("2) Filter (UeBERSCHRIFTEN)")

    df_processed = st.session_state["df_processed"]

    with st.form(key="form_filter_select_002"):
        col_options = ["-- Spalte waehlen --"] + list(df_processed.columns)
        filter_col = st.selectbox("Spalte waehlen", options=col_options, index=0, key="sel_filter_col_002")
        btn_prepare = st.form_submit_button("Werte anzeigen")

    if btn_prepare and filter_col and filter_col != "-- Spalte waehlen --":
        labels_map = _value_labels_with_counts(df_processed[filter_col])
        options_labels = [labels_map[k] for k in sorted(labels_map.keys(), key=lambda x: x.lower())]
        inv_map = {v: k for k, v in labels_map.items()}

        with st.form(key="form_filter_values_003"):
            sel_labels = st.multiselect(
                "Werte zum Entfernen (Zeilen mit diesen Werten werden gedroppt)",
                options=options_labels,
                default=[],
                help="Mehrfachauswahl moeglich. Verarbeitung erst bei 'Finalisieren'.",
                key="ms_values_003"
            )
            btn_finalize = st.form_submit_button("Finalisieren")

        if btn_finalize:
            selected_values_raw = [inv_map[lbl] for lbl in sel_labels]
            if selected_values_raw:
                mask_drop = df_processed[filter_col].astype(str).str.strip().isin(set(selected_values_raw))
                df_final = df_processed.loc[~mask_drop].reset_index(drop=True)
                st.session_state["df_final"] = df_final.copy()
                st.success(f"Finalisiert: {mask_drop.sum()} Zeilen entfernt.")
            else:
                st.warning("Keine Werte ausgewaehlt. Keine Aenderung vorgenommen.")

    # ===== Drei Vorschauen + Downloads =====
    st.markdown("---")
    st.subheader("3) Vorschauen & Downloads")

    if st.session_state["df_raw"] is not None:
        st.markdown("**Raw (15 Zeilen)**")
        st.dataframe(st.session_state["df_raw"].head(15), use_container_width=True)

    if st.session_state["df_processed"] is not None:
        st.markdown("**Bearbeitet / Prozessiert (15 Zeilen)**")
        st.dataframe(st.session_state["df_processed"].head(15), use_container_width=True)

        out_proc = io.BytesIO()
        with pd.ExcelWriter(out_proc, engine="openpyxl") as writer:
            st.session_state["df_processed"].to_excel(writer, index=False, sheet_name="Bereinigt")
        out_proc.seek(0)
        st.download_button(
            "Download: Bearbeitet (ohne Filter)",
            data=out_proc,
            file_name=f"{(supplement_name or '').strip() or 'default'}_bereinigt.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_bereinigt_only_001"
        )

    if st.session_state["df_final"] is not None:
        st.markdown("**Finalisiert (nach Filter, 15 Zeilen)**")
        st.dataframe(st.session_state["df_final"].head(15), use_container_width=True)

        out_final = io.BytesIO()
        with pd.ExcelWriter(out_final, engine="openpyxl") as writer:
            st.session_state["df_final"].to_excel(writer, index=False, sheet_name="Bereinigt_Final")
        out_final.seek(0)
        st.download_button(
            "Download: Finalisiert",
            data=out_final,
            file_name=f"{(supplement_name or '').strip() or 'default'}_bereinigt_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_bereinigt_final_003"
        )


# Lokaler/Online Start
if __name__ == "__main__":
    app(supplement_name="export", delete_enabled=True, custom_chars="")
