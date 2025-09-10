import io
import re
import json
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, List, Dict, Any

import pandas as pd
import streamlit as st

# Eigene Utilities (muessen vorhanden sein)
from excel_utils import clean_columns_values, rename_columns_to_standard, convert_quantity_columns


# ========= Text-Normalisierung (Diakritik & Schweizer 'ss') =========
def _fold_text(t: Any) -> str:
    """Trim, lower, Diakritik entfernen, ss-Schreibweise."""
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


# ========= Regeln laden (JSON-only) =========
def parse_rules_text(text: str) -> list:
    """Erwartet JSON: Liste von Regeln ODER Objekt mit Key 'rules' (Liste)."""
    if not text or not str(text).strip():
        return []
    t = str(text).strip()
    try:
        data = json.loads(t)
    except Exception as e:
        if "st" in globals():
            st.error(f"rules.json ist kein gueltiges JSON: {e}")
        return []

    rules = []
    if isinstance(data, list):
        rules = data
    elif isinstance(data, dict) and isinstance(data.get("rules"), list):
        rules = data["rules"]
    else:
        if "st" in globals():
            st.warning("rules.json geladen, aber keine Liste oder 'rules'-Liste gefunden.")
        return []

    valid_ops = {"eq","equals","neq","not_equals","contains","icontains","in","regex","matches","lt","le","gt","ge","checked"}
    cleaned = []
    for i, r in enumerate(rules):
        if not isinstance(r, dict):
            continue
        when = r.get("when", [])
        then = r.get("then", {})
        if not isinstance(when, list) or not isinstance(then, dict):
            continue
        ok = True
        for c in when:
            if not isinstance(c, dict):
                ok = False; break
            if "col" not in c or "op" not in c or "value" not in c:
                ok = False; break
            if str(c["op"]).lower() not in valid_ops:
                ok = False; break
        if ok:
            cleaned.append(r)
        else:
            if "st" in globals():
                st.caption(f"Regel {i} uebersprungen: ungueltiges Format/Operator.")
    return cleaned


def load_rules_from_repo(filename: str = "rules.json") -> list:
    """Laedt nur JSON-Regeln aus dem lokalen Repo."""
    try:
        base_dir = Path(__file__).parent
        path = base_dir / filename
        if not path.exists():
            if "st" in globals():
                st.info(f"{filename} nicht gefunden in {base_dir}.")
            return []
        if path.suffix.lower() != ".json":
            if "st" in globals():
                st.warning(f"{filename} ist keine .json-Datei. Bitte JSON verwenden.")
            return []

        txt = path.read_text(encoding="utf-8")
        rules = parse_rules_text(txt)
        if "st" in globals():
            st.caption(f"rules.json geladen: {len(rules)} gueltige Regeln.")
        return rules
    except Exception as e:
        if "st" in globals():
            st.error(f"Fehler beim Laden von rules.json: {e}")
        return []


# ========= Materialisierungs-Regeln =========
@dataclass
class _Cond:
    col: str
    op: str
    value: Any
    case_insensitive: bool = True  # durch _fold_text abgedeckt


def _apply_single_condition(df: pd.DataFrame, cond: _Cond) -> pd.Series:
    """Wendet eine Einzelbedingung an. Unterstuetzt String-, Regex- und numerische Operatoren."""
    col = cond.col
    if not col or col not in df.columns:
        return pd.Series(False, index=df.index)

    s_fold = _norm_series(df[col])
    op = str(cond.op or "equals").lower()
    val = cond.value
    val_folded = _fold_text(val) if isinstance(val, str) else val

    # Numerische Vergleiche
    if op in ("lt", "le", "gt", "ge"):
        s_num = pd.to_numeric(df[col], errors="coerce")
        try:
            v = float(val)
        except Exception:
            return pd.Series(False, index=df.index)
        if op == "lt":
            return s_num.lt(v)
        if op == "le":
            return s_num.le(v)
        if op == "gt":
            return s_num.gt(v)
        if op == "ge":
            return s_num.ge(v)

    # Häkchen/Ja/True/1
    if op in ("checked",):
        s_raw = df[col]
        s_fold2 = _norm_series(s_raw)
        s_num = pd.to_numeric(s_raw, errors="coerce")
        return (
            s_fold2.isin({"x", "ja", "true", "wahr", "y", "1"})
            | s_num.eq(1)
            | (s_raw == True)  # noqa: E712
        )

    # Gleichheit
    if op in ("eq", "equals"):
        if isinstance(val, str) and val_folded == "":
            return s_fold.eq("") | df[col].isna()
        if isinstance(val, str) and val_folded in {"x", "ja", "true", "wahr", "y", "1"}:
            s_num = pd.to_numeric(df[col], errors="coerce")
            return (
                s_fold.isin({"x", "ja", "true", "wahr", "y", "1"})
                | s_num.eq(1)
                | (df[col] == True)  # noqa: E712
            )
        return s_fold.eq(val_folded)

    # Ungleichheit
    if op in ("neq", "not_equals"):
        if isinstance(val, str) and val_folded == "":
            return ~(s_fold.eq("") | df[col].isna())
        return ~s_fold.eq(val_folded)

    # Teilstring
    if op in ("contains", "icontains"):
        if val is None:
            return pd.Series(False, index=df.index)
        return s_fold.str.contains(re.escape(val_folded), na=False)

    # In-Liste
    if op in ("in",):
        vals = [(_fold_text(v) if isinstance(v, str) else v) for v in (val or [])]
        return s_fold.isin(vals)

    # Regex
    if op in ("regex", "matches"):
        try:
            return s_fold.str.contains(val, flags=re.IGNORECASE, regex=True, na=False)
        except Exception:
            return pd.Series(False, index=df.index)

    return pd.Series(False, index=df.index)


def _build_condition_mask(df: pd.DataFrame, conds: List[Dict[str, Any]]) -> pd.Series:
    """UND-Verknuepfung ueber mehrere Bedingungen."""
    if df is None or df.empty or not conds:
        return pd.Series(True, index=df.index)
    mask = pd.Series(True, index=df.index)
    for c in conds:
        if not isinstance(c, dict):
            continue
        cond = _Cond(
            col=c.get("col", ""),
            op=c.get("op", "equals"),
            value=c.get("value", None),
            case_insensitive=True
        )
        m = _apply_single_condition(df, cond)
        if not isinstance(m, pd.Series) or len(m) != len(df):
            m = pd.Series(False, index=df.index)
        mask &= m
        if not mask.any():
            return mask
    return mask


def apply_materialization_rules(
    df: pd.DataFrame,
    rules: List[Dict[str, Any]],
    first_match_wins: bool = False
) -> pd.DataFrame:
    """
    Wendet Regeln auf einen DataFrame an.
    - Unterstuetzt action=drop / drop=true und then.set{...}
    - first_match_wins: Zeilen werden nach erstem Treffer fuer weitere Regeln gesperrt
    """
    if df is None or df.empty or not rules:
        return df

    out = df.copy()
    active_mask = pd.Series(True, index=out.index)

    for rule in rules:
        conds = rule.get("when", []) or []
        then_raw = rule.get("then", {}) or {}
        set_map: Dict[str, Any] = (then_raw.get("set") or {})

        # Drop-Action
        drop_action = False
        act = then_raw.get("action")
        if isinstance(act, str) and act.strip().lower() == "drop":
            drop_action = True
        if bool(then_raw.get("drop", False)):
            drop_action = True

        if not conds and not drop_action and not set_map:
            continue

        # Maske bilden
        try:
            mask = _build_condition_mask(out, conds) if conds else pd.Series(True, index=out.index)
            if first_match_wins:
                mask = mask & active_mask
        except Exception:
            continue

        idx = out.index[mask]
        if len(idx) == 0:
            continue

        # Drop zuerst
        if drop_action:
            out = out.drop(index=idx)
            if first_match_wins:
                active_mask = active_mask.reindex(out.index).fillna(False)
            continue

        # Set-Operationen
        if set_map:
            cols_present = [k for k in set_map.keys() if k in out.columns]
            for k in cols_present:
                v = set_map[k]
                if v == "__KEEP__":
                    continue
                out.loc[idx, k] = v

        if first_match_wins:
            active_mask.loc[idx] = False

    return out


# ========= Debug-Helfer =========
def _evaluate_rules_debug(df: pd.DataFrame, rules: List[Dict[str, Any]], sample_rows: int = 10):
    """
    Liefert:
      - summary_df: je Regel Anzahl Treffer, Drop/Set-Typ, betroffene Spalten, Fehler
      - samples: dict von {sheetname: DataFrame} mit Beispielzeilen je Regel
    """
    rows = []
    samples = {}
    for i, rule in enumerate(rules):
        then_raw = rule.get("then", {}) or {}
        set_map = then_raw.get("set") or {}
        drop_action = (str(then_raw.get("action", "")).strip().lower() == "drop") or bool(then_raw.get("drop", False))
        action = "drop" if drop_action else ("set" if set_map else "noop")
        set_cols = ",".join(k for k in set_map.keys()) if set_map else ""
        err = None
        try:
            mask = _build_condition_mask(df, rule.get("when", []) or [])
        except Exception as e:
            mask = pd.Series(False, index=df.index)
            err = f"{type(e).__name__}: {e}"
        count = int(mask.sum())
        rows.append({"rule_idx": i, "action": action, "set_cols": set_cols, "match_count": count, "error": err})
        if count > 0:
            samples[f"rule_{i:02d}_matches"] = df.loc[mask].head(sample_rows).copy()
    return pd.DataFrame(rows, columns=["rule_idx","action","set_cols","match_count","error"]), samples


def _export_rules_debug_xlsx(df_before: pd.DataFrame,
                             df_after: pd.DataFrame,
                             summary_df: pd.DataFrame,
                             samples: Dict[str, pd.DataFrame]) -> bytes:
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as w:
        df_before.to_excel(w, index=False, sheet_name="before_rules")
        summary_df.to_excel(w, index=False, sheet_name="rules_summary")
        for name, sdf in samples.items():
            w_sheet = name[:31] if len(name) > 31 else name
            sdf.to_excel(w, index=False, sheet_name=w_sheet)
        df_after.to_excel(w, index=False, sheet_name="after_rules")
    bio.seek(0)
    return bio.read()


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
    - GUID-Logik bereinigt: nur 'GUID' (eigene ID; bei Subs aus 'GUID Sub') und 'GUID Gruppe' (immer Mutter-GUID).
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
        # Mutter: GUID Gruppe = eigene GUID
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

    # === GUID-Logik ab hier konsolidieren ===
    # 1) Bei Subs: GUID aus 'GUID Sub' uebernehmen (falls vorhanden)
    if "GUID Sub" in promoted.columns:
        guid_sub_ok = promoted["GUID Sub"].notna() & promoted["GUID Sub"].astype(str).str.strip().ne("")
        promoted["GUID"] = promoted["GUID"].where(~guid_sub_ok, promoted["GUID Sub"])

    # 2) Flags setzen
    promoted["Mehrschichtiges Element"] = False
    promoted["Promoted"] = True

    # 3) Mutter-GUID an Subs als 'GUID Gruppe' vererben
    if "GUID" in cols:
        mother_guid_map = df.loc[is_mother, ["GUID"]].assign(grp=grp_id[is_mother]).set_index("grp")["GUID"]
        promoted["GUID Gruppe"] = grp_id[keep_sub_mask].map(mother_guid_map).values
    else:
        mother_guid_map = pd.Series(dtype=object)

    # 4) Muetter/Subs droppen und Promoted anhaengen
    to_drop_idx = df.index[drop_mother_mask | drop_sub_mask]
    if len(to_drop_idx) > 0:
        df = df.drop(index=to_drop_idx)
    if not promoted.empty:
        df = pd.concat([df, promoted], ignore_index=True)

    # 5) GUID Gruppe fuer uebrige Subs (falls noch nicht gesetzt) vererben
    if not mother_guid_map.empty:
        # Alle Zeilen mit Gruppen-ID, die keine Mutter sind, erhalten Mutter-GUID
        mask_need_group_guid = (~is_mother.reindex(df.index, fill_value=False)) & grp_id.reindex(df.index).notna()
        df.loc[mask_need_group_guid, "GUID Gruppe"] = grp_id.reindex(df.index)[mask_need_group_guid].map(mother_guid_map)

    # 6) GUID-Konsolidierung: Nur 'GUID' und 'GUID Gruppe' behalten
    if "GUID Sub" in df.columns:
        has_guid_sub = df["GUID Sub"].notna() & df["GUID Sub"].astype(str).str.strip().ne("")
        df["GUID"] = df["GUID"].where(~has_guid_sub, df["GUID Sub"])
        df.drop(columns=["GUID Sub"], inplace=True, errors="ignore")

    # ---------- (4) Sub-Spalten entfernen ----------
    subs_to_drop = [c for c in df.columns if c.endswith(" Sub") and c not in ("GUID Sub",)]
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

    # Sicherheit: Nur 'GUID' + 'GUID Gruppe' als GUID-Spalten behalten
    guid_like = [c for c in df.columns if c.lower().startswith("guid")]
    extra = [c for c in guid_like if c not in ("GUID", "GUID Gruppe")]
    if extra:
        df.drop(columns=extra, inplace=True, errors="ignore")

    return df


# ========= Streamlit App (3 Schritte) =========
def app(supplement_name: str, delete_enabled: bool, custom_chars: str):
    st.set_page_config(page_title="Vererbung & Regeln", layout="wide")
    st.header("Vererbung & Mengenuebernahme")

    # Session-State
    for key in ("df_raw", "df_step1", "df_step2", "df_final"):
        if key not in st.session_state:
            st.session_state[key] = None

    # Datei laden
    uploaded_file = st.file_uploader("Excel-Datei laden", type=["xlsx", "xls"], key="upl_file")
    if not uploaded_file:
        st.stop()

    try:
        df_raw = pd.read_excel(uploaded_file, engine="openpyxl")
        st.session_state["df_raw"] = df_raw.copy()
    except Exception as e:
        st.error(f"Fehler beim Einlesen: {e}")
        st.stop()

    st.markdown("""
**Ablauf**
1) **Einlesen & Bereinigung** (Vererbung/Promotion, GUID-Konsolidierung, Standardisierung).  
2) **Subs droppen (eBKP-H exakt)** + **Material-Filter** *(ohne Regeln)*.  
3) **Regeln (rules.json) anwenden** + Debug-Feedback + Export.
    """)

    # ===== Schritt 1 =====
    st.subheader("1) Einlesen & Bereinigung")
    with st.form(key="form_step1"):
        btn_step1 = st.form_submit_button("Schritt 1 starten (Bereinigung)")
    if btn_step1:
        with st.spinner("Schritt 1 laeuft ..."):
            df_step1 = _process_df(df_raw.copy(), drop_sub_values=[])   # kein Sub-Drop hier
            df_step1 = convert_quantity_columns(df_step1)
        st.session_state["df_step1"] = df_step1.copy()
        st.session_state["df_step2"] = None
        st.session_state["df_final"] = None
        st.success("Schritt 1 abgeschlossen.")

    if st.session_state["df_step1"] is None:
        st.info("Bitte Schritt 1 ausfuehren.")
        st.stop()

    df_step1 = st.session_state["df_step1"]
    st.markdown("**Schritt 1 – Bereinigt (Top 15)**")
    st.dataframe(df_step1.head(15), width="stretch")
    out_step1 = io.BytesIO()
    with pd.ExcelWriter(out_step1, engine="openpyxl") as writer:
        df_step1.to_excel(writer, index=False, sheet_name="Bereinigt_Step1")
    out_step1.seek(0)
    st.download_button(
        "Download: Schritt 1",
        data=out_step1,
        file_name="export_bereinigt_step1.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_step1"
    )

    # ===== Schritt 2 =====
    st.markdown("---")
    st.subheader("2) Subs droppen (eBKP-H) + Material-Filter (ohne Regeln)")

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
        btn_step2 = st.form_submit_button("Schritt 2 starten (ohne Regeln)")
    if btn_step2:
        with st.spinner("Schritt 2 laeuft ..."):
            df_after_subdrop = _process_df(df_raw.copy(), drop_sub_values=sel_drop_values)
            df_after_subdrop = convert_quantity_columns(df_after_subdrop)

            selected_materials = [materials_inv[lbl] for lbl in sel_material_labels]
            if selected_materials:
                mat_eff = df_after_subdrop.get("Material", pd.Series("", index=df_after_subdrop.index)).astype(str).str.strip()
                df_step2 = df_after_subdrop.loc[~mat_eff.isin(set(selected_materials))].reset_index(drop=True)
            else:
                df_step2 = df_after_subdrop

        st.session_state["df_step2"] = df_step2.copy()
        st.session_state["df_final"] = None
        st.success("Schritt 2 abgeschlossen (ohne Regeln).")

    if st.session_state["df_step2"] is None:
        st.info("Bitte Schritt 2 ausfuehren.")
        st.stop()

    df_step2 = st.session_state["df_step2"]
    st.markdown("**Schritt 2 – nach Sub-Drop & Material-Filter (Top 15)**")
    st.dataframe(df_step2.head(15), width="stretch")
    out_step2 = io.BytesIO()
    with pd.ExcelWriter(out_step2, engine="openpyxl") as writer:
        df_step2.to_excel(writer, index=False, sheet_name="Step2_no_rules")
    out_step2.seek(0)
    st.download_button(
        "Download: Schritt 2 (ohne Regeln)",
        data=out_step2,
        file_name="export_step2_no_rules.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="dl_step2"
    )

    # ===== Schritt 3: Regeln =====
    st.markdown("---")
    st.subheader("3) Regeln anwenden (rules.json)")

    with st.form(key="form_step3"):
        first_match_wins = st.checkbox("Materialisierungs-Regeln: erste Regel gewinnt (Stop nach Match)", value=False)
        debug_rules = st.checkbox("Regel-Debug aktivieren (Zusammenfassung & Export)", value=True)
        btn_step3 = st.form_submit_button("Schritt 3 starten (Regeln anwenden)")

    if btn_step3:
        df_input = df_step2.copy()
        rules_all: List[dict] = load_rules_from_repo("rules.json")

        total_rules = len(rules_all)
        if debug_rules:
            dbg_summary, dbg_samples = _evaluate_rules_debug(df_input, rules_all, sample_rows=10)
            st.caption(f"Regel-Debug: {total_rules} Regeln geladen. Gesamt-Treffer: {int(dbg_summary['match_count'].sum())}.")
            st.dataframe(dbg_summary, width="stretch")

        before = len(df_input)
        df_final = apply_materialization_rules(
            df_input.copy(),
            rules_all,
            first_match_wins=bool(first_match_wins)
        ) if rules_all else df_input.copy()
        after = len(df_final)

        # Debug-Export
        if debug_rules:
            xbytes = _export_rules_debug_xlsx(df_input, df_final, dbg_summary, dbg_samples)
            st.download_button(
                "Download Regel-Debug (Excel)",
                data=xbytes,
                file_name="regel_debug.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_rules_debug_step3"
            )

        st.session_state["df_final"] = df_final.copy()
        st.success(f"Schritt 3 abgeschlossen. Regeln angewendet: Differenz {before - after:+d} (vorher {before}, nachher {after}).")

    # Final-Ansicht & Download
    if st.session_state["df_final"] is not None:
        df_final = st.session_state["df_final"]
        st.markdown("**Finalisiert (Top 15)**")
        st.dataframe(df_final.head(15), width="stretch")
        out_final = io.BytesIO()
        with pd.ExcelWriter(out_final, engine="openpyxl") as writer:
            df_final.to_excel(writer, index=False, sheet_name="Final_Step3")
        out_final.seek(0)
        st.download_button(
            "Download: Final (nach Regeln)",
            data=out_final,
            file_name="export_final_step3.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_final_step3"
        )
