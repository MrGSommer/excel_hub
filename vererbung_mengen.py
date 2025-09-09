import streamlit as st
import pandas as pd
import io
from typing import Optional, List, Dict
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


# ========= Kernverarbeitung (vektorisiert) =========
def _process_df(
    df: pd.DataFrame,
    drop_sub_values: Optional[List[str]] = None,  # eBKP-H exakte Werte: nur Sub-Zeilen droppen
) -> pd.DataFrame:
    """
    Ablauf (performant, unveraenderte Funktionalitaet):
    - Master-Kontext an Subs vererben; eBKP-H der Mutter vererben, wenn 'eBKP-H Sub' undefiniert.
    - Generisch fuer Basis/'... Sub': 'Sub bevorzugen, sonst Mutter'.
    - Subs promoten; wenn mind. 1 Sub bleibt → Mutter droppen.
    - 'GUID' der Zeile bleibt Sub-GUID (Fallback GUID); 'GUID Gruppe' = GUID der Mutter.
    - 'GUID Sub' bleibt als Spalte erhalten; keine Deduplizierung.
    """
    drop_set = {str(v).strip().lower() for v in (drop_sub_values or []) if str(v).strip()}
    cols = pd.Index(df.columns)

    # Master-Kontext & Sub-Paare
    master_cols = [c for c in ["Teilprojekt", "Gebäude", "Baufeld", "Geschoss", "Umbaustatus", "Unter Terrain", "Typ"] if c in cols]
    sub_pairs = sorted({c[:-4] for c in cols if c.endswith(" Sub") and c[:-4] in cols and c[:-4] != "GUID"})

    # Mutter/Sub-Flags
    if master_cols:
        is_mother = df[master_cols].notna().all(axis=1)
        is_sub = df[master_cols].isna().all(axis=1)
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

        # eBKP-H der Mutter, wenn 'eBKP-H Sub' undefiniert
        if "eBKP-H" in cols:
            mother_ebkp_map = df.loc[is_mother, ["eBKP-H"]].assign(grp=grp_id[is_mother]).set_index("grp")["eBKP-H"]
            if "eBKP-H Sub" in cols:
                sub_ebkp = df["eBKP-H Sub"].astype(str).str.strip()
                undef_mask = df["eBKP-H Sub"].isna() | sub_ebkp.eq("") | sub_ebkp.str.contains(
                    r"(?i)(icht klassifiziert|eine zuordnung|icht verfügbar)", na=False
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
    if drop_set:
        ebkp_sub_norm = df["eBKP-H Sub"].astype(str).str.strip().str.lower() if "eBKP-H Sub" in cols else pd.Series("", index=df.index)
        ebkp_norm = df["eBKP-H"].astype(str).str.strip().str.lower() if "eBKP-H" in cols else pd.Series("", index=df.index)
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
            "nicht klassifiziert|keine zuordnung|nicht verfügbar", na=True
        )
        df = df[~mask_invalid]
    for c in ["Einzelteile", "Farbe"]:
        if c in df.columns:
            df.drop(columns=c, inplace=True)

    df.reset_index(drop=True, inplace=True)

    # (6) keine Deduplizierung ueber GUID

    # ---------- (7) Standardisieren & Werte bereinigen ----------
    df = rename_columns_to_standard(df)
    df = clean_columns_values(df, delete_enabled=True, custom_chars="")

    return df


# ========= Streamlit App =========
def app(supplement_name: str, delete_enabled: bool, custom_chars: str):
    st.header("Vererbung & Mengenuebernahme")

    st.markdown("""
    **Ablauf**
    1) Nur **Subs** anhand eBKP-H **ignorieren** (droppen) → **Verarbeitung starte**.  
    2) **Filter (UeBERSCHRIFTEN)**: Spalte waehlen → Dropdown mit zusammengefassten Werten erscheint.  
    3) **Finalisieren**: Auswahl anwenden (Zeilen droppen), Vorschau aktualisieren, **Download** der finalisierten Datei.  
    """)

    # --- Datei laden ---
    uploaded_file = st.file_uploader("Excel-Datei laden", type=["xlsx", "xls"], key="upl_file")
    if not uploaded_file:
        st.stop()

    try:
        df_raw = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Fehler beim Einlesen: {e}")
        st.stop()

    st.subheader("Originale Daten (15 Zeilen)")
    st.dataframe(df_raw.head(15), width="stretch")

    # eBKP-H Auswahl
    ebkp_options = sorted(
        pd.Series(
            pd.concat([
                df_raw.get("eBKP-H", pd.Series(dtype=object)),
                df_raw.get("eBKP-H Sub", pd.Series(dtype=object))
            ])
        ).dropna().astype(str).str.strip().unique()
    )

    # ===== Formular 1: Verarbeitung (nur Subs droppen) =====
    with st.form(key="form_process_001"):
        sel_drop_values = st.multiselect(
            "Subs ignorieren (droppen), wenn eBKP-H exakt gleich ist",
            options=ebkp_options,
            default=[],
            help="Exakter Textvergleich; wirkt nur auf Sub-Zeilen waehrend der Promotion."
        )
        btn_process = st.form_submit_button("Verarbeitung starte")

    if btn_process:
        with st.spinner("Verarbeitung laeuft ..."):
            df_processed = _process_df(df_raw.copy(), drop_sub_values=sel_drop_values)
            df_processed = convert_quantity_columns(df_processed)
        st.session_state["df_active"] = df_processed.copy()
        st.session_state["finalized"] = False  # Reset Finalisierungs-Flag

    # Verarbeitung muss erfolgt sein
    if "df_active" not in st.session_state:
        st.info("Bitte zuerst die Verarbeitung starten.")
        st.stop()

    df_active = st.session_state["df_active"]

    # Nur zeigen, wenn noch nicht finalisiert
    if not st.session_state.get("finalized", False):
        st.subheader("Bereinigte Daten (15 Zeilen)")
        st.dataframe(df_active.head(15), width="stretch")

        out_proc = io.BytesIO()
        with pd.ExcelWriter(out_proc, engine="openpyxl") as writer:
            df_active.to_excel(writer, index=False, sheet_name="Bereinigt")
        out_proc.seek(0)
        st.download_button(
            "Download: Bereinigte Datei (ohne weitere Verarbeitung)",
            data=out_proc,
            file_name=f"{(supplement_name or '').strip() or 'default'}_bereinigt.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_bereinigt_only_001"
        )

    # ===== Bereich 2: Filter (UeBERSCHRIFTEN) =====
    st.markdown("---")
    st.subheader("Filter (UeBERSCHRIFTEN)")

    with st.form(key="form_filter_select_002"):
        col_options = ["-- Spalte waehlen --"] + list(df_active.columns)
        filter_col = st.selectbox("Spalte waehlen", options=col_options, index=0, key="sel_filter_col_002")
        btn_prepare = st.form_submit_button("Werte anzeigen")

    if btn_prepare:
        st.session_state["finalized"] = False  # Reset, falls neuer Filterlauf

    if btn_prepare and filter_col and filter_col != "-- Spalte waehlen --":
        labels_map = _value_labels_with_counts(df_active[filter_col])
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
                mask_drop = df_active[filter_col].astype(str).str.strip().isin(set(selected_values_raw))
                df_after_filter = df_active.loc[~mask_drop].reset_index(drop=True)
                st.session_state["df_active"] = df_after_filter.copy()
                st.session_state["finalized"] = True
                st.success(f"Finalisiert: {mask_drop.sum()} Zeilen entfernt.")
            else:
                st.warning("Keine Werte ausgewaehlt. Keine Aenderung vorgenommen.")

    # --- Finalisierte Ausgabe unten, wenn finalisiert ---
    if st.session_state.get("finalized", False):
        df_final = st.session_state["df_active"]
        st.subheader("Finalisierte Vorschau (15 Zeilen)")
        st.dataframe(df_final.head(15), width="stretch")

        out_final = io.BytesIO()
        with pd.ExcelWriter(out_final, engine="openpyxl") as writer:
            df_final.to_excel(writer, index=False, sheet_name="Bereinigt_Final")
        out_final.seek(0)
        st.download_button(
            "Download: Bereinigte Datei (finalisiert)",
            data=out_final,
            file_name=f"{(supplement_name or '').strip() or 'default'}_bereinigt_final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_bereinigt_final_003"
        )
