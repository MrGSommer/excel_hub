import streamlit as st
import pandas as pd
import io
from typing import Optional, List, Dict
from excel_utils import clean_columns_values, rename_columns_to_standard, convert_quantity_columns


# ========= Hilfen =========
def _has_value(x) -> bool:
    return pd.notna(x) and str(x).strip() != ""


def _is_undef(val: any) -> bool:
    """eBKP-H-Wert gilt als 'nicht definiert'."""
    if not _has_value(val):
        return True
    txt = str(val).strip().lower()
    return (
        txt == ""
        or "icht klassifiziert" in txt
        or "eine zuordnung" in txt
        or "icht verfügbar" in txt
    )


def _is_na(x) -> bool:
    """Sicherer NA-Check (verhindert bool-Kontext von pd.NA)."""
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


# ========= Kernverarbeitung =========
def _process_df(
    df: pd.DataFrame,
    drop_sub_values: Optional[List[str]] = None,  # eBKP-H exakte Werte: nur Sub-Zeilen droppen
) -> pd.DataFrame:
    """
    Ablauf:
    - Master-Kontext an Subs vererben; eBKP-H der Mutter vererben, wenn 'eBKP-H Sub' undefiniert.
    - Generisch fuer Basis/'... Sub': 'Sub bevorzugen, sonst Mutter'.
    - Subs promoten; wenn mind. 1 Sub bleibt → Mutter droppen.
    - GUID der Zeile bleibt Sub-GUID (falls vorhanden); 'GUID Gruppe' = GUID der Mutter.
    - Nie 'GUID' oder 'GUID Sub' als Spalte loeschen; keine Deduplizierung ueber GUID.
    """
    drop_set = {str(v).strip().lower() for v in (drop_sub_values or []) if str(v).strip()}

    def _matches_drop_values(val: any) -> bool:
        if _is_na(val):
            return False
        return str(val).strip().lower() in drop_set if drop_set else False

    # Master-Kontextspalten (falls vorhanden)
    master_cols = ["Teilprojekt", "Gebäude", "Baufeld", "Geschoss", "Umbaustatus", "Unter Terrain", "Typ"]
    master_cols = [c for c in master_cols if c in df.columns]

    # Paare Basis / Sub (GUID bewusst ausschliessen)
    sub_pairs = sorted({
        base for c in df.columns
        if c.endswith(" Sub") and (base := c[:-4]) in df.columns and base != "GUID"
    })

    # Zeilentypen markieren
    df["Mehrschichtiges Element"] = df.apply(
        lambda row: all(pd.isna(row.get(c)) for c in master_cols), axis=1
    )
    if "GUID Gruppe" not in df.columns:
        df["GUID Gruppe"] = pd.NA  # immer GUID der Mutter
    df["Promoted"] = False

    # ===== 1) Mutter-Kontext + eBKP-H vererben =====
    i = 0
    while i < len(df):
        if all(pd.notna(df.at[i, c]) for c in master_cols):  # Mutter
            mother_guid = df.at[i, "GUID"] if "GUID" in df.columns else pd.NA
            df.at[i, "GUID Gruppe"] = mother_guid  # Mutter ist eigene Gruppe

            j = i + 1
            while j < len(df) and df.at[j, "Mehrschichtiges Element"]:
                # a) Master-Kontext vererben
                for c in master_cols:
                    df.at[j, c] = df.at[i, c]
                # b) eBKP-H vererben, falls 'eBKP-H Sub' nicht definiert
                if "eBKP-H" in df.columns:
                    mother_ebkp = df.at[i, "eBKP-H"]
                    sub_ebkp_sub = df.at[j, "eBKP-H Sub"] if "eBKP-H Sub" in df.columns else pd.NA
                    if _has_value(mother_ebkp) and _is_undef(sub_ebkp_sub):
                        df.at[j, "eBKP-H"] = mother_ebkp
                j += 1
            i = j
        else:
            i += 1

    # ===== 2) Sub bevorzugen, sonst Mutter =====
    i = 0
    while i < len(df):
        if all(pd.notna(df.at[i, c]) for c in master_cols):
            j = i + 1
            sub_idxs = []
            while j < len(df) and df.at[j, "Mehrschichtiges Element"]:
                sub_idxs.append(j)
                j += 1
            for idx in sub_idxs:
                for base in sub_pairs:
                    sub_col = f"{base} Sub"
                    if sub_col in df.columns and _has_value(df.at[idx, sub_col]):
                        df.at[idx, base] = df.at[idx, sub_col]
                    else:
                        df.at[idx, base] = df.at[i, base]
            i = j
        else:
            i += 1

    # ===== 3) Subs promoten; Mutter ggf. droppen (nur Subs gem. eBKP-H droppen) =====
    new_rows = []
    drop_idx = []
    i = 0
    while i < len(df):
        if all(pd.notna(df.at[i, c]) for c in master_cols):  # Mutter
            mother_guid = df.at[i, "GUID"] if "GUID" in df.columns else pd.NA
            j = i + 1
            sub_idxs = []
            while j < len(df) and df.at[j, "Mehrschichtiges Element"]:
                sub_idxs.append(j)
                j += 1

            # a) Nur Sub-Zeilen gem. eBKP-H-Liste droppen (exakt)
            if drop_set:
                for idx in list(sub_idxs):
                    ebkp_sub = df.at[idx, "eBKP-H Sub"] if "eBKP-H Sub" in df.columns else None
                    ebkp     = df.at[idx, "eBKP-H"] if "eBKP-H" in df.columns else None
                    if _matches_drop_values(ebkp_sub) or _matches_drop_values(ebkp):
                        drop_idx.append(idx)   # Sub wird physisch entfernt (gewollt)
                        sub_idxs.remove(idx)

            # b) Wenn mind. 1 Sub bleibt → Mutter droppen und Subs promoten
            if sub_idxs:
                drop_idx.append(i)  # Mutter droppen ist erlaubt
                for idx in sub_idxs:
                    new = df.loc[idx].copy()

                    # GUID der neuen (promoteten) Zeile = GUID Sub, sonst Fallback
                    if "GUID Sub" in df.columns and _has_value(df.at[idx, "GUID Sub"]):
                        new["GUID"] = df.at[idx, "GUID Sub"]
                    elif "GUID" in df.columns:
                        new["GUID"] = df.at[idx, "GUID"]

                    # Sub/Basis-Paare finalisieren
                    for base in sub_pairs:
                        sub_col = f"{base} Sub"
                        if sub_col in df.columns and _has_value(df.at[idx, sub_col]):
                            new[base] = df.at[idx, sub_col]
                        else:
                            new[base] = df.at[i, base]

                    # Metadaten
                    new["Mehrschichtiges Element"] = False
                    new["Promoted"] = True
                    new["GUID Gruppe"] = mother_guid  # Gruppen-ID = Mutter-GUID
                    for c in master_cols:
                        new[c] = df.at[i, c]

                    new_rows.append(new)
            else:
                # Mutter bleibt alleine bestehen
                df.at[i, "Mehrschichtiges Element"] = False
                df.at[i, "Promoted"] = False
                df.at[i, "GUID Gruppe"] = mother_guid

            i = j
        else:
            i += 1

    # Physische Drops anwenden (nur Mutter + explizit ignorierte Subs)
    if drop_idx:
        df.drop(index=drop_idx, inplace=True)
        df.reset_index(drop=True, inplace=True)
    # Promotete Subs anhaengen
    if new_rows:
        df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)

    # ===== 4) ... Sub-Spalten entfernen, aber 'GUID Sub' behalten =====
    subs_to_drop = [c for c in df.columns if c.endswith(" Sub") and c != "GUID Sub"]
    df.drop(columns=subs_to_drop, inplace=True, errors="ignore")

    # ===== 5) Restbereinigung =====
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

    # ===== 6) KEIN Deduplizieren ueber GUID (GUIDs nie verlieren) =====

    # ===== 7) Standardisieren & Werte bereinigen =====
    df = rename_columns_to_standard(df)
    df = clean_columns_values(df, delete_enabled=True, custom_chars="")

    return df


# ========= Streamlit App =========
def app(supplement_name: str, delete_enabled: bool, custom_chars: str):
    st.set_page_config(page_title="Vererbung & Mengenuebernahme", layout="wide")
    st.header("Vererbung & Mengenuebernahme")

    st.markdown("""
    **Ablauf**
    1) Nur **Subs** anhand eBKP-H **ignorieren** (droppen) → **Verarbeitung starte**.  
    2) **Filter (ÜBERSCHRIFTEN)**: Spalte waehlen → Dropdown mit zusammengefassten Werten erscheint.  
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
    st.dataframe(df_raw.head(15), use_container_width=True)

    # eBKP-H Auswahl fuer das Ignorieren von Subs
    ebkp_candidates = []
    if "eBKP-H" in df_raw.columns:
        ebkp_candidates.extend(df_raw["eBKP-H"].dropna().astype(str).str.strip().tolist())
    if "eBKP-H Sub" in df_raw.columns:
        ebkp_candidates.extend(df_raw["eBKP-H Sub"].dropna().astype(str).str.strip().tolist())
    ebkp_options = sorted({v for v in ebkp_candidates if v})

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

        # Session: verarbeitete Tabelle (Start fuer Filter)
        st.session_state["df_active"] = df_processed.copy()

    # Wenn noch nicht verarbeitet wurde, stoppen
    if "df_active" not in st.session_state:
        st.info("Bitte zuerst die Verarbeitung starten.")
        st.stop()

    # Aktueller Stand
    df_active = st.session_state["df_active"]

    st.subheader("Bereinigte Daten (15 Zeilen)")
    st.dataframe(df_active.head(15), use_container_width=True)

    # Direkter Download der verarbeiteten Datei (ohne weitere Filter)
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

    # ===== Bereich 2: Filter (ÜBERSCHRIFTEN) =====
    st.markdown("---")
    st.subheader("Filter von Elementen und droppen")

    # Spaltenauswahl fuer den Filter (keine Live-Verarbeitung)
    with st.form(key="form_filter_select_002"):
        col_options = ["-- Spalte waehlen --"] + list(df_active.columns)
        filter_col = st.selectbox("Spalte waehlen", options=col_options, index=0, key="sel_filter_col_002")
        btn_prepare = st.form_submit_button("Werte anzeigen")

    # Werte-Dropdown erst nach Klick auf "Werte anzeigen"
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
                df_active = df_after_filter
                st.success(f"Finalisiert: {mask_drop.sum()} Zeilen entfernt.")

                st.subheader("Finalisierte Vorschau (15 Zeilen)")
                st.dataframe(df_active.head(15), use_container_width=True)

                # Download fuer finalisierte (gefilterte) Datei – Pflichtbutton nach letztem Schritt
                out_final = io.BytesIO()
                with pd.ExcelWriter(out_final, engine="openpyxl") as writer:
                    df_active.to_excel(writer, index=False, sheet_name="Bereinigt_Final")
                out_final.seek(0)
                st.download_button(
                    "Download: Bereinigte Datei (finalisiert)",
                    data=out_final,
                    file_name=f"{(supplement_name or '').strip() or 'default'}_bereinigt_final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_bereinigt_final_003"
                )
            else:
                st.warning("Keine Werte ausgewaehlt. Keine Aenderung vorgenommen.")
