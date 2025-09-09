import streamlit as st
import pandas as pd
import io
import re
from typing import Tuple, Optional
from excel_utils import clean_columns_values, rename_columns_to_standard, convert_quantity_columns


# --------- Hilfen ---------
def _has_value(x) -> bool:
    return pd.notna(x) and str(x).strip() != ""


def _is_undef(val: any) -> bool:
    """eBKP-H als 'nicht definiert' erkennen."""
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
    """Sicherer Check, ob Wert NA ist (verhindert bool-Kontext von pd.NA)."""
    try:
        return pd.isna(x)
    except Exception:
        return False


# --------- Kernbereinigung ---------
def _process_df(
    df: pd.DataFrame,
    drop_sub_values: Optional[list[str]] = None,          # exakte eBKP-H-Werte fuer Sub-Zeilen droppen
    drop_by_type: Optional[Tuple[str, list[str]]] = None, # (spaltenname, werte) nach Verarbeitung loeschen
    drop_scope: str = "all",                               # "all" | "subs_only" | "promoted_subs_only"
) -> pd.DataFrame:
    """
    Grundregel:
    - Sub-Zeile zur Hauptzeile befoerdern: vorhandene '... Sub' Werte ersetzen Basiswerte; sonst Mutterwerte.
    - GUID: bevorzugt 'GUID Sub', sonst 'GUID'.
    - Master-Kontextspalten werden an Subs vererbt.
    """
    drop_sub_values = {str(v).strip().lower() for v in (drop_sub_values or []) if str(v).strip()}

    def _matches_drop_values(val: any) -> bool:
        if _is_na(val):
            return False
        return str(val).strip().lower() in drop_sub_values if drop_sub_values else False

    # Master-Kontextspalten
    master_cols = ["Teilprojekt", "Gebäude", "Baufeld", "Geschoss", "Umbaustatus", "Unter Terrain", "Typ"]
    master_cols = [c for c in master_cols if c in df.columns]

    # Paare Basis / Sub (GUID ausschliessen)
    sub_pairs = sorted({
        base for c in df.columns
        if c.endswith(" Sub") and (base := c[:-4]) in df.columns and base != "GUID"
    })

    # Mehrschichtiges Element flaggen: wenn alle master_cols leer → Sub-Zeile
    df["Mehrschichtiges Element"] = df.apply(
        lambda row: all(pd.isna(row.get(c)) for c in master_cols), axis=1
    )
    # Promoted-Flag fuer spaetere Filterung
    df["Promoted"] = False

    # 1) Mutter-Kontext an Subs vererben (inkl. eBKP-H, wenn Sub fehlt/unklassifiziert)
    i = 0
    while i < len(df):
        if all(pd.notna(df.at[i, c]) for c in master_cols):  # Mutter
            j = i + 1
            while j < len(df) and df.at[j, "Mehrschichtiges Element"]:
                for c in master_cols:
                    df.at[j, c] = df.at[i, c]
                if "eBKP-H" in df.columns:
                    mother_ebkp = df.at[i, "eBKP-H"]
                    sub_ebkp_sub = df.at[j, "eBKP-H Sub"] if "eBKP-H Sub" in df.columns else pd.NA
                    if _has_value(mother_ebkp) and _is_undef(sub_ebkp_sub):
                        df.at[j, "eBKP-H"] = mother_ebkp
                j += 1
            i = j
        else:
            i += 1

    # 2) Generisch Sub bevorzugen, sonst Mutter
    i = 0
    while i < len(df):
        if all(pd.notna(df.at[i, c]) for c in master_cols):  # Mutter
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

    # 3) Subs zu Hauptzeilen promoten; Mutter ggf. droppen
    new_rows = []
    drop_idx = []
    i = 0
    while i < len(df):
        if all(pd.notna(df.at[i, c]) for c in master_cols):  # Mutter
            j = i + 1
            sub_idxs = []
            while j < len(df) and df.at[j, "Mehrschichtiges Element"]:
                sub_idxs.append(j)
                j += 1

            # Sub-Zeilen anhand exakter eBKP-H-Liste droppen
            if drop_sub_values:
                for idx in list(sub_idxs):
                    ebkp_sub = df.at[idx, "eBKP-H Sub"] if "eBKP-H Sub" in df.columns else None
                    ebkp     = df.at[idx, "eBKP-H"] if "eBKP-H" in df.columns else None
                    if _matches_drop_values(ebkp_sub) or _matches_drop_values(ebkp):
                        drop_idx.append(idx)
                        sub_idxs.remove(idx)

            if sub_idxs:
                # Mutter droppen und Subs promoten
                drop_idx.append(i)
                for idx in sub_idxs:
                    new = df.loc[idx].copy()

                    # GUID der Sub bevorzugen; Fallback auf GUID
                    if "GUID Sub" in df.columns and _has_value(df.at[idx, "GUID Sub"]):
                        new["GUID"] = df.at[idx, "GUID Sub"]
                    elif "GUID" in df.columns:
                        new["GUID"] = df.at[idx, "GUID"]

                    # Fuer alle Basis/Sub-Paare: Sub bevorzugen, sonst Mutter
                    for base in sub_pairs:
                        sub_col = f"{base} Sub"
                        if sub_col in df.columns and _has_value(df.at[idx, sub_col]):
                            new[base] = df.at[idx, sub_col]
                        else:
                            new[base] = df.at[i, base]

                    new["Mehrschichtiges Element"] = False
                    new["Promoted"] = True
                    for c in master_cols:
                        new[c] = df.at[i, c]

                    new_rows.append(new)
            else:
                df.at[i, "Mehrschichtiges Element"] = False

            i = j
        else:
            i += 1

    if drop_idx:
        df.drop(index=drop_idx, inplace=True)
        df.reset_index(drop=True, inplace=True)
    if new_rows:
        df = pd.concat([df, pd.DataFrame(new_rows)], ignore_index=True)

    # 4) ... Sub-Spalten entfernen
    df.drop(columns=[c for c in df.columns if c.endswith(" Sub")], inplace=True, errors="ignore")

    # 5) Restbereinigung
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

    # 5a) Optional: Zeilen nach Typ (Spalte/Werte) loeschen (Scope beachtend)
    if drop_by_type:
        type_col, type_values = drop_by_type
        if type_col in df.columns and type_values:
            _vals = {str(v).strip().lower() for v in type_values if str(v).strip()}
            def _match(v):
                if _is_na(v):
                    return False
                return str(v).strip().lower() in _vals
            mask = df[type_col].apply(_match)
            if drop_scope == "subs_only":
                # urspruengliche Subs existieren nach dem Promoten nicht mehr als solche → hier keine Wirkung
                # Hinweis: Falls gewuenscht, vor dem Promoten markieren und separat loeschen.
                mask = mask & (df["Mehrschichtiges Element"] == True)
            elif drop_scope == "promoted_subs_only":
                mask = mask & (df["Promoted"] == True)
            df = df.loc[~mask].reset_index(drop=True)

    # 6) Deduplizieren (GUID-identische Voll-Duplikate)
    def _remove_exact_duplicates(d: pd.DataFrame) -> pd.DataFrame:
        if "GUID" not in d.columns:
            return d
        drop = []
        for guid, grp in d.groupby("GUID"):
            if len(grp) > 1 and all(n <= 1 for n in grp.nunique().values):
                drop.extend(grp.index.tolist()[1:])
        return d.drop(index=drop).reset_index(drop=True)

    df = _remove_exact_duplicates(df)

    # 7) Standardisieren & Werte bereinigen
    df = rename_columns_to_standard(df)
    df = clean_columns_values(df, delete_enabled=True, custom_chars="")

    return df


# --------- Finale Auswertung ---------
def summarize_column(df: pd.DataFrame, col: str) -> pd.DataFrame:
    """Werte in Spalte zaehlen und prozentual ausweisen; leere ausblenden."""
    if col not in df.columns:
        return pd.DataFrame()
    s = df[col].astype(str).str.strip()
    s = s[s != ""]
    summary = (
        s.value_counts(dropna=True)
         .rename_axis(col)
         .reset_index(name="Anzahl")
    )
    if len(s) > 0:
        summary["Anteil %"] = (summary["Anzahl"] / len(s) * 100).round(2)
    else:
        summary["Anteil %"] = 0.0
    return summary


# --------- Streamlit App ---------
def app(supplement_name: str, delete_enabled: bool, custom_chars: str):
    st.header("Vererbung & Mengenuebernahme")

    st.markdown("""
    **Logik**  
    1) eBKP-H der Mutter → an Subs vererben, wenn `eBKP-H Sub` nicht definiert ist.  
    2) Generisch: Fuer jedes Basis/`... Sub`-Paar gilt **Sub bevorzugen, sonst Mutter**.  
    3) Subs als Hauptzeilen promoten; vorhandene `... Sub`-Werte uebernehmen; Mutter droppen.  
    4) 'Nicht klassifiziert', 'Keine Zuordnung', 'Nicht verfügbar' gelten als nicht definiert.

    **Steuerung**  
    - Auswahl: Sub-Zeilen droppen anhand eBKP-H.  
    - Optionaler Typ-Filter (Spalte + Werte) nach Verarbeitung mit Scope.  
    - Finale Auswertung: Spalte auswaehlen, Werte zusammenfassen → Button *Finalisieren*.  
    """)

    uploaded_file = st.file_uploader("Excel-Datei laden", type=["xlsx", "xls"], key="vererbung_file_uploader")
    if not uploaded_file:
        return

    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Fehler beim Einlesen: {e}")
        return

    st.subheader("Originale Daten (15 Zeilen)")
    st.dataframe(df.head(15), use_container_width=True)

    # Kandidaten fuer eBKP-Dropdown
    ebkp_candidates = []
    if "eBKP-H" in df.columns:
        ebkp_candidates.extend(df["eBKP-H"].dropna().astype(str).str.strip().tolist())
    if "eBKP-H Sub" in df.columns:
        ebkp_candidates.extend(df["eBKP-H Sub"].dropna().astype(str).str.strip().tolist())
    ebkp_options = sorted({v for v in ebkp_candidates if v})

    # Kandidaten fuer Typ-Filter-Spalte
    type_col_options = ["-- keine --"] + [c for c in df.columns]

    # ---------------- Formular: Verarbeitung (kein Auto-Run) ----------------
    with st.form(key="vererbung_form"):
        sel_drop_values = st.multiselect(
            "Sub-Zeilen droppen, wenn eBKP-H exakt gleich einem der folgenden Werte ist",
            options=ebkp_options,
            default=[],
            help="Mehrfachauswahl moeglich. Exakter Textvergleich."
        )

        sel_type_col = st.selectbox(
            "Spalte fuer Typ-Filter (optional, z. B. Material)",
            options=type_col_options,
            index=0
        )

        type_value_options = []
        if sel_type_col and sel_type_col != "-- keine --" and sel_type_col in df.columns:
            type_value_options = sorted({
                v for v in df[sel_type_col].dropna().astype(str).str.strip().tolist() if v
            })
        sel_type_values = st.multiselect(
            "Zeilen loeschen, wenn Typ-Wert gleich ist (nach Verarbeitung angewandt)",
            options=type_value_options,
            default=[],
            help="Exakter Textvergleich; mehrere moeglich."
        )

        scope_label = st.radio(
            "Geltungsbereich des Typ-Filters",
            options=["alle Zeilen", "nur urspruengliche Subs", "nur promotete Subs"],
            horizontal=True,
            index=0
        )
        scope_map = {
            "alle Zeilen": "all",
            "nur urspruengliche Subs": "subs_only",
            "nur promotete Subs": "promoted_subs_only",
        }

        # Finale Auswertung: Spalte fuer Zusammenfassung (wird NACH Verarbeitung genutzt)
        final_summary_col = st.selectbox(
            "Spalte fuer finale Zusammenfassung (Dropdown, optional)",
            options=["-- keine --"] + list(df.columns),
            index=0
        )

        run = st.form_submit_button("Verarbeitung starte")

    # ---------------- Ausfuehrung bei Button-Klick ----------------
    if run:
        with st.spinner("Verarbeitung laeuft ..."):
            drop_by_type_arg = None
            if sel_type_col and sel_type_col != "-- keine --" and sel_type_values:
                drop_by_type_arg = (sel_type_col, sel_type_values)

            df_clean = _process_df(
                df.copy(),
                drop_sub_values=sel_drop_values,
                drop_by_type=drop_by_type_arg,
                drop_scope=scope_map.get(scope_label, "all"),
            )
            df_clean = convert_quantity_columns(df_clean)

        st.subheader("Bereinigte Daten (15 Zeilen)")
        st.dataframe(df_clean.head(15), use_container_width=True)

        # Export bereinigte Daten
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_clean.to_excel(writer, index=False, sheet_name="Bereinigt")
        output.seek(0)
        file_name = f"{(supplement_name or '').strip() or 'default'}_vererbung_mengen.xlsx"

        st.download_button(
            "Bereinigte Datei herunterladen",
            data=output,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # ---------------- Finale Zusammenfassung (separater Button) ----------------
        st.markdown("---")
        st.subheader("Finale Auswertung")
        with st.form(key="finalize_form"):
            # Wahl der Spalte fuer die finale Zusammenfassung fest halten oder neu auswaehlen
            final_col = st.selectbox(
                "Spalte fuer Zusammenfassung",
                options=["-- keine --"] + list(df_clean.columns),
                index=(["-- keine --"] + list(df_clean.columns)).index(final_summary_col) if final_summary_col in df_clean.columns else 0,
                help="Werte aus dieser Spalte werden gezaehlt und prozentual ausgewiesen."
            )
            do_finalize = st.form_submit_button("Finalisieren")

        if do_finalize:
            if final_col and final_col != "-- keine --":
                summary_df = summarize_column(df_clean, final_col)
                if not summary_df.empty:
                    st.markdown(f"**Zusammenfassung fuer: {final_col}**")
                    st.dataframe(summary_df, use_container_width=True)

                    # Export Summary + Clean als Excel
                    out2 = io.BytesIO()
                    with pd.ExcelWriter(out2, engine="openpyxl") as writer:
                        df_clean.to_excel(writer, index=False, sheet_name="Bereinigt")
                        summary_df.to_excel(writer, index=False, sheet_name="Auswertung")
                    out2.seek(0)
                    st.download_button(
                        "Bereinigte Datei + Auswertung herunterladen",
                        data=out2,
                        file_name=f"{(supplement_name or '').strip() or 'default'}_final.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.info("Keine Daten fuer die gewaehlte Spalte vorhanden.")
            else:
                st.warning("Bitte eine gueltige Spalte fuer die Zusammenfassung auswaehlen.")
