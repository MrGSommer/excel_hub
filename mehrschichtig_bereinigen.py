import streamlit as st
import pandas as pd
import io
import re
from excel_utils import clean_columns_values, rename_columns_to_standard, convert_quantity_columns




def clean_dataframe(df, delete_enabled=False, custom_chars="", match_sub_toggle=False, drop_treppe_sub=False):
    master_cols = ["Teilprojekt", "Gebäude", "Baufeld", "Geschoss", "Umbaustatus", "Unter Terrain"]
    master_cols = [col for col in master_cols if col in df.columns]

    # Flag mehrschichtig
    df["Mehrschichtiges Element"] = df.apply(
        lambda row: all(pd.isna(row[col]) for col in master_cols), axis=1
    )

    # Werte aus Mutterzeile füllen
    i = 0
    while i < len(df):
        if all(pd.notna(df.at[i, col]) for col in master_cols):
            j = i + 1
            while j < len(df) and df.at[j, "Mehrschichtiges Element"]:
                for col in master_cols:
                    df.at[j, col] = df.at[i, col]
                j += 1
            i = j
        else:
            i += 1

    # Entfernen nicht klassifizierter Hauptelemente
    drop_idx = []
    i = 0
    while i < len(df):
        if all(pd.notna(df.at[i, col]) for col in master_cols):
            j = i + 1
            sub_idxs = []
            while j < len(df) and df.at[j, "Mehrschichtiges Element"]:
                sub_idxs.append(j)
                j += 1
            if not sub_idxs and df.at[i, "eBKP-H"] == "Nicht klassifiziert":
                drop_idx.append(i)
            i = j
        else:
            i += 1
    if drop_idx:
        df.drop(index=drop_idx, inplace=True)
        df.reset_index(drop=True, inplace=True)

    # Aufschlüsselung mehrschichtiger Elemente
    new_rows = []
    drop_idx = []
    i = 0
    while i < len(df):
        if all(pd.notna(df.at[i, col]) for col in master_cols):
            j = i + 1
            sub_idxs = []
            while j < len(df) and df.at[j, "Mehrschichtiges Element"]:
                sub_idxs.append(j)
                j += 1

            if sub_idxs and match_sub_toggle and "eBKP-H Sub" in df.columns:
                mother_val = df.at[i, "eBKP-H"]
                for idx in sub_idxs:
                    sub_val = df.at[idx, "eBKP-H Sub"]
                    # 1) Treppe-Subs droppen, Mutterzeile bleibt erhalten
                    if "Treppe" in str(mother_val) and "Treppe" in str(sub_val):
                        drop_idx.append(idx)
                    # 2) alle anderen klassifizierten Subs (z.B. Podest) als neue Mutterzeile übernehmen
                    elif pd.notna(sub_val) and sub_val not in ["Nicht klassifiziert", "", "Keine Zuordnung"]:
                        new = df.loc[idx].copy()
                        for col in master_cols:
                            new[col] = df.at[i, col]
                        new["Mehrschichtiges Element"] = False
                        new_rows.append(new)
                # Weiter mit der nächsten Mutterzeile
                i = j
                continue


            # Standardfall ohne Sub-Toggle
            if sub_idxs:
                valid = any(
                    pd.notna(df.at[idx, "eBKP-H Sub"]) and
                    df.at[idx, "eBKP-H Sub"] not in ["Nicht klassifiziert", "", "Keine Zuordnung"]
                    for idx in sub_idxs
                )
                if valid:
                    drop_idx.append(i)
                    for idx in sub_idxs:
                        sub_val = df.at[idx, "eBKP-H Sub"]
                        if pd.notna(sub_val) and sub_val not in ["Nicht klassifiziert", "", "Keine Zuordnung"]:
                            new = df.loc[idx].copy()
                            new["Mehrschichtiges Element"] = True
                            new_rows.append(new)
                else:
                    drop_idx.extend(sub_idxs)
                    df.at[i, "Mehrschichtiges Element"] = False
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

    # Sub-Spalten mappen, löschen
    for idx, row in df.iterrows():
        if row.get("Mehrschichtiges Element", False):
            for col in df.columns:
                if col.endswith(" Sub"):
                    main = col.replace(" Sub", "")
                    if pd.notna(row[col]) and row[col] != "":
                        df.at[idx, main] = row[col]
    df.drop(columns=[c for c in df.columns if c.endswith(" Sub")], inplace=True, errors='ignore')

    # Entferne restliche Spalten
    for col in ["Einzelteile", "Farbe"]:
        if col in df.columns:
            df.drop(columns=col, inplace=True)

    # 'oi' bereinigen
    if "Unter Terrain" in df.columns:
        df.loc[df["Unter Terrain"] == "oi", "Unter Terrain"] = pd.NA

    df = df[~df["eBKP-H"].isin(["Keine Zuordnung", "Nicht klassifiziert"])]
    df.reset_index(drop=True, inplace=True)

    # Duplikate entfernen
    def remove_exact_duplicates(d):
        drop = []
        for guid, grp in d.groupby("GUID"):
            if len(grp) > 1 and all(n <= 1 for n in grp.nunique().values):
                drop.extend(grp.index.tolist()[1:])
        return d.drop(index=drop).reset_index(drop=True)
    df = remove_exact_duplicates(df)

    # Final clean
    df = rename_columns_to_standard(df)
    df = clean_columns_values(df, delete_enabled, custom_chars)
    return df


def app(supplement_name, delete_enabled, custom_chars, convert_quantity_columns):

    state = st.session_state
    supplement = supplement_name or (
        state.get("selected_sheet_values")
        or (state.uploaded_file_values.name.rsplit(".", 1)[0]
            if state.get("uploaded_file_values") else "")
    )
    
    st.header("Mehrschichtig Bereinigen")
    st.markdown("""
    **Einleitung:**  
    Bereinigt Excel-Dateien mehrschichtig mit optionalen Toggles.
    """ )

    uploaded_file = st.file_uploader("Excel-Datei laden", type=["xlsx", "xls"], key="bereinigen_file_uploader")
    if not uploaded_file:
        return

    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as e:
        st.error(f"Fehler beim Einlesen: {e}")
        return

    st.subheader("Originale Daten (15 Zeilen)")
    st.dataframe(df.head(15))

    use_match = st.checkbox("Sub-eBKP-H aufschlüsseln", value=False)
    drop_treppe = st.checkbox(
        "Treppe-Sub identisch zur Mutter droppen", value=False
    )
    if use_match:
        drop_treppe
    with st.spinner("Daten werden bereinigt ..."):
        df_clean = clean_dataframe(
            df,
            delete_enabled=delete_enabled,
            custom_chars=custom_chars,
            match_sub_toggle=use_match,
            drop_treppe_sub=drop_treppe
        )

    st.subheader("Bereinigte Daten (15 Zeilen)")
    st.dataframe(df_clean.head(15))

    output = io.BytesIO()
    df_export = convert_quantity_columns(df_clean.copy())
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_export.to_excel(writer, index=False)
    output.seek(0)
    file_name = f"{supplement_name.strip() or 'default'}_bereinigt.xlsx"
    st.download_button(
        "Bereinigte Datei herunterladen", data=output,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
