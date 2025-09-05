import pandas as pd
import streamlit as st
import re

# Preset für Mengenspalten
COLUMN_PRESET = {
    "Fläche (m2)": ["Fläche", "Flaeche", "Fläche BQ", "Fläche Total", "Fläche Solibri"],
    "Volumen (m3)": ["Volumen", "Volumen BQ", "Volumen Total", "Volumen Solibri"],
    "Länge (m)":   ["Länge", "Laenge", "Länge BQ", "Länge Solibri"],
    "Dicke (m)":   ["Dicke", "Dicke BQ", "Stärke", "Dicke Solibri"],
    "Höhe (m)":    ["Höhe", "Hoehe", "Höhe BQ", "Höhe Solibri"]
}

def convert_size_to_m(x):
    """
    Wandelt Strings mit Einheiten (mm, cm, dm, m sowie mm2, cm2, dm2, m2, mm3, cm3, dm3, m3)
    korrekt in Meter (bzw. m2, m3) um.
    Gibt bei '0 mm', '0 cm2' etc. direkt pd.NA zurück.
    """
    if pd.isna(x):
        return pd.NA
    s = str(x).strip()
    m = re.match(
        r"^([\d\.,]+(?:[eE][\+\-]?\d+)?)\s*(mm2|cm2|dm2|m2|mm3|cm3|dm3|m3|mm|cm|dm|m)$",
        s,
        flags=re.IGNORECASE
    )
    if m:
        num_str, unit = m.groups()
        num = float(num_str.replace(",", "."))
        # Neu: bei num == 0 direkt None
        if num == 0:
            return pd.NA
        unit = unit.lower()
        if unit.endswith(("2", "3")):
            base, exp = unit[:-1], int(unit[-1])
        else:
            base, exp = unit, 1
        factor = {"mm": 0.001, "cm": 0.01, "dm": 0.1, "m": 1}[base] ** exp
        return num * factor

    # 1) Einheit hinten (falls vorhanden) abtrennen und Dezimal vereinheitlichen
    s1 = s.replace("\xa0", " ").strip()
    s1 = re.sub(r"\s*(mm2|cm2|dm2|m2|mm3|cm3|dm3|m3|mm|cm|dm|m)\s*$", "", s1, flags=re.IGNORECASE)
    s1 = s1.replace("’", "").replace("'", "").replace(" ", "").replace(",", ".")
    
    # 2) Direktversuch: erlaubt auch 2.999E-4
    try:
        val = float(s1)
        return pd.NA if val == 0 else val
    except Exception:
        pass
    
    # 3) Harte Bereinigung: nur Ziffern, Punkt, Vorzeichen und Exponentenzeichen
    cleaned = re.sub(r"[^0-9eE\.\+\-]", "", s1)
    try:
        val = float(cleaned)
        return pd.NA if val == 0 else val
    except Exception:
        st.warning(f"Ungültiges Format in Zelle: '{s}'")
        return pd.NA



def convert_quantity_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Findet typische Mengenspalten (m, m2, m3, Stk., Stück, kg, lm, lfm, qm, cbm, cm, dm, Menge, Anzahl)
    anhand des Spaltennamens und konvertiert deren Werte robust zu float.
    Handhabt Tausendertrennzeichen (., ', Leerzeichen) und Dezimaltrennzeichen (., ,).
    """
    unit_patterns = [
        r"\bmenge\b", r"\banzahl\b",
        r"\bm\b", r"\bm2\b", r"\bm3\b",
        r"\bqm\b", r"\bcbm\b", r"\blm\b", r"\blfm\b",
        r"\bkg\b", r"\bt\b",
        r"\bcm\b", r"\bdm\b",
        r"\bstk\b", r"\bstueck\b", r"\bstück\b"
    ]
    unit_regex = re.compile("|".join(unit_patterns), flags=re.IGNORECASE)

    def parse_num(x):
        if pd.isna(x):
            return pd.NA
        s = str(x).strip()
        if s == "":
            return pd.NA
    
        # Tausenderzeichen entfernen, Dezimal vereinheitlichen
        s = s.replace("\xa0", "").replace(" ", "").replace("’", "").replace("'", "")
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")
        else:
            if "," in s:
                s = s.replace(",", ".")
    
        # Einheiten am Ende entfernen (verhindert 'E-43' durch 'm3')
        s = re.sub(r"(mm2|cm2|dm2|m2|mm3|cm3|dm3|m3|mm|cm|dm|m|qm|cbm|lm|lfm|kg|t|stk|stueck|stück|l)$",
                   "", s, flags=re.IGNORECASE)
    
        # Direktversuch: unterstützt 1.23E-4
        try:
            return float(s)
        except ValueError:
            pass
    
        # Fallback: harte Bereinigung, Exponenten zulassen
        s = re.sub(r"[^0-9eE\.\+\-]", "", s)
        try:
            return float(s) if s not in ("", "-", ".", "+", "e", "E") else pd.NA
        except ValueError:
            return pd.NA


    target_cols = [c for c in df.columns if unit_regex.search(str(c).lower())]
    for c in target_cols:
        df[c] = df[c].map(parse_num).astype("Float64")
    return df



def clean_columns_values(df: pd.DataFrame,
                         delete_enabled: bool = False,
                         custom_chars: str = "") -> pd.DataFrame:
    """
    1) 'Nicht klassifiziert' ⇒ None
    2) Nur Spalten aus COLUMN_PRESET keys: Einheitserkennung & Konvertierung
    3) Ganze 0-Werte ⇒ None
    4) Zusätzliche Zeichen (kommagetrennt) löschen, falls aktiviert
    5) Warnung bei komplett leeren Mengenspalten
    """
    # 1)
    df = df.replace("Nicht klassifiziert", pd.NA)

    # 1.1) Spalten mit Farbangaben löschen
    for drop_col in ("Farbe", "Color"):
        if drop_col in df.columns:
            df.drop(columns=drop_col, inplace=True)

     # 'oi' in "Unter Terrain" ⇒ None
    if "Unter Terrain" in df.columns:
        df.loc[df["Unter Terrain"] == "oi", "Unter Terrain"] = pd.NA

    # 2) Konvertierung nur für die Standard-Mengenspalten
    preset_cols = [col for col in COLUMN_PRESET.keys() if col in df.columns]
    empty_cols = []
    for col in preset_cols:
        df[col] = df[col].apply(convert_size_to_m)
        # 3) Null-Werte
        df[col] = df[col].mask(df[col] == 0, pd.NA)
        if df[col].isna().all():
            empty_cols.append(col)

    # 4) Optional: Zusätzliche Zeichen löschen
    if delete_enabled and custom_chars:
        chars = [c.strip() for c in custom_chars.split(",") if c.strip()]
        for col in df.columns:
            if df[col].dtype == object:
                for ch in chars:
                    df[col] = df[col].str.replace(ch, "", regex=False)
                df[col] = df[col].mask(
                    df[col].str.strip().isin(["", "0", "0.0"]),
                    pd.NA
                )

    # 5) Warnung
    if empty_cols:
        st.warning(
            "Folgende Mengenspalten nach Bereinigung komplett leer: "
            + ", ".join(empty_cols)
        )

    return df

def detect_header_row(df_raw: pd.DataFrame, suchbegriff: str = "Teilprojekt") -> int:
    """
    Index der ersten Zeile, die suchbegriff enthält. Fallback 0.
    """
    for idx, row in df_raw.iterrows():
        if row.astype(str).str.contains(suchbegriff, case=False, na=False).any():
            return idx
    return 0

def apply_preset_hierarchy(df: pd.DataFrame,
                           existing_hierarchy: dict,
                           preset: dict = None) -> dict:
    """
    Füllt existing_hierarchy nur, wenn noch leer, anhand COLUMN_PRESET.
    """
    if preset is None:
        preset = {
            "Flaeche": COLUMN_PRESET["Fläche (m2)"],
            "Volumen": COLUMN_PRESET["Volumen (m3)"],
            "Laenge":  COLUMN_PRESET["Länge (m)"],
            "Dicke":   COLUMN_PRESET["Dicke (m)"],
            "Hoehe":   COLUMN_PRESET["Höhe (m)"]
        }
    # vorhandene validieren
    for m, defaults in existing_hierarchy.items():
        existing_hierarchy[m] = [c for c in defaults if c in df.columns]
    # nur wenn alle noch leer
    if all(not v for v in existing_hierarchy.values()):
        for measure, keywords in preset.items():
            detected = []
            detected += [
                col for col in df.columns
                if "bq" in col.lower() and any(k.lower() in col.lower() for k in keywords)
            ]
            detected += [
                col for col in keywords if col in df.columns and col not in detected
            ]
            detected += [
                col for col in df.columns
                if any(k.lower() in col.lower() for k in keywords) and col not in detected
            ]
            sol = [c for c in detected if "solibri" in c.lower()]
            detected = [c for c in detected if c not in sol] + sol
            if detected:
                existing_hierarchy[measure] = detected
    return existing_hierarchy

def _as_lower_str_or_none(x):
    # Nur echte Strings zulassen; alles andere ignorieren
    return x.lower() if isinstance(x, str) else None

def rename_columns_to_standard(df: pd.DataFrame) -> pd.DataFrame:
    # 0) MultiIndex-Spalten aufloesen -> eindeutige Stringnamen
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = ["__".join([str(p) for p in tup if p is not None]) for tup in df.columns]

    renamed = {}

    # 1) Aliaslisten von None/Non-Strings saeubern (falls COLUMN_PRESET unsauber)
    clean_preset = {
        standard: [a for a in (alts or []) if isinstance(a, str) and a.strip() != ""]
        for standard, alts in COLUMN_PRESET.items()
    }

    # 2) Spaltennamen, die keine Strings sind, ignorieren (nicht zwangsweise in Strings casten)
    valid_cols = [c for c in df.columns if isinstance(c, str) and c.strip() != ""]

    for standard, alts in clean_preset.items():
        matches = []
        for c in valid_cols:
            c_l = _as_lower_str_or_none(c)
            # Wenn c kein String ist, ueberspringen
            if c_l is None:
                continue
            # Sobald ein Alias als Teilstring vorkommt, ist es ein Match
            for alt in alts:
                alt_l = _as_lower_str_or_none(alt)
                if alt_l and alt_l in c_l:
                    matches.append(c)
                    break

        if matches:
            if len(matches) > 1:
                st.warning(f"Mehrfach: {matches}. Nutze '{matches[0]}' fuer '{standard}'.")
            renamed[matches[0]] = standard

    if renamed:
        df = df.rename(columns=renamed)

    return df

def prepend_values_cleaning(df: pd.DataFrame,
                            delete_enabled: bool = False,
                            custom_chars: str = "") -> pd.DataFrame:
    """
    Nur Werte bereinigen, keine Spalten umbenennen.
    """
    return clean_columns_values(df, delete_enabled, custom_chars)
