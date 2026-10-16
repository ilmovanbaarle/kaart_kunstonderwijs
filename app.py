from copy import deepcopy
from io import BytesIO
from pathlib import Path
import math
import re

import pandas as pd
import plotly.express as px
import streamlit as st


# =============================================================================
# INSTELLINGEN
# =============================================================================

DEFAULT_WORKBOOK = "foto_2025_python_met_uitvragen.xlsx"

LABOUR_COLUMNS = [
    "Werkloosheid (%)",
    "Aantal maanden werkloos voor eerste baan",
    "Vaste aanstelling bij huidige functie (%)",
    "Dienstverband huidige functie: zelfstandige/freelance (%)",
    "Huidige functie binnen eigen of verwante opleidingsrichting (%)",
    "Bruto uurinkomen huidige functie (euro)",
]

DROP_COLUMNS = [
    "Bronbestand uitvraag",
    "Vraag 2: Aantal inschrijvingen correct?",
]

VARIANT_RENAMES = {
    "Onderwijsinstelling": "INSTELLINGSNAAM ACTUEEL",
    "Gemeentenaam": "GEMEENTENAAM",
    "ISAT": "OPLEIDINGSCODE ACTUEEL",
    "Naam studie": "OPLEIDINGSNAAM ACTUEEL",
}

SCHOOL_COORDS = {
    ("Amsterdamse Hogeschool voor de Kunsten", "Amsterdam"): (52.3676, 4.9041),

    ("ArtEZ", "Arnhem"): (51.9851, 5.8987),
    ("ArtEZ", "Enschede"): (52.2215, 6.8937),
    ("ArtEZ", "Zwolle"): (52.5168, 6.0830),

    ("Avans Hogeschool", "Breda"): (51.5719, 4.7683),
    ("Avans Hogeschool", "'s-Hertogenbosch"): (51.6978, 5.3037),
    ("Avans Hogeschool", "Den Bosch"): (51.6978, 5.3037),

    ("Codarts, Hogeschool voor de Kunsten", "Rotterdam"): (51.9225, 4.4792),

    ("Design Academy Eindhoven", "Eindhoven"): (51.4416, 5.4697),

    ("Fontys Hogeschool", "Tilburg"): (51.5555, 5.0913),

    ("Gerrit Rietveld Academie", "Amsterdam"): (52.3676, 4.9041),

    ("Hanzehogeschool Groningen", "Amsterdam"): (52.3676, 4.9041),
    ("Hanzehogeschool Groningen", "Groningen"): (53.2194, 6.5665),
    ("Hanzehogeschool Groningen", "Leeuwarden"): (53.2012, 5.7999),

    ("Hogeschool Inholland", "Haarlem"): (52.3874, 4.6462),

    ("Hogeschool Leiden", "Leiden"): (52.1601, 4.4970),

    ("Hogeschool Rotterdam", "Rotterdam"): (51.9225, 4.4792),

    ("Hogeschool der Kunsten Den Haag", "'s-Gravenhage"): (52.0705, 4.3007),
    ("Hogeschool der Kunsten Den Haag", "Den Haag"): (52.0705, 4.3007),

    ("Hogeschool voor de Kunsten Utrecht", "Utrecht"): (52.0907, 5.1214),

    ("NHL Stenden Hogeschool", "Leeuwarden"): (53.2012, 5.7999),

    ("Saxion Hogeschool", "Deventer"): (52.2661, 6.1552),

    ("Zuyd Hogeschool", "Maastricht"): (50.8514, 5.6910),
}


# =============================================================================
# ALGEMENE HULPFUNCTIES
# =============================================================================

def clean_text(value):
    """Maak tekstwaarden schoon."""
    if pd.isna(value):
        return ""

    text = str(value).strip()

    if text.lower() in {"nan", "none", "<na>"}:
        return ""

    return text


def normalize_text(value):
    """Normaliseer tekst uitsluitend voor vergelijkingen."""
    text = clean_text(value)
    text = re.sub(r"\s+", " ", text)

    return text.casefold()


def normalize_city(value):
    """Normaliseer enkele verschillende schrijfwijzen van gemeenten."""
    city = clean_text(value)

    aliases = {
        "den bosch": "'s-Hertogenbosch",
        "s-hertogenbosch": "'s-Hertogenbosch",
        "'s hertogenbosch": "'s-Hertogenbosch",
        "den haag": "'s-Gravenhage",
        "s-gravenhage": "'s-Gravenhage",
        "'s gravenhage": "'s-Gravenhage",
    }

    return aliases.get(city.casefold(), city)


def clean_code(value):
    """Verwijder een eventuele door Excel toegevoegde .0."""
    text = clean_text(value)

    return re.sub(r"\.0$", "", text)


def format_general_value(value):
    """Formatteer algemene waarden voor hoverinformatie."""
    if pd.isna(value):
        return "Niet beschikbaar"

    text = clean_text(value)

    if not text:
        return "Niet beschikbaar"

    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))

        return f"{value:g}".replace(".", ",")

    return text


def format_student_count(value):
    """Formatteer aantallen studenten zo leesbaar mogelijk."""
    if pd.isna(value) or clean_text(value) == "":
        return "Niet beschikbaar"

    try:
        number = float(value)

        if number.is_integer():
            return f"{int(number):,}".replace(",", ".")

        return (
            f"{number:,.1f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

    except (TypeError, ValueError):
        return clean_text(value)


def format_value(value, column):
    """Formatteer arbeidsmarktwaarden."""
    if pd.isna(value) or clean_text(value) == "":
        return "Niet beschikbaar"

    if column == "Bruto uurinkomen huidige functie (euro)":
        try:
            return f"EUR {float(value):.2f}".replace(".", ",")
        except (TypeError, ValueError):
            return clean_text(value)

    if column.endswith("(%)"):
        try:
            return f"{float(value):g}%"
        except (TypeError, ValueError):
            return clean_text(value)

    return format_general_value(value)


def safe_filename(value):
    """Maak van tekst een geldige bestandsnaam."""
    return "".join(
        character
        if character.isalnum() or character in "._- "
        else "_"
        for character in value
    )


def unique_clean_values(series):
    """Geef unieke niet-lege waarden uit een kolom."""
    values = {}

    for value in series:
        text = clean_text(value)

        if not text:
            continue

        key = normalize_text(text)

        if key not in values:
            values[key] = text

    return sorted(
        values.values(),
        key=lambda x: x.casefold(),
    )


# =============================================================================
# CLUSTERS
# =============================================================================

def split_cluster_values(value):
    """
    Splits eventueel meerdere clusterwaarden binnen één cel.

    Nieuwe regels, puntkomma's en pipes gelden als scheiding.
    Komma's bewust niet, omdat die onderdeel van namen kunnen zijn.
    """
    text = clean_text(value)

    if not text:
        return []

    parts = re.split(r"[\n\r;|]+", text)

    result = []
    seen = set()

    for part in parts:
        part = re.sub(
            r"^[•\-–—]+\s*",
            "",
            part,
        )

        part = re.sub(
            r"\s+",
            " ",
            part,
        ).strip()

        if not part:
            continue

        key = normalize_text(part)

        if key not in seen:
            seen.add(key)
            result.append(part)

    return result


def get_cluster_values(df):
    """
    Combineer alle unieke waarden uit Cluster en Cluster 2
    tot één alfabetische keuzelijst.
    """
    values = {}

    for column in ["Cluster", "Cluster 2"]:

        if column not in df.columns:
            continue

        for cell in df[column]:

            for value in split_cluster_values(cell):

                key = normalize_text(value)

                if key and key not in values:
                    values[key] = value

    return sorted(
        values.values(),
        key=lambda x: x.casefold(),
    )


def row_in_cluster(row, selected_cluster):
    """
    Een rij hoort bij het cluster als de gekozen waarde voorkomt
    in Cluster OF Cluster 2.
    """
    selected_key = normalize_text(selected_cluster)

    for column in ["Cluster", "Cluster 2"]:

        if column not in row.index:
            continue

        for value in split_cluster_values(row.get(column, "")):

            if normalize_text(value) == selected_key:
                return True

    return False


def filter_cluster(df, selected_cluster):
    """Filter dataframe op Cluster OF Cluster 2."""
    if not selected_cluster:
        return df.iloc[0:0].copy()

    mask = df.apply(
        lambda row: row_in_cluster(
            row,
            selected_cluster,
        ),
        axis=1,
    )

    return df.loc[mask].copy()


# =============================================================================
# COORDINATEN
# =============================================================================

def add_coordinates(df):
    """Voeg Latitude en Longitude toe."""
    df = df.copy()

    def lookup(row):

        school = clean_text(
            row.get(
                "INSTELLINGSNAAM ACTUEEL",
                "",
            )
        )

        city = normalize_city(
            row.get(
                "GEMEENTENAAM",
                "",
            )
        )

        # Exacte combinatie instelling + gemeente.
        exact_match = SCHOOL_COORDS.get(
            (school, city)
        )

        if exact_match:
            return exact_match

        # Indien instelling maar één bekende locatie heeft,
        # kan die veilig als fallback worden gebruikt.
        matches = [
            coordinates
            for (school_name, _), coordinates
            in SCHOOL_COORDS.items()
            if school_name == school
        ]

        unique_matches = list(
            dict.fromkeys(matches)
        )

        if len(unique_matches) == 1:
            return unique_matches[0]

        return None, None

    coordinates = df.apply(
        lookup,
        axis=1,
        result_type="expand",
    )

    df["Latitude"] = coordinates[0]
    df["Longitude"] = coordinates[1]

    return df


def apply_coordinate_offsets(df):
    """
    Als verschillende instellingen exact dezelfde coördinaten hebben,
    worden de markers licht uit elkaar geschoven.

    Records van dezelfde instelling blijven wel op één marker staan.
    """
    df = df.copy()

    if df.empty:
        return df

    df["Map Latitude"] = df["Latitude"].astype(float)
    df["Map Longitude"] = df["Longitude"].astype(float)

    groups = df.groupby(
        ["Latitude", "Longitude"],
        dropna=False,
        sort=False,
    )

    for _, indices in groups.groups.items():

        indices = list(indices)

        if len(indices) <= 1:
            continue

        institutions = sorted(
            {
                clean_text(
                    df.at[
                        index,
                        "INSTELLINGSNAAM ACTUEEL",
                    ]
                )
                for index in indices
            }
        )

        if len(institutions) <= 1:
            continue

        positions = {
            institution: position
            for position, institution
            in enumerate(institutions)
        }

        total = len(institutions)

        if total <= 4:
            radius = 0.018
        else:
            radius = 0.025

        for index in indices:

            institution = clean_text(
                df.at[
                    index,
                    "INSTELLINGSNAAM ACTUEEL",
                ]
            )

            position = positions[institution]

            angle = (
                2
                * math.pi
                * position
                / total
            )

            df.at[
                index,
                "Map Latitude",
            ] += radius * math.sin(angle)

            df.at[
                index,
                "Map Longitude",
            ] += radius * math.cos(angle)

    return df


# =============================================================================
# KOPPELSLEUTELS
# =============================================================================

def prepare_dataframe_keys(df):
    """Maak interne sleutels voor het koppelen van opleidingen en varianten."""
    df = df.copy()

    df["_institution_key"] = (
        df["INSTELLINGSNAAM ACTUEEL"]
        .map(normalize_text)
    )

    df["_code_key"] = (
        df["OPLEIDINGSCODE ACTUEEL"]
        .map(clean_code)
        .map(normalize_text)
    )

    df["_study_key"] = (
        df["OPLEIDINGSNAAM ACTUEEL"]
        .map(normalize_text)
    )

    df["_city_key"] = (
        df["GEMEENTENAAM"]
        .map(normalize_city)
        .map(normalize_text)
    )

    return df


# =============================================================================
# EXCEL INLEZEN
# =============================================================================

@st.cache_data(show_spinner=False)
def read_workbook(source):

    opleidingen = pd.read_excel(
        source,
        sheet_name="opleidingen",
        engine="openpyxl",
    )

    varianten = pd.read_excel(
        source,
        sheet_name="varianten",
        engine="openpyxl",
    )

    # Kolomnamen opschonen.
    opleidingen.columns = opleidingen.columns.map(
        lambda x: str(x).strip()
    )

    varianten.columns = varianten.columns.map(
        lambda x: str(x).strip()
    )

    opleidingen = opleidingen.drop(
        columns=DROP_COLUMNS,
        errors="ignore",
    )

    varianten = (
        varianten
        .drop(
            columns=DROP_COLUMNS,
            errors="ignore",
        )
        .rename(
            columns=VARIANT_RENAMES
        )
    )

    required_opleiding_columns = [
        "INSTELLINGSNAAM ACTUEEL",
        "OPLEIDINGSCODE ACTUEEL",
        "OPLEIDINGSNAAM ACTUEEL",
        "GEMEENTENAAM",
        "Cluster",
        "Cluster 2",
    ]

    required_variant_columns = [
        "INSTELLINGSNAAM ACTUEEL",
        "OPLEIDINGSCODE ACTUEEL",
        "OPLEIDINGSNAAM ACTUEEL",
        "GEMEENTENAAM",
        "Variant",
        "Aantal inschrijvingen 2025",
        "Cluster",
        "Cluster 2",
    ]

    missing_opleiding_columns = [
        column
        for column in required_opleiding_columns
        if column not in opleidingen.columns
    ]

    missing_variant_columns = [
        column
        for column in required_variant_columns
        if column not in varianten.columns
    ]

    if missing_opleiding_columns:
        raise ValueError(
            "In het tabblad opleidingen ontbreken de volgende kolommen: "
            + ", ".join(missing_opleiding_columns)
        )

    if missing_variant_columns:
        raise ValueError(
            "In het tabblad varianten ontbreken de volgende kolommen: "
            + ", ".join(missing_variant_columns)
        )

    # -------------------------------------------------------------
    # Basisvelden opschonen
    # -------------------------------------------------------------

    for df in [opleidingen, varianten]:

        df["OPLEIDINGSCODE ACTUEEL"] = (
            df["OPLEIDINGSCODE ACTUEEL"]
            .map(clean_code)
            .astype("string")
        )

        df["INSTELLINGSNAAM ACTUEEL"] = (
            df["INSTELLINGSNAAM ACTUEEL"]
            .astype("string")
            .str.strip()
        )

        df["OPLEIDINGSNAAM ACTUEEL"] = (
            df["OPLEIDINGSNAAM ACTUEEL"]
            .astype("string")
            .str.strip()
        )

        df["GEMEENTENAAM"] = (
            df["GEMEENTENAAM"]
            .map(normalize_city)
        )

    opleidingen = prepare_dataframe_keys(
        opleidingen
    )

    varianten = prepare_dataframe_keys(
        varianten
    )

    # -------------------------------------------------------------
    # Arbeidsmarktinformatie uit opleidingen aan varianten koppelen
    # -------------------------------------------------------------

    available_labour_columns = [
        column
        for column in LABOUR_COLUMNS
        if column in opleidingen.columns
    ]

    lookup_columns = [
        "INSTELLINGSNAAM ACTUEEL",
        "OPLEIDINGSCODE ACTUEEL",
        "OPLEIDINGSNAAM ACTUEEL",
        "GEMEENTENAAM",
        "_institution_key",
        "_code_key",
        "_study_key",
        "_city_key",
        *available_labour_columns,
    ]

    opleiding_lookup = (
        opleidingen[lookup_columns]
        .drop_duplicates()
        .copy()
    )

    # Eventuele reeds aanwezige arbeidsmarktvelden uit varianten verwijderen.
    varianten = varianten.drop(
        columns=LABOUR_COLUMNS,
        errors="ignore",
    )

    # Varianten mét gemeente koppelen zo specifiek mogelijk.
    variants_with_city = varianten[
        varianten["_city_key"].ne("")
    ].copy()

    variants_without_city = varianten[
        varianten["_city_key"].eq("")
    ].copy()

    if not variants_with_city.empty:

        city_lookup_columns = [
            "_institution_key",
            "_code_key",
            "_study_key",
            "_city_key",
            *available_labour_columns,
        ]

        city_lookup = (
            opleiding_lookup[
                city_lookup_columns
            ]
            .drop_duplicates()
        )

        variants_with_city = (
            variants_with_city.merge(
                city_lookup,
                how="left",
                on=[
                    "_institution_key",
                    "_code_key",
                    "_study_key",
                    "_city_key",
                ],
            )
        )

    # Varianten zonder gemeente krijgen de gemeente uit tabblad opleidingen.
    if not variants_without_city.empty:

        no_city_lookup_columns = [
            "_institution_key",
            "_code_key",
            "_study_key",
            "GEMEENTENAAM",
            "_city_key",
            *available_labour_columns,
        ]

        no_city_lookup = (
            opleiding_lookup[
                no_city_lookup_columns
            ]
            .drop_duplicates()
        )

        variants_without_city = (
            variants_without_city
            .drop(
                columns=[
                    "GEMEENTENAAM",
                    "_city_key",
                ],
                errors="ignore",
            )
            .merge(
                no_city_lookup,
                how="left",
                on=[
                    "_institution_key",
                    "_code_key",
                    "_study_key",
                ],
            )
        )

    varianten = pd.concat(
        [
            variants_with_city,
            variants_without_city,
        ],
        ignore_index=True,
        sort=False,
    )

    # -------------------------------------------------------------
    # Coördinaten toevoegen
    # -------------------------------------------------------------

    opleidingen = add_coordinates(
        opleidingen
    )

    varianten = add_coordinates(
        varianten
    )

    return opleidingen, varianten


# =============================================================================
# ARBEIDSMARKTINFORMATIE
# =============================================================================

def show_labour_panel(df):

    st.sidebar.markdown(
        "### Arbeidsmarktinformatie"
    )

    for column in LABOUR_COLUMNS:

        if column not in df.columns:
            values = []

        else:
            values = (
                df[column]
                .dropna()
                .drop_duplicates()
                .tolist()
            )

        if len(values) == 0:
            text = "Niet beschikbaar"

        else:
            text = " / ".join(
                format_value(
                    value,
                    column,
                )
                for value in values
            )

        st.sidebar.metric(
            column,
            text,
        )


# =============================================================================
# STUDIES ONDERDEEL VAN CLUSTER
# =============================================================================

def show_studies_in_cluster(df):

    st.sidebar.markdown(
        "### Studies onderdeel van cluster"
    )

    if (
        df.empty
        or "OPLEIDINGSNAAM ACTUEEL"
        not in df.columns
    ):
        st.sidebar.caption(
            "Geen studies gevonden."
        )
        return

    studies = unique_clean_values(
        df["OPLEIDINGSNAAM ACTUEEL"]
    )

    if not studies:
        st.sidebar.caption(
            "Geen studies gevonden."
        )
        return

    for study in studies:
        st.sidebar.markdown(
            f"- {study}"
        )

def show_variants_in_cluster(df):
    """Toon alle unieke varianten binnen het geselecteerde variantcluster."""

    st.sidebar.markdown(
        "### Varianten onderdeel van cluster"
    )

    if (
        df.empty
        or "Variant" not in df.columns
    ):
        st.sidebar.caption(
            "Geen varianten gevonden."
        )
        return

    variants = unique_clean_values(
        df["Variant"]
    )

    if not variants:
        st.sidebar.caption(
            "Geen varianten gevonden."
        )
        return

    for variant in variants:
        st.sidebar.markdown(
            f"- {variant}"
        )

# =============================================================================
# EXCEL DOWNLOAD
# =============================================================================

def subset_to_excel(df, sheet_name):

    output = BytesIO()

    # Interne technische kolommen niet exporteren.
    export_df = df.drop(
        columns=[
            column
            for column in df.columns
            if (
                column.startswith("_")
                or column in {
                    "Map Latitude",
                    "Map Longitude",
                }
            )
        ],
        errors="ignore",
    )

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:

        safe_sheet_name = sheet_name[:31]

        export_df.to_excel(
            writer,
            sheet_name=safe_sheet_name,
            index=False,
        )

        worksheet = writer.book[
            safe_sheet_name
        ]

        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = (
            worksheet.dimensions
        )

        for cells in worksheet.columns:

            maximum_length = max(
                (
                    len(str(cell.value))
                    if cell.value is not None
                    else 0
                )
                for cell in cells[:200]
            )

            width = min(
                max(
                    maximum_length + 2,
                    10,
                ),
                45,
            )

            worksheet.column_dimensions[
                cells[0].column_letter
            ].width = width

    return output.getvalue()


# =============================================================================
# VARIANTEN KOPPELEN AAN OPLEIDINGSCLUSTER
# =============================================================================

def create_selected_study_keys(
    filtered_opleidingen,
):

    code_keys = set()
    name_keys = set()

    for _, row in filtered_opleidingen.iterrows():

        institution = row.get(
            "_institution_key",
            "",
        )

        code = row.get(
            "_code_key",
            "",
        )

        study = row.get(
            "_study_key",
            "",
        )

        if institution and code:
            code_keys.add(
                (institution, code)
            )

        if institution and study:
            name_keys.add(
                (institution, study)
            )

    return code_keys, name_keys


def variants_for_education_cluster(
    opleidingen,
    varianten,
    selected_education_cluster,
):

    # Eerst opleidingen binnen gekozen opleidingscluster bepalen.
    selected_opleidingen = filter_cluster(
        opleidingen,
        selected_education_cluster,
    )

    code_keys, name_keys = (
        create_selected_study_keys(
            selected_opleidingen
        )
    )

    def belongs_to_selection(row):

        institution = row.get(
            "_institution_key",
            "",
        )

        code = row.get(
            "_code_key",
            "",
        )

        study = row.get(
            "_study_key",
            "",
        )

        # Primaire koppeling:
        # instelling + opleidingscode
        if institution and code:

            if (
                institution,
                code,
            ) in code_keys:
                return True

        # Fallback:
        # instelling + opleidingsnaam
        if institution and study:

            if (
                institution,
                study,
            ) in name_keys:
                return True

        return False

    mask = varianten.apply(
        belongs_to_selection,
        axis=1,
    )

    related_variants = (
        varianten.loc[mask]
        .copy()
    )

    return (
        selected_opleidingen,
        related_variants,
    )


# =============================================================================
# HOVERINFORMATIE OPLEIDINGEN
# =============================================================================

def build_education_hover(group):

    institution = clean_text(
        group[
            "INSTELLINGSNAAM ACTUEEL"
        ].iloc[0]
    )

    city = clean_text(
        group[
            "GEMEENTENAAM"
        ].iloc[0]
    )

    # Instellingsnaam als kop.
    lines = [
        f"<b>{institution}</b>"
    ]

    if city:
        lines.append(city)

    lines.append("")

    sort_columns = [
        column
        for column in [
            "OPLEIDINGSNAAM ACTUEEL",
            "OPLEIDINGSVORM",
        ]
        if column in group.columns
    ]

    if sort_columns:
        sorted_group = group.sort_values(
            sort_columns,
            kind="stable",
        )
    else:
        sorted_group = group

    for position, (_, row) in enumerate(
        sorted_group.iterrows(),
        start=1,
    ):

        study_name = clean_text(
            row.get(
                "OPLEIDINGSNAAM ACTUEEL",
                "",
            )
        )

        if not study_name:
            study_name = "Onbekende opleiding"

        lines.append(
            f"<b>{position}. {study_name}</b>"
        )

        lines.append(
            "Opleidingsvorm: "
            + format_general_value(
                row.get("OPLEIDINGSVORM")
            )
        )

        lines.append(
            "Aantal studenten 2023: "
            + format_student_count(
                row.get("2023")
            )
        )

        lines.append(
            "Aantal studenten 2024: "
            + format_student_count(
                row.get("2024")
            )
        )

        lines.append(
            "Aantal studenten 2025: "
            + format_student_count(
                row.get("2025")
            )
        )

        lines.append(
            "3-jarig gemiddelde studenten: "
            + format_student_count(
                row.get(
                    "3 jarig gemiddelde studenten"
                )
            )
        )

        lines.append(
            "Varianten studie: "
            + format_general_value(
                row.get("Varianten")
            )
        )

        if position < len(sorted_group):
            lines.append("")

    return "<br>".join(lines)


# =============================================================================
# HOVERINFORMATIE VARIANTEN
# =============================================================================

def build_variant_hover(group):

    institution = clean_text(
        group[
            "INSTELLINGSNAAM ACTUEEL"
        ].iloc[0]
    )

    city = clean_text(
        group[
            "GEMEENTENAAM"
        ].iloc[0]
    )

    # Instellingsnaam als kop.
    lines = [
        f"<b>{institution}</b>"
    ]

    if city:
        lines.append(city)

    lines.append("")

    sort_columns = [
        column
        for column in [
            "OPLEIDINGSNAAM ACTUEEL",
            "Variant",
        ]
        if column in group.columns
    ]

    if sort_columns:
        sorted_group = group.sort_values(
            sort_columns,
            kind="stable",
        )
    else:
        sorted_group = group

    for position, (_, row) in enumerate(
        sorted_group.iterrows(),
        start=1,
    ):

        study_name = clean_text(
            row.get(
                "OPLEIDINGSNAAM ACTUEEL",
                "",
            )
        )

        variant_name = clean_text(
            row.get(
                "Variant",
                "",
            )
        )

        if not study_name:
            study_name = "Onbekende studie"

        if not variant_name:
            variant_name = "Onbekende variant"

        lines.append(
            f"<b>{position}. {variant_name}</b>"
        )

        lines.append(
            f"Naam studie: {study_name}"
        )

        lines.append(
            "Aantal inschrijvingen 2025: "
            + format_student_count(
                row.get(
                    "Aantal inschrijvingen 2025"
                )
            )
        )

        if position < len(sorted_group):
            lines.append("")

    return "<br>".join(lines)


# =============================================================================
# RECORDS SAMENVOEGEN PER INSTELLING + LOCATIE
# =============================================================================

def aggregate_map_data(
    df,
    mode,
):

    mappable = df.dropna(
        subset=[
            "Latitude",
            "Longitude",
        ]
    ).copy()

    if mappable.empty:
        return pd.DataFrame()

    group_columns = [
        "INSTELLINGSNAAM ACTUEEL",
        "GEMEENTENAAM",
        "Latitude",
        "Longitude",
    ]

    records = []

    for group_key, group in mappable.groupby(
        group_columns,
        dropna=False,
        sort=True,
    ):

        (
            institution,
            city,
            latitude,
            longitude,
        ) = group_key

        # ----------------------------------------------------------
        # OPLEIDINGEN
        # ----------------------------------------------------------

        if mode == "opleidingen":

            deduplication_columns = [
                column
                for column in [
                    "INSTELLINGSNAAM ACTUEEL",
                    "GEMEENTENAAM",
                    "OPLEIDINGSCODE ACTUEEL",
                    "OPLEIDINGSNAAM ACTUEEL",
                    "OPLEIDINGSVORM",
                    "2023",
                    "2024",
                    "2025",
                    "3 jarig gemiddelde studenten",
                    "Varianten",
                ]
                if column in group.columns
            ]

            display_group = (
                group.drop_duplicates(
                    subset=deduplication_columns
                )
            )

            hover_text = (
                build_education_hover(
                    display_group
                )
            )

            item_count = len(
                display_group
            )

            if item_count == 1:
                item_type = "opleiding"
            else:
                item_type = "opleidingen"

        # ----------------------------------------------------------
        # VARIANTEN
        # ----------------------------------------------------------

        else:

            deduplication_columns = [
                column
                for column in [
                    "INSTELLINGSNAAM ACTUEEL",
                    "GEMEENTENAAM",
                    "OPLEIDINGSCODE ACTUEEL",
                    "OPLEIDINGSNAAM ACTUEEL",
                    "Variant",
                    "Aantal inschrijvingen 2025",
                ]
                if column in group.columns
            ]

            display_group = (
                group.drop_duplicates(
                    subset=deduplication_columns
                )
            )

            hover_text = (
                build_variant_hover(
                    display_group
                )
            )

            item_count = len(
                display_group
            )

            if item_count == 1:
                item_type = "variant"
            else:
                item_type = "varianten"

        records.append(
            {
                "INSTELLINGSNAAM ACTUEEL": clean_text(
                    institution
                ),
                "GEMEENTENAAM": clean_text(
                    city
                ),
                "Latitude": float(
                    latitude
                ),
                "Longitude": float(
                    longitude
                ),
                "Aantal onderdelen": item_count,
                "Aantal label": (
                    f"{item_count} {item_type}"
                ),
                "Marker label": (
                    str(item_count)
                    if item_count > 1
                    else ""
                ),
                "Hover": hover_text,
            }
        )

    aggregated = pd.DataFrame(
        records
    )

    # Instellingen die dezelfde coördinaten hebben
    # licht uit elkaar schuiven.
    aggregated = apply_coordinate_offsets(
        aggregated
    )

    return aggregated


# =============================================================================
# KAART
# =============================================================================

def create_map_figure(
    map_data,
    title,
):

    institutions = sorted(
        unique_clean_values(
            map_data[
                "INSTELLINGSNAAM ACTUEEL"
            ]
        ),
        key=lambda value: value.casefold(),
    )

    max_count = max(
        int(
            map_data[
                "Aantal onderdelen"
            ].max()
        ),
        1,
    )

    if max_count > 1:
        size_max = 30
    else:
        size_max = 22

    fig = px.scatter_map(
        map_data,
        lat="Map Latitude",
        lon="Map Longitude",
        color="INSTELLINGSNAAM ACTUEEL",
        size="Aantal onderdelen",
        size_max=size_max,
        category_orders={
            "INSTELLINGSNAAM ACTUEEL": institutions
        },
        hover_name=None,
        hover_data={
            "Map Latitude": False,
            "Map Longitude": False,
            "INSTELLINGSNAAM ACTUEEL": False,
            "Aantal onderdelen": False,
            "GEMEENTENAAM": False,
            "Hover": False,
        },
        custom_data=[
            "Hover",
            "Marker label",
            "Aantal onderdelen",
        ],
        height=760,
        title=title,
    )

    # -------------------------------------------------------------
    # Zwarte achtergrondmarker toevoegen voor rand
    # -------------------------------------------------------------

    original_traces = list(
        fig.data
    )

    fig.data = ()

    for trace in original_traces:

        outline = deepcopy(
            trace
        )

        outline.name = (
            f"{trace.name} outline"
        )

        outline.showlegend = False
        outline.hoverinfo = "skip"
        outline.hovertemplate = None

        original_sizes = (
            trace.marker.size
        )

        if original_sizes is None:

            outline_sizes = 26

        else:

            try:
                outline_sizes = [
                    float(size) + 7
                    for size in original_sizes
                ]

            except TypeError:
                outline_sizes = (
                    float(original_sizes)
                    + 7
                )

        outline.marker = {
            "size": outline_sizes,
            "color": "black",
            "opacity": 0.95,
        }

        trace.marker.opacity = 0.90

        # Alleen onze zelf opgebouwde hover tonen.
        trace.hovertemplate = (
            "%{customdata[0]}"
            "<extra></extra>"
        )

        fig.add_trace(
            outline
        )

        fig.add_trace(
            trace
        )

    # -------------------------------------------------------------
    # Nummer toevoegen bij locaties met meerdere onderdelen
    # -------------------------------------------------------------

    multiple_items = map_data[
        map_data[
            "Aantal onderdelen"
        ] > 1
    ].copy()

    if not multiple_items.empty:

        fig.add_scattermap(
            lat=multiple_items[
                "Map Latitude"
            ],
            lon=multiple_items[
                "Map Longitude"
            ],
            mode="text",
            text=multiple_items[
                "Marker label"
            ],
            textfont={
                "size": 11,
                "color": "white",
                "family": "Arial Black",
            },
            hoverinfo="skip",
            showlegend=False,
            name="Aantal onderdelen",
        )

    # -------------------------------------------------------------
    # Kaartopmaak
    # -------------------------------------------------------------

    fig.update_layout(
        paper_bgcolor="#eeeeee",
        plot_bgcolor="#eeeeee",
        map={
            "style": "carto-positron",
            "center": {
                "lat": 52.20,
                "lon": 5.30,
            },
            "zoom": 5.35,
            "pitch": 0,
        },
        margin={
            "r": 0,
            "t": 55,
            "l": 0,
            "b": 0,
        },
        legend_title_text=(
            "Onderwijsinstelling"
        ),
        hoverlabel={
            "bgcolor": "white",
            "font_size": 13,
            "font_family": "Arial",
            "align": "left",
        },
    )

    return fig


# =============================================================================
# SELECTIE OPLEIDINGEN
# =============================================================================

def select_education_data(
    opleidingen,
):

    clusters = get_cluster_values(
        opleidingen
    )

    if not clusters:

        st.warning(
            "In het tabblad opleidingen zijn geen "
            "clusterwaarden gevonden."
        )

        st.stop()

    # Dezelfde session_state key wordt ook gebruikt
    # bij de variantenweergave.
    if (
        "selected_education_cluster" not in st.session_state
        or st.session_state.selected_education_cluster not in clusters
    ):
        st.session_state.selected_education_cluster = clusters[0]

    st.sidebar.markdown("### Cluster")

    selected_cluster = st.sidebar.selectbox(
        "Cluster",
        clusters,
        key="selected_education_cluster",
        label_visibility="collapsed",
    )

    filtered = filter_cluster(
        opleidingen,
        selected_cluster,
    )

    title = (
        f"Opleidingen binnen cluster: "
        f"{selected_cluster}"
    )

    return (
        filtered,
        selected_cluster,
        title,
    )


# =============================================================================
# SELECTIE VARIANTEN
# =============================================================================

def select_variant_data(
    opleidingen,
    varianten,
):

    # -------------------------------------------------------------
    # STAP 1: opleidingscluster
    # -------------------------------------------------------------

    education_clusters = get_cluster_values(
        opleidingen
    )

    if not education_clusters:

        st.warning(
            "In het tabblad opleidingen zijn geen "
            "clusterwaarden gevonden."
        )

        st.stop()

    # Dezelfde session_state key als bij opleidingen.
    # Daardoor blijft de selectie behouden wanneer je wisselt
    # tussen opleidingen en varianten.
    if (
        "selected_education_cluster" not in st.session_state
        or st.session_state.selected_education_cluster not in education_clusters
    ):
        st.session_state.selected_education_cluster = education_clusters[0]

    selected_education_cluster = st.sidebar.selectbox(
        "Cluster opleiding",
        education_clusters,
        key="selected_education_cluster",
    )

    # -------------------------------------------------------------
    # Bepalen welke opleidingen hieronder vallen
    # en welke varianten daarbij horen
    # -------------------------------------------------------------

    (
        selected_opleidingen,
        related_variants,
    ) = variants_for_education_cluster(
        opleidingen,
        varianten,
        selected_education_cluster,
    )

    if related_variants.empty:

        st.warning(
            "Voor het gekozen opleidingscluster "
            "zijn geen gekoppelde varianten gevonden."
        )

        st.stop()

    # -------------------------------------------------------------
    # STAP 2: variantcluster
    # -------------------------------------------------------------

    variant_clusters = get_cluster_values(
        related_variants
    )

    if not variant_clusters:

        st.warning(
            "Voor de varianten binnen dit "
            "opleidingscluster zijn geen "
            "variantclusters gevonden."
        )

        st.stop()

    # Controleer ook hier of de eerder gekozen variantcluster
    # nog geldig is na wijziging van het opleidingscluster.
    if (
        "selected_variant_cluster" not in st.session_state
        or st.session_state.selected_variant_cluster not in variant_clusters
    ):
        st.session_state.selected_variant_cluster = variant_clusters[0]

    selected_variant_cluster = st.sidebar.selectbox(
        "Cluster variant",
        variant_clusters,
        key="selected_variant_cluster",
    )

    filtered_variants = filter_cluster(
        related_variants,
        selected_variant_cluster,
    )

    title = (
        f"Varianten binnen "
        f"{selected_education_cluster}"
        f" - "
        f"{selected_variant_cluster}"
    )

    return (
        filtered_variants,
        selected_opleidingen,
        selected_education_cluster,
        selected_variant_cluster,
        title,
    )


# =============================================================================
# HOOFDPROGRAMMA
# =============================================================================

def main():

    st.set_page_config(
        page_title="Kaart kunstopleidingen",
        layout="wide",
    )

    st.title(
        "Locaties van opleidingen en varianten"
    )

    # -------------------------------------------------------------------------
    # Bestand selecteren
    # -------------------------------------------------------------------------

    default_file = (
        Path(__file__)
        .with_name(
            DEFAULT_WORKBOOK
        )
    )

    source = default_file

    # -------------------------------------------------------------------------
    # Data inlezen
    # -------------------------------------------------------------------------

    try:

        opleidingen, varianten = (
            read_workbook(
                source
            )
        )

    except Exception as error:

        st.error(
            "Het Excelbestand kon niet "
            "correct worden ingelezen."
        )

        st.exception(
            error
        )

        st.stop()

    # -------------------------------------------------------------------------
    # Gegevensbron kiezen
    # -------------------------------------------------------------------------

    st.sidebar.markdown("### Kies gegevensbron")

    sheet = st.sidebar.radio(
        "Kies gegevensbron",
        [
            "opleidingen",
            "varianten",
        ],
        horizontal=True,
        key="selected_sheet",
        label_visibility="collapsed",
    )

    st.sidebar.divider()

    # =========================================================================
    # OPLEIDINGEN
    # =========================================================================

    if sheet == "opleidingen":

        (
            filtered,
            selected_cluster,
            title,
        ) = select_education_data(
            opleidingen
        )

        if filtered.empty:

            st.warning(
                "Voor het gekozen cluster "
                "zijn geen opleidingen gevonden."
            )

            st.stop()

        # ---------------------------------------------------------------------
        # Studies in cluster
        # ---------------------------------------------------------------------

        st.sidebar.divider()

        show_studies_in_cluster(
            filtered
        )

        # ---------------------------------------------------------------------
        # Arbeidsmarktinformatie
        # ---------------------------------------------------------------------

        st.sidebar.divider()

        show_labour_panel(
            filtered
        )

        # ---------------------------------------------------------------------
        # Download
        # ---------------------------------------------------------------------

        st.sidebar.divider()

        download_name = (
            safe_filename(
                f"subset_opleidingen_"
                f"{selected_cluster}.xlsx"
            )
        )

        st.sidebar.download_button(
            label=(
                "Download selectie als Excel"
            ),
            data=subset_to_excel(
                filtered,
                "opleidingen",
            ),
            file_name=download_name,
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

        # ---------------------------------------------------------------------
        # Kaartdata
        # ---------------------------------------------------------------------

        map_data = aggregate_map_data(
            filtered,
            mode="opleidingen",
        )

        total_records = len(
            filtered
        )

        if total_records == 1:
            record_label = (
                "opleidingsregel"
            )
        else:
            record_label = (
                "opleidingsregels"
            )

    # =========================================================================
    # VARIANTEN
    # =========================================================================

    else:

        (
            filtered,
            selected_opleidingen,
            selected_education_cluster,
            selected_variant_cluster,
            title,
        ) = select_variant_data(
            opleidingen,
            varianten,
        )

        if filtered.empty:

            st.warning(
                "Voor het gekozen variantcluster "
                "zijn geen varianten gevonden."
            )

            st.stop()

        # ---------------------------------------------------------------------
        # Studies in opleidingscluster
        # ---------------------------------------------------------------------

        st.sidebar.divider()

        show_studies_in_cluster(
            selected_opleidingen
        )

        st.sidebar.divider()

        show_variants_in_cluster(
            filtered
        )

        # ---------------------------------------------------------------------
        # Arbeidsmarktinformatie
        # ---------------------------------------------------------------------

        st.sidebar.divider()

        show_labour_panel(
            filtered
        )

        # ---------------------------------------------------------------------
        # Download
        # ---------------------------------------------------------------------

        st.sidebar.divider()

        download_name = (
            safe_filename(
                "subset_varianten_"
                f"{selected_education_cluster}_"
                f"{selected_variant_cluster}.xlsx"
            )
        )

        st.sidebar.download_button(
            label=(
                "Download selectie als Excel"
            ),
            data=subset_to_excel(
                filtered,
                "varianten",
            ),
            file_name=download_name,
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

        # ---------------------------------------------------------------------
        # Kaartdata
        # ---------------------------------------------------------------------

        map_data = aggregate_map_data(
            filtered,
            mode="varianten",
        )

        total_records = len(
            filtered
        )

        if total_records == 1:
            record_label = "variant"
        else:
            record_label = "varianten"

    # =========================================================================
    # KAART
    # =========================================================================

    if map_data.empty:

        st.warning(
            "Voor deze selectie zijn geen "
            "locaties met coördinaten beschikbaar."
        )

        st.stop()

    figure = create_map_figure(
        map_data,
        title,
    )

    st.plotly_chart(
        figure,
        use_container_width=True,
    )

    # =========================================================================
    # ONDERSCHRIFT
    # =========================================================================

    location_count = len(
        map_data
    )

    if location_count == 1:
        location_label = "locatie"
    else:
        location_label = "locaties"

    st.caption(
        f"{total_records} {record_label} "
        f"verdeeld over "
        f"{location_count} {location_label}. "
        "Een cijfer in een marker geeft aan dat "
        "meerdere opleidingen of varianten van "
        "dezelfde instelling op deze locatie "
        "binnen de huidige selectie vallen."
    )


# =============================================================================
# START PROGRAMMA
# =============================================================================

if __name__ == "__main__":
    main()
