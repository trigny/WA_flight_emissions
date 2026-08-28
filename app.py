from pathlib import Path
import re
import unicodedata
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


st.set_page_config(
    page_title="WA Flight Emission Dashboard",
    page_icon="📊",
    layout="wide",
)


# -------------------------------------------------------------------
# Repository files
# -------------------------------------------------------------------
BASE = Path(__file__).resolve().parent

EXCEL_FILE = BASE / "Flight Emissions Dashboard.xlsx"
CUSTOM_FIELDS_FILE = BASE / "Custom_Fields_2026-06.xlsx"

# Project numbers valid through 2025.
PROJECT_OPTIONS_FILE = BASE / "Custom field options.csv"

# Project numbers valid from 2026 onward.
PROJECT_OPTIONS_2026_FILE = BASE / "Custom field options_2026.xlsx"


# -------------------------------------------------------------------
# Target pathway
# -------------------------------------------------------------------
TARGET_BASE_YEAR = 2024
TARGET_YEAR = 2030
TARGET_REDUCTION = 0.50


# -------------------------------------------------------------------
# General helper functions
# -------------------------------------------------------------------
def clean_key(value):
    """Convert a source value into a clean matching key."""
    if pd.isna(value):
        return ""

    value = str(value).strip()

    if value.lower() in {
        "",
        "0",
        "0.0",
        "nan",
        "none",
        "<na>",
    }:
        return ""

    return value


def clean_series(series, blank="Unassigned"):
    """Clean a pandas Series and replace empty values with a label."""
    result = series.fillna("").astype(str).str.strip()

    result = result.replace(
        {
            "0": "",
            "0.0": "",
            "nan": "",
            "None": "",
            "<NA>": "",
        }
    )

    return result.mask(
        result.eq(""),
        blank,
    )

def normalize_match_text(value):
    """
    Normalize a name or other text for cross-source matching.

    Accents, capitalization, spaces, punctuation, and hyphens are removed.
    For example, 'Joërg Leumann' and 'Joerg Leumann' become comparable.
    """
    value = clean_key(value)

    if not value:
        return ""

    value = unicodedata.normalize(
        "NFKD",
        value,
    )

    value = value.encode(
        "ascii",
        "ignore",
    ).decode(
        "ascii"
    )

    return re.sub(
        r"[^a-z0-9]",
        "",
        value.lower(),
    )


def normalize_match_date(value):
    """
    Convert a date or datetime into a consistent YYYY-MM-DD value.
    """
    value = pd.to_datetime(
        value,
        errors="coerce",
    )

    if pd.isna(value):
        return ""

    return value.strftime(
        "%Y-%m-%d"
    )


def normalize_match_cabin(value):
    """
    Normalize cabin terminology between Traveler Manifest and legacy data.
    """
    value = normalize_match_text(
        value
    )

    if "business" in value:
        return "business"

    if "first" in value:
        return "first"

    if "premium" in value:
        return "premiumeconomy"

    if "economy" in value:
        return "economy"

    return value



def unique_map(frame, key, value):
    """
    Create a mapping only when one identifier corresponds to exactly
    one distinct project value.
    """
    data = frame[
        [
            key,
            value,
        ]
    ].copy()

    data[key] = data[key].map(clean_key)
    data[value] = data[value].map(clean_key)

    data = data[
        (data[key] != "")
        & (data[value] != "")
    ]

    grouped = (
        data.groupby(key)[value]
        .agg(
            lambda values: sorted(
                set(values)
            )
        )
    )

    return {
        key_value: values[0]
        for key_value, values in grouped.items()
        if len(values) == 1
    }


def canonical_project(value, valid_codes):
    """
    Validate a project number against the applicable project list.

    If a value contains an SP suffix, such as 2.1.4-SP1, the parent
    project number is accepted only when the parent exists in the
    applicable project list.
    """
    code = clean_key(value)

    if code in valid_codes:
        return code

    parent = re.sub(
        r"-SP\d+$",
        "",
        code,
        flags=re.IGNORECASE,
    )

    if parent in valid_codes:
        return parent

    return ""


def target_for_year(year, base_value):
    """
    Calculate the annual emissions-per-FTE target.

    The pathway starts from the 2024 actual value and declines
    linearly to a 50 percent reduction in 2030.
    """
    if base_value is None or pd.isna(base_value):
        return None

    if year <= TARGET_BASE_YEAR:
        return float(base_value)

    target_2030 = float(base_value) * (
        1.0 - TARGET_REDUCTION
    )

    if year >= TARGET_YEAR:
        return target_2030

    fraction = (
        year - TARGET_BASE_YEAR
    ) / (
        TARGET_YEAR - TARGET_BASE_YEAR
    )

    return float(base_value) + (
        target_2030 - float(base_value)
    ) * fraction


def prepare_project_options(
    option_data,
    option_file,
):
    """
    Validate and standardize one project-options dataset.
    """
    option_data = option_data.copy()

    option_data.columns = (
        option_data.columns
        .astype(str)
        .str.strip()
    )

    required_columns = {
        "Name",
        "Description",
    }

    missing_columns = required_columns.difference(
        option_data.columns
    )

    if missing_columns:
        raise ValueError(
            f"{option_file.name} is missing required columns: "
            f"{sorted(missing_columns)}"
        )

    option_data["Name"] = (
        option_data["Name"]
        .map(clean_key)
    )

    option_data["Description"] = (
        option_data["Description"]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    option_data = (
        option_data[
            option_data["Name"] != ""
        ]
        .drop_duplicates(
            subset="Name",
            keep="first",
        )
        .copy()
    )

    return option_data


# -------------------------------------------------------------------
# Load and integrate the repository data
# -------------------------------------------------------------------
@st.cache_data(
    show_spinner="Reading repository files..."
)
def load_data(
    workbook_mtime,
    custom_mtime,
    options_mtime,
    options_2026_mtime,
):
    """
    Load the source files and assign a validated project number
    according to the year of each individual flight record.

    The timestamp parameters are used by Streamlit to invalidate
    the cache when any source file changes.
    """

    # ---------------------------------------------------------------
    # Load the source tables
    # ---------------------------------------------------------------
    all_data = pd.read_excel(
        EXCEL_FILE,
        sheet_name="All Integrated Data",
        engine="openpyxl",
    )

    traveler = pd.read_excel(
        EXCEL_FILE,
        sheet_name="Traveler Manifest",
        header=8,
        engine="openpyxl",
    )

    legacy = pd.read_excel(
        EXCEL_FILE,
        sheet_name="Legacy MyClimate Import",
        engine="openpyxl",
    )

    fte = pd.read_excel(
        EXCEL_FILE,
        sheet_name="FTE Data",
        engine="openpyxl",
    )

    custom = pd.read_excel(
        CUSTOM_FIELDS_FILE,
        header=6,
        engine="openpyxl",
    )

    # ---------------------------------------------------------------
    # Load both project-option lists
    # ---------------------------------------------------------------
    options = pd.read_csv(
        PROJECT_OPTIONS_FILE,
        dtype=str,
    )

    options_2026 = pd.read_excel(
        PROJECT_OPTIONS_2026_FILE,
        dtype=str,
        engine="openpyxl",
    )

    options = prepare_project_options(
        options,
        PROJECT_OPTIONS_FILE,
    )

    options_2026 = prepare_project_options(
        options_2026,
        PROJECT_OPTIONS_2026_FILE,
    )

    # Project codes and descriptions valid through 2025.
    valid_codes = set(
        options["Name"]
    )

    descriptions = (
        options
        .set_index("Name")["Description"]
        .to_dict()
    )

    # Project codes and descriptions valid from 2026 onward.
    valid_codes_2026 = set(
        options_2026["Name"]
    )

    descriptions_2026 = (
        options_2026
        .set_index("Name")["Description"]
        .to_dict()
    )

    # ---------------------------------------------------------------
    # Validate the Custom Fields structure
    # ---------------------------------------------------------------
    required_custom = {
        "Custom Question",
        "Travel Data Answer",
        "Travel Data Transaction Key",
        "Trip ID",
        "Spotnana PNR ID",
        "Confirmation Number",
    }

    missing_custom = required_custom.difference(
        custom.columns
    )

    if missing_custom:
        raise ValueError(
            "Missing Custom Fields columns: "
            f"{sorted(missing_custom)}"
        )

    # ---------------------------------------------------------------
    # Extract raw project answers from Custom Fields
    # ---------------------------------------------------------------
    project_rows = custom[
        custom["Custom Question"]
        .fillna("")
        .astype(str)
        .str.strip()
        .eq("(UD15) Project Codes")
    ].copy()

    # Important:
    # Keep the raw project value here. Do not validate it against
    # either project list until the year of the integrated flight
    # record is known.
    project_rows["Project"] = (
        project_rows["Travel Data Answer"]
        .map(clean_key)
    )

    project_rows["TX"] = (
        project_rows[
            "Travel Data Transaction Key"
        ]
        .map(clean_key)
        .str.replace(
            r"-Q\d+$",
            "",
            regex=True,
        )
    )

    project_rows["TRIP"] = (
        project_rows["Trip ID"]
        .map(clean_key)
    )

    project_rows["PNR"] = (
        project_rows["Spotnana PNR ID"]
        .map(clean_key)
    )

    project_rows["TICKET"] = (
        project_rows["Confirmation Number"]
        .map(clean_key)
    )

    project_rows = project_rows[
        project_rows["Project"] != ""
    ].copy()

    project_maps = {
        key: unique_map(
            project_rows,
            key,
            "Project",
        )
        for key in [
            "TX",
            "TRIP",
            "PNR",
            "TICKET",
        ]
    }


    # ---------------------------------------------------------------
    # Prepare conservative legacy project fallbacks
    # ---------------------------------------------------------------

    # Keep the raw legacy project value until the integrated
    # flight record's year is known.
    #
    # The applicable project whitelist is applied later:
    #   through 2025 -> Custom field options.csv
    #   from 2026    -> Custom field options_2026.xlsx
    legacy["Resolved Project"] = (
        legacy["Projektnummer"]
        .map(clean_key)
    )

    legacy_lookup = legacy[
        legacy["Resolved Project"] != ""
    ].copy()

    # Normalize the matching fields in the legacy source.
    legacy_lookup["Match Person"] = (
        legacy_lookup["Name"]
        .map(normalize_match_text)
    )

    legacy_lookup["Match Date"] = (
        legacy_lookup["Date"]
        .map(normalize_match_date)
    )

    legacy_lookup["Match Departure"] = (
        legacy_lookup["DepartureAirport"]
        .map(clean_key)
        .str.upper()
    )

    legacy_lookup["Match Arrival"] = (
        legacy_lookup["ArrivalAirport"]
        .map(clean_key)
        .str.upper()
    )

    legacy_lookup["Match Cabin"] = (
        legacy_lookup["Class"]
        .map(normalize_match_cabin)
    )

    # Composite keys are stored as strings because the existing unique_map()
    # function cleans and compares string keys.
    legacy_lookup["Exact Match Key"] = (
        legacy_lookup["Match Person"]
        + "|"
        + legacy_lookup["Match Date"]
        + "|"
        + legacy_lookup["Match Departure"]
        + "|"
        + legacy_lookup["Match Arrival"]
        + "|"
        + legacy_lookup["Match Cabin"]
    )

    legacy_lookup["Person Date Key"] = (
        legacy_lookup["Match Person"]
        + "|"
        + legacy_lookup["Match Date"]
    )

    # Keep a fallback only when every matching legacy record agrees
    # on one raw project value.
    legacy_exact_project_map = unique_map(
        legacy_lookup,
        "Exact Match Key",
        "Resolved Project",
    )

    legacy_person_date_project_map = unique_map(
        legacy_lookup,
        "Person Date Key",
        "Resolved Project",
    )


    # ---------------------------------------------------------------
    # Prepare equivalent Traveler Manifest matching keys
    # ---------------------------------------------------------------
    traveler["Match Person"] = (
        traveler["Traveler Name"]
        .map(normalize_match_text)
    )

    traveler["Match Date"] = (
        traveler["Departure Date & Time"]
        .map(normalize_match_date)
    )

    traveler["Match Departure"] = (
        traveler["Departure Airport Code"]
        .map(clean_key)
        .str.upper()
    )

    traveler["Match Arrival"] = (
        traveler["Arrival Airport Code"]
        .map(clean_key)
        .str.upper()
    )

    traveler["Match Cabin"] = (
        traveler["Cabin"]
        .map(normalize_match_cabin)
    )

    traveler["Exact Match Key"] = (
        traveler["Match Person"]
        + "|"
        + traveler["Match Date"]
        + "|"
        + traveler["Match Departure"]
        + "|"
        + traveler["Match Arrival"]
        + "|"
        + traveler["Match Cabin"]
    )

    traveler["Person Date Key"] = (
        traveler["Match Person"]
        + "|"
        + traveler["Match Date"]
    )


    # ---------------------------------------------------------------
    # Match Traveler Manifest rows to a raw project value
    # ---------------------------------------------------------------
    def traveler_project(row):
        """
        Resolve a raw project value in descending order of reliability.

        Custom Fields identifiers are preferred. Legacy data is used only
        when Custom Fields produces no project and the legacy key maps
        uniquely to one raw project value.
        """
        candidates = [
            # 1. Exact Custom Fields transaction key
            project_maps["TX"].get(
                clean_key(
                    row.get("Transaction Key")
                ),
                "",
            ),

            # 2. Unique Custom Fields PNR
            project_maps["PNR"].get(
                clean_key(
                    row.get("Spotnana PNR ID")
                ),
                "",
            ),

            # 3. Unique Custom Fields ticket or confirmation number
            project_maps["TICKET"].get(
                clean_key(
                    row.get("Ticket Number")
                ),
                "",
            ),

            # 4. Unique Custom Fields Trip ID
            project_maps["TRIP"].get(
                clean_key(
                    row.get("Trip ID")
                ),
                "",
            ),

            # 5. Exact legacy flight:
            # traveler + date + departure + arrival + cabin
            legacy_exact_project_map.get(
                clean_key(
                    row.get("Exact Match Key")
                ),
                "",
            ),

            # 6. Conservative legacy fallback:
            # traveler + date, only when all matching records agree
            legacy_person_date_project_map.get(
                clean_key(
                    row.get("Person Date Key")
                ),
                "",
            ),
        ]

        return next(
            (
                candidate
                for candidate in candidates
                if clean_key(candidate)
            ),
            "",
        )


    traveler["Resolved Project"] = (
        traveler.apply(
            traveler_project,
            axis=1,
        )
    )


    # ---------------------------------------------------------------
    # Transfer the raw project value to All Integrated Data
    # ---------------------------------------------------------------
    def integrated_project(row):
        try:
            source_index = int(
                float(
                    row["Calc_or_Source_Row"]
                )
            ) - 2
        except (
            ValueError,
            TypeError,
            KeyError,
        ):
            return ""

        source = clean_key(
            row.get("Record_Source")
        )

        if (
            source == "Traveler Manifest"
            and 0 <= source_index < len(traveler)
        ):
            return clean_key(
                traveler.iloc[source_index][
                    "Resolved Project"
                ]
            )

        if (
            source == "Legacy MyClimate Import"
            and 0 <= source_index < len(legacy)
        ):
            return clean_key(
                legacy.iloc[source_index][
                    "Resolved Project"
                ]
            )

        return ""

    all_data["Project Number"] = (
        all_data.apply(
            integrated_project,
            axis=1,
        )
    )

    # The year must be numeric before the applicable project list
    # can be selected.
    all_data["Year"] = pd.to_numeric(
        all_data["Year"],
        errors="coerce",
    )

    # ---------------------------------------------------------------
    # Select project resources according to each flight's year
    # ---------------------------------------------------------------
    def project_resources_for_year(year):
        """
        Return project codes and descriptions applicable to
        one individual flight year.
        """
        if (
            pd.notna(year)
            and int(year) >= 2026
        ):
            return (
                valid_codes_2026,
                descriptions_2026,
            )

        return (
            valid_codes,
            descriptions,
        )

    def validate_project_for_year(row):
        """
        Validate the matched raw project value against the
        project list applicable to the flight year.
        """
        year_codes, _ = (
            project_resources_for_year(
                row["Year"]
            )
        )

        return canonical_project(
            row["Project Number"],
            year_codes,
        )

    all_data["Project Number"] = (
        all_data.apply(
            validate_project_for_year,
            axis=1,
        )
    )

    def describe_project_for_year(row):
        """
        Retrieve the project description from the same list
        used to validate the project number.
        """
        _, year_descriptions = (
            project_resources_for_year(
                row["Year"]
            )
        )

        return year_descriptions.get(
            row["Project Number"],
            "",
        )

    all_data["Project Description"] = (
        all_data.apply(
            describe_project_for_year,
            axis=1,
        )
    )

    # ---------------------------------------------------------------
    # Prepare the final flight dataset
    # ---------------------------------------------------------------
    data = all_data[
        all_data["Include_Final"]
        .astype(str)
        .str.strip()
        .str.lower()
        .eq("yes")
    ].copy()

    data["Date"] = pd.to_datetime(
        data["Date"],
        errors="coerce",
    )

    data["Year"] = pd.to_numeric(
        data["Year"],
        errors="coerce",
    )

    data = data.dropna(
        subset=["Year"]
    )

    data["Year"] = (
        data["Year"]
        .astype(int)
    )

    data["Emissions"] = pd.to_numeric(
        data["Final_RFI3_tCO2e"],
        errors="coerce",
    ).fillna(0)

    data["Distance"] = pd.to_numeric(
        data["Distance_km"],
        errors="coerce",
    ).fillna(0)

    data["Cabin"] = (
        clean_series(
            data["Class"],
            "Unknown",
        )
        .str.lower()
    )

    data["Team"] = clean_series(
        data["Team"],
        "External",
    )

    data["Flight Type"] = (
        data["Flight_Type"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.lower()
        .replace(
            {
                "very_short_haul": (
                    "Very short haul"
                ),
                "short_haul": (
                    "Short haul"
                ),
                "medium_haul": (
                    "Medium haul"
                ),
                "long_haul": (
                    "Long haul"
                ),
            }
        )
    )

    data["Project Number"] = clean_series(
        data["Project Number"],
        "Unassigned",
    )

    data["Month"] = (
        data["Date"]
        .dt.month
    )

    data["Month Name"] = (
        data["Date"]
        .dt.strftime("%b")
    )

    # ---------------------------------------------------------------
    # Prepare FTE data
    # ---------------------------------------------------------------
    fte["Year"] = pd.to_numeric(
        fte["Year"],
        errors="coerce",
    )

    fte["FTE"] = pd.to_numeric(
        fte["FTE"],
        errors="coerce",
    )

    fte = fte.dropna(
        subset=["Year"]
    )

    fte["Year"] = (
        fte["Year"]
        .astype(int)
    )

    return data, fte


# -------------------------------------------------------------------
# Verify required repository files
# -------------------------------------------------------------------
for repository_file in [
    EXCEL_FILE,
    CUSTOM_FIELDS_FILE,
    PROJECT_OPTIONS_FILE,
    PROJECT_OPTIONS_2026_FILE,
]:
    if not repository_file.exists():
        st.error(
            "Missing repository file: "
            f"{repository_file.name}"
        )
        st.stop()


# -------------------------------------------------------------------
# Load dashboard data
# -------------------------------------------------------------------
try:
    flights, fte = load_data(
        EXCEL_FILE.stat().st_mtime,
        CUSTOM_FIELDS_FILE.stat().st_mtime,
        PROJECT_OPTIONS_FILE.stat().st_mtime,
        PROJECT_OPTIONS_2026_FILE.stat().st_mtime,
    )
except Exception as exc:
    st.error(
        "The dashboard could not integrate "
        "the repository data files."
    )
    st.exception(exc)
    st.stop()


# -------------------------------------------------------------------
# Annual data and target pathway
# -------------------------------------------------------------------
annual_all = (
    flights.groupby(
        "Year",
        as_index=False,
    )
    .agg(
        Flights=(
            "Emissions",
            "size",
        ),
        Emissions=(
            "Emissions",
            "sum",
        ),
        Distance=(
            "Distance",
            "sum",
        ),
    )
    .merge(
        fte[
            [
                "Year",
                "FTE",
            ]
        ],
        on="Year",
        how="left",
    )
)

annual_all["Emissions per FTE"] = (
    annual_all["Emissions"]
    / annual_all["FTE"]
)

base_rows = annual_all.loc[
    annual_all["Year"].eq(
        TARGET_BASE_YEAR
    ),
    "Emissions per FTE",
]

# -------------------------------------------------------------------
# Baseline and target pathway
# -------------------------------------------------------------------

# The baseline is the simple average of the annual emissions-per-FTE
# results for 2023 and 2024.
BASELINE_YEARS = [
    2023,
    2024,
]

baseline_rows = annual_all[
    annual_all["Year"].isin(
        BASELINE_YEARS
    )
].copy()

baseline_rows = baseline_rows.dropna(
    subset=[
        "Emissions per FTE",
    ]
)

if len(baseline_rows) == len(BASELINE_YEARS):
    target_base_value = (
        baseline_rows["Emissions per FTE"]
        .mean()
    )
else:
    target_base_value = None


# The target line starts in 2024 at the 2023–2024 average
# and reaches half of that baseline in 2030.
target_pathway = {
    year: target_for_year(
        year,
        target_base_value,
    )
    for year in range(
        TARGET_BASE_YEAR,
        TARGET_YEAR + 1,
    )
}

target_pathway = {
    year: target_for_year(
        year,
        target_base_value,
    )
    for year in range(
        TARGET_BASE_YEAR,
        TARGET_YEAR + 1,
    )
}


# -------------------------------------------------------------------
# Dashboard header and sidebar
# -------------------------------------------------------------------
st.title(
    "📊 Wyss Academy Flight Emissions Dashboard"
)

years = sorted(
    flights["Year"].unique()
)

with st.sidebar:
    selected_year = st.selectbox(
        "Analysis year",
        years,
        index=len(years) - 1,
    )

    cabins = sorted(
        flights["Cabin"].unique()
    )

    selected_cabins = st.multiselect(
        "Cabin class",
        cabins,
        default=cabins,
    )

    teams = sorted(
        flights["Team"].unique()
    )

    selected_teams = st.multiselect(
        "Teams",
        teams,
        default=teams,
    )

    if st.button(
        "Clear cache and reload data"
    ):
        st.cache_data.clear()
        st.rerun()


# -------------------------------------------------------------------
# Apply dashboard filters
# -------------------------------------------------------------------
filtered = flights[
    flights["Cabin"].isin(
        selected_cabins
    )
    & flights["Team"].isin(
        selected_teams
    )
]

selected = filtered[
    filtered["Year"].eq(
        selected_year
    )
]

fte_map = (
    fte.dropna(
        subset=["FTE"]
    )
    .drop_duplicates(
        "Year",
        keep="last",
    )
    .set_index(
        "Year"
    )["FTE"]
    .to_dict()
)

selected_fte = fte_map.get(
    selected_year
)

selected_emissions = (
    selected["Emissions"]
    .sum()
)


# -------------------------------------------------------------------
# Main dashboard metrics
# -------------------------------------------------------------------
metrics = st.columns(4)

metrics[0].metric(
    "Flights",
    f"{len(selected):,}",
)

metrics[1].metric(
    "Emissions",
    f"{selected_emissions:.1f} tCO₂e",
)

metrics[2].metric(
    "Distance",
    f"{selected['Distance'].sum():,.0f} km",
)

metrics[3].metric(
    "Emissions per FTE",
    (
        "n/a"
        if not selected_fte
        else (
            f"{selected_emissions / selected_fte:.2f} "
            "tCO₂e/FTE"
        )
    ),
)

# -------------------------------------------------------------------
# Project planning calculator
# -------------------------------------------------------------------
st.divider()
st.subheader("Project planning estimate")
planning_years = list(range(min(years), TARGET_YEAR + 1))
planning_default = selected_year if selected_year in planning_years else TARGET_YEAR
c1, c2 = st.columns(2)
with c1:
    planning_year = st.selectbox(
        "Planning year",
        planning_years,
        index=planning_years.index(planning_default),
    )

default_fte = float(fte_map.get(planning_year, max(fte_map.values())))
with c2:
    planned_fte = st.number_input(
        "Planned project FTE",
        min_value=0.1,
        max_value=1000.0,
        value=default_fte,
        step=5.0,
        format="%.1f",
        key=f"planned_fte_{planning_year}",
        help=(
            "Defaults to FTE Data for the selected year. "
            "The plus and minus controls change the scenario by 5 FTE."
        ),
    )

valid_types = ["Very short haul", "Short haul", "Medium haul", "Long haul"]
valid_cabins = ["economy", "premiumeconomy", "business"]
reference = flights[
    flights["Flight Type"].isin(valid_types) & flights["Cabin"].isin(valid_cabins)
]
factors = reference.groupby(["Flight Type", "Cabin"])["Emissions"].agg(
    mean="mean", records="size"
)
labels = {
    "Very short haul": "Very short haul (<500 km)",
    "Short haul": "Short haul (500–1,500 km)",
    "Medium haul": "Medium haul (1,500–4,000 km)",
    "Long haul": "Long haul (>4,000 km)",
}

heading = st.columns([2.2, 1, 1, 1, 1.4])
for column, title in zip(
    heading,
    ["Flight distance", "Economy", "Premium economy", "Business", "Estimated tCO₂e"],
):
    column.markdown(f"**{title}**")

planned_emissions = 0.0
planned_segments = 0
for flight_type in valid_types:
    row = st.columns([2.2, 1, 1, 1, 1.4])
    row[0].write(labels[flight_type])
    row_total = 0.0
    for column, cabin in zip(row[1:4], valid_cabins):
        key = (flight_type, cabin)
        factor = float(factors.loc[key, "mean"]) if key in factors.index else 0.0
        records = int(factors.loc[key, "records"]) if key in factors.index else 0
        with column:
            number = st.selectbox(
                f"{labels[flight_type]} {cabin}",
                range(501),
                key=f"plan_{flight_type}_{cabin}",
                label_visibility="collapsed",
                help=(
                    f"All-years mean: {factor:.3f} tCO₂e per segment, "
                    f"based on {records} records."
                ),
            )
        planned_segments += number
        row_total += number * factor
    planned_emissions += row_total
    row[4].write(f"{row_total:.2f}")

target_rate = target_pathway.get(
    planning_year, target_for_year(planning_year, target_base_value)
)
planned_per_fte = planned_emissions / planned_fte
year_target = target_rate * planned_fte if target_rate is not None else None
remaining = year_target - planned_emissions if year_target is not None else None

kpis = st.columns(4)
kpis[0].metric("Planned one-way flights", f"{planned_segments:,}")
kpis[1].metric("Estimated project emissions", f"{planned_emissions:.2f} tCO₂e")
kpis[2].metric("Estimated emissions per FTE", f"{planned_per_fte:.2f} tCO₂e/FTE")
kpis[3].metric(
    f"{planning_year} target per FTE",
    "n/a" if target_rate is None else f"{target_rate:.2f} tCO₂e/FTE",
    delta=(
        None
        if target_rate is None
        else f"{target_rate - planned_per_fte:+.2f} tCO₂e/FTE remaining"
    ),
)

budget_columns = st.columns(2)
budget_columns[0].metric(
    f"{planning_year} emissions target",
    "n/a" if year_target is None else f"{year_target:.2f} tCO₂e",
)
budget_columns[1].metric(
    f"Remaining allowance for {planning_year}",
    "n/a" if remaining is None else f"{remaining:.2f} tCO₂e",
    delta=(
        None
        if remaining is None
        else "Within target"
        if remaining >= 0
        else f"{abs(remaining):.2f} tCO₂e over target"
    ),
    delta_color="normal" if remaining is None or remaining >= 0 else "inverse",
)

if target_rate is not None:
    st.caption(
        f"{planning_year} target calculation: {target_rate:.2f} tCO₂e/FTE × "
        f"{planned_fte:.1f} FTE = {year_target:.2f} tCO₂e. "
        "Flight factors use the unfiltered mean across all available years. "
        "The per-FTE target follows a linear pathway from the 2024 actual value "
        "to a 50% reduction in 2030."
    )

# -------------------------------------------------------------------
# Project emissions
# -------------------------------------------------------------------
st.divider()
st.subheader(f"Emissions by project number ({selected_year})")
assigned = selected[selected["Project Number"].ne("Unassigned")]
unassigned = selected[selected["Project Number"].eq("Unassigned")]
project_summary = (
    assigned.groupby(["Project Number", "Project Description"], as_index=False)
    .agg(
        Flights=("Emissions", "size"),
        Emissions=("Emissions", "sum"),
        Distance=("Distance", "sum"),
    )
    .sort_values("Emissions", ascending=False)
)

project_metrics = st.columns(3)
project_metrics[0].metric(
    "Assigned flight records",
    f"{len(assigned):,}",
    f"{len(assigned) / len(selected):.1%} of selected" if len(selected) else "0%",
)
project_metrics[1].metric(
    "Assigned emissions", f"{assigned['Emissions'].sum():.2f} tCO₂e"
)
project_metrics[2].metric(
    "Unassigned emissions", f"{unassigned['Emissions'].sum():.2f} tCO₂e"
)

if len(project_summary):
    project_chart = project_summary.head(15).sort_values("Emissions")
    project_figure = px.bar(
        project_chart,
        x="Emissions",
        y="Project Number",
        orientation="h",
        text="Emissions",
        custom_data=["Project Description", "Flights", "Distance"],
        title=f"Highest-emitting validated projects in {selected_year}",
    )
    project_figure.update_traces(
        texttemplate="%{text:.2f}",
        textposition="outside",
        hovertemplate=(
            "<b>%{y}</b><br>%{customdata[0]}<br>"
            "Emissions: %{x:.2f} tCO₂e<br>"
            "Flights: %{customdata[1]:,.0f}<br>"
            "Distance: %{customdata[2]:,.0f} km<extra></extra>"
        ),
    )
    st.plotly_chart(project_figure, use_container_width=True)
    st.dataframe(
        project_summary.round({"Emissions": 2, "Distance": 0}),
        use_container_width=True,
        hide_index=True,
    )
else:
    st.info("No validated project assignments for the selected filters.")

# -------------------------------------------------------------------
# Restored flight-distance pie charts
# -------------------------------------------------------------------
st.divider()
st.subheader(f"Flights and emissions by flight distance ({selected_year})")

flight_type_order = [
    "Very short haul (<500 km)",
    "Short haul (500–1,500 km)",
    "Medium haul (1,500–4,000 km)",
    "Long haul (>4,000 km)",
]
flight_type_display = {
    "Very short haul": "Very short haul (<500 km)",
    "Short haul": "Short haul (500–1,500 km)",
    "Medium haul": "Medium haul (1,500–4,000 km)",
    "Long haul": "Long haul (>4,000 km)",
}
flight_type_colors = {
    "Very short haul (<500 km)": "#E45745",
    "Short haul (500–1,500 km)": "#F2B84B",
    "Medium haul (1,500–4,000 km)": "#4C93C3",
    "Long haul (>4,000 km)": "#2F7D78",
}

pie_data = selected[selected["Flight Type"].isin(flight_type_display)].copy()
pie_data["Flight distance"] = pie_data["Flight Type"].map(flight_type_display)
by_flight_type = (
    pie_data.groupby("Flight distance", as_index=False, observed=True)
    .agg(Flights=("Emissions", "size"), Emissions=("Emissions", "sum"))
)
by_flight_type["Flight distance"] = pd.Categorical(
    by_flight_type["Flight distance"],
    categories=flight_type_order,
    ordered=True,
)
by_flight_type = by_flight_type.sort_values("Flight distance")

pie_left, pie_right = st.columns(2)
with pie_left:
    total_flights = int(by_flight_type["Flights"].sum())
    flights_pie = px.pie(
        by_flight_type,
        names="Flight distance",
        values="Flights",
        title=f"Flights ({total_flights:,})",
        category_orders={"Flight distance": flight_type_order},
        color="Flight distance",
        color_discrete_map=flight_type_colors,
    )
    flights_pie.update_traces(
        sort=False,
        direction="clockwise",
        textposition="outside",
        texttemplate="%{percent:.1%}<br>(%{value:,.0f})",
        hovertemplate=(
            "%{label}<br>Flights: %{value:,.0f}<br>"
            "Share: %{percent:.1%}<extra></extra>"
        ),
    )
    flights_pie.update_layout(
        height=500,
        margin=dict(l=35, r=35, t=80, b=35),
        legend_title_text="Flight distance",
        uniformtext_minsize=10,
        uniformtext_mode="hide",
    )
    st.plotly_chart(flights_pie, use_container_width=True)

with pie_right:
    total_emissions = float(by_flight_type["Emissions"].sum())
    emissions_pie = px.pie(
        by_flight_type,
        names="Flight distance",
        values="Emissions",
        title=f"Emissions ({total_emissions:,.2f} tCO₂e)",
        category_orders={"Flight distance": flight_type_order},
        color="Flight distance",
        color_discrete_map=flight_type_colors,
    )
    emissions_pie.update_traces(
        sort=False,
        direction="clockwise",
        textposition="outside",
        texttemplate="%{percent:.1%}<br>(%{value:,.2f} tCO₂e)",
        hovertemplate=(
            "%{label}<br>Emissions: %{value:,.2f} tCO₂e<br>"
            "Share: %{percent:.1%}<extra></extra>"
        ),
    )
    emissions_pie.update_layout(
        height=500,
        margin=dict(l=35, r=35, t=80, b=35),
        showlegend=False,
        uniformtext_minsize=10,
        uniformtext_mode="hide",
    )
    st.plotly_chart(emissions_pie, use_container_width=True)

# -------------------------------------------------------------------
# Annual emissions-per-FTE chart
# -------------------------------------------------------------------
st.divider()
st.subheader("Annual emissions per FTE")
annual_figure = go.Figure()
annual_view = annual_all.dropna(subset=["Emissions per FTE"])
annual_figure.add_bar(
    x=annual_view["Year"],
    y=annual_view["Emissions per FTE"],
    name="Actual",
    text=annual_view["Emissions per FTE"].round(2),
    textposition="outside",
)
pathway_years = list(target_pathway)
annual_figure.add_scatter(
    x=pathway_years,
    y=[target_pathway[year] for year in pathway_years],
    name="Target",
    mode="lines+markers+text",
    text=[f"{target_pathway[year]:.2f}" for year in pathway_years],
    textposition="top center",
    line=dict(dash="dot"),
)
annual_figure.update_yaxes(title="tCO₂e/FTE", rangemode="tozero")
st.plotly_chart(annual_figure, use_container_width=True)
