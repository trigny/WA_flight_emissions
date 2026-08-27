from pathlib import Path
import re

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="MyClimate Dashboard", page_icon="📊", layout="wide")

# -------------------------------------------------------------------
# Repository files
# -------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "Flight Emissions Dashboard.xlsx"
CUSTOM_FIELDS_FILE = BASE_DIR / "Custom_Fields_2026-06.xlsx"
PROJECT_OPTIONS_FILE = BASE_DIR / "Custom field options_new.csv"

ALL_DATA_SHEET = "All Integrated Data"
TRAVELER_SHEET = "Traveler Manifest"
LEGACY_SHEET = "Legacy MyClimate Import"
FTE_SHEET = "FTE Data"
DASHBOARD_SHEET = "Dashboard"

PROJECT_QUESTION = "(UD15) Project Codes"
PROJECT_OPTIONS_HEADER = 0
CUSTOM_FIELDS_HEADER = 6       # Excel row 7
TRAVELER_HEADER = 8            # Excel row 9
TRAVELER_FIRST_SOURCE_ROW = 2  # Calc_or_Source_Row = 2 is first Traveler Manifest data row
LEGACY_FIRST_SOURCE_ROW = 2    # Calc_or_Source_Row = 2 is first legacy data row

# Target pathway: 2024 actual emissions/FTE, declining linearly to 50% by 2030.
TARGET_BASE_YEAR = 2024
TARGET_YEAR = 2030
TARGET_REDUCTION = 0.50


# -------------------------------------------------------------------
# Helpers
# -------------------------------------------------------------------
def clean_text(series: pd.Series, blank_label: str = "Unassigned") -> pd.Series:
    cleaned = series.fillna("").astype(str).str.strip()
    cleaned = cleaned.replace({"0": "", "0.0": "", "nan": "", "None": "", "<NA>": ""})
    return cleaned.mask(cleaned.eq(""), blank_label)


def clean_key(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    return "" if text.lower() in {"", "0", "0.0", "nan", "none", "<na>"} else text


def unique_mapping(frame: pd.DataFrame, key_col: str, value_col: str) -> dict:
    """Return keys having exactly one distinct nonblank value."""
    usable = frame[[key_col, value_col]].copy()
    usable[key_col] = usable[key_col].map(clean_key)
    usable[value_col] = usable[value_col].map(clean_key)
    usable = usable[(usable[key_col] != "") & (usable[value_col] != "")]
    grouped = usable.groupby(key_col)[value_col].agg(lambda values: sorted(set(values)))
    return {key: values[0] for key, values in grouped.items() if len(values) == 1}


def canonical_project_code(raw_code, valid_codes: set[str]) -> str:
    """Validate a project code against the authoritative options list.

    Spotnana sometimes exports work-package suffixes such as 2.1.4-SP1.
    If the full code is not in the options list, the suffix is removed and the
    parent code is accepted only when that parent exists in the options list.
    Operational/cost-centre codes such as LIM5000 therefore remain unassigned.
    """
    code = clean_key(raw_code)
    if not code:
        return ""
    if code in valid_codes:
        return code
    parent = re.sub(r"-SP\d+$", "", code, flags=re.IGNORECASE)
    return parent if parent in valid_codes else ""


def target_for_year(year: int, base_value: float | None) -> float | None:
    if base_value is None or pd.isna(base_value):
        return None
    if year <= TARGET_BASE_YEAR:
        return float(base_value)
    target_value = float(base_value) * (1.0 - TARGET_REDUCTION)
    if year >= TARGET_YEAR:
        return target_value
    fraction = (year - TARGET_BASE_YEAR) / (TARGET_YEAR - TARGET_BASE_YEAR)
    return float(base_value) + (target_value - float(base_value)) * fraction


def format_annual_table(annual: pd.DataFrame) -> pd.DataFrame:
    table = annual.copy()
    table["Year"] = table["Year"].astype(int).astype(str)
    table["Flights"] = table["Flights"].fillna(0).astype(int)
    table["Emissions_tCO2e"] = table["Emissions_tCO2e"].round(2)
    table["Distance_km"] = table["Distance_km"].round(0)
    table["FTE"] = table["FTE"].round(1)
    table["Relative_tCO2e_per_FTE"] = table["Relative_tCO2e_per_FTE"].round(2)
    return table


def make_annual_summary(data: pd.DataFrame, fte_data: pd.DataFrame) -> pd.DataFrame:
    annual = (
        data.groupby("Year", as_index=False)
        .agg(
            Flights=("Emissions_tCO2e", "size"),
            Emissions_tCO2e=("Emissions_tCO2e", "sum"),
            Distance_km=("Distance_km", "sum"),
        )
        .merge(fte_data[["Year", "FTE"]], on="Year", how="left")
    )
    annual["Relative_tCO2e_per_FTE"] = annual["Emissions_tCO2e"] / annual["FTE"]
    return annual.sort_values("Year")


@st.cache_data(show_spinner="Reading repository data files...")
def load_data(workbook_mtime: float, custom_mtime: float, options_mtime: float):
    # Main emissions records
    all_data = pd.read_excel(EXCEL_FILE, sheet_name=ALL_DATA_SHEET, engine="openpyxl")
    traveler = pd.read_excel(
        EXCEL_FILE, sheet_name=TRAVELER_SHEET, header=TRAVELER_HEADER, engine="openpyxl"
    )
    legacy = pd.read_excel(EXCEL_FILE, sheet_name=LEGACY_SHEET, engine="openpyxl")
    fte = pd.read_excel(EXCEL_FILE, sheet_name=FTE_SHEET, engine="openpyxl")

    # Separate project sources
    custom = pd.read_excel(
        CUSTOM_FIELDS_FILE, header=CUSTOM_FIELDS_HEADER, engine="openpyxl"
    )
    options = pd.read_csv(PROJECT_OPTIONS_FILE, dtype=str)
    options.columns = options.columns.astype(str).str.strip()
    if not {"Name", "Description"}.issubset(options.columns):
        raise ValueError("Custom field options_new.csv must contain Name and Description columns.")
    options["Name"] = options["Name"].map(clean_key)
    options["Description"] = options["Description"].fillna("").astype(str).str.strip()
    options = options[options["Name"] != ""].drop_duplicates("Name")
    valid_codes = set(options["Name"])
    descriptions = options.set_index("Name")["Description"].to_dict()

    # Build validated Custom Fields mappings.
    required_custom = {
        "Custom Question", "Travel Data Answer", "Trip ID", "Spotnana PNR ID",
        "Confirmation Number", "Travel Data Transaction Key"
    }
    missing_custom = sorted(required_custom.difference(custom.columns))
    if missing_custom:
        raise ValueError(f"Missing Custom Fields columns: {missing_custom}")

    project_rows = custom[
        custom["Custom Question"].fillna("").astype(str).str.strip().eq(PROJECT_QUESTION)
    ].copy()
    project_rows["Project Number"] = project_rows["Travel Data Answer"].map(
        lambda value: canonical_project_code(value, valid_codes)
    )
    project_rows["Transaction Key"] = (
        project_rows["Travel Data Transaction Key"].map(clean_key)
        .str.replace(r"-Q\d+$", "", regex=True)
    )
    project_rows["Trip ID key"] = project_rows["Trip ID"].map(clean_key)
    project_rows["PNR key"] = project_rows["Spotnana PNR ID"].map(clean_key)
    project_rows["Ticket key"] = project_rows["Confirmation Number"].map(clean_key)
    project_rows = project_rows[project_rows["Project Number"] != ""].copy()

    custom_maps = {
        "transaction": unique_mapping(project_rows, "Transaction Key", "Project Number"),
        "pnr": unique_mapping(project_rows, "PNR key", "Project Number"),
        "ticket": unique_mapping(project_rows, "Ticket key", "Project Number"),
        "trip": unique_mapping(project_rows, "Trip ID key", "Project Number"),
    }

    required_traveler = {
        "Transaction Key", "Spotnana PNR ID", "Ticket Number", "Trip ID"
    }
    missing_traveler = sorted(required_traveler.difference(traveler.columns))
    if missing_traveler:
        raise ValueError(f"Missing Traveler Manifest columns: {missing_traveler}")

    def resolve_traveler_project(row) -> str:
        candidates = (
            custom_maps["transaction"].get(clean_key(row.get("Transaction Key")), ""),
            custom_maps["pnr"].get(clean_key(row.get("Spotnana PNR ID")), ""),
            custom_maps["ticket"].get(clean_key(row.get("Ticket Number")), ""),
            custom_maps["trip"].get(clean_key(row.get("Trip ID")), ""),
        )
        return next((candidate for candidate in candidates if candidate), "")

    traveler["Resolved Project Number"] = traveler.apply(resolve_traveler_project, axis=1)
    legacy["Resolved Project Number"] = legacy["Projektnummer"].map(
        lambda value: canonical_project_code(value, valid_codes)
    )

    # Attach a project number to each integrated row in Python, without Excel column L.
    def integrated_project(row) -> str:
        source = clean_key(row.get("Record_Source"))
        try:
            source_row = int(float(row.get("Calc_or_Source_Row")))
        except (TypeError, ValueError):
            return ""
        if source == "Traveler Manifest":
            pos = source_row - TRAVELER_FIRST_SOURCE_ROW
            if 0 <= pos < len(traveler):
                return clean_key(traveler.iloc[pos]["Resolved Project Number"])
        if source == "Legacy MyClimate Import":
            pos = source_row - LEGACY_FIRST_SOURCE_ROW
            if 0 <= pos < len(legacy):
                return clean_key(legacy.iloc[pos]["Resolved Project Number"])
        return ""

    all_data["Project Number"] = all_data.apply(integrated_project, axis=1)
    all_data["Project Description"] = all_data["Project Number"].map(descriptions).fillna("")

    # Support current and older workbook headers.
    cabin_col = "Cabin Class" if "Cabin Class" in all_data.columns else "Class"
    team_col = "Team" if "Team" in all_data.columns else None
    required = [
        "Date", "Year", "DepartureAirport", "ArrivalAirport", cabin_col,
        "Flight_Type", "Distance_km", "Final_RFI3_tCO2e", "Include_Final",
        "Project Number", "Project Description",
    ]
    missing = [column for column in required if column not in all_data.columns]
    if missing:
        raise ValueError(f"Missing expected columns in '{ALL_DATA_SHEET}': {missing}")

    keep = required + ([team_col] if team_col else [])
    flights = all_data[keep].copy()
    flights = flights[
        flights["Include_Final"].astype(str).str.strip().str.lower().eq("yes")
    ].copy()
    flights["Date"] = pd.to_datetime(flights["Date"], errors="coerce")
    flights["Year"] = pd.to_numeric(flights["Year"], errors="coerce")
    flights = flights.dropna(subset=["Year"])
    flights["Year"] = flights["Year"].astype(int)
    flights["DepartureAirport"] = clean_text(flights["DepartureAirport"], "UNK").str.upper()
    flights["ArrivalAirport"] = clean_text(flights["ArrivalAirport"], "UNK").str.upper()
    flights["Route"] = flights["DepartureAirport"] + " → " + flights["ArrivalAirport"]
    flights["Cabin Class"] = clean_text(flights[cabin_col], "Unknown").str.lower()
    flights["Teams"] = clean_text(flights[team_col], "External") if team_col else "External"
    flights["Flight Type"] = (
        flights["Flight_Type"].fillna("").astype(str).str.strip().str.lower()
        .replace({
            "very_short_haul": "Very short haul", "short_haul": "Short haul",
            "medium_haul": "Medium haul", "long_haul": "Long haul",
        })
    )
    flights["Distance_km"] = pd.to_numeric(flights["Distance_km"], errors="coerce").fillna(0)
    flights["Emissions_tCO2e"] = pd.to_numeric(
        flights["Final_RFI3_tCO2e"], errors="coerce"
    ).fillna(0)
    flights["Project Number"] = clean_text(flights["Project Number"], "Unassigned")
    flights["Project Description"] = flights["Project Description"].fillna("").astype(str).str.strip()
    flights.loc[flights["Project Number"].eq("Unassigned"), "Project Description"] = "No validated project match"
    flights["Project"] = flights.apply(
        lambda row: (
            "Unassigned" if row["Project Number"] == "Unassigned"
            else f'{row["Project Number"]} · {row["Project Description"]}'
        ), axis=1
    )
    flights["Month"] = flights["Date"].dt.month
    flights["Month_name"] = flights["Date"].dt.strftime("%b")

    fte = fte[["Year", "Employees", "FTE"]].copy()
    fte["Year"] = pd.to_numeric(fte["Year"], errors="coerce")
    fte = fte.dropna(subset=["Year"])
    fte["Year"] = fte["Year"].astype(int)
    fte["Employees"] = pd.to_numeric(fte["Employees"], errors="coerce")
    fte["FTE"] = pd.to_numeric(fte["FTE"], errors="coerce")

    try:
        dash = pd.read_excel(EXCEL_FILE, sheet_name=DASHBOARD_SHEET, header=None, engine="openpyxl")
        default_year = int(dash.iloc[7, 1]) if pd.notna(dash.iloc[7, 1]) else None
        rfi_factor = dash.iloc[9, 1] if pd.notna(dash.iloc[9, 1]) else 3
    except Exception:
        default_year, rfi_factor = None, 3

    diagnostics = {
        "custom_project_rows": len(project_rows),
        "validated_projects": len(valid_codes),
        "assigned_flights": int(flights["Project Number"].ne("Unassigned").sum()),
        "total_flights": len(flights),
    }
    return flights, fte, default_year, rfi_factor, diagnostics


# -------------------------------------------------------------------
# Load data
# -------------------------------------------------------------------
required_files = [EXCEL_FILE, CUSTOM_FIELDS_FILE, PROJECT_OPTIONS_FILE]
missing_files = [path.name for path in required_files if not path.exists()]
if missing_files:
    st.error("Missing repository file(s): " + ", ".join(missing_files))
    st.stop()

try:
    flights, fte, workbook_default_year, rfi_factor, diagnostics = load_data(
        EXCEL_FILE.stat().st_mtime,
        CUSTOM_FIELDS_FILE.stat().st_mtime,
        PROJECT_OPTIONS_FILE.stat().st_mtime,
    )
except Exception as exc:
    st.error("The dashboard could not read or integrate the repository data files.")
    st.exception(exc)
    st.stop()

available_years = sorted(flights["Year"].dropna().unique().tolist())
if not available_years:
    st.warning("No valid flight records found after applying Include_Final = Yes.")
    st.stop()

# Unfiltered annual data drives targets and planning factors.
annual_all = make_annual_summary(flights, fte)
base_series = annual_all.loc[
    annual_all["Year"].eq(TARGET_BASE_YEAR), "Relative_tCO2e_per_FTE"
]
target_base_value = (
    float(base_series.iloc[0]) if len(base_series) and pd.notna(base_series.iloc[0]) else None
)

# -------------------------------------------------------------------
# Header and controls
# -------------------------------------------------------------------
st.title("📊 Wyss Academy Flight Emissions Dashboard")
default_year = workbook_default_year if workbook_default_year in available_years else max(available_years)

with st.sidebar:
    st.header("Controls")
    selected_year = st.selectbox(
        "Analysis year", available_years, index=available_years.index(default_year)
    )
    baseline_years = st.multiselect(
        "Baseline years", available_years,
        default=[year for year in [2023, 2024] if year in available_years],
    )
    st.divider()
    st.subheader("Dashboard filters")
    cabin_values = sorted(flights["Cabin Class"].dropna().unique().tolist())
    selected_cabins = st.multiselect("Cabin class", cabin_values, default=cabin_values)
    team_values = sorted(flights["Teams"].dropna().unique().tolist())
    selected_teams = st.multiselect("Teams", team_values, default=team_values)
    st.divider()
    st.caption(f"Workbook: `{EXCEL_FILE.name}`")
    st.caption(f"Custom fields: `{CUSTOM_FIELDS_FILE.name}`")
    st.caption(f"Project options: `{PROJECT_OPTIONS_FILE.name}`")
    if st.button("Clear cache and reload data"):
        st.cache_data.clear()
        st.rerun()

filtered_all_years = flights[
    flights["Cabin Class"].isin(selected_cabins) & flights["Teams"].isin(selected_teams)
].copy()
selected = filtered_all_years[filtered_all_years["Year"].eq(selected_year)].copy()
baseline = (
    filtered_all_years[filtered_all_years["Year"].isin(baseline_years)].copy()
    if baseline_years else filtered_all_years.iloc[0:0].copy()
)
annual = make_annual_summary(filtered_all_years, fte)

# -------------------------------------------------------------------
# KPIs
# -------------------------------------------------------------------
selected_fte_series = annual.loc[annual["Year"].eq(selected_year), "FTE"]
selected_fte = (
    float(selected_fte_series.iloc[0])
    if len(selected_fte_series) and pd.notna(selected_fte_series.iloc[0]) else None
)
selected_emissions = selected["Emissions_tCO2e"].sum()
selected_flights = len(selected)
selected_distance = selected["Distance_km"].sum()
selected_relative = selected_emissions / selected_fte if selected_fte and selected_fte != 0 else None
baseline_abs = baseline.groupby("Year")["Emissions_tCO2e"].sum().mean() if len(baseline) else None

k1, k2, k3, k4 = st.columns(4)
k1.metric("Flights", f"{selected_flights:,.0f}")
k2.metric("Emissions", f"{selected_emissions:,.1f} tCO₂e")
k3.metric("Distance", f"{selected_distance:,.0f} km")
k4.metric("Emissions per FTE", "n/a" if selected_relative is None else f"{selected_relative:,.2f} tCO₂e/FTE")
if baseline_abs is not None:
    st.caption(
        f"Baseline absolute emissions: {baseline_abs:,.1f} tCO₂e based on "
        f"{', '.join(map(str, baseline_years))}."
    )

# -------------------------------------------------------------------
# Project planning estimate
# -------------------------------------------------------------------
st.divider()
st.subheader("Project planning estimate")

planning_years = list(range(min(available_years), TARGET_YEAR + 1))
planning_default = selected_year if selected_year in planning_years else planning_years[-1]
control_left, control_middle = st.columns([1, 1])
with control_left:
    planning_year = st.selectbox(
        "Planning year", planning_years, index=planning_years.index(planning_default),
        key="planning_year",
    )

# The FTE Data sheet is authoritative for years with an entered value:
# 2023: 76.6, 2024: 92.6, 2025: 94.7, 2026: 96.5.
# The dropdown retains these reference values, while Custom allows any
# future or hypothetical FTE scenario without changing the source workbook.
fte_reference = (
    fte.dropna(subset=["FTE"])
    .drop_duplicates("Year", keep="last")
    .set_index("Year")["FTE"]
    .astype(float)
    .to_dict()
)
fte_reference_values = list(dict.fromkeys(float(value) for value in fte_reference.values()))
default_fte = float(fte_reference.get(planning_year, fte_reference_values[-1]))
fte_choices = fte_reference_values + ["Custom"]
with control_middle:
    fte_choice = st.selectbox(
        "Planned project FTE",
        fte_choices,
        index=fte_choices.index(default_fte),
        format_func=lambda value: "Custom value" if value == "Custom" else f"{value:.1f}",
        key=f"planned_project_fte_choice_{planning_year}",
        help=(
            "Completed-year defaults come from the FTE Data sheet. "
            "Choose Custom value for a future or hypothetical scenario."
        ),
    )

if fte_choice == "Custom":
    planned_fte = st.number_input(
        "Custom planned project FTE",
        min_value=0.1,
        max_value=1000.0,
        value=default_fte,
        step=0.1,
        format="%.1f",
        key=f"planned_project_fte_custom_{planning_year}",
    )
else:
    planned_fte = float(fte_choice)

# The target is not displayed as a standalone card. It is used directly
# in the year-sensitive Target allowance metric below.
planning_target_per_fte = target_for_year(planning_year, target_base_value)

reference_year = 2025
valid_types = ["Very short haul", "Short haul", "Medium haul", "Long haul"]
valid_cabins = ["economy", "premiumeconomy", "business"]
planning_reference = flights[
    flights["Year"].eq(reference_year)
    & flights["Emissions_tCO2e"].notna()
    & flights["Flight Type"].isin(valid_types)
    & flights["Cabin Class"].isin(valid_cabins)
].copy()
factors = planning_reference.groupby(["Flight Type", "Cabin Class"])["Emissions_tCO2e"].agg(mean="mean", records="size")
labels = {
    "Very short haul": "Very short haul (<500 km)",
    "Short haul": "Short haul (500–1,500 km)",
    "Medium haul": "Medium haul (1,500–4,000 km)",
    "Long haul": "Long haul (>4,000 km)",
}
heading = st.columns([2.2, 1, 1, 1, 1.4])
for column, title in zip(heading, ["Flight distance", "Economy", "Premium economy", "Business", "Estimated tCO₂e"]):
    column.markdown(f"**{title}**")

planned_total_emissions, planned_segments = 0.0, 0
unavailable = []
for flight_type in valid_types:
    row = st.columns([2.2, 1, 1, 1, 1.4])
    row[0].write(labels[flight_type])
    row_total = 0.0
    for input_column, cabin in zip(row[1:4], valid_cabins):
        key = (flight_type, cabin)
        available = key in factors.index
        factor = float(factors.loc[key, "mean"]) if available else None
        records = int(factors.loc[key, "records"]) if available else 0
        with input_column:
            number = st.selectbox(
                f"{labels[flight_type]} {cabin}", range(501),
                key=f"plan_{flight_type}_{cabin}", label_visibility="collapsed",
                help=(
                    f"2025 mean: {factor:.3f} tCO₂e per one-way segment, based on {records} records."
                    if available else "No 2025 reference factor is available for this combination."
                ),
            )
        planned_segments += number
        if available:
            row_total += number * factor
        elif number:
            unavailable.append(f"{labels[flight_type]} {cabin}")
    planned_total_emissions += row_total
    row[4].write(f"{row_total:.2f}")

if unavailable:
    st.warning("Not included because no 2025 reference factor is available: " + ", ".join(unavailable))

planned_per_fte = planned_total_emissions / planned_fte if planned_fte else None
target_budget = planning_target_per_fte * planned_fte if planning_target_per_fte is not None else None
variance = planned_total_emissions - target_budget if target_budget is not None else None
r1, r2, r3, r4 = st.columns(4)
r1.metric("Planned one-way flights", f"{planned_segments:,}")
r2.metric("Estimated project emissions", f"{planned_total_emissions:.2f} tCO₂e")
r3.metric("Estimated emissions per FTE", "n/a" if planned_per_fte is None else f"{planned_per_fte:.2f} tCO₂e/FTE")
r4.metric(
    f"{planning_year} target allowance",
    "n/a" if target_budget is None else f"{target_budget:.2f} tCO₂e",
    delta=None if variance is None else f"{variance:+.2f} tCO₂e vs target",
    delta_color="inverse",
    help=(
        "Selected-year emissions target per FTE multiplied by the selected FTE. "
        "Both the planning year and the FTE scenario update this allowance."
    ),
)
st.caption(
    "Flight factors use the unfiltered 2025 mean for each distance and cabin combination. "
    "The per-FTE target follows a linear pathway from the 2024 actual value to a 50% reduction in 2030."
)

# -------------------------------------------------------------------
# Emissions by project
# -------------------------------------------------------------------
st.divider()
st.subheader(f"Emissions by project number ({selected_year})")
assigned = selected[selected["Project Number"].ne("Unassigned")].copy()
unassigned = selected[selected["Project Number"].eq("Unassigned")].copy()
assigned_emissions = assigned["Emissions_tCO2e"].sum()
unassigned_emissions = unassigned["Emissions_tCO2e"].sum()
coverage = len(assigned) / len(selected) if len(selected) else 0.0

p1, p2, p3 = st.columns(3)
p1.metric("Assigned flight records", f"{len(assigned):,}", f"{coverage:.1%} of selected records")
p2.metric("Assigned emissions", f"{assigned_emissions:.2f} tCO₂e")
p3.metric("Unassigned emissions", f"{unassigned_emissions:.2f} tCO₂e")

project_summary = (
    assigned.groupby(["Project Number", "Project Description"], as_index=False)
    .agg(
        Flights=("Emissions_tCO2e", "size"),
        Emissions_tCO2e=("Emissions_tCO2e", "sum"),
        Distance_km=("Distance_km", "sum"),
    )
    .sort_values("Emissions_tCO2e", ascending=False)
)
if len(project_summary):
    project_summary["Share of assigned emissions"] = (
        project_summary["Emissions_tCO2e"] / project_summary["Emissions_tCO2e"].sum()
    )
    top_limit = st.slider(
        "Projects shown in chart", min_value=5,
        max_value=max(5, min(30, len(project_summary))),
        value=min(15, max(5, len(project_summary))), step=1,
    ) if len(project_summary) >= 5 else len(project_summary)
    chart_data = project_summary.head(top_limit).sort_values("Emissions_tCO2e", ascending=True)
    fig_project = px.bar(
        chart_data,
        x="Emissions_tCO2e", y="Project Number", orientation="h",
        text="Emissions_tCO2e", custom_data=["Project Description", "Flights", "Distance_km"],
        labels={"Emissions_tCO2e": "Emissions (tCO₂e)", "Project Number": "Project number"},
        title=f"Highest-emitting validated projects in {selected_year}",
    )
    fig_project.update_traces(
        marker_color="#4F81BD", texttemplate="%{text:.2f}", textposition="outside",
        hovertemplate=(
            "<b>%{y}</b><br>%{customdata[0]}<br>"
            "Emissions: %{x:.2f} tCO₂e<br>Flights: %{customdata[1]:,.0f}<br>"
            "Distance: %{customdata[2]:,.0f} km<extra></extra>"
        ),
    )
    fig_project.update_layout(height=max(430, 30 * len(chart_data)), margin=dict(l=20, r=80, t=60, b=30))
    st.plotly_chart(fig_project, use_container_width=True)

    display_projects = project_summary.copy()
    display_projects["Emissions_tCO2e"] = display_projects["Emissions_tCO2e"].round(2)
    display_projects["Distance_km"] = display_projects["Distance_km"].round(0)
    st.dataframe(
        display_projects,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Project Number": st.column_config.TextColumn("Project number"),
            "Project Description": st.column_config.TextColumn("Project description", width="large"),
            "Flights": st.column_config.NumberColumn("Flight records", format="%d"),
            "Emissions_tCO2e": st.column_config.NumberColumn("Emissions (tCO₂e)", format="%.2f"),
            "Distance_km": st.column_config.NumberColumn("Distance (km)", format="%.0f"),
            "Share of assigned emissions": st.column_config.ProgressColumn(
                "Share of assigned emissions", min_value=0, max_value=1, format="%.1%%"
            ),
        },
    )
else:
    st.info("No flight records for the current filters could be matched to a validated project number.")

# -------------------------------------------------------------------
# Dashboard chart analysis
# -------------------------------------------------------------------
st.divider()
st.subheader("Dashboard chart analysis")
left, middle, right = st.columns((1.2, 1, 1))
with left:
    monthly = (
        selected.dropna(subset=["Date"])
        .groupby(["Month", "Month_name"], as_index=False)["Emissions_tCO2e"].sum()
        .sort_values("Month")
    )
    fig_month = px.line(monthly, x="Month_name", y="Emissions_tCO2e", markers=True,
                        title=f"Emissions over time ({selected_year})",
                        labels={"Month_name": "Month", "Emissions_tCO2e": "Emissions (tCO₂e)"})
    fig_month.update_layout(height=400, margin=dict(l=20, r=20, t=60, b=20))
    st.plotly_chart(fig_month, use_container_width=True)
with middle:
    by_cabin = selected.groupby("Cabin Class", as_index=False).agg(
        Flights=("Emissions_tCO2e", "size"), Emissions_tCO2e=("Emissions_tCO2e", "sum")
    ).sort_values("Emissions_tCO2e", ascending=False)
    fig_cabin = px.bar(by_cabin, x="Cabin Class", y="Emissions_tCO2e", text_auto=".1f",
                       title=f"Emissions by cabin class ({selected_year})",
                       labels={"Cabin Class": "Cabin class", "Emissions_tCO2e": "Emissions (tCO₂e)"})
    fig_cabin.update_layout(height=400, margin=dict(l=20, r=20, t=60, b=70), xaxis_tickangle=-25)
    st.plotly_chart(fig_cabin, use_container_width=True)
with right:
    by_team = selected.groupby("Teams", as_index=False).agg(
        Flights=("Emissions_tCO2e", "size"), Emissions_tCO2e=("Emissions_tCO2e", "sum")
    ).sort_values("Emissions_tCO2e", ascending=False)
    fig_team = px.bar(by_team, x="Teams", y="Emissions_tCO2e", text_auto=".1f",
                      title=f"Emissions by teams ({selected_year})",
                      labels={"Teams": "Teams", "Emissions_tCO2e": "Emissions (tCO₂e)"})
    fig_team.update_layout(height=400, margin=dict(l=20, r=20, t=60, b=90), xaxis_tickangle=-35)
    st.plotly_chart(fig_team, use_container_width=True)

# Annual emissions-per-FTE chart and table
left2, right2 = st.columns((1.1, 1))
with left2:
    annual_relative = annual.dropna(subset=["Relative_tCO2e_per_FTE"]).copy()
    annual_relative = annual_relative[annual_relative["FTE"].gt(0)]
    fig_annual = go.Figure()
    fig_annual.add_bar(
        x=annual_relative["Year"], y=annual_relative["Relative_tCO2e_per_FTE"],
        name="Emissions per FTE", marker_color="#4F81BD",
        text=annual_relative["Relative_tCO2e_per_FTE"].round(2), textposition="outside",
    )
    if target_base_value is not None:
        pathway_years = list(range(TARGET_BASE_YEAR, TARGET_YEAR + 1))
        pathway_values = [target_for_year(year, target_base_value) for year in pathway_years]
        fig_annual.add_scatter(
            x=pathway_years, y=pathway_values, mode="lines+markers+text",
            text=[f"{value:.2f}" for value in pathway_values], textposition="top center",
            name="Emissions per FTE target", line=dict(color="#2F5597", width=3, dash="dot"),
        )
    fig_annual.update_layout(
        title="Annual emissions per FTE", height=420,
        margin=dict(l=20, r=20, t=60, b=20),
        legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5),
        bargap=0.55,
    )
    fig_annual.update_yaxes(title="tCO₂e / FTE", rangemode="tozero")
    st.plotly_chart(fig_annual, use_container_width=True)
with right2:
    st.subheader("Annual summary table")
    st.dataframe(format_annual_table(annual), use_container_width=True, hide_index=True)

# -------------------------------------------------------------------
# Detailed summaries
# -------------------------------------------------------------------
st.divider()
st.subheader("Detailed summaries")
tab_projects, tab_teams, tab_cabin = st.tabs([
    "Projects summary", "Teams summary", "Cabin class summary"
])
with tab_projects:
    if len(project_summary):
        st.dataframe(display_projects, use_container_width=True, hide_index=True)
    else:
        st.info("No validated project assignments for the current filters.")
with tab_teams:
    teams_summary = selected.groupby("Teams", as_index=False).agg(
        Flights=("Emissions_tCO2e", "size"), Emissions_tCO2e=("Emissions_tCO2e", "sum"),
        Distance_km=("Distance_km", "sum")
    ).sort_values("Emissions_tCO2e", ascending=False)
    st.dataframe(teams_summary.round({"Emissions_tCO2e": 2, "Distance_km": 0}), use_container_width=True, hide_index=True)
with tab_cabin:
    cabin_summary = selected.groupby("Cabin Class", as_index=False).agg(
        Flights=("Emissions_tCO2e", "size"), Emissions_tCO2e=("Emissions_tCO2e", "sum"),
        Distance_km=("Distance_km", "sum")
    ).sort_values("Emissions_tCO2e", ascending=False)
    st.dataframe(cabin_summary.round({"Emissions_tCO2e": 2, "Distance_km": 0}), use_container_width=True, hide_index=True)
