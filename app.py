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
# Repository file
# -------------------------------------------------------------------
BASE = Path(__file__).resolve().parent
EXCEL_FILE = BASE / "Flight Emissions Dashboard v3.xlsx"

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


# -------------------------------------------------------------------
# Load the simplified Excel dashboard
# -------------------------------------------------------------------
@st.cache_data(show_spinner="Reading flight emissions workbook...")
def load_data(workbook_mtime):
    """Load dashboard-ready flight and FTE data from the main workbook.

    All Integrated Data is the single source for flight records. It already
    contains the fixed history through 2025, calculated newer records, project
    codes, duplicate decisions, distances and Final_RFI3_tCO2e values.
    """
    all_data = pd.read_excel(
        EXCEL_FILE,
        sheet_name="All Integrated Data",
        engine="openpyxl",
    )
    fte = pd.read_excel(
        EXCEL_FILE,
        sheet_name="FTE Data",
        engine="openpyxl",
    )

    required = {
        "Traveler",
        "Date",
        "Year",
        "DepartureAirport",
        "ArrivalAirport",
        "Class",
        "Project_ID_Code",
        "Project_Description",
        "Flight_Type",
        "Team",
        "Distance_km",
        "Final_RFI3_tCO2e",
        "Include_Final",
    }
    missing = required.difference(all_data.columns)
    if missing:
        raise ValueError(
            "All Integrated Data is missing required columns: "
            f"{sorted(missing)}"
        )

    # Use the duplicate decision already calculated in the Excel workflow.
    data = all_data[
        all_data["Include_Final"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.casefold()
        .eq("yes")
    ].copy()

    data["Date"] = pd.to_datetime(data["Date"], errors="coerce")
    data["Year"] = pd.to_numeric(data["Year"], errors="coerce")
    data = data.dropna(subset=["Year"])
    data["Year"] = data["Year"].astype(int)

    data["Emissions"] = pd.to_numeric(
        data["Final_RFI3_tCO2e"],
        errors="coerce",
    ).fillna(0.0)
    data["Distance"] = pd.to_numeric(
        data["Distance_km"],
        errors="coerce",
    ).fillna(0.0)
    cabin_key = (
        data["Class"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.casefold()
        .str.replace(r"[\s_-]+", "", regex=True)
    )
    data["Cabin"] = cabin_key.replace(
        {
            "premiumeconomy": "premiumeconomy",
            "unknown": "",
            "unknowncabin": "",
        }
    )
    data["Team"] = clean_series(
        data["Team"],
        "External",
    )

    # Keep the original dashboard grouping convention.
    team_normalized = data["Team"].astype(str).str.strip().str.casefold()
    external_mask = (
        team_normalized.str.startswith("guest")
        | team_normalized.eq("wa associate")
    )
    data.loc[external_mask, "Team"] = "External"

    data["Flight Type"] = (
        data["Flight_Type"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.casefold()
        .replace(
            {
                "very_short_haul": "Very short haul",
                "short_haul": "Short haul",
                "medium_haul": "Medium haul",
                "long_haul": "Long haul",
            }
        )
    )

    # Project codes are now native fields in the Excel integrated dataset.
    data["Project Number"] = clean_series(
        data["Project_ID_Code"],
        "Unassigned",
    )
    # Project description is the travel reason carried from Traveler Manifest
    # column C into All Integrated Data.Project_Description.
    data["Project Description"] = clean_series(
        data["Project_Description"],
        "Unassigned",
    )

    data["Month"] = data["Date"].dt.month
    data["Month Name"] = data["Date"].dt.strftime("%b")

    required_fte = {"Year", "FTE"}
    missing_fte = required_fte.difference(fte.columns)
    if missing_fte:
        raise ValueError(
            "FTE Data is missing required columns: "
            f"{sorted(missing_fte)}"
        )
    fte["Year"] = pd.to_numeric(fte["Year"], errors="coerce")
    fte["FTE"] = pd.to_numeric(fte["FTE"], errors="coerce")
    fte = fte.dropna(subset=["Year"])
    fte["Year"] = fte["Year"].astype(int)

    return data, fte

# -------------------------------------------------------------------
# Verify and load the workbook
# -------------------------------------------------------------------
if not EXCEL_FILE.exists():
    st.error(f"Missing repository file: {EXCEL_FILE.name}")
    st.stop()

try:
    flights, fte = load_data(EXCEL_FILE.stat().st_mtime)
except Exception as exc:
    st.error("The dashboard could not read the flight emissions workbook.")
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
        title=f"Highest emitting projects in {selected_year}",
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
    st.info("No project assignments for the selected filters.")



# -------------------------------------------------------------------
# Dashboard chart analysis
# -------------------------------------------------------------------
st.divider()

st.subheader("Dashboard chart analysis")


# -------------------------------------------------------------------
# Monthly emissions, cabin-class emissions, and team emissions
# -------------------------------------------------------------------
chart_left, chart_middle = st.columns(
    [
        1.2,
        1,
    ]
)


# -------------------------------------------------------------------
# Emissions over time
# -------------------------------------------------------------------
with chart_left:
    monthly_emissions = (
        selected.dropna(
            subset=[
                "Date",
            ]
        )
        .groupby(
            [
                "Month",
                "Month Name",
            ],
            as_index=False,
        )
        .agg(
            Emissions=(
                "Emissions",
                "sum",
            )
        )
        .sort_values(
            "Month"
        )
    )

    month_order = [
        "Jan",
        "Feb",
        "Mar",
        "Apr",
        "May",
        "Jun",
        "Jul",
        "Aug",
        "Sep",
        "Oct",
        "Nov",
        "Dec",
    ]

    monthly_emissions["Month Name"] = pd.Categorical(
        monthly_emissions["Month Name"],
        categories=month_order,
        ordered=True,
    )

    monthly_emissions = monthly_emissions.sort_values(
        "Month Name"
    )

    monthly_figure = px.line(
        monthly_emissions,
        x="Month Name",
        y="Emissions",
        markers=True,
        title=f"Emissions over time ({selected_year})",
        labels={
            "Month Name": "Month",
            "Emissions": "Emissions (tCO₂e)",
        },
    )

    monthly_figure.update_traces(
        line=dict(
            color="#0B70C9",
            width=2,
        ),
        marker=dict(
            color="#0B70C9",
            size=7,
        ),
        hovertemplate=(
            "<b>%{x}</b><br>"
            "Emissions: %{y:.2f} tCO₂e"
            "<extra></extra>"
        ),
    )

    monthly_figure.update_xaxes(
        categoryorder="array",
        categoryarray=month_order,
        title="Month",
    )

    monthly_figure.update_yaxes(
        title="Emissions (tCO₂e)",
        rangemode="tozero",
    )

    monthly_figure.update_layout(
        height=420,
        margin=dict(
            l=20,
            r=20,
            t=60,
            b=40,
        ),
        showlegend=False,
    )

    st.plotly_chart(
        monthly_figure,
        use_container_width=True,
    )


# -------------------------------------------------------------------
# Emissions by cabin class
# -------------------------------------------------------------------
with chart_middle:
    cabin_order = [
        "economy",
        "premiumeconomy",
        "business",
        "first",
    ]

    cabin_emissions = (
        selected.groupby(
            "Cabin",
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
        )
    )

    cabin_emissions["Cabin"] = pd.Categorical(
        cabin_emissions["Cabin"],
        categories=cabin_order,
        ordered=True,
    )

    cabin_emissions = cabin_emissions.sort_values(
        "Cabin"
    )

    cabin_figure = px.bar(
        cabin_emissions,
        x="Cabin",
        y="Emissions",
        text="Emissions",
        title=f"Emissions by cabin class ({selected_year})",
        labels={
            "Cabin": "Cabin class",
            "Emissions": "Emissions (tCO₂e)",
        },
        custom_data=[
            "Flights",
        ],
    )

    cabin_figure.update_traces(
        marker_color="#0B70C9",
        texttemplate="%{text:.1f}",
        textposition="inside",
        insidetextanchor="middle",
        textfont=dict(
            color="white",
        ),
        hovertemplate=(
            "<b>%{x}</b><br>"
            "Emissions: %{y:.2f} tCO₂e<br>"
            "Flight records: %{customdata,.0f}"
            "<extra></extra>"
        ),
    )

    cabin_figure.update_xaxes(
        categoryorder="array",
        categoryarray=cabin_order,
        title="Cabin class",
        tickangle=-25,
    )

    cabin_figure.update_yaxes(
        title="Emissions (tCO₂e)",
        rangemode="tozero",
    )

    cabin_figure.update_layout(
        height=420,
        margin=dict(
            l=20,
            r=20,
            t=60,
            b=70,
        ),
        showlegend=False,
    )

    st.plotly_chart(
        cabin_figure,
        use_container_width=True,
    )


# -------------------------------------------------------------------
# Cabin-class contribution within teams
# -------------------------------------------------------------------
team_cabin_emissions = (
    selected.groupby(
        [
            "Team",
            "Cabin",
        ],
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
    )
)

team_totals = (
    team_cabin_emissions.groupby(
        "Team",
        as_index=False,
    )
    .agg(
        Total_Emissions=(
            "Emissions",
            "sum",
        )
    )
    .sort_values(
        "Total_Emissions",
        ascending=False,
    )
)

stacked_team_order = team_totals[
    "Team"
].tolist()

cabin_colors = {
    "economy": "#5AAAE6",
    "premiumeconomy": "#F2B84B",
    "business": "#E83B3B",
    "first": "#8259C8",
}

team_cabin_figure = px.bar(
    team_cabin_emissions,
    x="Team",
    y="Emissions",
    color="Cabin",
    title=(
        "Cabin class contribution within teams "
        f"({selected_year})"
    ),
    labels={
        "Team": "Teams",
        "Emissions": "Emissions (tCO₂e)",
        "Cabin": "Cabin class",
    },
    category_orders={
        "Team": stacked_team_order,
        "Cabin": cabin_order,
    },
    color_discrete_map=cabin_colors,
    custom_data=[
        "Flights",
    ],
)

team_cabin_figure.update_traces(
    hovertemplate=(
        "<b>%{x}</b><br>"
        "Cabin class: %{fullData.name}<br>"
        "Emissions: %{y:.2f} tCO₂e<br>"
        "Flight records: %{customdata,.0f}"
        "<extra></extra>"
    ),
)

team_cabin_figure.update_xaxes(
    categoryorder="array",
    categoryarray=stacked_team_order,
    title="Teams",
    tickangle=-35,
)

team_cabin_figure.update_yaxes(
    title="Emissions (tCO₂e)",
    rangemode="tozero",
)

team_cabin_figure.update_layout(
    barmode="stack",
    height=500,
    margin=dict(
        l=20,
        r=20,
        t=60,
        b=130,
    ),
    legend=dict(
        title="Cabin class",
        orientation="v",
        yanchor="top",
        y=1,
        xanchor="left",
        x=1.01,
    ),
)

st.plotly_chart(
    team_cabin_figure,
    use_container_width=True,
)

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

# -------------------------------------------------------------------
# Detailed summaries
# -------------------------------------------------------------------
st.divider()
st.subheader("Detailed summaries")


# -------------------------------------------------------------------
# Prepare project summary
# -------------------------------------------------------------------
project_detail_summary = (
    selected[
        selected["Project Number"] != "Unassigned"
    ]
    .groupby(
        [
            "Project Number",
            "Project Description",
        ],
        as_index=False,
    )
    .agg(
        Flight_records=(
            "Emissions",
            "size",
        ),
        Emissions_tCO2e=(
            "Emissions",
            "sum",
        ),
        Distance_km=(
            "Distance",
            "sum",
        ),
    )
    .sort_values(
        "Emissions_tCO2e",
        ascending=False,
    )
)

total_assigned_project_emissions = (
    project_detail_summary[
        "Emissions_tCO2e"
    ].sum()
)

if total_assigned_project_emissions > 0:
    project_detail_summary[
        "Share_of_assigned_emissions"
    ] = (
        project_detail_summary[
            "Emissions_tCO2e"
        ]
        / total_assigned_project_emissions
    )
else:
    project_detail_summary[
        "Share_of_assigned_emissions"
    ] = 0.0


# -------------------------------------------------------------------
# Prepare team summary
# -------------------------------------------------------------------
team_detail_summary = (
    selected.groupby(
        "Team",
        as_index=False,
    )
    .agg(
        Flight_records=(
            "Emissions",
            "size",
        ),
        Emissions_tCO2e=(
            "Emissions",
            "sum",
        ),
        Distance_km=(
            "Distance",
            "sum",
        ),
    )
    .sort_values(
        "Emissions_tCO2e",
        ascending=False,
    )
)

total_team_emissions = (
    team_detail_summary[
        "Emissions_tCO2e"
    ].sum()
)

if total_team_emissions > 0:
    team_detail_summary[
        "Share_of_total_emissions"
    ] = (
        team_detail_summary[
            "Emissions_tCO2e"
        ]
        / total_team_emissions
    )
else:
    team_detail_summary[
        "Share_of_total_emissions"
    ] = 0.0


# -------------------------------------------------------------------
# Prepare cabin-class summary
# -------------------------------------------------------------------
cabin_detail_summary = (
    selected.groupby(
        "Cabin",
        as_index=False,
    )
    .agg(
        Flight_records=(
            "Emissions",
            "size",
        ),
        Emissions_tCO2e=(
            "Emissions",
            "sum",
        ),
        Distance_km=(
            "Distance",
            "sum",
        ),
    )
    .sort_values(
        "Emissions_tCO2e",
        ascending=False,
    )
)

cabin_detail_summary[
    "Cabin class"
] = (
    cabin_detail_summary["Cabin"]
    .map(
        {
            "economy": "Economy",
            "premiumeconomy": "Premium economy",
            "business": "Business",
            "first": "First",
        }
    )
    .fillna(
        cabin_detail_summary["Cabin"]
        .astype(str)
        .str.replace(
            "_",
            " ",
            regex=False,
        )
        .str.title()
    )
)

total_cabin_emissions = (
    cabin_detail_summary[
        "Emissions_tCO2e"
    ].sum()
)

if total_cabin_emissions > 0:
    cabin_detail_summary[
        "Share_of_total_emissions"
    ] = (
        cabin_detail_summary[
            "Emissions_tCO2e"
        ]
        / total_cabin_emissions
    )
else:
    cabin_detail_summary[
        "Share_of_total_emissions"
    ] = 0.0

cabin_detail_summary = (
    cabin_detail_summary[
        [
            "Cabin class",
            "Flight_records",
            "Emissions_tCO2e",
            "Distance_km",
            "Share_of_total_emissions",
        ]
    ]
)


# -------------------------------------------------------------------
# Prepare detailed flight-record table
# -------------------------------------------------------------------
flight_detail_columns = [
    "Date",
    "Traveler",
    "DepartureAirport",
    "ArrivalAirport",
    "Cabin",
    "Flight Type",
    "Team",
    "Project Number",
    "Project Description",
    "Distance",
    "Emissions",
]

# Keep only columns that exist in the current integrated dataset.
available_flight_detail_columns = [
    column
    for column in flight_detail_columns
    if column in selected.columns
]

flight_detail_table = selected[
    available_flight_detail_columns
].copy()

if "Date" in flight_detail_table.columns:
    flight_detail_table["Date"] = pd.to_datetime(
        flight_detail_table["Date"],
        errors="coerce",
    ).dt.date

if "Cabin" in flight_detail_table.columns:
    flight_detail_table["Cabin"] = (
        flight_detail_table["Cabin"]
        .map(
            {
                "economy": "Economy",
                "premiumeconomy": "Premium economy",
                "business": "Business",
                "first": "First",
                }
        )
        .fillna(
            flight_detail_table["Cabin"]
            .astype(str)
            .str.replace(
                "_",
                " ",
                regex=False,
            )
            .str.title()
        )
    )

if "Distance" in flight_detail_table.columns:
    flight_detail_table["Distance"] = (
        pd.to_numeric(
            flight_detail_table["Distance"],
            errors="coerce",
        )
        .round(0)
    )

if "Emissions" in flight_detail_table.columns:
    flight_detail_table["Emissions"] = (
        pd.to_numeric(
            flight_detail_table["Emissions"],
            errors="coerce",
        )
        .round(3)
    )

if "Date" in flight_detail_table.columns:
    flight_detail_table = (
        flight_detail_table.sort_values(
            "Date",
            ascending=False,
        )
    )


# -------------------------------------------------------------------
# Display detailed summary tabs
# -------------------------------------------------------------------
(
    project_tab,
    team_tab,
    cabin_tab,
    flight_tab,
) = st.tabs(
    [
        "Projects",
        "Teams",
        "Cabin classes",
        "Flight records",
    ]
)


# -------------------------------------------------------------------
# Project summary table
# -------------------------------------------------------------------
with project_tab:
    st.markdown(
        f"**Validated project assignments for {selected_year}**"
    )

    if len(project_detail_summary) > 0:
        st.dataframe(
            project_detail_summary,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Project Number": st.column_config.TextColumn(
                    "Project number",
                    width="small",
                ),
                "Project Description": st.column_config.TextColumn(
                    "Project description",
                    width="large",
                ),
                "Flight_records": st.column_config.NumberColumn(
                    "Flight records",
                    format="%d",
                ),
                "Emissions_tCO2e": st.column_config.NumberColumn(
                    "Emissions (tCO₂e)",
                    format="%.2f",
                ),
                "Distance_km": st.column_config.NumberColumn(
                    "Distance (km)",
                    format="%.0f",
                ),
                "Share_of_assigned_emissions": (
                    st.column_config.ProgressColumn(
                        "Share of assigned emissions",
                        min_value=0.0,
                        max_value=1.0,
                        format="%.1%%",
                    )
                ),
            },
        )
    else:
        st.info(
            "No validated project assignments are available "
            "for the selected year and filters."
        )


# -------------------------------------------------------------------
# Team summary table
# -------------------------------------------------------------------
with team_tab:
    st.markdown(
        f"**Team-level summary for {selected_year}**"
    )

    st.dataframe(
        team_detail_summary,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Team": st.column_config.TextColumn(
                "Team",
                width="large",
            ),
            "Flight_records": st.column_config.NumberColumn(
                "Flight records",
                format="%d",
            ),
            "Emissions_tCO2e": st.column_config.NumberColumn(
                "Emissions (tCO₂e)",
                format="%.2f",
            ),
            "Distance_km": st.column_config.NumberColumn(
                "Distance (km)",
                format="%.0f",
            ),
            "Share_of_total_emissions": (
                st.column_config.ProgressColumn(
                    "Share of total emissions",
                    min_value=0.0,
                    max_value=1.0,
                    format="%.1%%",
                )
            ),
        },
    )


# -------------------------------------------------------------------
# Cabin-class summary table
# -------------------------------------------------------------------
with cabin_tab:
    st.markdown(
        f"**Cabin-class summary for {selected_year}**"
    )

    st.dataframe(
        cabin_detail_summary,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Cabin class": st.column_config.TextColumn(
                "Cabin class",
                width="medium",
            ),
            "Flight_records": st.column_config.NumberColumn(
                "Flight records",
                format="%d",
            ),
            "Emissions_tCO2e": st.column_config.NumberColumn(
                "Emissions (tCO₂e)",
                format="%.2f",
            ),
            "Distance_km": st.column_config.NumberColumn(
                "Distance (km)",
                format="%.0f",
            ),
            "Share_of_total_emissions": (
                st.column_config.ProgressColumn(
                    "Share of total emissions",
                    min_value=0.0,
                    max_value=1.0,
                    format="%.1%%",
                )
            ),
        },
    )


# -------------------------------------------------------------------
# Individual flight-record table
# -------------------------------------------------------------------
with flight_tab:
    st.markdown(
        f"**Included flight records for {selected_year}**"
    )

    st.caption(
        "This table reflects the selected year, cabin-class filter, "
        "team filter, duplicate exclusions, and Include_Final selection."
    )

    flight_column_config = {}

    if "Date" in flight_detail_table.columns:
        flight_column_config[
            "Date"
        ] = st.column_config.DateColumn(
            "Date",
            format="YYYY-MM-DD",
        )

    if "Traveler" in flight_detail_table.columns:
        flight_column_config[
            "Traveler"
        ] = st.column_config.TextColumn(
            "Traveler",
            width="medium",
        )

    if "DepartureAirport" in flight_detail_table.columns:
        flight_column_config[
            "DepartureAirport"
        ] = st.column_config.TextColumn(
            "Departure",
            width="small",
        )

    if "ArrivalAirport" in flight_detail_table.columns:
        flight_column_config[
            "ArrivalAirport"
        ] = st.column_config.TextColumn(
            "Arrival",
            width="small",
        )

    if "Cabin" in flight_detail_table.columns:
        flight_column_config[
            "Cabin"
        ] = st.column_config.TextColumn(
            "Cabin class",
            width="medium",
        )

    if "Flight Type" in flight_detail_table.columns:
        flight_column_config[
            "Flight Type"
        ] = st.column_config.TextColumn(
            "Flight distance",
            width="medium",
        )

    if "Team" in flight_detail_table.columns:
        flight_column_config[
            "Team"
        ] = st.column_config.TextColumn(
            "Team",
            width="medium",
        )

    if "Project Number" in flight_detail_table.columns:
        flight_column_config[
            "Project Number"
        ] = st.column_config.TextColumn(
            "Project number",
            width="small",
        )

    if "Project Description" in flight_detail_table.columns:
        flight_column_config[
            "Project Description"
        ] = st.column_config.TextColumn(
            "Project description",
            width="large",
        )

    if "Distance" in flight_detail_table.columns:
        flight_column_config[
            "Distance"
        ] = st.column_config.NumberColumn(
            "Distance (km)",
            format="%.0f",
        )

    if "Emissions" in flight_detail_table.columns:
        flight_column_config[
            "Emissions"
        ] = st.column_config.NumberColumn(
            "Emissions (tCO₂e)",
            format="%.3f",
        )

    st.dataframe(
        flight_detail_table,
        use_container_width=True,
        hide_index=True,
        column_config=flight_column_config,
        height=600,
    )