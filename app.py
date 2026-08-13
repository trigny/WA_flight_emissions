from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

# ==========================================================
# Streamlit page setup
# ==========================================================
st.set_page_config(page_title="MyClimate Dashboard", page_icon="📊", layout="wide")

# Keep the original Excel workbook in the same folder as this app.
# When the workbook is updated, replace the file but keep the same filename.
EXCEL_FILE = Path(__file__).parent / "MyClimate Methodology Workbook_final (1).xlsx"

ALL_DATA_SHEET = "All Integrated Data"
FTE_SHEET = "FTE Data"
DASHBOARD_SHEET = "Dashboard"

# Required columns from All Integrated Data.
BASE_COLUMNS = [
    "Date",
    "Year",
    "DepartureAirport",
    "ArrivalAirport",
    "Class",
    "Distance_km",
    "Final_RFI3_tCO2e",
    "Include_Final",
]

# Optional granularity columns.
# If these columns are added to All Integrated Data, the dashboard will automatically use them.
OPTIONAL_DIMENSIONS = [
    "Team",
    "Equipe",
    "Équipe",
    "Project",
    "Projet",
    "Hub",
    "Location",
    "Site",
    "Programme",
    "Program",
    "Department",
    "Unit",
    "Cost_Center",
    "Cost Center",
    "Travel_Purpose",
    "Travel Purpose",
    "Funding_Source",
    "Funding Source",
]

# Nice display names for common dimensions.
DISPLAY_NAMES = {
    "Team": "Team",
    "Equipe": "Équipe",
    "Équipe": "Équipe",
    "Project": "Project",
    "Projet": "Projet",
    "Hub": "Hub",
    "Location": "Location",
    "Site": "Site",
    "Programme": "Programme",
    "Program": "Program",
    "Department": "Department",
    "Unit": "Unit",
    "Cost_Center": "Cost center",
    "Cost Center": "Cost center",
    "Travel_Purpose": "Travel purpose",
    "Travel Purpose": "Travel purpose",
    "Funding_Source": "Funding source",
    "Funding Source": "Funding source",
}


def label(col_name: str) -> str:
    """Clean label for display."""
    return DISPLAY_NAMES.get(col_name, col_name.replace("_", " "))


@st.cache_data(show_spinner="Reading Excel workbook...")
def load_from_excel(file_mtime: float):
    """Read the original workbook and keep only dashboard-relevant fields in memory."""
    all_data = pd.read_excel(EXCEL_FILE, sheet_name=ALL_DATA_SHEET, engine="openpyxl")
    fte = pd.read_excel(EXCEL_FILE, sheet_name=FTE_SHEET, engine="openpyxl")

    # Read dashboard defaults where possible.
    try:
        dash = pd.read_excel(EXCEL_FILE, sheet_name=DASHBOARD_SHEET, header=None, engine="openpyxl")
        default_year = int(dash.iloc[7, 1]) if pd.notna(dash.iloc[7, 1]) else None
        rfi_factor = dash.iloc[9, 1] if pd.notna(dash.iloc[9, 1]) else 3
    except Exception:
        default_year = None
        rfi_factor = 3

    missing = [col for col in BASE_COLUMNS if col not in all_data.columns]
    if missing:
        raise ValueError(f"Missing expected columns in '{ALL_DATA_SHEET}': {missing}")

    available_dimensions = [col for col in OPTIONAL_DIMENSIONS if col in all_data.columns]

    flights = all_data[BASE_COLUMNS + available_dimensions].copy()

    # Use only final integrated records, excluding duplicate copies according to workbook logic.
    flights = flights[flights["Include_Final"].astype(str).str.strip().str.lower().eq("yes")].copy()

    # Clean core fields.
    flights["Date"] = pd.to_datetime(flights["Date"], errors="coerce")
    flights["Year"] = pd.to_numeric(flights["Year"], errors="coerce")
    flights = flights.dropna(subset=["Year"])
    flights["Year"] = flights["Year"].astype(int)

    flights["DepartureAirport"] = flights["DepartureAirport"].fillna("UNK").astype(str).str.upper().str.strip()
    flights["ArrivalAirport"] = flights["ArrivalAirport"].fillna("UNK").astype(str).str.upper().str.strip()
    flights["Route"] = flights["DepartureAirport"] + " → " + flights["ArrivalAirport"]

    flights["Class"] = flights["Class"].fillna("Unknown").astype(str).str.lower().str.strip()
    flights["Distance_km"] = pd.to_numeric(flights["Distance_km"], errors="coerce").fillna(0)
    flights["Emissions_tCO2e"] = pd.to_numeric(flights["Final_RFI3_tCO2e"], errors="coerce").fillna(0)

    flights["Month"] = flights["Date"].dt.month
    flights["Month_name"] = flights["Date"].dt.strftime("%b")

    # Clean optional dimensions.
    for dim in available_dimensions:
        flights[dim] = flights[dim].fillna("Unassigned").astype(str).str.strip()
        flights.loc[flights[dim].eq(""), dim] = "Unassigned"

    # Keep only relevant fields in Python after reading from the full original workbook.
    keep_cols = [
        "Date",
        "Year",
        "DepartureAirport",
        "ArrivalAirport",
        "Route",
        "Class",
        "Distance_km",
        "Emissions_tCO2e",
        "Month",
        "Month_name",
    ] + available_dimensions
    flights = flights[keep_cols]

    fte = fte[["Year", "Employees", "FTE"]].copy()
    fte["Year"] = pd.to_numeric(fte["Year"], errors="coerce")
    fte = fte.dropna(subset=["Year"])
    fte["Year"] = fte["Year"].astype(int)
    fte["FTE"] = pd.to_numeric(fte["FTE"], errors="coerce")
    fte["Employees"] = pd.to_numeric(fte["Employees"], errors="coerce")

    return flights, fte, default_year, rfi_factor, available_dimensions


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


def format_annual_table(annual: pd.DataFrame) -> pd.DataFrame:
    table = annual.copy()
    table["Year"] = table["Year"].astype(int).astype(str)
    table["Flights"] = table["Flights"].fillna(0).astype(int)
    table["Emissions_tCO2e"] = table["Emissions_tCO2e"].round(2)
    table["Distance_km"] = table["Distance_km"].round(0)
    table["FTE"] = table["FTE"].round(1)
    table["Relative_tCO2e_per_FTE"] = table["Relative_tCO2e_per_FTE"].round(2)
    return table


def annual_table_config():
    return {
        "Year": st.column_config.TextColumn("Year"),
        "Flights": st.column_config.NumberColumn("Flights", format="%d"),
        "Emissions_tCO2e": st.column_config.NumberColumn("Emissions (tCO₂e)", format="%.2f"),
        "Distance_km": st.column_config.NumberColumn("Distance (km)", format="%.0f"),
        "FTE": st.column_config.NumberColumn("FTE", format="%.1f"),
        "Relative_tCO2e_per_FTE": st.column_config.NumberColumn("Relative emissions (tCO₂e/FTE)", format="%.2f"),
    }


# ==========================================================
# Load data
# ==========================================================
if not EXCEL_FILE.exists():
    st.error(f"Excel workbook not found: {EXCEL_FILE.name}")
    st.stop()

try:
    flights, fte, workbook_default_year, rfi_factor, available_dimensions = load_from_excel(EXCEL_FILE.stat().st_mtime)
except Exception as exc:
    st.error("The dashboard could not read the Excel workbook.")
    st.exception(exc)
    st.stop()

available_years = sorted(flights["Year"].dropna().unique().tolist())
if not available_years:
    st.warning("No valid flight records found after applying Include_Final = Yes.")
    st.stop()

# ==========================================================
# Header
# ==========================================================
st.title("📊 MyClimate Flight Emissions Dashboard")
st.caption(
    "The dashboard reads directly from the original Excel workbook. "
    "The Python app keeps only the fields needed for the dashboard after loading the workbook."
)

# ==========================================================
# Sidebar controls
# ==========================================================
default_year = workbook_default_year if workbook_default_year in available_years else max(available_years)

with st.sidebar:
    st.header("Controls")

    selected_year = st.selectbox(
        "Analysis year",
        available_years,
        index=available_years.index(default_year),
    )

    baseline_defaults = [y for y in [2023, 2024] if y in available_years]
    baseline_years = st.multiselect(
        "Baseline years",
        available_years,
        default=baseline_defaults,
    )

    st.divider()
    st.subheader("Granularity filters")

    dimension_filters = {}
    if available_dimensions:
        for dim in available_dimensions:
            values = sorted(flights[dim].dropna().unique().tolist())
            selected_values = st.multiselect(
                f"Filter by {label(dim)}",
                values,
                default=values,
            )
            dimension_filters[dim] = selected_values
    else:
        st.info("No Team / Project / Hub columns found yet in All Integrated Data.")

    st.divider()
    st.caption(f"Workbook: `{EXCEL_FILE.name}`")
    st.caption(f"RFI factor from workbook controls: {rfi_factor}")

    if st.button("Clear cache and reload workbook"):
        st.cache_data.clear()
        st.rerun()

# ==========================================================
# Filter data
# ==========================================================
filtered_all_years = flights.copy()
for dim, selected_values in dimension_filters.items():
    filtered_all_years = filtered_all_years[filtered_all_years[dim].isin(selected_values)]

selected = filtered_all_years[filtered_all_years["Year"] == selected_year].copy()
baseline = filtered_all_years[filtered_all_years["Year"].isin(baseline_years)].copy() if baseline_years else filtered_all_years.iloc[0:0].copy()
annual = make_annual_summary(filtered_all_years, fte)

# ==========================================================
# KPIs
# ==========================================================
selected_fte_series = annual.loc[annual["Year"] == selected_year, "FTE"]
selected_fte = float(selected_fte_series.iloc[0]) if len(selected_fte_series) and pd.notna(selected_fte_series.iloc[0]) else None
selected_emissions = selected["Emissions_tCO2e"].sum()
selected_flights = len(selected)
selected_distance = selected["Distance_km"].sum()
selected_relative = selected_emissions / selected_fte if selected_fte and selected_fte != 0 else None

baseline_abs = baseline.groupby("Year")["Emissions_tCO2e"].sum().mean() if len(baseline) else None
baseline_fte = fte[fte["Year"].isin(baseline_years)]["FTE"].mean() if baseline_years else None
baseline_rel = baseline_abs / baseline_fte if baseline_abs and baseline_fte and baseline_fte != 0 else None

col1, col2, col3, col4 = st.columns(4)
col1.metric("Flights", f"{selected_flights:,.0f}")
col2.metric("Emissions", f"{selected_emissions:,.1f} tCO₂e")
col3.metric("Distance", f"{selected_distance:,.0f} km")
col4.metric("Relative emissions", "n/a" if selected_relative is None else f"{selected_relative:,.2f} tCO₂e/FTE")

if baseline_abs is not None:
    st.caption(
        f"Baseline absolute emissions: {baseline_abs:,.1f} tCO₂e. "
        + (f"Baseline relative emissions: {baseline_rel:,.2f} tCO₂e/FTE." if baseline_rel is not None else "")
    )

st.divider()

# ==========================================================
# Main chart row
# ==========================================================
left, right = st.columns((1.15, 1))

with left:
    monthly = (
        selected.dropna(subset=["Date"])
        .groupby(["Month", "Month_name"], as_index=False)["Emissions_tCO2e"]
        .sum()
        .sort_values("Month")
    )
    fig_month = px.line(
        monthly,
        x="Month_name",
        y="Emissions_tCO2e",
        markers=True,
        title=f"Emissions over time ({selected_year})",
        labels={"Month_name": "Month", "Emissions_tCO2e": "Emissions (tCO₂e)"},
    )
    fig_month.update_layout(height=420, margin=dict(l=20, r=20, t=60, b=20))
    st.plotly_chart(fig_month, use_container_width=True)

with right:
    by_class = (
        selected.groupby("Class", as_index=False)["Emissions_tCO2e"]
        .sum()
        .sort_values("Emissions_tCO2e", ascending=False)
    )
    fig_class = px.bar(
        by_class,
        x="Class",
        y="Emissions_tCO2e",
        title=f"Emissions by cabin class ({selected_year})",
        labels={"Class": "Class", "Emissions_tCO2e": "Emissions (tCO₂e)"},
    )
    fig_class.update_layout(height=420, margin=dict(l=20, r=20, t=60, b=20))
    st.plotly_chart(fig_class, use_container_width=True)

# ==========================================================
# Annual chart and summary table
# ==========================================================
left2, right2 = st.columns((1.15, 1))

with left2:
    fig_annual = px.bar(
        annual,
        x="Year",
        y="Emissions_tCO2e",
        text_auto=".0f",
        title="Annual absolute emissions",
        labels={"Emissions_tCO2e": "Emissions (tCO₂e)", "Year": "Year"},
    )
    fig_annual.update_xaxes(
        tickmode="array",
        tickvals=annual["Year"].astype(int).tolist(),
        ticktext=[str(int(y)) for y in annual["Year"].tolist()],
    )
    if baseline_abs is not None:
        fig_annual.add_hline(y=baseline_abs, line_dash="dash", annotation_text="Baseline avg", annotation_position="top left")
    fig_annual.update_layout(height=420, margin=dict(l=20, r=20, t=60, b=20))
    st.plotly_chart(fig_annual, use_container_width=True)

with right2:
    st.subheader("Annual summary table")
    annual_display = format_annual_table(annual)
    st.dataframe(
        annual_display,
        use_container_width=True,
        hide_index=True,
        column_config=annual_table_config(),
    )

# ==========================================================
# Granularity section
# ==========================================================
st.divider()
st.subheader("Granularity: by team, project, hub, etc.")

if available_dimensions:
    selected_dimension = st.selectbox(
        "Group emissions by",
        available_dimensions,
        format_func=label,
    )

    by_dimension = (
        selected.groupby(selected_dimension, as_index=False)
        .agg(
            Flights=("Emissions_tCO2e", "size"),
            Emissions_tCO2e=("Emissions_tCO2e", "sum"),
            Distance_km=("Distance_km", "sum"),
        )
        .sort_values("Emissions_tCO2e", ascending=False)
    )

    top_n = min(15, len(by_dimension))
    fig_dimension = px.bar(
        by_dimension.head(top_n),
        x=selected_dimension,
        y="Emissions_tCO2e",
        text_auto=".1f",
        title=f"Emissions by {label(selected_dimension)} ({selected_year})",
        labels={selected_dimension: label(selected_dimension), "Emissions_tCO2e": "Emissions (tCO₂e)"},
    )
    fig_dimension.update_layout(height=450, margin=dict(l=20, r=20, t=60, b=90), xaxis_tickangle=-30)
    st.plotly_chart(fig_dimension, use_container_width=True)

    st.dataframe(
        by_dimension.round({"Emissions_tCO2e": 2, "Distance_km": 0}),
        use_container_width=True,
        hide_index=True,
        column_config={
            selected_dimension: st.column_config.TextColumn(label(selected_dimension)),
            "Flights": st.column_config.NumberColumn("Flights", format="%d"),
            "Emissions_tCO2e": st.column_config.NumberColumn("Emissions (tCO₂e)", format="%.2f"),
            "Distance_km": st.column_config.NumberColumn("Distance (km)", format="%.0f"),
        },
    )

    # Annual trend for top groups.
    top_groups = by_dimension[selected_dimension].head(8).tolist()
    trend = filtered_all_years[filtered_all_years[selected_dimension].isin(top_groups)]
    trend = trend.groupby(["Year", selected_dimension], as_index=False)["Emissions_tCO2e"].sum()

    if len(trend):
        fig_trend = px.line(
            trend,
            x="Year",
            y="Emissions_tCO2e",
            color=selected_dimension,
            markers=True,
            title=f"Annual emissions trend by {label(selected_dimension)}",
            labels={"Year": "Year", "Emissions_tCO2e": "Emissions (tCO₂e)", selected_dimension: label(selected_dimension)},
        )
        fig_trend.update_xaxes(
            tickmode="array",
            tickvals=annual["Year"].astype(int).tolist(),
            ticktext=[str(int(y)) for y in annual["Year"].tolist()],
        )
        fig_trend.update_layout(height=430, margin=dict(l=20, r=20, t=60, b=20))
        st.plotly_chart(fig_trend, use_container_width=True)
else:
    st.info(
        "To enable this section, add one or more columns to the 'All Integrated Data' sheet, for example: "
        "Team, Project, Hub, Programme, Department, Cost_Center, or Travel_Purpose. "
        "The dashboard will detect these columns automatically after the workbook is updated."
    )

# ==========================================================
# Top routes
# ==========================================================
st.divider()
st.subheader("Top routes")
top_routes = (
    selected.groupby("Route", as_index=False)
    .agg(
        Flights=("Emissions_tCO2e", "size"),
        Emissions_tCO2e=("Emissions_tCO2e", "sum"),
        Distance_km=("Distance_km", "sum"),
    )
    .sort_values("Emissions_tCO2e", ascending=False)
    .head(15)
)
st.dataframe(
    top_routes.round({"Emissions_tCO2e": 2, "Distance_km": 0}),
    use_container_width=True,
    hide_index=True,
    column_config={
        "Route": st.column_config.TextColumn("Route"),
        "Flights": st.column_config.NumberColumn("Flights", format="%d"),
        "Emissions_tCO2e": st.column_config.NumberColumn("Emissions (tCO₂e)", format="%.2f"),
        "Distance_km": st.column_config.NumberColumn("Distance (km)", format="%.0f"),
    },
)

with st.expander("Fields used by this dashboard"):
    st.write(
        "The app reads the original workbook, then keeps only the fields needed for dashboarding. "
        "Core fields: Date, Year, DepartureAirport, ArrivalAirport, Route, Class, Distance_km, Emissions_tCO2e, Month, Month_name."
    )
    if available_dimensions:
        st.write("Detected granularity fields:", ", ".join([label(dim) for dim in available_dimensions]))
    else:
        st.write("No granularity fields detected yet. Add Team, Project, Hub, etc. to All Integrated Data to activate granular filters and charts.")
