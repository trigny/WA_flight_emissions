
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="MyClimate Dashboard", page_icon="📊", layout="wide")

# Keep the original Excel workbook in the same folder as this app.
# When the workbook is updated, replace this file with the updated workbook using the same filename,
# or change EXCEL_FILE below to match the new filename.
EXCEL_FILE = Path(__file__).parent / "MyClimate Methodology Workbook_final (1).xlsx"
ALL_DATA_SHEET = "All Integrated Data"
FTE_SHEET = "FTE Data"
DASHBOARD_SHEET = "Dashboard"

RELEVANT_COLUMNS = [
    "Date",
    "Year",
    "DepartureAirport",
    "ArrivalAirport",
    "Class",
    "Distance_km",
    "Final_RFI3_tCO2e",
    "Include_Final",
]

@st.cache_data(show_spinner="Reading Excel workbook...")
def load_from_excel(file_mtime: float):
    """Read the original workbook and keep only dashboard-relevant fields in memory."""
    all_data = pd.read_excel(EXCEL_FILE, sheet_name=ALL_DATA_SHEET, engine="openpyxl")
    fte = pd.read_excel(EXCEL_FILE, sheet_name=FTE_SHEET, engine="openpyxl")

    # Optional: read dashboard settings from the Excel Dashboard tab.
    try:
        dash = pd.read_excel(EXCEL_FILE, sheet_name=DASHBOARD_SHEET, header=None, engine="openpyxl")
        default_year = int(dash.iloc[7, 1]) if pd.notna(dash.iloc[7, 1]) else None
        rfi_factor = dash.iloc[9, 1] if pd.notna(dash.iloc[9, 1]) else 3
    except Exception:
        default_year = None
        rfi_factor = 3

    missing = [col for col in RELEVANT_COLUMNS if col not in all_data.columns]
    if missing:
        raise ValueError(f"Missing expected columns in '{ALL_DATA_SHEET}': {missing}")

    flights = all_data[RELEVANT_COLUMNS].copy()

    # Use only final integrated records, excluding duplicate copies according to workbook logic.
    flights = flights[flights["Include_Final"].astype(str).str.strip().str.lower().eq("yes")].copy()

    # Clean and derive display columns.
    flights["Date"] = pd.to_datetime(flights["Date"], errors="coerce")
    flights["Year"] = pd.to_numeric(flights["Year"], errors="coerce")
    flights = flights.dropna(subset=["Year"])
    flights["Year"] = flights["Year"].astype(int)
    flights["Class"] = flights["Class"].fillna("unknown").astype(str).str.lower().str.strip()
    flights["DepartureAirport"] = flights["DepartureAirport"].fillna("UNK").astype(str).str.upper().str.strip()
    flights["ArrivalAirport"] = flights["ArrivalAirport"].fillna("UNK").astype(str).str.upper().str.strip()
    flights["Route"] = flights["DepartureAirport"] + " → " + flights["ArrivalAirport"]
    flights["Emissions_tCO2e"] = pd.to_numeric(flights["Final_RFI3_tCO2e"], errors="coerce").fillna(0)
    flights["Distance_km"] = pd.to_numeric(flights["Distance_km"], errors="coerce").fillna(0)
    flights["Month"] = flights["Date"].dt.month
    flights["Month_name"] = flights["Date"].dt.strftime("%b")

    # Keep only what the dashboard needs after reading from the full original Excel file.
    flights = flights[
        [
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
        ]
    ]

    fte = fte[["Year", "Employees", "FTE"]].copy()
    fte["Year"] = pd.to_numeric(fte["Year"], errors="coerce")
    fte = fte.dropna(subset=["Year"])
    fte["Year"] = fte["Year"].astype(int)
    fte["FTE"] = pd.to_numeric(fte["FTE"], errors="coerce")

    return flights, fte, default_year, rfi_factor

if not EXCEL_FILE.exists():
    st.error(f"Excel workbook not found: {EXCEL_FILE.name}")
    st.stop()

flights, fte, workbook_default_year, rfi_factor = load_from_excel(EXCEL_FILE.stat().st_mtime)

st.title("📊 MyClimate Flight Emissions Dashboard")
st.caption(
    "This dashboard reads directly from the original Excel workbook and keeps only dashboard-relevant fields in Python. "
    "Update the workbook file to refresh the dashboard data."
)

available_years = sorted(flights["Year"].dropna().unique().tolist())
if not available_years:
    st.warning("No valid flight records found after applying Include_Final = Yes.")
    st.stop()

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
    st.caption(f"Workbook: `{EXCEL_FILE.name}`")
    st.caption(f"RFI factor from workbook controls: {rfi_factor}")
    if st.button("Clear cache and reload workbook"):
        st.cache_data.clear()
        st.rerun()

selected = flights[flights["Year"] == selected_year].copy()
baseline = flights[flights["Year"].isin(baseline_years)].copy() if baseline_years else flights.iloc[0:0].copy()

annual = (
    flights.groupby("Year", as_index=False)
    .agg(
        Flights=("Emissions_tCO2e", "size"),
        Emissions_tCO2e=("Emissions_tCO2e", "sum"),
        Distance_km=("Distance_km", "sum"),
    )
    .merge(fte[["Year", "FTE"]], on="Year", how="left")
)
annual["Relative_tCO2e_per_FTE"] = annual["Emissions_tCO2e"] / annual["FTE"]

selected_fte_series = annual.loc[annual["Year"] == selected_year, "FTE"]
selected_fte = float(selected_fte_series.iloc[0]) if len(selected_fte_series) and pd.notna(selected_fte_series.iloc[0]) else None
selected_emissions = selected["Emissions_tCO2e"].sum()
selected_flights = len(selected)
selected_relative = selected_emissions / selected_fte if selected_fte and selected_fte != 0 else None

baseline_abs = baseline.groupby("Year")["Emissions_tCO2e"].sum().mean() if len(baseline) else None
baseline_fte = fte[fte["Year"].isin(baseline_years)]["FTE"].mean() if baseline_years else None
baseline_rel = baseline_abs / baseline_fte if baseline_abs and baseline_fte and baseline_fte != 0 else None

col1, col2, col3, col4 = st.columns(4)
col1.metric("Flights", f"{selected_flights:,.0f}")
col2.metric("Emissions", f"{selected_emissions:,.1f} tCO₂e")
col3.metric("Relative emissions", "n/a" if selected_relative is None else f"{selected_relative:,.2f} tCO₂e/FTE")
col4.metric("Baseline emissions", "n/a" if baseline_abs is None else f"{baseline_abs:,.1f} tCO₂e")

if selected_relative is not None and baseline_rel is not None:
    st.caption(f"Baseline relative emissions: {baseline_rel:,.2f} tCO₂e/FTE based on {', '.join(map(str, baseline_years))}.")

st.divider()

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
        tickvals=annual["Year"].tolist(),
        ticktext=[str(int(y)) for y in annual["Year"].tolist()]
    )
    if baseline_abs is not None:
        fig_annual.add_hline(y=baseline_abs, line_dash="dash", annotation_text="Baseline avg", annotation_position="top left")
    fig_annual.update_layout(height=420, margin=dict(l=20, r=20, t=60, b=20))
    st.plotly_chart(fig_annual, use_container_width=True)

with right2:
    st.subheader("Annual summary table")

    annual_display = annual.copy()
    annual_display["Year"] = annual_display["Year"].astype(int).astype(str)
    annual_display["Flights"] = annual_display["Flights"].astype(int)
    annual_display["Emissions_tCO2e"] = annual_display["Emissions_tCO2e"].round(2)
    annual_display["Distance_km"] = annual_display["Distance_km"].round(0)
    annual_display["FTE"] = annual_display["FTE"].round(1)
    annual_display["Relative_tCO2e_per_FTE"] = annual_display["Relative_tCO2e_per_FTE"].round(2)

    st.dataframe(
        annual_display,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Year": st.column_config.TextColumn("Year"),
            "Flights": st.column_config.NumberColumn("Flights", format="%d"),
            "Emissions_tCO2e": st.column_config.NumberColumn("Emissions (tCO₂e)", format="%.2f"),
            "Distance_km": st.column_config.NumberColumn("Distance (km)", format="%.0f"),
            "FTE": st.column_config.NumberColumn("FTE", format="%.1f"),
            "Relative_tCO2e_per_FTE": st.column_config.NumberColumn("Relative emissions (tCO₂e/FTE)", format="%.2f"),
        },
    )

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
st.dataframe(top_routes, use_container_width=True, hide_index=True)


with st.expander("Dashboard fields kept in Python"):
    st.write(
        "The app reads the original workbook, then keeps only: Date, Year, DepartureAirport, ArrivalAirport, "
        "Route, Class, Distance_km, Emissions_tCO2e, Month, and Month_name. Traveler names are not used in the app logic."
    )
