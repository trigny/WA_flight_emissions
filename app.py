from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="MyClimate Dashboard", page_icon="📊", layout="wide")

# -------------------------------------------------------------------
# Workbook source
# -------------------------------------------------------------------
# Put your current workbook in the same folder as this app.
# The app tries these names in order so you can keep either filename.
WORKBOOK_CANDIDATES = [
    "MyClimate Methodology Workbook_final_updated_department_blankJ.xlsx",
    "MyClimate Methodology Workbook_final_updated_department.xlsx",
    "MyClimate Methodology Workbook_final (1).xlsx",
]

EXCEL_FILE = None
for candidate in WORKBOOK_CANDIDATES:
    candidate_path = Path(__file__).parent / candidate
    if candidate_path.exists():
        EXCEL_FILE = candidate_path
        break

ALL_DATA_SHEET = "All Integrated Data"
FTE_SHEET = "FTE Data"
DASHBOARD_SHEET = "Dashboard"


# -------------------------------------------------------------------
# Helpers
# -------------------------------------------------------------------
def clean_text(series: pd.Series, blank_label: str = "Unassigned") -> pd.Series:
    cleaned = series.fillna("").astype(str).str.strip()
    cleaned = cleaned.replace({"0": "", "0.0": "", "nan": "", "None": ""})
    return cleaned.mask(cleaned.eq(""), blank_label)


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


@st.cache_data(show_spinner="Reading Excel workbook...")
def load_data(file_mtime: float):
    all_data = pd.read_excel(EXCEL_FILE, sheet_name=ALL_DATA_SHEET, engine="openpyxl")
    fte = pd.read_excel(EXCEL_FILE, sheet_name=FTE_SHEET, engine="openpyxl")

    # Support both old and new workbook headers.
    cabin_col = "Cabin Class" if "Cabin Class" in all_data.columns else "Class"
    hub_col = "Traveler Department" if "Traveler Department" in all_data.columns else None

    required = [
        "Date",
        "Year",
        "DepartureAirport",
        "ArrivalAirport",
        cabin_col,
        "Distance_km",
        "Final_RFI3_tCO2e",
        "Include_Final",
    ]
    missing = [col for col in required if col not in all_data.columns]
    if missing:
        raise ValueError(f"Missing expected columns in '{ALL_DATA_SHEET}': {missing}")

    keep = required + ([hub_col] if hub_col else [])
    flights = all_data[keep].copy()
    flights = flights[flights["Include_Final"].astype(str).str.strip().str.lower().eq("yes")].copy()

    flights["Date"] = pd.to_datetime(flights["Date"], errors="coerce")
    flights["Year"] = pd.to_numeric(flights["Year"], errors="coerce")
    flights = flights.dropna(subset=["Year"])
    flights["Year"] = flights["Year"].astype(int)

    flights["DepartureAirport"] = clean_text(flights["DepartureAirport"], "UNK").str.upper()
    flights["ArrivalAirport"] = clean_text(flights["ArrivalAirport"], "UNK").str.upper()
    flights["Route"] = flights["DepartureAirport"] + " → " + flights["ArrivalAirport"]

    flights["Cabin Class"] = clean_text(flights[cabin_col], "Unknown").str.lower()
    flights["Hubs"] = clean_text(flights[hub_col], "Unassigned") if hub_col else "Unassigned"
    flights["Distance_km"] = pd.to_numeric(flights["Distance_km"], errors="coerce").fillna(0)
    flights["Emissions_tCO2e"] = pd.to_numeric(flights["Final_RFI3_tCO2e"], errors="coerce").fillna(0)
    flights["Month"] = flights["Date"].dt.month
    flights["Month_name"] = flights["Date"].dt.strftime("%b")

    flights = flights[
        [
            "Date",
            "Year",
            "DepartureAirport",
            "ArrivalAirport",
            "Route",
            "Cabin Class",
            "Hubs",
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
    fte["Employees"] = pd.to_numeric(fte["Employees"], errors="coerce")
    fte["FTE"] = pd.to_numeric(fte["FTE"], errors="coerce")

    try:
        dash = pd.read_excel(EXCEL_FILE, sheet_name=DASHBOARD_SHEET, header=None, engine="openpyxl")
        default_year = int(dash.iloc[7, 1]) if pd.notna(dash.iloc[7, 1]) else None
        rfi_factor = dash.iloc[9, 1] if pd.notna(dash.iloc[9, 1]) else 3
    except Exception:
        default_year = None
        rfi_factor = 3

    return flights, fte, default_year, rfi_factor


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


# -------------------------------------------------------------------
# Load workbook
# -------------------------------------------------------------------
if EXCEL_FILE is None:
    st.error("No Excel workbook found in the app folder. Add the updated workbook next to app.py.")
    st.stop()

try:
    flights, fte, workbook_default_year, rfi_factor = load_data(EXCEL_FILE.stat().st_mtime)
except Exception as exc:
    st.error("The dashboard could not read the Excel workbook.")
    st.exception(exc)
    st.stop()

available_years = sorted(flights["Year"].dropna().unique().tolist())
if not available_years:
    st.warning("No valid flight records found after applying Include_Final = Yes.")
    st.stop()

# -------------------------------------------------------------------
# Header and controls
# -------------------------------------------------------------------
st.title("📊 MyClimate Flight Emissions Dashboard")
st.caption("Reads from the original Excel workflow. Traveler Department is shown as **Hubs** in this dashboard.")

default_year = workbook_default_year if workbook_default_year in available_years else max(available_years)

with st.sidebar:
    st.header("Controls")
    selected_year = st.selectbox("Analysis year", available_years, index=available_years.index(default_year))
    baseline_years = st.multiselect(
        "Baseline years",
        available_years,
        default=[y for y in [2023, 2024] if y in available_years],
    )

    st.divider()
    st.subheader("Dashboard filters")

    cabin_values = sorted(flights["Cabin Class"].dropna().unique().tolist())
    selected_cabins = st.multiselect("Cabin class", cabin_values, default=cabin_values)

    hub_values = sorted(flights["Hubs"].dropna().unique().tolist())
    selected_hubs = st.multiselect("Hubs", hub_values, default=hub_values)

    st.divider()
    st.caption(f"Workbook: `{EXCEL_FILE.name}`")
    st.caption(f"RFI factor from workbook controls: {rfi_factor}")
    if st.button("Clear cache and reload workbook"):
        st.cache_data.clear()
        st.rerun()

filtered_all_years = flights[
    flights["Cabin Class"].isin(selected_cabins) & flights["Hubs"].isin(selected_hubs)
].copy()
selected = filtered_all_years[filtered_all_years["Year"] == selected_year].copy()
baseline = filtered_all_years[filtered_all_years["Year"].isin(baseline_years)].copy() if baseline_years else filtered_all_years.iloc[0:0].copy()
annual = make_annual_summary(filtered_all_years, fte)

# -------------------------------------------------------------------
# KPIs
# -------------------------------------------------------------------
selected_fte_series = annual.loc[annual["Year"] == selected_year, "FTE"]
selected_fte = float(selected_fte_series.iloc[0]) if len(selected_fte_series) and pd.notna(selected_fte_series.iloc[0]) else None
selected_emissions = selected["Emissions_tCO2e"].sum()
selected_flights = len(selected)
selected_distance = selected["Distance_km"].sum()
selected_relative = selected_emissions / selected_fte if selected_fte and selected_fte != 0 else None
baseline_abs = baseline.groupby("Year")["Emissions_tCO2e"].sum().mean() if len(baseline) else None

k1, k2, k3, k4 = st.columns(4)
k1.metric("Flights", f"{selected_flights:,.0f}")
k2.metric("Emissions", f"{selected_emissions:,.1f} tCO₂e")
k3.metric("Distance", f"{selected_distance:,.0f} km")
k4.metric("Relative emissions", "n/a" if selected_relative is None else f"{selected_relative:,.2f} tCO₂e/FTE")
if baseline_abs is not None:
    st.caption(f"Baseline absolute emissions: {baseline_abs:,.1f} tCO₂e based on {', '.join(map(str, baseline_years))}.")

st.divider()

# -------------------------------------------------------------------
# Dashboard chart analysis
# -------------------------------------------------------------------
st.subheader("Dashboard chart analysis")

left, middle, right = st.columns((1.2, 1, 1))

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
    fig_month.update_layout(height=400, margin=dict(l=20, r=20, t=60, b=20))
    st.plotly_chart(fig_month, use_container_width=True)

with middle:
    by_cabin = (
        selected.groupby("Cabin Class", as_index=False)
        .agg(Flights=("Emissions_tCO2e", "size"), Emissions_tCO2e=("Emissions_tCO2e", "sum"))
        .sort_values("Emissions_tCO2e", ascending=False)
    )
    fig_cabin = px.bar(
        by_cabin,
        x="Cabin Class",
        y="Emissions_tCO2e",
        text_auto=".1f",
        title=f"Emissions by cabin class ({selected_year})",
        labels={"Cabin Class": "Cabin class", "Emissions_tCO2e": "Emissions (tCO₂e)"},
    )
    fig_cabin.update_layout(height=400, margin=dict(l=20, r=20, t=60, b=70), xaxis_tickangle=-25)
    st.plotly_chart(fig_cabin, use_container_width=True)

with right:
    by_hub = (
        selected.groupby("Hubs", as_index=False)
        .agg(Flights=("Emissions_tCO2e", "size"), Emissions_tCO2e=("Emissions_tCO2e", "sum"))
        .sort_values("Emissions_tCO2e", ascending=False)
        .head(12)
    )
    fig_hub = px.bar(
        by_hub,
        x="Hubs",
        y="Emissions_tCO2e",
        text_auto=".1f",
        title=f"Emissions by hubs ({selected_year})",
        labels={"Hubs": "Hubs", "Emissions_tCO2e": "Emissions (tCO₂e)"},
    )
    fig_hub.update_layout(height=400, margin=dict(l=20, r=20, t=60, b=90), xaxis_tickangle=-35)
    st.plotly_chart(fig_hub, use_container_width=True)

# Cabin class by hub stacked chart.
hub_cabin = (
    selected.groupby(["Hubs", "Cabin Class"], as_index=False)
    .agg(Flights=("Emissions_tCO2e", "size"), Emissions_tCO2e=("Emissions_tCO2e", "sum"))
)
top_hubs = (
    hub_cabin.groupby("Hubs", as_index=False)["Emissions_tCO2e"]
    .sum()
    .sort_values("Emissions_tCO2e", ascending=False)
    .head(12)["Hubs"]
    .tolist()
)
hub_cabin_top = hub_cabin[hub_cabin["Hubs"].isin(top_hubs)]

fig_stack = px.bar(
    hub_cabin_top,
    x="Hubs",
    y="Emissions_tCO2e",
    color="Cabin Class",
    title=f"Cabin class contribution within hubs ({selected_year})",
    labels={"Hubs": "Hubs", "Emissions_tCO2e": "Emissions (tCO₂e)", "Cabin Class": "Cabin class"},
)
fig_stack.update_layout(height=460, margin=dict(l=20, r=20, t=60, b=100), xaxis_tickangle=-35)
st.plotly_chart(fig_stack, use_container_width=True)

# -------------------------------------------------------------------
# Annual chart and annual table
# -------------------------------------------------------------------
left2, right2 = st.columns((1.1, 1))

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
    st.dataframe(
        format_annual_table(annual),
        use_container_width=True,
        hide_index=True,
        column_config=annual_table_config(),
    )

# -------------------------------------------------------------------
# Tables
# -------------------------------------------------------------------
st.divider()
st.subheader("Detailed summaries")

tab_routes, tab_hubs, tab_cabin = st.tabs(["Top routes", "Hubs summary", "Cabin class summary"])

with tab_routes:
    top_routes = (
        selected.groupby("Route", as_index=False)
        .agg(Flights=("Emissions_tCO2e", "size"), Emissions_tCO2e=("Emissions_tCO2e", "sum"), Distance_km=("Distance_km", "sum"))
        .sort_values("Emissions_tCO2e", ascending=False)
        .head(20)
    )
    st.dataframe(top_routes.round({"Emissions_tCO2e": 2, "Distance_km": 0}), use_container_width=True, hide_index=True)

with tab_hubs:
    hubs_summary = (
        selected.groupby("Hubs", as_index=False)
        .agg(Flights=("Emissions_tCO2e", "size"), Emissions_tCO2e=("Emissions_tCO2e", "sum"), Distance_km=("Distance_km", "sum"))
        .sort_values("Emissions_tCO2e", ascending=False)
    )
    st.dataframe(hubs_summary.round({"Emissions_tCO2e": 2, "Distance_km": 0}), use_container_width=True, hide_index=True)

with tab_cabin:
    cabin_summary = (
        selected.groupby("Cabin Class", as_index=False)
        .agg(Flights=("Emissions_tCO2e", "size"), Emissions_tCO2e=("Emissions_tCO2e", "sum"), Distance_km=("Distance_km", "sum"))
        .sort_values("Emissions_tCO2e", ascending=False)
    )
    st.dataframe(cabin_summary.round({"Emissions_tCO2e": 2, "Distance_km": 0}), use_container_width=True, hide_index=True)

with st.expander("Fields used by this dashboard"):
    st.write(
        "Core dashboard fields: Date, Year, DepartureAirport, ArrivalAirport, Route, Cabin Class, Hubs, Distance_km, Emissions_tCO2e, Month, Month_name."
    )
    st.write("Excel column `Traveler Department` is displayed as `Hubs` in the dashboard. Blank or zero department values are grouped only in the dashboard as `Unassigned`.")
