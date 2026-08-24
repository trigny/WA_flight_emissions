from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="MyClimate Dashboard", page_icon="📊", layout="wide")

# -------------------------------------------------------------------
# Workbook source
# -------------------------------------------------------------------
EXCEL_FILE = Path(__file__).parent / "Flight Emissions Dashboard.xlsx"

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
    team_col = "Team" if "Team" in all_data.columns else None

    required = [
        "Date",
        "Year",
        "DepartureAirport",
        "ArrivalAirport",
        cabin_col,
        "Flight_Type",
        "Distance_km",
        "Final_RFI3_tCO2e",
        "Include_Final",
    ]
    missing = [col for col in required if col not in all_data.columns]
    if missing:
        raise ValueError(f"Missing expected columns in '{ALL_DATA_SHEET}': {missing}")

    keep = required + ([team_col] if team_col else [])
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
    flights["Teams"] = clean_text(flights[team_col], "Unassigned") if team_col else "Unassigned"
    flights["Flight Type"] = clean_text(flights["Flight_Type"], "Unknown").str.lower().replace({"short_haul": "Short haul", "medium_haul": "Medium haul", "long_haul": "Long haul"})
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
            "Flight Type",
            "Teams",
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
if not EXCEL_FILE.exists():
    st.error("The workbook 'Flight Emissions Dashboard.xlsx' was not found. Place it in the same folder as app.py.")
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
st.caption("Reads the Team and Flight_Type fields from the updated Excel workflow.")

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

    team_values = sorted(flights["Teams"].dropna().unique().tolist())
    selected_teams = st.multiselect("Teams", team_values, default=team_values)

    st.divider()
    st.caption(f"Workbook: `{EXCEL_FILE.name}`")
    st.caption(f"RFI factor from workbook controls: {rfi_factor}")
    if st.button("Clear cache and reload workbook"):
        st.cache_data.clear()
        st.rerun()

filtered_all_years = flights[
    flights["Cabin Class"].isin(selected_cabins) & flights["Teams"].isin(selected_teams)
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
    by_team = (
        selected.groupby("Teams", as_index=False)
        .agg(Flights=("Emissions_tCO2e", "size"), Emissions_tCO2e=("Emissions_tCO2e", "sum"))
        .sort_values("Emissions_tCO2e", ascending=False)
        .head(12)
    )
    fig_team = px.bar(
        by_team,
        x="Teams",
        y="Emissions_tCO2e",
        text_auto=".1f",
        title=f"Emissions by teams ({selected_year})",
        labels={"Teams": "Teams", "Emissions_tCO2e": "Emissions (tCO₂e)"},
    )
    fig_team.update_layout(height=400, margin=dict(l=20, r=20, t=60, b=90), xaxis_tickangle=-35)
    st.plotly_chart(fig_team, use_container_width=True)

by_flight_type = selected.groupby("Flight Type", as_index=False).agg(Flights=("Emissions_tCO2e", "size"), Emissions_tCO2e=("Emissions_tCO2e", "sum"))
fig_flight_type = px.bar(by_flight_type, x="Flight Type", y="Emissions_tCO2e", text_auto=".1f", title=f"Emissions by flight type ({selected_year})", category_orders={"Flight Type": ["Short haul", "Medium haul", "Long haul", "Unknown"]}, labels={"Flight Type": "Flight type", "Emissions_tCO2e": "Emissions (tCO₂e)"})
fig_flight_type.update_layout(height=420, margin=dict(l=20, r=20, t=60, b=60))
st.plotly_chart(fig_flight_type, use_container_width=True)

# Cabin class by team stacked chart.
team_cabin = (
    selected.groupby(["Teams", "Cabin Class"], as_index=False)
    .agg(Flights=("Emissions_tCO2e", "size"), Emissions_tCO2e=("Emissions_tCO2e", "sum"))
)
top_teams = (
    team_cabin.groupby("Teams", as_index=False)["Emissions_tCO2e"]
    .sum()
    .sort_values("Emissions_tCO2e", ascending=False)
    .head(12)["Teams"]
    .tolist()
)
team_cabin_top = team_cabin[team_cabin["Teams"].isin(top_teams)]

fig_stack = px.bar(
    team_cabin_top,
    x="Teams",
    y="Emissions_tCO2e",
    color="Cabin Class",
    title=f"Cabin class contribution within teams ({selected_year})",
    labels={"Teams": "Teams", "Emissions_tCO2e": "Emissions (tCO₂e)", "Cabin Class": "Cabin class"},
)
fig_stack.update_layout(height=460, margin=dict(l=20, r=20, t=60, b=100), xaxis_tickangle=-35)
st.plotly_chart(fig_stack, use_container_width=True)

# -------------------------------------------------------------------
# Annual chart and annual table
# -------------------------------------------------------------------
left2, right2 = st.columns((1.1, 1))

with left2:
    pathway_start_year = 2024
    pathway_target_year = 2030
    pathway_target_value = 4.0
    annual_relative = annual.dropna(subset=["Relative_tCO2e_per_FTE"]).copy()
    annual_relative = annual_relative[annual_relative["FTE"].gt(0)].copy()

    # Same baseline definition as the Excel dashboard:
    # average 2023/2024 emissions divided by average 2023/2024 FTE.
    baseline_emissions = filtered_all_years[filtered_all_years["Year"].isin([2023, 2024])].groupby("Year")["Emissions_tCO2e"].sum().reindex([2023, 2024])
    baseline_fte = fte[fte["Year"].isin([2023, 2024])].set_index("Year")["FTE"].reindex([2023, 2024])
    pathway_start_value = float(baseline_emissions.mean() / baseline_fte.mean()) if baseline_emissions.notna().all() and baseline_fte.notna().all() and baseline_fte.mean() != 0 else None

    fig_annual = go.Figure()
    fig_annual.add_bar(x=annual_relative["Year"], y=annual_relative["Relative_tCO2e_per_FTE"], name="Relative emissions", marker_color="#4F81BD", text=annual_relative["Relative_tCO2e_per_FTE"].round(2), textposition="outside", hovertemplate="%{x}: %{y:.2f} tCO₂e/FTE<extra></extra>")
    if pathway_start_value is not None:
        pathway_years = list(range(pathway_start_year, pathway_target_year + 1))
        annual_step = (pathway_target_value - pathway_start_value) / (pathway_target_year - pathway_start_year)
        pathway_values = [pathway_start_value + annual_step * (year - pathway_start_year) for year in pathway_years]
        fig_annual.add_scatter(x=pathway_years, y=pathway_values, mode="lines", name="Relative emission target", line=dict(color="#2F5597", width=3, dash="dot"), hovertemplate="%{x}: %{y:.2f} tCO₂e/FTE<extra></extra>")
    chart_years = sorted(set(annual["Year"].astype(int).tolist() + list(range(2024, 2031))))
    fig_annual.update_xaxes(tickmode="array", tickvals=chart_years, ticktext=[str(year) for year in chart_years], title="Year")
    fig_annual.update_yaxes(title="tCO₂e / FTE", rangemode="tozero")
    fig_annual.update_layout(title="Annual relative emissions", height=420, margin=dict(l=20, r=20, t=60, b=20), legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5), bargap=0.55)
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

tab_teams, tab_cabin = st.tabs(["Teams summary", "Cabin class summary"])

with tab_teams:
    teams_summary = (
        selected.groupby("Teams", as_index=False)
        .agg(Flights=("Emissions_tCO2e", "size"), Emissions_tCO2e=("Emissions_tCO2e", "sum"), Distance_km=("Distance_km", "sum"))
        .sort_values("Emissions_tCO2e", ascending=False)
    )
    st.dataframe(teams_summary.round({"Emissions_tCO2e": 2, "Distance_km": 0}), use_container_width=True, hide_index=True)

with tab_cabin:
    cabin_summary = (
        selected.groupby("Cabin Class", as_index=False)
        .agg(Flights=("Emissions_tCO2e", "size"), Emissions_tCO2e=("Emissions_tCO2e", "sum"), Distance_km=("Distance_km", "sum"))
        .sort_values("Emissions_tCO2e", ascending=False)
    )
    st.dataframe(cabin_summary.round({"Emissions_tCO2e": 2, "Distance_km": 0}), use_container_width=True, hide_index=True)

with st.expander("Fields used by this dashboard"):
    st.write(
        "Core dashboard fields: Date, Year, DepartureAirport, ArrivalAirport, Route, Cabin Class, Flight Type, Teams, Distance_km, Emissions_tCO2e, Month, Month_name."
    )
    st.write("Excel column `Team` is displayed as `Teams`. Blank or zero team values are grouped as `Unassigned`.")
