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
        "Relative_tCO2e_per_FTE": st.column_config.NumberColumn("Emissions per FTE (tCO₂e/FTE)", format="%.2f"),
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
    flights["Teams"] = clean_text(flights[team_col], "External") if team_col else "External"
    flights["Flight Type"] = (
        flights["Flight_Type"].fillna("").astype(str).str.strip().str.lower()
        .replace({
            "very_short_haul": "Very short haul",
            "short_haul": "Short haul",
            "medium_haul": "Medium haul",
            "long_haul": "Long haul",
        })
    )
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
k4.metric("Emissions per FTE", "n/a" if selected_relative is None else f"{selected_relative:,.2f} tCO₂e/FTE")
if baseline_abs is not None:
    st.caption(f"Baseline absolute emissions: {baseline_abs:,.1f} tCO₂e based on {', '.join(map(str, baseline_years))}.")

# -------------------------------------------------------------------
# Project planning estimate
# -------------------------------------------------------------------
st.divider()
st.subheader("Project planning estimate")
planning_years = list(range(2024, 2031))
baseline_emissions = (
    flights[flights["Year"].isin([2023, 2024])]
    .groupby("Year")["Emissions_tCO2e"].sum().reindex([2023, 2024])
)
baseline_fte = (
    fte[fte["Year"].isin([2023, 2024])]
    .set_index("Year")["FTE"].reindex([2023, 2024])
)

if not (baseline_emissions.notna().all() and baseline_fte.notna().all() and baseline_fte.mean() > 0):
    st.warning("Valid 2023 and 2024 emissions and FTE data are required for the planning target.")
else:
    baseline_per_fte = float(baseline_emissions.mean() / baseline_fte.mean())
    target_2030 = baseline_per_fte / 2
    target_step = (target_2030 - baseline_per_fte) / 6
    yearly_targets = {
        year: baseline_per_fte + target_step * (year - 2024)
        for year in planning_years
    }

    planning_year = st.selectbox(
        "Target year", planning_years, index=2, key="planning_target_year"
    )
    selected_target = yearly_targets[planning_year]

    selected_year_fte = fte.loc[fte["Year"].eq(planning_year), "FTE"].dropna()
    latest_fte = fte.sort_values("Year")["FTE"].dropna()
    default_planning_fte = float(selected_year_fte.iloc[0] if not selected_year_fte.empty else latest_fte.iloc[-1])
    planning_fte = st.number_input(
        "Planned FTE", min_value=0.1, value=default_planning_fte, step=1.0,
        help="Flight inputs are totals. Total estimated emissions are divided by this FTE value.",
    )

    reference_years = sorted(
        flights.loc[flights["Emissions_tCO2e"].gt(0), "Year"].dropna().unique().tolist()
    )
    if not reference_years:
        st.warning("No positive flight-emission records are available for planning factors.")
    else:
        reference_year = 2025 if 2025 in reference_years else max(reference_years)
        valid_types = ["Very short haul", "Short haul", "Medium haul", "Long haul"]
        valid_cabins = ["economy", "premiumeconomy", "business"]
        factors = (
            flights[
                flights["Year"].eq(reference_year)
                & flights["Emissions_tCO2e"].gt(0)
                & flights["Flight Type"].isin(valid_types)
                & flights["Cabin Class"].isin(valid_cabins)
            ]
            .groupby(["Flight Type", "Cabin Class"])["Emissions_tCO2e"]
            .agg(["mean", "size"])
        )

        labels = {
            "Very short haul": "Very short haul (<500 km)",
            "Short haul": "Short haul (500–1,500 km)",
            "Medium haul": "Medium haul (1,500–4,000 km)",
            "Long haul": "Long haul (>4,000 km)",
        }
        st.markdown("#### Planned one-way flights")
        st.caption(
            f"Average emission factors are calculated from included {reference_year} flights."
        )
        heading = st.columns([2.2, 1, 1, 1, 1.4])
        heading[0].markdown("**Flight distance**")
        heading[1].markdown("**Economy**")
        heading[2].markdown("**Premium economy**")
        heading[3].markdown("**Business**")
        heading[4].markdown("**Estimated tCO₂e**")

        planned_total_emissions = 0.0
        planned_segments = 0
        unavailable = []
        for flight_type in valid_types:
            row = st.columns([2.2, 1, 1, 1, 1.4])
            row[0].write(labels[flight_type])
            row_total = 0.0
            for input_col, cabin in zip(row[1:4], valid_cabins):
                key = (flight_type, cabin)
                available = key in factors.index
                with input_col:
                    number = st.selectbox(
                        f"{labels[flight_type]} {cabin}",
                        range(501),
                        key=f"plan_{flight_type}_{cabin}",
                        label_visibility="collapsed",
                        help=(
                            f"Total one-way {cabin} segments for the planned FTE."
                            if available else
                            f"No {reference_year} emission factor is available."
                        ),
                    )
                planned_segments += number
                if available:
                    row_total += number * float(factors.loc[key, "mean"])
                elif number:
                    unavailable.append(f"{labels[flight_type]} {cabin}")
            planned_total_emissions += row_total
            row[4].write(f"{row_total:.2f}")

        if unavailable:
            st.warning(
                "Not included because no reference factor is available: "
                + ", ".join(unavailable)
            )

        planned_emissions_per_fte = planned_total_emissions / planning_fte
        remaining = selected_target - planned_emissions_per_fte
        share = planned_emissions_per_fte / selected_target if selected_target > 0 else 0
        r1, r2, r3 = st.columns(3)
        r1.metric("Estimated emissions", f"{planned_emissions_per_fte:.2f} tCO₂e per FTE")
        r2.metric(f"{planning_year} target", f"{selected_target:.2f} tCO₂e per FTE")
        r3.metric(
            "Remaining allowance" if remaining >= 0 else "Above target by",
            f"{abs(remaining):.2f} tCO₂e per FTE",
        )
        st.progress(min(max(share, 0.0), 1.0))
        st.caption(f"This configuration uses {share:.0%} of the {planning_year} target.")
        if remaining >= 0:
            st.success(f"Within target. Remaining: {remaining:.2f} tCO₂e per FTE.")
        else:
            st.error(f"Above target by {abs(remaining):.2f} tCO₂e per FTE.")

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

# Only valid distance categories are plotted. Blank and unknown records are excluded.
flight_type_selected = selected[selected["Flight Type"].isin(flight_type_display)].copy()
flight_type_selected["Flight distance"] = flight_type_selected["Flight Type"].map(flight_type_display)
by_flight_type = (
    flight_type_selected.groupby("Flight distance", as_index=False, observed=True)
    .agg(Flights=("Emissions_tCO2e", "size"), Emissions_tCO2e=("Emissions_tCO2e", "sum"))
)
by_flight_type["Flight distance"] = pd.Categorical(
    by_flight_type["Flight distance"], categories=flight_type_order, ordered=True
)
by_flight_type = by_flight_type.sort_values("Flight distance")

pie_left, pie_right = st.columns(2)
with pie_left:
    total_flights = int(by_flight_type["Flights"].sum())
    fig_flights_pie = px.pie(
        by_flight_type, names="Flight distance", values="Flights",
        title=f"Flights ({total_flights:,})",
        category_orders={"Flight distance": flight_type_order},
        color="Flight distance", color_discrete_map=flight_type_colors,
    )
    fig_flights_pie.update_traces(
        sort=False, direction="clockwise", textposition="outside",
        texttemplate="%{percent:.1%}<br>(%{value:,.0f})",
        hovertemplate="%{label}<br>Flights: %{value:,.0f}<br>Share: %{percent:.1%}<extra></extra>",
    )
    fig_flights_pie.update_layout(
        height=500, margin=dict(l=35, r=35, t=80, b=35),
        legend_title_text="Flight distance", uniformtext_minsize=10,
        uniformtext_mode="hide",
    )
    st.plotly_chart(fig_flights_pie, use_container_width=True)

with pie_right:
    total_emissions = float(by_flight_type["Emissions_tCO2e"].sum())
    fig_emissions_pie = px.pie(
        by_flight_type, names="Flight distance", values="Emissions_tCO2e",
        title=f"Emissions ({total_emissions:,.2f} tCO₂e)",
        category_orders={"Flight distance": flight_type_order},
        color="Flight distance", color_discrete_map=flight_type_colors,
    )
    fig_emissions_pie.update_traces(
        sort=False, direction="clockwise", textposition="outside",
        texttemplate="%{percent:.1%}<br>(%{value:,.2f} tCO₂e)",
        hovertemplate="%{label}<br>Emissions: %{value:,.2f} tCO₂e<br>Share: %{percent:.1%}<extra></extra>",
    )
    fig_emissions_pie.update_layout(
        height=500, margin=dict(l=35, r=35, t=80, b=35),
        showlegend=False, uniformtext_minsize=10, uniformtext_mode="hide",
    )
    st.plotly_chart(fig_emissions_pie, use_container_width=True)

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
    annual_relative = annual.dropna(subset=["Relative_tCO2e_per_FTE"]).copy()
    annual_relative = annual_relative[annual_relative["FTE"].gt(0)].copy()

    # Baseline is average 2023/2024 emissions divided by average 2023/2024 FTE.
    baseline_emissions = filtered_all_years[filtered_all_years["Year"].isin([2023, 2024])].groupby("Year")["Emissions_tCO2e"].sum().reindex([2023, 2024])
    baseline_fte = fte[fte["Year"].isin([2023, 2024])].set_index("Year")["FTE"].reindex([2023, 2024])
    pathway_start_value = float(baseline_emissions.mean() / baseline_fte.mean()) if baseline_emissions.notna().all() and baseline_fte.notna().all() and baseline_fte.mean() != 0 else None
    pathway_target_value = pathway_start_value / 2 if pathway_start_value is not None else None

    fig_annual = go.Figure()
    fig_annual.add_bar(x=annual_relative["Year"], y=annual_relative["Relative_tCO2e_per_FTE"], name="Emissions per FTE", marker_color="#4F81BD", text=annual_relative["Relative_tCO2e_per_FTE"].round(2), textposition="outside", hovertemplate="%{x}: %{y:.2f} tCO₂e/FTE<extra></extra>")
    if pathway_start_value is not None and pathway_target_value is not None:
        pathway_years = list(range(pathway_start_year, pathway_target_year + 1))
        annual_step = (pathway_target_value - pathway_start_value) / (pathway_target_year - pathway_start_year)
        pathway_values = [pathway_start_value + annual_step * (year - pathway_start_year) for year in pathway_years]
        pathway_labels = [f"{value:.2f}" for value in pathway_values]
        fig_annual.add_scatter(x=pathway_years, y=pathway_values, mode="lines+text", text=pathway_labels, textposition="top center", textfont=dict(color="#2F5597", size=11), cliponaxis=False, name="Emissions per FTE target", line=dict(color="#2F5597", width=3, dash="dot"), hovertemplate="%{x}: %{y:.2f} tCO₂e/FTE<extra></extra>")
    chart_years = sorted(set(annual["Year"].astype(int).tolist() + list(range(2024, 2031))))
    fig_annual.update_xaxes(tickmode="array", tickvals=chart_years, ticktext=[str(year) for year in chart_years], title="Year")
    fig_annual.update_yaxes(title="tCO₂e / FTE", rangemode="tozero")
    fig_annual.update_layout(title="Annual emissions per FTE", height=420, margin=dict(l=20, r=20, t=60, b=20), legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5), bargap=0.55)
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

