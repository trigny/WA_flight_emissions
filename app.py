from pathlib import Path
import re
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="MyClimate Dashboard", page_icon="📊", layout="wide")
BASE = Path(__file__).resolve().parent
EXCEL_FILE = BASE / "Flight Emissions Dashboard.xlsx"
CUSTOM_FIELDS_FILE = BASE / "Custom_Fields_2026-06.xlsx"
PROJECT_OPTIONS_FILE = BASE / "Custom field options_new.csv"
TARGET_BASE_YEAR, TARGET_YEAR, TARGET_REDUCTION = 2024, 2030, 0.50


def clean_key(value):
    if pd.isna(value):
        return ""
    value = str(value).strip()
    return "" if value.lower() in {"", "0", "0.0", "nan", "none", "<na>"} else value


def clean_series(series, blank="Unassigned"):
    result = series.fillna("").astype(str).str.strip()
    result = result.replace({"0": "", "0.0": "", "nan": "", "None": "", "<NA>": ""})
    return result.mask(result.eq(""), blank)


def unique_map(frame, key, value):
    data = frame[[key, value]].copy()
    data[key] = data[key].map(clean_key)
    data[value] = data[value].map(clean_key)
    data = data[(data[key] != "") & (data[value] != "")]
    grouped = data.groupby(key)[value].agg(lambda x: sorted(set(x)))
    return {k: v[0] for k, v in grouped.items() if len(v) == 1}


def target_for_year(year, base_value):
    """Linear pathway: 2024 actual to exactly 50% lower in 2030."""
    if base_value is None or pd.isna(base_value):
        return None
    if year <= TARGET_BASE_YEAR:
        return float(base_value)
    end_value = float(base_value) * (1 - TARGET_REDUCTION)
    if year >= TARGET_YEAR:
        return end_value
    fraction = (year - TARGET_BASE_YEAR) / (TARGET_YEAR - TARGET_BASE_YEAR)
    return float(base_value) + (end_value - float(base_value)) * fraction


def canonical_project(value, valid_codes):
    code = clean_key(value)
    if code in valid_codes:
        return code
    parent = re.sub(r"-SP\d+$", "", code, flags=re.I)
    return parent if parent in valid_codes else ""


@st.cache_data(show_spinner="Reading repository files...")
def load_data(workbook_mtime, custom_mtime, options_mtime):
    all_data = pd.read_excel(EXCEL_FILE, sheet_name="All Integrated Data", engine="openpyxl")
    traveler = pd.read_excel(EXCEL_FILE, sheet_name="Traveler Manifest", header=8, engine="openpyxl")
    legacy = pd.read_excel(EXCEL_FILE, sheet_name="Legacy MyClimate Import", engine="openpyxl")
    fte = pd.read_excel(EXCEL_FILE, sheet_name="FTE Data", engine="openpyxl")
    custom = pd.read_excel(CUSTOM_FIELDS_FILE, header=6, engine="openpyxl")
    options = pd.read_csv(PROJECT_OPTIONS_FILE, dtype=str)

    options.columns = options.columns.astype(str).str.strip()
    if not {"Name", "Description"}.issubset(options.columns):
        raise ValueError("Project options CSV requires Name and Description columns")
    options["Name"] = options["Name"].map(clean_key)
    options = options[options["Name"] != ""].drop_duplicates("Name")
    valid_codes = set(options["Name"])
    descriptions = options.set_index("Name")["Description"].fillna("").to_dict()

    project_rows = custom[
        custom["Custom Question"].fillna("").astype(str).str.strip().eq("(UD15) Project Codes")
    ].copy()
    project_rows["Project"] = project_rows["Travel Data Answer"].map(
        lambda x: canonical_project(x, valid_codes)
    )
    project_rows["TX"] = project_rows["Travel Data Transaction Key"].map(clean_key).str.replace(
        r"-Q\d+$", "", regex=True
    )
    for source, key in [("Trip ID", "TRIP"), ("Spotnana PNR ID", "PNR"), ("Confirmation Number", "TICKET")]:
        project_rows[key] = project_rows[source].map(clean_key)
    project_rows = project_rows[project_rows["Project"] != ""]
    maps = {key: unique_map(project_rows, key, "Project") for key in ["TX", "TRIP", "PNR", "TICKET"]}

    def traveler_project(row):
        candidates = [
            maps["TX"].get(clean_key(row.get("Transaction Key")), ""),
            maps["PNR"].get(clean_key(row.get("Spotnana PNR ID")), ""),
            maps["TICKET"].get(clean_key(row.get("Ticket Number")), ""),
            maps["TRIP"].get(clean_key(row.get("Trip ID")), ""),
        ]
        return next((x for x in candidates if x), "")

    traveler["Resolved Project"] = traveler.apply(traveler_project, axis=1)
    legacy["Resolved Project"] = legacy["Projektnummer"].map(lambda x: canonical_project(x, valid_codes))

    def integrated_project(row):
        try:
            index = int(float(row["Calc_or_Source_Row"])) - 2
        except (ValueError, TypeError, KeyError):
            return ""
        source = clean_key(row.get("Record_Source"))
        if source == "Traveler Manifest" and 0 <= index < len(traveler):
            return clean_key(traveler.iloc[index]["Resolved Project"])
        if source == "Legacy MyClimate Import" and 0 <= index < len(legacy):
            return clean_key(legacy.iloc[index]["Resolved Project"])
        return ""

    all_data["Project Number"] = all_data.apply(integrated_project, axis=1)
    all_data["Project Description"] = all_data["Project Number"].map(descriptions).fillna("")
    data = all_data[all_data["Include_Final"].astype(str).str.strip().str.lower().eq("yes")].copy()
    data["Date"] = pd.to_datetime(data["Date"], errors="coerce")
    data["Year"] = pd.to_numeric(data["Year"], errors="coerce")
    data = data.dropna(subset=["Year"])
    data["Year"] = data["Year"].astype(int)
    data["Emissions"] = pd.to_numeric(data["Final_RFI3_tCO2e"], errors="coerce").fillna(0)
    data["Distance"] = pd.to_numeric(data["Distance_km"], errors="coerce").fillna(0)
    data["Cabin"] = clean_series(data["Class"], "Unknown").str.lower()
    data["Team"] = clean_series(data["Team"], "External")
    data["Flight Type"] = data["Flight_Type"].fillna("").astype(str).str.strip().str.lower().replace({
        "very_short_haul": "Very short haul", "short_haul": "Short haul",
        "medium_haul": "Medium haul", "long_haul": "Long haul",
    })
    data["Project Number"] = clean_series(data["Project Number"], "Unassigned")
    data["Month"] = data["Date"].dt.month
    data["Month Name"] = data["Date"].dt.strftime("%b")

    fte["Year"] = pd.to_numeric(fte["Year"], errors="coerce")
    fte["FTE"] = pd.to_numeric(fte["FTE"], errors="coerce")
    fte = fte.dropna(subset=["Year"])
    fte["Year"] = fte["Year"].astype(int)
    return data, fte


for file in [EXCEL_FILE, CUSTOM_FIELDS_FILE, PROJECT_OPTIONS_FILE]:
    if not file.exists():
        st.error(f"Missing repository file: {file.name}")
        st.stop()

try:
    flights, fte = load_data(
        EXCEL_FILE.stat().st_mtime, CUSTOM_FIELDS_FILE.stat().st_mtime, PROJECT_OPTIONS_FILE.stat().st_mtime
    )
except Exception as exc:
    st.error("The dashboard could not integrate the repository data files.")
    st.exception(exc)
    st.stop()

annual_all = flights.groupby("Year", as_index=False).agg(
    Flights=("Emissions", "size"), Emissions=("Emissions", "sum"), Distance=("Distance", "sum")
).merge(fte[["Year", "FTE"]], on="Year", how="left")
annual_all["Emissions per FTE"] = annual_all["Emissions"] / annual_all["FTE"]
base_rows = annual_all.loc[annual_all["Year"].eq(TARGET_BASE_YEAR), "Emissions per FTE"]
target_base_value = float(base_rows.iloc[0]) if len(base_rows) and pd.notna(base_rows.iloc[0]) else None
target_pathway = {year: target_for_year(year, target_base_value) for year in range(2024, 2031)}

st.title("📊 Wyss Academy Flight Emissions Dashboard")
years = sorted(flights["Year"].unique())
with st.sidebar:
    selected_year = st.selectbox("Analysis year", years, index=len(years) - 1)
    cabins = sorted(flights["Cabin"].unique())
    selected_cabins = st.multiselect("Cabin class", cabins, default=cabins)
    teams = sorted(flights["Team"].unique())
    selected_teams = st.multiselect("Teams", teams, default=teams)
    if st.button("Clear cache and reload data"):
        st.cache_data.clear()
        st.rerun()

filtered = flights[flights["Cabin"].isin(selected_cabins) & flights["Team"].isin(selected_teams)]
selected = filtered[filtered["Year"].eq(selected_year)]
fte_map = fte.dropna(subset=["FTE"]).drop_duplicates("Year", keep="last").set_index("Year")["FTE"].to_dict()
selected_fte = fte_map.get(selected_year)
selected_emissions = selected["Emissions"].sum()
metrics = st.columns(4)
metrics[0].metric("Flights", f"{len(selected):,}")
metrics[1].metric("Emissions", f"{selected_emissions:.1f} tCO₂e")
metrics[2].metric("Distance", f"{selected['Distance'].sum():,.0f} km")
metrics[3].metric("Emissions per FTE", "n/a" if not selected_fte else f"{selected_emissions/selected_fte:.2f} tCO₂e/FTE")

st.divider()
st.subheader("Project planning estimate")
planning_years = list(range(min(years), TARGET_YEAR + 1))
c1, c2 = st.columns(2)
with c1:
    planning_year = st.selectbox("Planning year", planning_years, index=planning_years.index(selected_year if selected_year in planning_years else TARGET_YEAR))
default_fte = float(fte_map.get(planning_year, max(fte_map.values())))
with c2:
    planned_fte = st.number_input(
        "Planned project FTE", min_value=0.1, max_value=1000.0, value=default_fte,
        step=5.0, format="%.1f", key=f"planned_fte_{planning_year}",
        help="Defaults to FTE Data for the selected year. Plus/minus changes the scenario by 5 FTE.",
    )

valid_types = ["Very short haul", "Short haul", "Medium haul", "Long haul"]
valid_cabins = ["economy", "premiumeconomy", "business"]
reference = flights[flights["Flight Type"].isin(valid_types) & flights["Cabin"].isin(valid_cabins)]
factors = reference.groupby(["Flight Type", "Cabin"])["Emissions"].agg(mean="mean", records="size")
labels = {"Very short haul": "Very short haul (<500 km)", "Short haul": "Short haul (500–1,500 km)",
          "Medium haul": "Medium haul (1,500–4,000 km)", "Long haul": "Long haul (>4,000 km)"}
head = st.columns([2.2, 1, 1, 1, 1.4])
for col, title in zip(head, ["Flight distance", "Economy", "Premium economy", "Business", "Estimated tCO₂e"]):
    col.markdown(f"**{title}**")
planned_emissions, planned_segments = 0.0, 0
for flight_type in valid_types:
    row = st.columns([2.2, 1, 1, 1, 1.4])
    row[0].write(labels[flight_type])
    row_total = 0.0
    for col, cabin in zip(row[1:4], valid_cabins):
        key = (flight_type, cabin)
        factor = float(factors.loc[key, "mean"]) if key in factors.index else 0.0
        records = int(factors.loc[key, "records"]) if key in factors.index else 0
        with col:
            number = st.selectbox(labels[flight_type] + cabin, range(501), key=f"plan_{flight_type}_{cabin}", label_visibility="collapsed",
                                  help=f"All-years mean: {factor:.3f} tCO₂e per segment, based on {records} records.")
        planned_segments += number
        row_total += number * factor
    planned_emissions += row_total
    row[4].write(f"{row_total:.2f}")

target_rate = target_pathway.get(planning_year, target_for_year(planning_year, target_base_value))
planned_per_fte = planned_emissions / planned_fte
year_target = target_rate * planned_fte if target_rate is not None else None
remaining = year_target - planned_emissions if year_target is not None else None
k = st.columns(4)
k[0].metric("Planned one-way flights", f"{planned_segments:,}")
k[1].metric("Estimated project emissions", f"{planned_emissions:.2f} tCO₂e")
k[2].metric("Estimated emissions per FTE", f"{planned_per_fte:.2f} tCO₂e/FTE")
k[3].metric(f"{planning_year} target per FTE", "n/a" if target_rate is None else f"{target_rate:.2f} tCO₂e/FTE",
            delta=None if target_rate is None else f"{target_rate-planned_per_fte:+.2f} tCO₂e/FTE remaining")
b = st.columns(2)
b[0].metric(f"{planning_year} emissions target", "n/a" if year_target is None else f"{year_target:.2f} tCO₂e")
b[1].metric(f"Remaining allowance for {planning_year}", "n/a" if remaining is None else f"{remaining:.2f} tCO₂e",
            delta=None if remaining is None else ("Within target" if remaining >= 0 else f"{abs(remaining):.2f} tCO₂e over target"),
            delta_color="normal" if remaining is None or remaining >= 0 else "inverse")
if target_rate is not None:
    st.caption(f"{planning_year} target calculation: {target_rate:.2f} tCO₂e/FTE × {planned_fte:.1f} FTE = {year_target:.2f} tCO₂e. "
               "Flight factors use the unfiltered mean across all available years. The per-FTE target follows a linear pathway from the 2024 actual value to a 50% reduction in 2030.")

st.divider()
st.subheader(f"Emissions by project number ({selected_year})")
assigned = selected[selected["Project Number"].ne("Unassigned")]
unassigned = selected[selected["Project Number"].eq("Unassigned")]
project_summary = assigned.groupby(["Project Number", "Project Description"], as_index=False).agg(
    Flights=("Emissions", "size"), Emissions=("Emissions", "sum"), Distance=("Distance", "sum")
).sort_values("Emissions", ascending=False)
p = st.columns(3)
p[0].metric("Assigned flight records", f"{len(assigned):,}", f"{len(assigned)/len(selected):.1%} of selected" if len(selected) else "0%")
p[1].metric("Assigned emissions", f"{assigned['Emissions'].sum():.2f} tCO₂e")
p[2].metric("Unassigned emissions", f"{unassigned['Emissions'].sum():.2f} tCO₂e")
if len(project_summary):
    chart = project_summary.head(15).sort_values("Emissions")
    fig = px.bar(chart, x="Emissions", y="Project Number", orientation="h", text="Emissions",
                 custom_data=["Project Description", "Flights", "Distance"], title=f"Highest-emitting validated projects in {selected_year}")
    fig.update_traces(texttemplate="%{text:.2f}", textposition="outside",
                      hovertemplate="<b>%{y}</b><br>%{customdata[0]}<br>Emissions: %{x:.2f} tCO₂e<br>Flights: %{customdata[1]}<extra></extra>")
    st.plotly_chart(fig, use_container_width=True)
    st.dataframe(project_summary.round({"Emissions": 2, "Distance": 0}), use_container_width=True, hide_index=True)
else:
    st.info("No validated project assignments for the selected filters.")

st.divider()
st.subheader("Annual emissions per FTE")
fig = go.Figure()
annual_view = annual_all.dropna(subset=["Emissions per FTE"])
fig.add_bar(x=annual_view["Year"], y=annual_view["Emissions per FTE"], name="Actual", text=annual_view["Emissions per FTE"].round(2), textposition="outside")
path_years = list(target_pathway)
fig.add_scatter(x=path_years, y=[target_pathway[y] for y in path_years], name="Target", mode="lines+markers+text",
                text=[f"{target_pathway[y]:.2f}" for y in path_years], textposition="top center", line=dict(dash="dot"))
fig.update_yaxes(title="tCO₂e/FTE", rangemode="tozero")
st.plotly_chart(fig, use_container_width=True)
