# MyClimate Streamlit Dashboard From Original Excel

This app keeps the original Excel workbook and reads directly from it each time the dashboard loads.

## Files

- `app.py`: Streamlit dashboard
- `MyClimate Methodology Workbook_final (1).xlsx`: original Excel workbook
- `requirements.txt`: Python dependencies

## How the app uses the workbook

The app reads:

- `All Integrated Data`
- `FTE Data`
- `Dashboard` for the default year/RFI control values where available

Then, inside Python, it keeps only dashboard-relevant fields:

- Date
- Year
- DepartureAirport
- ArrivalAirport
- Route
- Class
- Distance_km
- Emissions_tCO2e
- Month
- Month_name

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Update the dashboard data

Replace `MyClimate Methodology Workbook_final (1).xlsx` with the updated workbook, keeping the same filename. Then rerun the app or click **Clear cache and reload workbook** in the sidebar.

## Deploy online

Upload these files to a GitHub repository, then deploy `app.py` on Streamlit Community Cloud.

Important: if the GitHub repository is public, the full Excel workbook will also be public. Use a private repository if the workbook contains sensitive information.
