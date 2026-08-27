# WA Emissions Streamlit Dashboard 


## Files

- `app.py`: Streamlit dashboard
- 3 main excel files 
- `requirements.txt`: Python dependencies

## How the app uses the workbook

The app reads from main excel file:

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

...

## Deploy online

Upload these files to a GitHub repository, then deploy `app.py` on Streamlit Community Cloud.
