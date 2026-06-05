# Remittance Review Dashboard

Streamlit dashboard for reviewing remittance Excel workbooks with:

- Balance remittance analysis
- No CusDec / not-written-off risk indicators
- BNM-CNM and consignee-declarant relationship mapping
- BNM vs imports review by customs key
- Exportable CSV tables

## Run

```powershell
py -m pip install -r requirements.txt
py -m streamlit run app.py
```

Use the sidebar uploader to load an Excel workbook, or configure the default workbook path in `app.py` for local use.
