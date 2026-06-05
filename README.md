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

## Google Colab

Upload this repository to Colab, then run:

```python
!pip install -r requirements.txt
!pip install streamlit
```

Start the app with a tunnel such as localtunnel:

```python
!npm install -g localtunnel
!streamlit run app.py --server.headless=true --server.port=8501 & npx localtunnel --port 8501
```

Then upload the Excel workbook from the Streamlit sidebar. You can also upload a workbook to `/content` with Colab's file upload tool or mount Google Drive and enter the workbook path, such as `/content/drive/MyDrive/remittances.xlsx`.
