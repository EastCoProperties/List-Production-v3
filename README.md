
# Underwriting Engine V2.1

A Streamlit web app for processing vacant land acquisition lists and generating a confidence-weighted offer model.

## Key V2.1 Change

This version improves how extremely low county assessed values are handled.

If **Assessed Reliability Ratio < 0.40**, the assessed value is treated as likely deferred/ag/timber/conservation-distorted.

When this happens:

- County assessed value receives **0% weight** in Offer MV Used
- County assessed value is **excluded from the minimum offer floor**
- County assessed value is **excluded from the 600% assessed review flag**
- The row is flagged as **Review - Likely Deferred Assessed**
- The model places more reliance on:
  - TLP Estimate
  - Parent Median Value

This prevents unrealistically low assessed values from creating unrealistically low offers.

## Manual Adjustment

The output workbook is designed so that **Offer MV Used** can be manually edited in Excel.

When Offer MV Used changes, the following recalculate automatically:

- Development Cost Used
- 1st Offer
- 2nd Offer
- 3rd Offer
- % MV
- Profit
- ROI
- Review Flag
- Deal Quality Score

## Acreage Buckets

- 1.00–1.99
- 2.00–2.99
- 3.00–4.99
- 5.00–7.49
- 7.50–9.99
- 10.00–14.99
- 15.00–24.99
- 25.00–49.99
- 50.00–99.99
- 100+

## Run Locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy to Streamlit Cloud

Upload these files to your GitHub repo root:

- `app.py`
- `requirements.txt`
- `README.md`
- `.gitignore`

Then deploy with main file path:

- `app.py`
