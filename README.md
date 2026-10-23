# SportsCom Reimbursement Generator

A small Streamlit app to generate standardized `.docx` cover letters for DoMS IIT Madras Sports and Cultural Committee expenses.

## Letter types
- **Payment to Self** — includes fixed reimbursement bank details.
- **Payment to Vendor** — includes the vendor name and asks for direct payment; bank details are omitted.

## Run locally
Requires Python 3.10 or newer.

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Publish on GitHub and deploy
1. Create a new GitHub repository.
2. Upload `app.py`, `requirements.txt`, and `README.md`.
3. On [Streamlit Community Cloud](https://share.streamlit.io/), connect your GitHub account and choose this repository.
4. Set the main file path to `app.py` and deploy.

## Fixed details
The sender, recipient, signature, committee name, and self-reimbursement bank details are constants near the top of `app.py`. Edit them there if authorized details change.

## Checks included
- Requires an event/purpose and at least one complete expense line.
- Requires a vendor name for vendor payments.
- Checks that the sum of line items matches the entered total.
- Generates the DOCX in memory; it does not save uploaded bills or generated documents on the app server.

## Important
Review the generated letter and attach the relevant invoices/bills before submitting. The app generates the cover letter only; it does not attach receipts automatically.
