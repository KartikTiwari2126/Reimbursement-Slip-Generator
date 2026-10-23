# SportsCom Reimbursement Generator

A Streamlit app to generate standardized `.docx` cover letters for DoMS IIT Madras Sports and Cultural Committee expenses.

## Features
- **Payment to Self** — editable bank details, pre-filled with Aadarsh Kumar's sample details.
- **Payment to Vendor** — direct payment request to the entered vendor; no bank details included.
- Total amount is calculated from the line items whenever an item amount changes, and remains editable manually. If you override the total, changing an item amount recalculates it again; the **Recalculate total from items** button is also available.
- Generated letters use Times New Roman, 10 pt, single line spacing, compact margins, and zero paragraph spacing to help fit on one page.

## Run locally
Requires Python 3.10 or newer.

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

On Windows, if `python` is not recognized, try `py -m pip install -r requirements.txt` and `py -m streamlit run app.py`.

## Deploy from GitHub with Streamlit Community Cloud
This is a Streamlit app, not a static website. **GitHub Pages cannot run the Python app.**

1. Create a GitHub repository.
2. Extract this ZIP and upload the contents of the `sportscom_reimbursement_generator` folder (`app.py`, `requirements.txt`, and `README.md`) to the repository root. Do not upload only the ZIP.
3. Open [Streamlit Community Cloud](https://share.streamlit.io/) and sign in with GitHub.
4. Create an app, select the repository and branch, and set the main file path to `app.py`.
5. If files are in a subfolder, use that relative path, e.g. `sportscom_reimbursement_generator/app.py`.

## Troubleshooting
- If Streamlit cannot find `app.py`, check the main file path and repository folder structure.
- If deployment fails, open the app's Manage app / Logs view.
- Confirm `requirements.txt` is beside `app.py`.

## Fixed details and privacy
Sender, recipient, signature, and committee name are constants near the top of `app.py`. Bank details are editable in the app and included only for Payment to Self.
The app generates the document in memory; it does not save bills or generated letters on the server.

## Before submitting
Review the generated letter and attach the relevant invoices/bills separately. The app creates the cover letter only.
