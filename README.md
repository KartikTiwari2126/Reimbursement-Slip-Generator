# SportsCom Reimbursement Generator

A Streamlit app to generate standardized `.docx` cover letters for DoMS IIT Madras Sports and Cultural Committee expenses.

## Letter types
- **Payment to Self** — includes editable bank details, pre-filled with Aadarsh Kumar's sample details.
- **Payment to Vendor** — asks for direct payment to the entered vendor and does not include bank details.

## Run locally (Windows, macOS, or Linux)
Install Python 3.10 or newer. Extract the project ZIP, open a terminal in the extracted project folder, then run:

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

If `python` is not recognized on Windows, try `py -m pip install -r requirements.txt` and `py -m streamlit run app.py`.

## Deploy from GitHub with Streamlit Community Cloud
**Important:** This is a Streamlit app, not a static GitHub Pages website. GitHub Pages cannot run the Python app. Use Streamlit Community Cloud to deploy it.

1. Sign in to GitHub and create a new repository (Public is simplest for a first deployment).
2. Extract this ZIP on your computer.
3. Open the repository on GitHub, select **Add file → Upload files**, and upload the *contents* of the `sportscom_reimbursement_generator` folder: `app.py`, `requirements.txt`, and `README.md`. Do not upload only the ZIP.
4. Commit the files to the repository's root. After upload, `app.py` and `requirements.txt` should be visible on the repository's main **Code** page, not inside another nested folder.
5. Open [Streamlit Community Cloud](https://share.streamlit.io/) and sign in with GitHub.
6. Select **Create app** (or **Deploy an app**) → choose the repository and branch → set **Main file path** to exactly `app.py` → deploy.
7. If the files are in a folder inside the repository rather than the root, set the main file path to `sportscom_reimbursement_generator/app.py` instead.

## Troubleshooting deployment
- **"File does not exist" / cannot find `app.py`:** the main file path must match where `app.py` actually lives in the repository. If `app.py` is at repository root, use `app.py`.
- **Blank or failed build:** open the app's deployment page and check **Manage app → Logs**. Confirm `requirements.txt` is in the same folder as `app.py`.
- **Trying to deploy on GitHub Pages:** this won't work for Streamlit/Python. Deploy through Streamlit Community Cloud instead.
- **ModuleNotFoundError:** check that `python-docx` and `streamlit` are listed in `requirements.txt`, then redeploy.
- **App doesn't update after a change:** commit/push the updated files and reboot the app from its Streamlit management menu.

## Fixed details and privacy
Sender, recipient, signature, and committee name are constants near the top of `app.py`. Change them there if authorized details change.
Bank details are entered in the UI for each self-reimbursement letter and pre-filled with Aadarsh Kumar's sample details. They are not included in vendor-payment letters.
The app generates the document in memory; it does not save bills or generated letters to the server.

## Before submitting
Review the generated letter and attach the relevant invoices/bills separately. This app creates the cover letter only.
