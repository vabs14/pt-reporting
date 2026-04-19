# PT Quarterly Report Generator — A&G Physiotherapy Inc.

Generates a professional Word (`.docx`) and PDF quarterly report from an Excel data file.
Can be run as a command-line script or hosted as a Streamlit web app.

---

## File naming convention

Excel files **must** follow this naming pattern:

```
PT_<Quarter>_<Year>_<HomeName>.xlsx
```

Examples:
- `PT_Q1_2026_BurtonManor.xlsx`
- `PT_Q2_2026_burton-manor.xlsx`
- `PT_Q3_2026_Sunrise_LTC.xlsx`

Place the file in the `data/` folder before running locally.

---

## Run locally (command line)

### 1. Prerequisites

Python 3.10+ and Homebrew are recommended on macOS.

```bash
# Install dependencies
/opt/homebrew/bin/pip3 install pandas openpyxl python-docx matplotlib pillow
```

For PDF export, install LibreOffice:
- Download from [libreoffice.org](https://www.libreoffice.org/download/download/) and install normally.

### 2. Run the script

```bash
cd src

# Auto-discovers the newest .xlsx in ../data/
/opt/homebrew/bin/python3 generate_pt_report.py

# Or pass a specific file
/opt/homebrew/bin/python3 generate_pt_report.py ../data/PT_Q1_2026_BurtonManor.xlsx
```

Output files are saved in `src/`:
```
src/PT_Q1_2026_Burton_Manor_Report.docx
src/PT_Q1_2026_Burton_Manor_Report.pdf
```

---

## Run locally (Streamlit web app)

### 1. Install Streamlit

```bash
/opt/homebrew/bin/pip3 install streamlit
```

### 2. Create the secrets file

```bash
mkdir -p .streamlit
```

Create `.streamlit/secrets.toml` with your users:

```toml
[users]
vabs    = "your_password"
brother = "their_password"
```

> This file is in `.gitignore` and will never be committed.

### 3. Start the app

```bash
/opt/homebrew/bin/python3 -m streamlit run app.py
```

Open [http://localhost:8501](http://localhost:8501) in your browser, sign in, upload the Excel file and click **Generate Report**.

---

## Deploy on Streamlit Community Cloud

### 1. Push to GitHub

```bash
git add .
git commit -m "deploy streamlit app"
git push origin main
```

### 2. Create the app

1. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
2. Click **New app**.
3. Fill in the fields:

| Field | Value |
|---|---|
| Repository | `your-github-username/pt-reporting` |
| Branch | `main` |
| Main file path | `app.py` |

4. Click **Advanced settings** and paste your secrets into the **Secrets** field:

```toml
[users]
vabs    = "your_password"
brother = "their_password"
```

5. Click **Deploy**.

### 3. Managing users after deployment

Go to your app → **⋮ menu → Settings → Secrets** and edit the `[users]` block directly. Changes take effect within seconds — no redeployment needed.

---

## Project structure

```
pt-reporting/
├── app.py                  # Streamlit web app
├── requirements.txt        # Python dependencies for Streamlit Cloud
├── packages.txt            # System packages (LibreOffice) for Streamlit Cloud
├── .gitignore
├── .streamlit/
│   └── secrets.toml        # Local credentials (never committed)
├── data/
│   └── PT_Q1_2026_*.xlsx   # Excel data files
└── src/
    ├── generate_pt_report.py   # Core report generation script
    └── logo_clean.png          # A&G Physiotherapy logo
```

---

## Adding or removing users

Edit `.streamlit/secrets.toml` locally, or update secrets in the Streamlit Cloud dashboard:

```toml
[users]
vabs    = "password"     # add a line to add a user
gaurav  = "password"     # remove a line to remove a user
```
