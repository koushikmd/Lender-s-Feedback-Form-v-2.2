# Lender Feedback Tool — Desktop Edition

**Offline, self-contained Windows application** for automating IDLC Finance's Lender's Feedback Form generation.

---

## 🔒 Security & Privacy

- **Fully offline** — runs on `127.0.0.1` (localhost) only
- **Never accessible from the network** — bound to loopback address
- **No data transmission** — all PDF processing happens on your machine
- **No persistent storage** — files are processed in memory and temporary files are deleted immediately
- **No Python installation required** for end users — the `.exe` bundles everything

---

## 📦 What's in this folder

| File | Purpose |
|---|---|
| `app.py` | Main Flask application (the tool) |
| `pdf_extractor.py` | PDF extraction engine |
| `docx_generator.py` | Word document generator |
| `build.spec` | PyInstaller build configuration |
| `build_windows.bat` | Build script — creates `LenderFeedbackTool.exe` |
| `requirements.txt` | Python dependencies |
| `README.md` | This file |

---

## 🚀 How to build the .exe (one-time, on any Windows PC with Python)

1. **Install Python 3.10+** from <https://python.org> (check "Add Python to PATH" during install)
2. **Copy this entire folder** to the Windows PC
3. **Double-click `build_windows.bat`**
4. Wait 2–5 minutes. The build creates `dist\LenderFeedbackTool.exe` (~40–50 MB)

---

## 👩‍💻 How end users run the tool

1. Copy `LenderFeedbackTool.exe` to any Windows PC (no installation needed)
2. **Double-click it**
3. A console window opens saying `Starting local server at http://127.0.0.1:5050/`
4. The browser automatically opens to the tool
5. Upload Appraisal + ISBS PDFs → Review/edit → Download DOCX
6. When done, **close the console window** to stop the server

---

## 🛠 How to run from source (if Python is available)

```bash
pip install -r requirements.txt
python app.py
```

The tool opens at `http://127.0.0.1:5050` automatically.

---

## 📋 Workflow

1. **Step 1 — Upload PDFs**: Drop the Appraisal and ISBS PDFs
2. **Step 2 — Review & Edit**: 18 fields are extracted automatically; edit any inaccurate values
3. **Step 3 — Download DOCX**: Generate and download the filled Lender's Feedback Form

---

## 🏦 Supported Banks & FIs

Bank names are extracted dynamically from PDF table columns — **no hardcoded list**. Confirmed working with:

- IDLC Finance Limited
- BRAC Bank Limited
- Mutual Trust Bank Limited
- The City Bank Limited
- Uttara Bank Limited
- Eastern Bank Limited
- Al Arafah Islami Bank Limited
- Agrani Bank Limited
- Standard Bank Limited
- Standard Chartered Bank
- United Commercial Bank Limited
- …and any other FI that appears in the liability tables.

---

## ❓ Troubleshooting

**"Windows protected your PC" popup on first run**
→ Click "More info" → "Run anyway". This happens because the .exe isn't signed with a commercial certificate. This is safe — it's your own build.

**Browser doesn't open automatically**
→ Manually visit `http://127.0.0.1:5050` (or the port shown in the console)

**Port already in use**
→ The app auto-finds a free port between 5050–5100

**Extraction results have minor formatting issues** (e.g., "Person al Loan" instead of "Personal Loan")
→ These come from PDF text splitting. Edit the field in Step 2 before generating the DOCX.
