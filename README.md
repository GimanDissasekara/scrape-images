# 💊 Medicine Photo Scraper & Document Inserter

A Python tool that automatically extracts medicine names from a Word document (`.docx`), downloads their photos from the internet, and inserts the images back into the document.

## 📋 Overview

This script reads a **BNF-style Drug Classification** document that contains photo placeholders like:

```
[ Photo: Warfarin tablets – 4 cm × 4 cm ]
```

It then:

1. **Extracts** all medicine names from these placeholders
2. **Downloads** a photo of each medicine from Bing/DuckDuckGo Images
3. **Saves** all photos into a `medicine_photos/` folder
4. **Inserts** the photos into a new copy of the document (4 cm × 4 cm, centered)

---

## 🛠️ Setup Instructions

### Prerequisites

- **Python 3.10+** installed on your system
- **pip** (Python package manager)

### Step 1 — Clone or Download the Project

Download the project folder to your machine.

### Step 2 — Open a Terminal

Open **PowerShell** (Windows) or **Terminal** (Mac/Linux) and navigate to the project folder:

```bash
cd path/to/photo-scrape-to-doc
```

### Step 3 — Create a Python Virtual Environment

```bash
# Create the virtual environment
python -m venv venv
```

### Step 4 — Activate the Virtual Environment

**Windows (PowerShell):**

```powershell
.\venv\Scripts\Activate.ps1
```

> ⚠️ If you get an execution policy error, run this first:
> ```powershell
> Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
> ```

**Windows (Command Prompt):**

```cmd
venv\Scripts\activate.bat
```

**Mac / Linux:**

```bash
source venv/bin/activate
```

You should see `(venv)` appear at the beginning of your terminal prompt.

### Step 5 — Install Dependencies

```bash
pip install -r requirements.txt
```

---

## 🚀 Usage

### Step 1 — Place Your Document

Make sure `Drug_Classification_BNF.docx` is in the project folder.

### Step 2 — Run the Script

```bash
python medicine_photo_scraper.py
```

### Step 3 — Wait for Completion

The script will:

- Extract all medicine names from the document
- Download photos (this takes a few minutes for ~75 medicines)
- Insert photos into the document

Progress is displayed in the terminal:

```
[1/75] Searching: Warfarin tablets
  ✓ Downloaded successfully
[2/75] Searching: Apixaban / Thromban pack
  ✓ Downloaded successfully
...
```

### Step 4 — Check the Output

Once complete, you will find:

| File/Folder | Description |
|---|---|
| `Drug_Classification_BNF_with_photos.docx` | The new document with photos inserted |
| `medicine_photos/` | Folder containing all downloaded images |
| `scraper.log` | Detailed log of the entire process |

> 📝 The original document is **never modified**. A new copy is created.

---

## 📁 Project Structure

```
photo-scrape-to-doc/
├── Drug_Classification_BNF.docx          # Input document (your source file)
├── Drug_Classification_BNF_with_photos.docx  # Output document (generated)
├── medicine_photo_scraper.py             # Main Python script
├── requirements.txt                      # Python dependencies
├── README.md                             # This file
├── scraper.log                           # Execution log (generated)
├── .env                                  # Environment variables (optional)
└── medicine_photos/                      # Downloaded photos (generated)
    ├── warfarin_tablets/
    │   └── photo.jpg
    ├── apixaban_thromban_pack/
    │   └── 000001.jpg
    └── ...
```

---

## ⚙️ Configuration

You can modify these settings at the top of `medicine_photo_scraper.py`:

| Variable | Default | Description |
|---|---|---|
| `DOCX_FILE` | `Drug_Classification_BNF.docx` | Input document filename |
| `OUTPUT_DOCX` | `Drug_Classification_BNF_with_photos.docx` | Output document filename |
| `PHOTO_DIR` | `medicine_photos` | Folder to save downloaded photos |
| `IMAGE_SIZE_CM` | `4` | Image size in the document (cm) |

---

## 🔁 Re-running the Script

The script is **idempotent** — if you run it again:

- Already downloaded photos are **skipped** (no re-download)
- The output document is **regenerated** from scratch

To force a fresh download of all photos, delete the `medicine_photos/` folder first.

---

## ❓ Troubleshooting

### `venv\Scripts\Activate.ps1 cannot be loaded`

Run this in PowerShell:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### `ModuleNotFoundError: No module named 'docx'`

Make sure you:

1. Activated the virtual environment (`(venv)` should be visible in your prompt)
2. Installed dependencies: `pip install -r requirements.txt`

### Some medicines show "No image found"

The script tries **Bing Images** first, then **DuckDuckGo** as a fallback. If both fail, the original placeholder text is kept in the document. You can manually add photos to the `medicine_photos/<medicine_name>/` folder and re-run.

### Deactivating the Virtual Environment

When you're done, simply run:

```bash
deactivate
```

---

## 📄 License

This project is for educational purposes only. Downloaded images may be subject to copyright.
