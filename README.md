## Word Wizard Setup Guide

### 🖥️ Windows Instructions

```
git clone https://github.com/Sarthak-Sahu-1409/Word-Wizard.git
cd Word-Wizard

python -m venv venv
.\venv\Scripts\Activate.ps1    # or .\venv\Scripts\activate for cmd

pip install --upgrade pip
pip install customtkinter python-docx pypdf docx2pdf pywin32
python app.py
```

### 🍏 MacOS\Linux Instructions

```
git clone https://github.com/Sarthak-Sahu-1409/Word-Wizard.git
cd Word-Wizard

python3 -m venv venv
source venv/bin/activate

pip3 install --upgrade pip
pip3 install customtkinter python-docx pypdf docx2pdf
# On macOS, docx2pdf requires Microsoft Word for reliable conversion.
# On Linux, install LibreOffice if you want docx→pdf conversion support.
python3 app.py
```

![App Preview](App%20Preview.png "App Preview Screenshot")

## Key Features

1. **Live connect to active Word document (Windows only):** Detects and uses the active MS Word document as the base for appendices.  
2. **Add PDF appendices & select pages:** Add PDFs and specify page ranges like `1-3,6,9-10` to append only selected pages.  
3. **Reorder appendices (Move Up / Move Down):** Move a selected appendix up or down with list and selection kept in sync.  
4. **Rename appendix titles & auto heading pages:** Double-click to rename an appendix; generates a centered title page for each appendix.  
5. **Generate merged PDF with progress & validation:** Converts base doc to PDF, merges heading pages and selected PDF pages, shows progress, and prompts to save.
