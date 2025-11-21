# Safe2Day Document Processing System

This system processes Word documents with the S2D prefix, allowing you to:

1. Analyze logo positions in documents
2. Select client logos
3. Generate branded document copies and PDFs

**IMPORTANT: Source documents in Forms folder are NEVER modified. All edits are done on copies.**

## Project Structure

```
Safe2Day/
├── Forms/              # SOURCE documents with S2D prefix (NEVER modified)
├── Client_Forms/       # Client copies with replaced logos (created automatically)
├── Client_Logos/       # Client logo images
├── PDFs/              # Generated PDF output (created automatically)
├── .venv/             # Python virtual environment
├── main.py            # Main orchestrator script
├── markup_document.py # Analyzes logo positions (does NOT modify source)
├── select_logo.py     # Logo selection interface
├── format_documents.py # Creates copies with logos and PDFs
└── requirements.txt   # Python dependencies
```

## Setup

1. **Activate virtual environment:**
   ```powershell
   .venv\Scripts\Activate.ps1
   ```

2. **Install dependencies (already done):**
   ```powershell
   pip install -r requirements.txt
   ```

## Workflow

### One-Time Setup: Analyze Logo Positions

Run this once to analyze existing logo positions in your documents:

```powershell
python markup_document.py
```

This will:

- Scan all `S2D*.docx` files in the Forms folder
- Identify existing logo positions in headers/footers
- Save position data to `logo_positions.json`
- **Does NOT modify any source documents**

### Regular Use: Generate Branded PDFs

Run the main script to create branded documents:

```powershell
python main.py
```

This will:

1. Display available logos from `Client_Logos` folder
2. Let you select a client logo
3. Copy source documents to `Client_Forms` folder
4. Replace logos in the copies (NOT source documents)
5. Create PDF versions in the `PDFs` folder

## Manual Component Scripts

You can also run components individually:

### Select a Logo
```powershell
python select_logo.py
```

### Generate PDFs with Selected Logo
```powershell
python format_documents.py
```

## Requirements

- Python 3.13
- python-docx
- docx2pdf
- Pillow
- Microsoft Word (required for docx2pdf conversion on Windows)

## Notes

- All Word documents must start with "S2D" prefix
- Supported logo formats: .png, .jpg, .jpeg, .bmp, .gif, .tiff
- **Source documents in Forms folder are NEVER modified**
- Client copies are saved to `Client_Forms` folder
- PDFs are automatically saved to the `PDFs` folder
- Logo selection is saved to `selected_logo.json`
- Logo positions are saved to `logo_positions.json`
