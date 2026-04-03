# PMRSI-PDF-to-pdf-partlists

A desktop application that extracts structured parts table data from Porsche-style catalog PDFs and exports it to professionally formatted Excel workbooks.

---

## Features

- **PDF Parsing** -- Extracts Category, Item Number, Description, Material designation, Quantity, Material code, and Page numbers
- **Drag & Drop** -- Load PDFs by dragging them into the application window
- **Formatted Excel Output** -- Generates styled `.xlsx` files with headers, table formatting, and frozen panes
- **Auto-Open** -- Optionally opens the generated Excel file after conversion
- **Threaded Processing** -- UI stays responsive during conversion with real-time log output

## Screenshot

<!-- Add a screenshot of the application here -->
<!-- ![App Screenshot](docs/screenshot.png) -->

## Prerequisites

- Python 3.6 or higher
- pip package manager

## Installation

1. **Clone the repository**

   ```bash
   git clone https://github.com/<your-username>/pdf-to-excel.git
   cd pdf-to-excel
   ```

2. **Create a virtual environment**

   ```bash
   python -m venv .venv
   ```

3. **Activate the virtual environment**

   - **Windows:**
     ```bash
     .venv\Scripts\activate
     ```
   - **macOS / Linux:**
     ```bash
     source .venv/bin/activate
     ```

4. **Install dependencies**

   ```bash
   pip install -r requirements.txt
   ```

## Usage

Run the application:

```bash
python pdf_to_excel.py
```

### Steps

1. **Load a PDF** -- Drag and drop a Porsche parts catalog PDF onto the drop zone, or click **Browse** to select one.
2. **Convert** -- Click **Convert** to parse the PDF. Progress and logs are shown in real time.
3. **Save** -- The Excel file is saved to the `Excels/` folder by default. Use **Save As** to choose a different location.
4. **Auto-Open** -- Check the "Open Excel after save" option to automatically open the file when conversion completes.

## Project Structure

```
pdf-to-excel/
├── pdf_to_excel.py     # Main application (GUI, parser, Excel export)
├── requirements.txt    # Python dependencies
├── Excels/             # Default output directory for generated .xlsx files
└── README.md
```

## Dependencies

| Package | Purpose |
|---|---|
| [PyMuPDF](https://pymupdf.readthedocs.io/) | PDF text extraction |
| [openpyxl](https://openpyxl.readthedocs.io/) | Excel file creation and formatting |
| [tkinterdnd2](https://github.com/pmgagne/tkinterdnd2) | Native drag-and-drop support for Tkinter |

## Limitations

- Designed specifically for Porsche parts catalog PDF layouts. Other PDF formats may require adjustments to the parsing logic.
- Processes one PDF at a time (no batch mode).

## License

This project is provided as-is. See [LICENSE](LICENSE) for details.
