# PDF to Excel Converter

A GUI application that converts PDF shipping reports to Excel spreadsheets with formatted data.

## Features

- 🖥️ User-friendly GUI built with ttkbootstrap
- 📄 Extracts data from PDF shipping reports
- 📊 Generates formatted Excel files with multiple sheets
- 🎨 Professional styling with colored headers
- 🔄 Real-time conversion progress tracking

## Installation

### Prerequisites

- Python 3.10 or higher
- pip package manager

### Setup

1. Clone the repository:
```bash
git clone https://github.com/apayziev/pdf-to-excel-converter.git
cd pdf-to-excel-converter
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Running the Application

**GUI Mode:**
```bash
python gui_settings.py
```

**Command Line Mode:**
```bash
python extract_to_excel.py "path/to/your/file.pdf"
```

### Building Executable

To create a standalone executable:

**Windows:**
```cmd
build.bat
```

**Linux/Mac:**
```bash
./build.sh
```

The executable will be created in the `dist/` directory as a **single file** - no `_internal` folder needed!

## How It Works

1. Select a PDF file using the "Browse PDF" button
2. Click "Start Conversion"
3. Monitor the progress in the log area
4. The generated Excel file path will appear at the bottom
5. Application will close automatically after 3 seconds

## Troubleshooting

### Issue: New window appears during conversion

**Fixed:** The application now properly suppresses subprocess console windows on Windows.

### Issue: Slow startup time

**Note:** The single-file executable extracts dependencies to a temporary folder on startup, which may cause a slight delay on first run. This is normal behavior.

### Issue: Missing imports error

**Fixed:** All required imports have been added to the codebase.

## GitHub Actions

The project includes automated builds for Windows:

- Automatically builds on push to main/master branch
- Can be triggered manually via workflow_dispatch
- Produces downloadable artifacts

## License

MIT License

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.
