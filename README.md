# Document to PDF Converter

This script converts various document files (HWP, DOCX, PPTX, XLSX) to PDF format. It can also copy other specified file types directly to the destination folder.

## Important Note

**This script is designed to run on a Windows environment.** It relies on `pywin32` to interact with COM objects for Microsoft Office (Word, Excel, PowerPoint) and Hancom Office (HWP). These applications must be installed on the Windows machine where the script is executed.

## Features

- Converts HWP, HWPX, DOC, DOCX, PPT, PPTX, XLS, XLSX files to PDF.
- Copies JPG, JPEG, PNG, GIF, BMP, TXT, and existing PDF files to the destination.
- Recreates the source folder's directory structure in the destination folder.
- Logs conversion successes and failures.
- Provides a progress bar during file processing.

## Prerequisites

- Windows Operating System
- Microsoft Office (Word, Excel, PowerPoint) installed
- Hancom Office (HWP) installed
- Python 3.x

## Setup and Usage

1.  **Clone the repository:**
    ```bash
    git clone <repository_url>
    cd <repository_directory>
    ```

2.  **Create and activate a virtual environment using `uv`:**
    ```bash
    uv venv
    source .venv/bin/activate  # On Windows, use: .venv\Scripts\activate
    ```

3.  **Install dependencies:**
    ```bash
    uv pip install -r requirements.txt
    ```
    The `requirements.txt` file includes:
    - `pywin32`: For COM automation of Office applications (Windows-specific).
    - `tqdm`: For displaying progress bars.

4.  **Run the script:**
    ```bash
    python dtp.py
    ```
    The script will prompt you to enter:
    - Source folder path
    - Destination folder path
    - Optionally, specific extensions to convert (e.g., `hwp,docx,pptx`)

## Logging

A log file named `conversion_log_YYYYMMDD_HHMMSS.log` will be created in the script's directory, recording the details of the conversion process.

## Contributing

Contributions are welcome! Please feel free to submit a pull request or open an issue.
