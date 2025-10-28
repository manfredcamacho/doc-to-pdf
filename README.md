# Doc/Docx to PDF Converter

A simple, cross-platform Python script to batch convert `.doc` and `.docx` files to PDF using the LibreOffice engine.

It automatically finds your LibreOffice installation (on Windows, macOS, and Linux) and can process single files, multiple files, or entire directories at once.

## Requirements

1.  **Python 3**: (Usually pre-installed on Linux/macOS. For Windows, get it from [python.org](https://www.python.org/)).
2.  **LibreOffice**: The script *requires* LibreOffice to be installed. You can download it from [libreoffice.org](https://www.libreoffice.org/).

## How to Use

1.  **Install LibreOffice** (if you haven't already).
2.  **Download the Script**: Save the `doc2pdf.py` file to your computer.
3.  **Open Your Terminal** (Command Prompt, PowerShell, or Terminal).
4.  **Run the script**:
    ```bash
    python doc2pdf.py [OPTIONS] <input_file_or_dir_1> [input_2] ...
    ```

### Options

* `inputs`: (Required) One or more paths to the files or directories you want to convert.
* `-o, --output-dir <PATH>`: (Optional) Specify a directory to save all converted PDFs. If omitted, PDFs are saved in the same directory as their original source file.

## Examples

### 1. Convert a single file

Converts `report.docx` and saves it as `report.pdf` in the *same folder*.

```bash
python doc2pdf.py "C:\Users\MyUser\Documents\report.docx"
```

### 2. Convert multiple specific files
Converts both files, saving each PDF in its respective original folder.

```bash
python doc2pdf.py ./docs/file_one.doc ./other/file_two.docx
```

### 3. Convert all files in a directory
Finds all .doc and .docx files inside the Monthly_Reports folder and saves the new PDFs inside that same folder.

```bash
python doc2pdf.py "/home/myuser/Documents/Monthly_Reports"
```

### 4. Convert files to a specific output directory
Converts report.docx and saves the resulting report.pdf inside the C:\Converted_PDFs folder.

```bash
python doc2pdf.py -o "C:\Converted_PDFs" "C:\Users\MyUser\Documents\report.docx"
```

### 5. Convert mixed inputs to one output directory
Converts file.doc and all documents inside the Old_Reports folder, saving all resulting PDFs into the All_PDFs folder.

```bash
python doc2pdf.py -o ./All_PDFs ./file.doc ./Old_Reports
```
