
# PowerPoint to PDF Converter

A simple and efficient Python script to automate the batch conversion of PowerPoint presentations (`.pptx`, `.ppt`) into PDF files. This tool utilizes the local Microsoft PowerPoint application to ensure high-fidelity output.

## Features
- **Batch Processing:** Automatically converts all PowerPoint files in the directory.
- **Background Execution:** Minimizes UI interruption during the conversion process.
- **Organized Output:** Saves all generated PDFs into a dedicated `converted_pdfs` folder to keep your directory clean.

## Prerequisites
- **Python 3.x** installed on your system.
- **Microsoft PowerPoint** must be installed (Windows only).
- Install the required library:
  ```bash
  pip install comtypes
```

## How to Use

1. Place the `converter.py` script in the folder containing your PowerPoint files.

2. Run the script using your terminal:

   Bash

   ```
   python converter.py
   ```

3. The converted PDF files will appear in the `converted_pdfs` folder.

## Troubleshooting

- **Permission/Access Errors:** Ensure that no PowerPoint instances are frozen before running the script.
- **Performance:** If you have a very large number of files, the script may take some time. It is recommended to process files in batches if necessary.

## License

This project is open-source and available for personal use.
