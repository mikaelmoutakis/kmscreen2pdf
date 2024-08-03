# kmscreen2pdf
**kmscreen2pdf** is a small Windows utility designed to create PDF files from images exported from KM Screen. The resulting PDF file will include hidden text extracted from a text file.

## Installation and Usage:

1. Install [ImageMagick](https://imagemagick.org/script/download.php) on your system.
2. Download the [kmscreen2pdf.exe](https://github.com/mikaelmoutakis/kmscreen2pdf/releases/download/release/kmscreen2pdf_version_0.1.1_win11.zip) Windows executable.
3. Place the executable in a directory where you have read and write access.
4. Execute the application from FileMaker or run it directly.
5. The application will create a directory named 'kmscreen2pdf' in the same location as the executable. Within this directory, the following subdirectories will be generated:
    ````
    kmscreen2pdf
    ├── error
    ├── input
    ├── logs
    └── output
    ````
6. Place the WMF and TXT files in the 'input' subdirectory. Ensure the files have identical names, differing only in their extensions (e.g., 'export_2024-07-20.wmf' and 'export_2024-07-20.txt'). Note that the file extensions should be in lowercase (e.g., 'txt' and not 'TXT').
7. Run the kmscreen2pdf.exe program. The program will: a) Create a PDF file in the 'output' directory. b) Delete the WMF and TXT files. c) Log information in the 'logs' directory. If an error occurs, the input files will be moved to the 'error' directory.


## Options:
The utility has the following command line options:

```
kmscreen2pdf.exe

Usage:
  kmscreen2pdf.exe [--basedir=<dir>] [--filetype=<ext>] [--do-not-convert-text]
  kmscreen2pdf.exe -h

Options:
  -h --help                 Show this screen.
  --version                 Show version.
  --basedir=<dir>           Select the base directory for the input and output.
  --filetype=<ext>          Select the file type for the image [default: wmf].
  --do-not-convert-text     Do not convert text files from UTF-16 LE to UTF-8.

Description:
    This Windows program converts images and text files to PDFs with invisible text.
    Requires ImageMagick to be installed.
```

