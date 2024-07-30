# kmscreen2pdf
A small Windows utility for creating PDF files from images exported from KM Screen. 
The PDF file will include hidden text read from a text file.

## Installation and Usage:
1. Install [ImageMagick](https://imagemagick.org/script/download.php) on the system. 
2. Download the [kmscreen2pdf.exe](https://github.com/mikaelmoutakis/kmscreen2pdf/releases/) windows executable.  
3. Place it in a directory that the user has write and read access to. 
4. Call the application from FileMaker or run it directly. 
5. The application will then create directory 'kmscreen2pdf' in the same directory as the
executable. The program will then create the following subdirectories:
    ````
    kmscreen2pdf
    ├── error
    ├── input
    ├── logs
    └── output
    ````
6. Place the WMF and TXT files in the 'input' subdirectory. The files should have the same names, except for the file ending. E.g. 'export_2024-07-20.wmf' and 'export_2024-07-20.txt'. Note that the file endings are lowercase letters (e.g. 'txt' and not 'TXT').
7. Run the kmscreen2pdf.exe program. The program will then a) create a pdf file in the 'output' directory, b) delete the wmf and txt files, c) add information to the log file in the 'logs' directory. If something goes wrong, the input files will be moved to the 'error' directory. 

## Options:
The utility has the following command line options:

```
kmscreen2pdf.exe

Usage:
  kmscreen2pdf.exe [OPTIONS]
  kmscreen2pdf.exe -h

Options:
  -h --help         Show this screen.
  --version         Show version.
  --basedir=<dir>   Select the base directory for the input and output.
  --filetype=<ext>  Select the file type for the image [default: wmf].

Description:
    This Windows program converts images and text files to PDFs with invisible text.
    Requires ImageMagick to be installed.
```

