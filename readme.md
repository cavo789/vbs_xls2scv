![Banner](images/banner.jpg)

# Excel - Export worksheets to csv files

> Convert Excel worksheets to csv files in batch

This script will process each `.xlsx` file located in a specific folder (f.i. `c:\temp\input`), start MS Excel, open each file in Excel and save each visible worksheets as a `.csv` file in a second folder (like f.i. `c:\temp\output`).

Notes:

- The original file will remains unchanged
- Only visible worksheets will be saved as `.csv`

## Table of Contents

- [Install](#install)
- [Usage](#usage)
- [License](#license)

## Install

Get a raw copy of `xls2csv.vbs` and save it onto your disk.

If you want, get also a copy of `xls2csv.cmd` and save it in the same folder.

## Usage

### With the xls2csv.vbs only

1. Start a DOS prompt
2. Go to the folder where you've saved the `xls2csv.vbs` script
3. Call `cscript xls2csv.vbs` and specify the folder where you've your `.xlsx` files and where you want to save the generated `.csv` files like, f.i., `cscript xls2csv.vbs c:\temp\input c:\temp\output`

### With the xls2csv.cmd

1. Edit the `xls2csv.cmd` file in Notepad
2. Modify the path to `xls2csv.cmd`the input and output folders
3. Save the file
4. From now, just double-click on the `xls2csv.cmd` to execute it.

### License

[MIT](LICENSE)
