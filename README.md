# Excel Unlocker

A simple command-line tool to remove protection from Excel (.xlsx) files.

## Overview

Excel Unlocker is a C# utility that removes sheet protection and workbook structure locks from Excel files. It works by extracting the Excel file's internal XML structure, removing protection elements, and repackaging the file.

## Features

- Remove sheet protection from all worksheets
- Unlock workbook structure (lockStructure parameter)
- Preserve original file formatting and data
- Simple command-line interface
- Colorful ASCII banner

## How It Works

The tool operates by:

1. Renaming the .xlsx file to .zip (Excel files are compressed archives)
2. Extracting the archive contents
3. Removing `<sheetProtection>` elements from worksheet XML files
4. Changing `lockStructure="1"` to `lockStructure="0"` in workbook.xml
5. Recompressing the modified files
6. Outputting the unlocked file as `unlocked_file.xlsx`

## Requirements

- .NET 6.0 or higher
- Windows, Linux, or macOS

## Usage

1. Run the application:
```bash
   dotnet run
```

2. Enter the name of the Excel file you want to unlock:
```
   Input file name (es Example.xlsx): 
   MyProtectedFile.xlsx
```

3. The unlocked file will be saved as `unlocked_file.xlsx` in the same directory

## Building
```bash
dotnet build
```

## Installation

Clone the repository and build the project:
```bash
git clone <repository-url>
cd ExcelUnlocker
dotnet build
```

## Important Notes

- This tool is intended for unlocking your own Excel files where you've forgotten the password
- Always keep a backup of your original file
- The output file is always named `unlocked_file.xlsx`
- Only works with .xlsx files (Excel 2007 and newer)

## Legal Disclaimer

This tool is provided for educational purposes and to help users recover access to their own files. Users are responsible for ensuring they have the right to unlock any files they process with this tool.

## License

All rights reserved © 2025

## Author

Made by Simo
