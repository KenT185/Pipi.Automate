# Pipi.Automate
This repository contains multiple C# console applications aimed at automating my wife's work tasks. Each project focuses on a specific functionality.

## 1) WordTablesMerger
This project merges tables from multiple .docx files into one, with the following features:

- Takes all .docx files from a specified path (provided by the user).
- Merges all tables from the selected files into one table.
- Saves the merged table in a new .docx file.
- Removes the header (first row) and the second row from all tables in all files.
- Colors the first row yellow in each file to indicate a new file.
- Performs specific modifications for cells such as removing tabulators.

### How to Use
1. Download or clone the repository.
2. Run the executable (.exe) file stored in \Pipi.Automate\WordTablesMerger\bin\Release\net8.0\win-x64\WordTablesMerger.exe
4. Follow the prompts to specify the path containing the .docx files.
5. The application will process the files and generate the merged table in a new .docx file.

### Notes
- This application currently supports only .docx files.
- Ensure that the input files have tables formatted consistently for the best results.

## 2) SalaryCalculator (TODO)

In progress...

