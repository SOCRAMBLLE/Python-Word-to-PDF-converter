# Word to PDF File Converter

Welcome to the Word to PDF File Converter project! This Python script allows you to convert Word documents (.docx) to PDF format. Follow the instructions below to use the script.

## Prerequisites
Before you begin, make sure you have Python installed.

- [Python](https://www.python.org/)

## Usage
1. Place your Word documents (.docx) in the `word` folder.

2. Run the script to convert the Word documents to PDF:

   ```bash
   python convert.py
   ```
3. Converted PDF files will be saved in the pdf folder (automatically created if it doesn't exist).

## File Conversion
The convert.py script uses the win32com.client module to interact with Microsoft Word and convert the documents. Each Word document (.docx) in the word folder will be converted to a corresponding PDF file in the pdf folder.

## Permissions
The script also includes a function to change file permissions using the os module. This function is used to change permissions for the Word documents before conversion.

## Contribution
Contributions are welcome! If you'd like to contribute to this project, feel free to fork the repository, make your changes, and submit a pull request.