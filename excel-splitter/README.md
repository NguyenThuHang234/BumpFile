# Excel Splitter Tool

This project is a Python tool designed to extract the first column from an XLSX file, split the data into sub-files with 300 lines each, and save all the sub-files into a specified directory.

## Features

- Extracts the first column from an XLSX file.
- Splits the extracted data into multiple sub-files, each containing up to 300 lines.
- Saves the sub-files in a specified directory.

## Requirements

To run this project, you need to have Python installed along with the following dependencies:

- `openpyxl`: A library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files.

You can install the required dependencies using pip:

```
pip install -r requirements.txt
```

## Usage

1. Place your XLSX file in the desired directory.
2. Update the `src/excel_splitter.py` file with the path to your XLSX file and the output directory.
3. Run the script:

```
python src/excel_splitter.py
```

4. Check the specified output directory for the generated sub-files.

## License

This project is licensed under the MIT License.