# JsonExcelConverter

`JsonExcelConverter` is a Python class designed to facilitate the conversion between JSON and Excel file formats. It provides methods to convert JSON data to Excel spreadsheets and vice versa, while ensuring that the data structure is consistent and correctly formatted.📊🔄

## Features🌟

- Convert JSON data to Excel spreadsheets with hierarchical levels flattened.📈
- Convert Excel spreadsheets back to nested JSON structures.🔄
- Validate that JSON data is consistent in depth before conversion.✅
- Ensure that Excel data rows are consistent in depth before conversion.📏
- Save converted data to specified output files with proper formatting.💾

## Requirements🛠️

- Python 3.x
- Pandas
- OpenPyXL

You can install the required Python packages using pip:

```
pip install pandas openpyxl
```

## Installation🚀

Clone this repository or download the source code. Ensure you have Python 3.x installed and the required packages.

```
git clone https://github.com/yourusername/JsonExcelConverter.git
cd JsonExcelConverter
pip install -r requirements.txt
```

## Usage📚

### Initializing the Converter

You can initialize the `JsonExcelConverter` with various combinations of input and output files.

```
from JsonExcelConverter import JsonExcelConverter

# Initialize with JSON file
converter = JsonExcelConverter(json_file='data.json')

# Initialize with Excel file and output JSON file
converter = JsonExcelConverter(excel_file='data.xlsx', output_json_file='output.json')

```

### Converting JSON to Excel

To convert a JSON file to an Excel file:

```
converter = JsonExcelConverter(json_file='data.json')
converter.json_to_excel()  # Converts JSON to Excel
```

Alternatively, if you have a dictionary:
```
data_dict = { ... }  # Your dictionary data
converter = JsonExcelConverter(data_dict=data_dict, excel_file='output.xlsx')
converter.json_to_excel()  # Converts dictionary to Excel
```

### Converting Excel to JSON

To convert an Excel file to a JSON file:

```
converter = JsonExcelConverter(excel_file='data.xlsx', output_json_file='output.json')
converter.excel_to_json()  # Converts Excel to JSON
```

## Methods🛠️

### `json_to_excel()`

Converts JSON data or a dictionary to an Excel file. The JSON data must be consistent in depth.

**Raises:**
- `ValueError` if JSON data is not consistent in depth.

### `excel_to_json()`

Converts an Excel file to a JSON file. The Excel file must have rows that are consistent in depth.

**Raises:**
- `ValueError` if Excel data is not consistent in depth.

### `_validate_json_depth(data)`

Validates that all JSON data is at the same depth.❌

**Args:**
- `data` (dict): JSON data to be validated.

**Raises:**
- `ValueError` if JSON data is not consistent in depth.❌

### `_validate_excel_depth(data)`

Validates that all rows in the Excel data have the same depth.❌

**Args:**
- `data` (dict): Nested dictionary to be validated.

**Raises:**
- `ValueError` if Excel rows are not consistent in depth.❌

## Contributing🤝

Feel free to contribute to this small project by submitting issues, bug reports, or feature requests.🚀