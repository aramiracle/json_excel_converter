import json
from json_excel_converter import JsonExcelConverter

def main():
    original_json_file = 'categories.json'
    excel_file = 'categories.xlsx'
    reconstructed_json_file = 'reconstructed_categories.json'

    # Create an instance of JsonExcelConverter for JSON to Excel conversion
    converter = JsonExcelConverter(json_file=original_json_file, excel_file=excel_file)
    converter.json_to_excel()

    # Create an instance of JsonExcelConverter for Excel to JSON conversion
    converter = JsonExcelConverter(excel_file=excel_file, output_json_file=reconstructed_json_file)
    converter.excel_to_json()

    # Load and compare the original and reconstructed JSON files
    with open(original_json_file, 'r') as f:
        original_dict = json.load(f)

    with open(reconstructed_json_file, 'r') as f:
        reconstructed_dict = json.load(f)

    if reconstructed_dict == original_dict:
        print('Conversion is ok.')
    else:
        print('Conversion failed. Here is the reconstructed dictionary:')
        print(json.dumps(reconstructed_dict, indent=4))

if __name__ == "__main__":
    main()
