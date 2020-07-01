import os
import json
import argparse

import pandas as pd


class Error(Exception):
    pass


def convert_nested_json(df):
    """
    Converts a JSON string in cells to a dict.
    """
    object_cols = df.select_dtypes(include='object')

    def try_get_json(o):
        try:
            return json.loads(o)
        except:
            return o

    for col_name, col_values in object_cols.items():
        df[col_name] = list(map(try_get_json, col_values))


def gen_json(sheets):
    jsoned_sheets = []

    for sheet_name, sheet in sheets.items():
        convert_nested_json(sheet)

        # Create a JSON string with sheet data
        jsoned_sheet = sheet.to_json(orient='records', force_ascii=False,
                                     date_format='iso', indent=4)

        # Replace 4 spaces with 8 for better formatting
        jsoned_sheet = jsoned_sheet.replace('\n    ', '\n        ')

        # And also add 4 spaces to the end
        jsoned_sheet = jsoned_sheet[:-1] + '    ]'

        jsoned_sheets.append(jsoned_sheet)

    # Create a JSON layout with the names of the sheets
    # {
    #     "sheet_1": %s,
    #     "sheet_2": %s
    #     ...
    # }
    json_str = json.dumps({i: None for i in sheets.keys()}, indent=4, ensure_ascii=False)
    json_str = json_str.replace('null', '%s')

    # Substituting the created JSON with sheet data in the final JSON
    json_str = json_str % tuple(jsoned_sheets)

    return json_str


def main():
    parser = argparse.ArgumentParser(description='Convert Excel workbook to JSON file')

    parser.add_argument('xl_path', type=str, help='Path to the Excel file')
    parser.add_argument('-o', type=str, default='.',
                        help='Path to the output JSON file', dest='json_path')

    args = parser.parse_args()

    if args.json_path == '.':
        args.json_path = os.path.splitext(args.xl_path)[0] + '.json'

    if os.path.isfile(args.json_path):
        raise Error('JSON file already exists')

    print('Reading an Excel file, please wait...')
    sheets = pd.read_excel(args.xl_path, sheet_name=None)
    print('Converting...')

    result = gen_json(sheets)

    # Writing the created JSON to a file
    with open(args.json_path, 'w', encoding='utf-8') as f:
        f.write(result)

    print('Done. JSON was created successfully')


if __name__ == '__main__':
    try:
        main()
    except Error as e:
        print(f'Error: {e}')
        exit(1)
