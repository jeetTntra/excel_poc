# Read the excel file and convert it to a dataframe
import re
import sys

import openpyxl as op
import numpy as np

# Regex to find the string "X.XXX To X.XXX"
DECIMAL_TO_DECIMAL = r"(\d+\.\d+) To (\d+\.\d+)"

# Regex to find the string "STRING-X.XXX-X.XXX"
STRING_DECIMAL_DECIMAL = r"(\w+)-(\d+\.\d+)-(\d+\.\d+)"


def excel_to_csv(file_path, table):
    wb = op.load_workbook(file_path)
    sheet_names = wb.sheetnames

    print("Total number of sheets: ", len(sheet_names))
    # print("Sheet names: ", sheet_names)
    filtered_list = []
    for sheet_name in sheet_names:
        if re.match(DECIMAL_TO_DECIMAL, sheet_name) or re.match(STRING_DECIMAL_DECIMAL, sheet_name):
            filtered_list.append(wb[sheet_name])

    print("Length of the filtered_list: ", len(filtered_list))
    # only keep 1 sheet for testing
    filtered_list = filtered_list[0:1]
    parse_data(filtered_list, table)


def convert_to_csv(sheet_table, table):
    # Transform the sheet_table and save it to csv
    csv_table = []
    if table == "1":
        header = ["Pointer", "Clarity", "Color", "Price", "Font"]
        # Get the data type of the sheet_table
        for sheet in sheet_table:
            for key, value in sheet.items():
                # print("Key: ", key)
                # print("Value: ", value)
                for item in value:
                    csv_table.append([item["Pointer"], item["Clarity"], item["Color"], item["Price"], item["Font"]])
        # save_csv(csv_table, header, table)


# Save the csv_table to csv file with table_{table}.csv with header and
def save_csv(csv_table, header, table):
    np.savetxt(f"table_{table}.csv", csv_table, delimiter=",", header=",".join(header), fmt="%s", comments='')
    print("Data inserted successfully")


def parse_table_one(header, pointer, sheet, sheet_items):
    for row in sheet.iter_rows(min_row=1, max_row=1, values_only=True):
        for cell in row:
            if cell is not None:
                header.append(cell)
        pointer = header[0]
        header = header[1:]
    for row in sheet.iter_rows(min_row=2, values_only=False):
        if row[0].value is None:
            print("Empty row found", row[0].row)
            break
        for cell in row[1:]:
            # skip the empty cells
            if cell.value is None:
                continue
            if cell is not None:
                text = cell
                style = sheet.cell(row=cell.row, column=cell.column).font
                font_style = ""
                if style.b:
                    font_style = "bold".upper()
                elif style.i:
                    font_style = "italic".upper()
                else:
                    font_style = "normal".upper()

                row_dict = {
                    "Pointer": pointer,
                    "Clarity": sheet.cell(row=1, column=cell.column).value,
                    "Color": row[0].value,
                    "Price": text.value,
                    "Font": font_style
                }
                sheet_items.append(row_dict)


def parse_table_two(header, pointer, sheet, sheet_items):
    # clarity_header is set of clarity headers
    pointer_header_index = []
    clarity_header = set()
    clarity_header_row_index = 0
    cut_header = set()
    cut_header_row_index = 0
    florescence_header = ["None", "Faint", "Medium", "Strong"]
    # the dictionary to store the table 2 is like this
    # {
    #     "Pointer": 0.20 To 0.229
    #     "Clarity": "IF",
    #     "Cut": "3EX",
    #     "Fluorescence": "None",
    #     "Color": "D",
    #     "Value": "-38",
    #     "Font": "BOLD"
    #     "Value_Color": "Red"
    # }
    for row in sheet.iter_rows(min_row=1, values_only=False):
        # Loop through the first cell of the row is "Range =>"
        if row[0].value == "Range =>":
            for cell in row[1:]:  # skip the first cell of the row and loop through the rest of the cells in the row
                if cell.value is None:
                    continue
                pointer_header_index.append(row[0].row)
        else:
            pass

    # Loop through the sheet and get the data
    for row in sheet.iter_rows(min_row=pointer_header_index[0], max_row=pointer_header_index[1] - 1, values_only=False):
        color_header_index = []
        florescence_header_index = []
        if row[0].value == "Clarity =>":
            for cell in row[1:]:
                if cell.value is None:
                    continue
                clarity_header.add(cell.value)
                clarity_header_row_index = row[0].row
        elif row[0].value == "Cut =>":
            for cell in row[1:]:
                if cell.value is None:
                    continue
                cut_header.add(cell.value)
                cut_header_row_index = row[0].row
        elif row[0].value == "Color":
            for cell in row:
                if cell.value is None:
                    continue
                if cell.value == "Color":
                    color_header_index.append(row[0].row)
                    color_header_index.append(cell.column)
                elif cell.value == "Florescence":
                    florescence_header_index.append(row[0].row)
                    florescence_header_index.append(cell.column)
            print("Color header index: ", color_header_index)
            print("Florescence header index: ", florescence_header_index)
        else:
            pass

        for cell in row:
            print(cell.value)

    print("Header: ", header)
    print("Clarity Header: ", clarity_header, clarity_header_row_index)
    print("Cut Header: ", cut_header, cut_header_row_index)
    print("Pointer Header: ", pointer_header_index)


def parse_data(list, table):
    sheet_table = []
    for sheet in list:
        print("Sheet name: ", sheet.title)
        pointer = []
        header = []
        sheet_items = []

        if table == "1":
            parse_table_one(header, pointer, sheet, sheet_items)
        elif table == "2":
            parse_table_two(header, pointer, sheet, sheet_items)

        sheet_table.append({sheet.title: sheet_items})
    convert_to_csv(sheet_table, table)


if __name__ == '__main__':
    # get file path from command line
    file_path = sys.argv[1]
    table = sys.argv[2]
    # convert excel file to dataframe
    excel_to_csv(file_path, table)

# In the terminal, run the following command
# python main.py "path/to/excel/file" "table_number"
