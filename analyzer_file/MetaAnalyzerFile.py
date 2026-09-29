import os

import pandas as pd
from openpyxl import load_workbook

class MetaAnalyzerFile:

    def __init__(self, path_file: str):
        self.path_file = path_file


    def get_workbook_info(self) -> dict:
        full_file_name = os.path.basename(self.path_file)
        extension = os.path.splitext(full_file_name)[1]

        workbook = load_workbook(filename=self.path_file, read_only=False)
        list_sheets = workbook.sheetnames

        return {
            'nameFile': full_file_name,
            'extension': extension,
            'sheetsCount': len(list_sheets)
        }


    def get_profile(self) -> dict:
        response_data = {}
        response_data['sheets'] = []

        workbook = load_workbook(self.path_file, read_only=False)

        list_sheets = workbook.sheetnames
        for element in list_sheets:
            title_list = element
            row_count = workbook[element].max_row
            col_count = workbook[element].max_column
            cells = row_count * col_count

            meta_list_data = {
                'nameList': title_list,
                'dimensions': {
                    'rows': row_count,
                    'columns': col_count,
                    'cells': cells
                }
            }

            response_data['sheets'].append(meta_list_data)

        return response_data


    def get_dataframe_info(self) -> dict:
        response_data = {}
        response_data['sheets'] = []

        file_excel = pd.ExcelFile(self.path_file)
        list_sheets = file_excel.sheet_names

        for element in list_sheets:
            dataframe = pd.read_excel(self.path_file, sheet_name=element)
            struct_response = {
                'nameList': element,
                'dimensions': {
                    'rows': dataframe.shape[0],
                    'columns': dataframe.shape[1],
                    'size': dataframe.size
                }
            }

            response_data['sheets'].append(struct_response)

        return response_data

    # -> Данный метод получает информацию о листе(получает 5 х 5 матрицу, где видно какие ячейки пустые, а какие имеют значение)
    # -> Структура такая: {sheet: name_sheet, preview: [row x col]}
    def get_preview_list(self) -> dict:
        response_data = {}
        response_data['sheets'] = []

        work_book = load_workbook(self.path_file, read_only=False)
        list_sheets = work_book.sheetnames

        for element_sheet in list_sheets:
            list_row = []

            for row_element in work_book[element_sheet].iter_rows(min_row=1, max_row=5, min_col=1, max_col=5):
                list_cell = []

                for cell_element in row_element:
                    value = cell_element.value

                    if value is not None:
                        value = str(value)

                    list_cell.append(value)

                list_row.append(list_cell)

            struct = {
                'nameList': element_sheet,
                'preview': list_row
            }
            response_data['sheets'].append(struct)

        return response_data


    def find_merged_cells(self) -> dict:
        response_data = {}
        response_data['sheets'] = []

        work_book = load_workbook(self.path_file)
        sheets = work_book.sheetnames

        for element_sheet in sheets:
            sheet = work_book[element_sheet]
            list_merged_cells = []

            for element_merged in sheet.merged_cells.ranges:
                list_merged_cells.append(str(element_merged))

            struct_response = {
                'sheet': element_sheet,
                'mergedCells': list_merged_cells
            }

            response_data['sheets'].append(struct_response)

        return response_data


    def find_empty_row(self) -> dict:
        response_data = {}
        response_data['sheets'] = []

        work_book = load_workbook(self.path_file, read_only=False)
        list_sheets = work_book.sheetnames

        for element_sheet in list_sheets:
            sheet = work_book[element_sheet]
            list_empty_rows = []

            for row in range(1, sheet.max_row + 1):
                if all(cell.value is None for cell in sheet[row]):
                    list_empty_rows.append(row)

            struct_response = {
                'sheet': element_sheet,
                'emptyRows': list_empty_rows
            }

            response_data['sheets'].append(struct_response)
        return response_data


    def find_empty_col(self) -> dict:
        response_data = {}
        response_data['sheets'] = []

        work_book = load_workbook(self.path_file, read_only=False)
        sheets = work_book.sheetnames

        for element_list in sheets:
            sheet = work_book[element_list]
            col_empty_list = []

            for element_col in sheet.iter_cols():
                if all(cell.value is None for cell in element_col):
                    col_empty_list.append(element_col[0].column_letter)

            struct_response = {
                'sheet': element_list,
                'emptyColumns': col_empty_list
            }
            response_data['sheets'].append(struct_response)
        return response_data


    def find_hidden_rows(self) -> dict:
        response_data = {}
        response_data['sheets'] = []

        work_book = load_workbook(self.path_file, read_only=False)
        sheets = work_book.sheetnames

        for element_sheet in sheets:
            sheet = work_book[element_sheet]
            list_row_hidden = []

            for row_element in sheet.row_dimensions:
                if sheet.row_dimensions[row_element].hidden:
                    list_row_hidden.append(row_element)

            struct_response = {
                'sheet': element_sheet,
                'hiddenRows': list_row_hidden
            }

            response_data['sheets'].append(struct_response)
        return response_data


    def find_hidden_columns(self) -> dict:
        response_data = {}
        response_data['sheets'] = []

        work_book = load_workbook(self.path_file, read_only=False)
        sheets = work_book.sheetnames

        for element in sheets:
            sheet = work_book[element]
            list_columns_hidden = []

            for col_element in sheet.column_dimensions:
                if sheet.column_dimensions[col_element].hidden:
                    list_columns_hidden.append(col_element)

            struct_response = {
                'sheet': element,
                'hiddenColumns': list_columns_hidden
            }
            response_data['sheets'].append(struct_response)
        return response_data


    def find_formulas(self) -> dict:
        response_data = {}
        response_data['sheets'] = []

        work_book = load_workbook(self.path_file, read_only=False)
        list_sheets = work_book.sheetnames

        for element_sheet in list_sheets:
            sheet = work_book[element_sheet]
            list_formulas = []

            for element_row in sheet.iter_rows():
                for element_cell in element_row:
                    if isinstance(element_cell.value, str) and element_cell.value.startswith('='):
                        list_formulas.append(element_cell.value)

            struct_response = {
                'sheet': element_sheet,
                'formulas': list_formulas
            }
            response_data['sheets'].append(struct_response)
        return response_data


    def find_column_types(self) -> dict:
        response_data = {}
        response_data['sheets'] = []

        work_book = load_workbook(self.path_file, read_only=False)
        list_sheets = work_book.sheetnames

        for element_sheet in list_sheets:
            sheet = work_book[element_sheet]
            columns_list = []

            for col_element in sheet.iter_cols():
                types = {}

                for element_cell in col_element:
                    types[type(element_cell.value).__name__] = types.get(type(element_cell.value).__name__, 0) + 1

                columns_list.append({
                    'column': col_element[0].column_letter,
                    'types': types,
                    'total': len(col_element),
                    'empty': types.get('NoneType', 0),
                    'mixedType': len(types) - ('NoneType' in types) > 1
                })

            struct_response = {
                'sheet': element_sheet,
                'columns': columns_list
            }
            response_data['sheets'].append(struct_response)
        return response_data


    def create_profiler_file(self) -> dict:
        workbook = self.get_workbook_info()
        profile = self.get_profile()
        preview = self.get_preview_list()
        merged_cells = self.find_merged_cells()
        empty_rows = self.find_empty_row()
        empty_columns = self.find_empty_col()
        hidden_rows = self.find_hidden_rows()
        hidden_columns = self.find_hidden_columns()
        formulas = self.find_formulas()
        column_types = self.find_column_types()

        return {
            'workbook': workbook,
            'profile': profile,
            'preview': preview,
            'mergedCells': merged_cells,
            'emptyRows': empty_rows,
            'emptyColumns': empty_columns,
            'hiddenRows': hidden_rows,
            'hiddenColumns': hidden_columns,
            'formulas': formulas,
            'columnTypes': column_types
        }