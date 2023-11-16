from RPA.Cloud.Google import Google
from robot.libraries.BuiltIn import BuiltIn
from robot.libraries.Collections import Collections

class GoogleSheetsInteractions:
    def __init__(self):
        self.string_lib = BuiltIn().get_library_instance('String')
        self.google_lib = Google(vault_name='GoogleSheets', vault_secret_key='service_account')
        self.collections_lib = Collections()
        self.gSheetObject = {}

    def compile_google_sheet_object(self, sheet_id, sheet_range, values_required, row_offset=0, row_length=0):
        sheet_exported_list = self.get_values_from_google_sheet(sheet_id, sheet_range, values_required)
        gSheetObject = {
            'sheetId': sheet_id,
            'sheetRange': sheet_range,
            'valuesRequired': values_required,
            'rowOffset': row_offset,
            'rowLength': row_length,
            'values': sheet_exported_list
        }
        return gSheetObject

    def get_values_from_google_sheet(self, sheet_id, sheet_range, values_required):
        sheet_values = self.google_lib.wait_until_keyword_succeeds(4, '1 minute', self.google_lib.get_sheet_values,
                                                                   sheet_id, sheet_range)
        if 'values' not in sheet_values:
            if values_required:
                raise ValueError(f"List with ID '{sheet_id}' and range '{sheet_range}' is empty")
            return []
        return sheet_values['values']

    def update_values_google_sheet_object(self, gSheetObject):
        new_values = self.get_values_from_google_sheet(gSheetObject['sheetId'], gSheetObject['sheetRange'],
                                                       values_required=False)
        gSheetObject.pop('values', None)
        gSheetObject['values'] = new_values
        return gSheetObject

    def get_first_row_from_google_sheet(self, gSheetObject):
        gSheetObject = self.update_values_google_sheet_object(gSheetObject)
        values = gSheetObject['values']
        if not values:
            return []
        
        first_item = values[0]
        normalized_row = [self.string_lib.convert_to_string(cell) if self.string_lib.get_regexp_matches(
            self.string_lib.convert_to_string(cell), pattern='\\S') else '' for cell in first_item]
        row_length = gSheetObject['rowLength']
        while len(normalized_row) < row_length:
            normalized_row.append('')
        return normalized_row

    def remove_row_google_sheet(self, gSheetObject, row_idx):
        gSheetObject = self.update_values_google_sheet_object(gSheetObject)
        if row_idx < 0 or row_idx >= len(gSheetObject['values']):
            raise ValueError(f"Invalid row index: {row_idx}")
        
        gSheetObject['values'].pop(row_idx)
        self.google_lib.clear_sheet_values(gSheetObject['sheetId'], gSheetObject['sheetRange'])
        self.google_lib.insert_sheet_values(gSheetObject['sheetId'], gSheetObject['sheetRange'], gSheetObject['values'], 'ROWS')
        return gSheetObject

    def append_row_google_sheet(self, gSheetObject, new_row):
        gSheetObject = self.update_values_google_sheet_object(gSheetObject)
        gSheetObject['values'].append(new_row)
        self.google_lib.insert_sheet_values(gSheetObject['sheetId'], gSheetObject['sheetRange'], gSheetObject['values'], 'ROWS')
        return gSheetObject

    def get_cell_from_google_sheet(self, gSheetObject, row_idx, col_idx):
        gSheetObject = self.update_values_google_sheet_object(gSheetObject)
        values = gSheetObject['values']
        if row_idx < 0 or row_idx >= len(values) or col_idx < 0 or col_idx >= len(values[row_idx]):
            return ''

        return values[row_idx][col_idx]
