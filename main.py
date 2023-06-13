from excel_file_updater import ExcelFileUpdater

_source_file = 'files/source_file.xlsx'
_source_sheet = 'Sheet1'
_source_value_column_index = 0  # column 'A'
_source_content_column_index = 1  # column 'B'

_target_file = 'files/target_file.xlsx'
_target_sheet = 'BillingOffer'
_target_value_column_index = 8  # column 'I'
_target_content_column_index = 33  # column 'AH'

updater = ExcelFileUpdater(
    _source_file,
    _source_sheet,
    _source_value_column_index,
    _source_content_column_index,
    _target_file,
    _target_sheet,
    _target_value_column_index,
    _target_content_column_index)

updater.update_target_file()
