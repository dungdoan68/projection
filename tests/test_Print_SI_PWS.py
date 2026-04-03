from data import data_source
from src import ilp2026, map_SI_ILP2026
import xlwings as xw
import os
print(os.path.dirname("tests/"))

ilp2026.get_Print_Public_SI_By_List(
    source_SI_wb=xw.Book(data_source.SOURCE_FILE_ILP_2026_PUBLIC_SI),
    source_sheet_name=data_source.SOURCE_SHEET,
    target_sheet_name_3y=data_source.TARGET_SHEET_3Y,
    target_sheet_name_f=data_source.TARGET_SHEET_FULL,
    map_SI=map_SI_ILP2026,
)
#
#
# ilp2026.get_Result_Projection_ILP_By_Case(
#     source_wb_path=xw.Book(data_source.SOURCE_FILE_ILP_2026_PUBLIC_SI),
#     source_sheet_name = data_source.SOURCE_SHEET,
#     target_sheet_name_3y = data_source.TARGET_SHEET_3Y,
#     target_sheet_name_f = data_source.TARGET_SHEET_FULL,
#     map_SI = map_SI_ILP2026,
#     testCase='1972',
#     tc_Source= TESTCASE_FILE
# )
