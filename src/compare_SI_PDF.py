import time
from datetime import datetime

import utils_helper

import xlwings as xw

# from spire.xls import *
# from spire.xls.common import *

# excel_app = win32com.client.Dispatch("Excel.Application")
# excel_app.Visible = False  # Run Excel in the background
# # excel_app.DisplayAlerts = False

TESTCASE_NO = 'B3'
Scenario_Int = 'B54'
POL_INFO = 'Pol_Info'
Prem_Term = 'H54'
SOURCE_FILE_UL = '../data/SI file/UL_Copay0.xlsb'
SOURCE_FILE_Term = '../data/SI file/TermBase_Copay0.xlsm'
SOURCE_FILE_ILP_2026 = '../Resource/SI file/2026ILP Illustration Spreadsheet V6.5.1.xlsb'
SOURCE_FILE_ILP_Package = '../data/SI file/RUV11_V11_Copay0.xlsb'
TARGET_FILE = "../Resource/2.xlsb"
SOURCE_SHEET = 'Interim'
TARGET_SHEET_3Y = 'Final_3yrs_Prem'
TARGET_SHEET_FULL = 'Final_flex_Prem'
ORIGINAL_FILE = '../data/Original_ILP2026.xlsb'
TESTCASE_FILE = '../data/testCases.xlsx'
TESTCASE_SHEET = 'TC'
Result = 'Result'


def get_Values_From_Interim_ILP(source_path, source_sheet_name, target_sheet_name_3y, target_sheet_name_f, map_SI):
    try:
        source_wb = xw.Book(source_path)
        source_sheet = source_wb.sheets[source_sheet_name]
        # ws.clear_contents()
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        print('Open the workbooks')
    list_Tcs = utils_helper.get_List_Tcs()
    i = 2
    utils_helper.get_Create_Tcs(ORIGINAL_FILE, list_Tcs)
    TC_SOURCE = xw.Book(TESTCASE_FILE)
    for item in list_Tcs:
        elapsed_time = 0
        start_time = time.perf_counter()
        print(datetime.now())
        testcase = 'Resource/' + str(item[0]) + '.xlsb'
        target_wb = xw.Book(testcase)
        utils_helper.get_Input_Value(source_wb, POL_INFO, TESTCASE_NO, item)
        utils_helper.get_Input_Value(source_wb, POL_INFO, Prem_Term, 3)
        utils_helper.get_Input_Value(source_wb, POL_INFO, Scenario_Int, 'L')
        utils_helper.get_Copy_Paste_Fund(source_wb, map_SI.get_Fund_List, map_SI.Low_Fund_Source,
                                         map_SI.Low_Fund_Target)
        target_sheet_3y = target_wb.sheets[target_sheet_name_3y]
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_COI_Benefit_Low, target_sheet_3y)
        utils_helper.get_Input_Value(source_wb, POL_INFO, Scenario_Int, 'H')
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_COI_Benefit_High, target_sheet_3y)
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Fund_Table, target_sheet_3y)
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Key_Values_High, target_sheet_3y)
        utils_helper.get_Input_Value(source_wb, POL_INFO, Scenario_Int, 'L')
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Key_Values_Low, target_sheet_3y)
        utils_helper.get_Input_Value(source_wb, POL_INFO, Prem_Term, 20)
        utils_helper.get_Input_Value(source_wb, POL_INFO, Scenario_Int, 'L')
        utils_helper.get_Copy_Paste_Fund(source_wb, map_SI.get_Fund_List, map_SI.Low_Fund_Source,
                                         map_SI.Low_Fund_Target)
        target_sheet_full = target_wb.sheets[target_sheet_name_f]
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_COI_Benefit_Low, target_sheet_full)
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_PW_BAV_PW_TAV_Value, target_sheet_full)
        utils_helper.get_Input_Value(source_wb, POL_INFO, Scenario_Int, 'H')
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Fund_Table, target_sheet_full)
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_COI_Benefit_High, target_sheet_full)
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Key_Values_High, target_sheet_full)
        utils_helper.get_Input_Value(source_wb, POL_INFO, Scenario_Int, 'L')
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Key_Values_Low, target_sheet_full)
        utils_helper.get_Copy_Paste_From_TC_Sheet_To_SI_File(target_wb.sheets[target_sheet_name_3y], 'A8:IC106', 'A8',
                                                             source_wb.sheets[target_sheet_name_3y])
        utils_helper.get_Copy_Paste_From_TC_Sheet_To_SI_File(target_wb.sheets[target_sheet_name_f], 'A8:IC106', 'A8',
                                                             source_wb.sheets[target_sheet_name_f])
        print('Start compare SI')
        source_wb.api.RefreshAll()
        result_Sheet = source_wb.sheets['Sum_Validation']
        value = utils_helper.get_Cell_Value(source_wb, result_Sheet, 'H4')
        result_Sheet_target = target_wb.sheets[Result]
        result_Sheet_target.range('A1').value = value
        get_Log_TC_Result(i)
        # item_index = f'B{i}'
        # value_index = f'C{i}'
        # utils_helper.get_Input_Value(TC_SOURCE, TESTCASE_SHEET, item_index, item)
        # utils_helper.get_Input_Value(TC_SOURCE, TESTCASE_SHEET, value_index, value)
        # print(f"Verify Tc: {testcase} done with result {value}")
        # print(f"Close file {testcase}")
        # target_wb.save()
        # target_wb.close()
        # end_time = time.perf_counter()
        # print(datetime.now())
        # elapsed_time = end_time - start_time
        # elapsed_time_index = f'D{i}'
        # utils_helper.get_Input_Value(TC_SOURCE, TESTCASE_SHEET, elapsed_time_index, elapsed_time)
        # print(f"Execution took:{start_time} - {end_time} {elapsed_time:.6f} seconds")
        # i = i + 1


def get_Log_TC_Result(index):
    item_index = f'B{index}'
    value_index = f'C{index}'
    utils_helper.get_Input_Value(TC_SOURCE, TESTCASE_SHEET, item_index, item)
    utils_helper.get_Input_Value(TC_SOURCE, TESTCASE_SHEET, value_index, value)
    print(f"Verify Tc: {testcase} done with result {value}")
    print(f"Close file {testcase}")
    target_wb.save()
    target_wb.close()
    end_time = time.perf_counter()
    print(datetime.now())
    elapsed_time = end_time - start_time
    elapsed_time_index = f'D{index}'
    utils_helper.get_Input_Value(TC_SOURCE, TESTCASE_SHEET, elapsed_time_index, elapsed_time)
    print(f"Execution took:{start_time} - {end_time} {elapsed_time:.6f} seconds")
    index = index + 1



#
# get_Values_From_Interim_ILP(
#     source_path=SOURCE_FILE_ILP_2026,
#     # target_path=TARGET_FILE,
#     source_sheet_name=SOURCE_SHEET,
#     target_sheet_name_3y=TARGET_SHEET_3Y,
#     target_sheet_name_f=TARGET_SHEET_FULL,
#     map_SI=map_SI_ILP2026
# )




