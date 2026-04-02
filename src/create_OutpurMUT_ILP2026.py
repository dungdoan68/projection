import time
from datetime import datetime

import map_SI_ILP2026
import utils_helper

import xlwings as xw

TESTCASE_NO = 'B3'
Scenario_Int = 'B56'
POL_INFO = 'Pol_Info'
Prem_Term = 'H56'
SOURCE_FILE_ILP_2026 = '../Resource/SI file/2026ILP Illustration Spreadsheet V6.5.4 4.xlsb'
TARGET_FILE = "../Resource/2.xlsb"
SOURCE_SHEET = 'Interim'
TARGET_SHEET_3Y = 'Final_3yrs_Prem'
TARGET_SHEET_FULL = 'Final_flex_Prem'
ORIGINAL_MUT_FILE = '../data/MUT_Sample - Copy.xlsx'
ORIGINAL_UAT_FILE = '../data/Original_ILP2026.xlsb'
TESTCASE_FILE = '../data/testCases.xlsx'
TESTCASE_SHEET = 'TC'
SOURCE_SHEET_MUT_MUSTPAY = 'Output MUT_3yrs_Prem'
SOURCE_SHEET_MUT_FULL = 'Output MUT_flex_Prem'
MUT_SHEET_NAME = 'BaseAndRider'
#COPY MUT DATA OUTPUT
MUSTPAY_COPY = 'C16'
FULL_MUSTPAY_PASTE = 'A1'
FULL_COPY = 'C11'
CONTROL_SHEET = 'Control'


def get_Values_From_Interim_ILP(source_path, source_sheet_name, target_sheet_name_3y, target_sheet_name_f, map_SI,
                                source_Sheet_Mustpay, source_Sheet_Full):
    # try:
    #     source_wb = xw.Book(source_path)
    #     # ws.clear_contents()
    # except Exception as e:
    #     print(f"An error occurred: {e}")
    # finally:
    #     print('Open the workbooks')
    source_wb = utils_helper.get_Open_Excel_File(source_path)
    source_sheet = source_wb.sheets[source_sheet_name]
    list_Tcs = utils_helper.get_List_Tcs()
    i = 2
    utils_helper.get_Create_Tcs_MUT(ORIGINAL_MUT_FILE, list_Tcs)
    utils_helper.get_Create_Tcs(ORIGINAL_UAT_FILE, list_Tcs)
    TC_SOURCE = xw  .Book(TESTCASE_FILE)
    for item in list_Tcs:
        app = xw.App(visible=True)
        # app.quit()
        elapsed_time = 0
        start_time = time.perf_counter()
        print(datetime.now())
        testcase = 'Resource/UAT_Output/' + str(item[0]) + '.xlsb'
        utils_helper.get_Input_Value(source_wb, POL_INFO, TESTCASE_NO, item)
        utils_helper.get_Input_Value(source_wb, POL_INFO, Prem_Term, 3)
        utils_helper.get_Input_Value(source_wb, POL_INFO, Scenario_Int, 'L')
        target_wb = xw.Book(testcase)
        target_sheet_3y = target_wb.sheets[target_sheet_name_3y]
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_COI_Benefit_Low, target_sheet_3y)  #COPY PAGE I - PREMIUM && COPY PAGE II - BENEFIT
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_PW_BAV_PW_TAV_Value, target_sheet_3y)  #Copy PW BAV and PW TAV for Low Interest Rate
        utils_helper.get_Copy_Paste_Fund(source_wb, map_SI.get_Fund_List, map_SI.Low_Fund_Source, map_SI.Low_Fund_Target)  # Copy low fund to another space
        utils_helper.get_Input_Value(source_wb, POL_INFO, Scenario_Int, 'H')
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_COI_Benefit_High, target_sheet_3y)  #COPY PAGE I - PREMIUM  &&  COPY PAGE II - BENEFIT
        utils_helper.get_Clear_Content(target_sheet_3y, 'S:S')
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Premium_Benefit_High, target_sheet_3y)  #PREM TABLE BENEFIT TABLE
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Fund_Table, target_sheet_3y)  #'FUND TABLE
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Fund_High, target_sheet_3y)
        utils_helper.get_Input_Value(source_wb, POL_INFO, Scenario_Int, 'L')
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Premium_Benefit_Low, target_sheet_3y)  # PREM TABLE BENEFIT TABLE
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Fund_Low, target_sheet_3y)  #'Copy Low Fund interest
        target_sheet_full = target_wb.sheets[target_sheet_name_f]
        utils_helper.get_Input_Value(source_wb, POL_INFO, Prem_Term, 20)
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_COI_Benefit_Low, target_sheet_full)  # COPY PAGE I - PREMIUM && COPY PAGE II - BENEFIT
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_PW_BAV_PW_TAV_Value, target_sheet_full)  # Copy PW BAV and PW TAV for Low Interest Rate
        utils_helper.get_Copy_Paste_Fund(source_wb, map_SI.get_Fund_List, map_SI.Low_Fund_Source, map_SI.Low_Fund_Target)  # Copy low fund to another space
        utils_helper.get_Input_Value(source_wb, POL_INFO, Scenario_Int, 'H')
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_COI_Benefit_High, target_sheet_full)  # COPY PAGE I - PREMIUM && COPY PAGE II - BENEFIT
        utils_helper.get_Clear_Content(target_sheet_full, 'S:S')
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Premium_Benefit_High, target_sheet_full)  # PREM TABLE BENEFIT TABLE
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Withdrawal_High, target_sheet_full)
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Fund_Table, target_sheet_full)  # 'FUND TABLE
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Fund_High, target_sheet_full)
        utils_helper.get_Input_Value(source_wb, POL_INFO, Scenario_Int, 'L')
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Premium_Benefit_Low, target_sheet_full)  # PREM TABLE BENEFIT TABLE
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Withdrawal_Low, target_sheet_full)
        utils_helper.get_Clear_Content(target_sheet_full, 'AI:AI')
        utils_helper.get_Clear_Content(target_sheet_full, 'AY:AY')
        utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Fund_Low, target_sheet_full)  # 'Copy Low Fund interest
        print('Start copy from UAT to SI file')
        utils_helper.get_Copy_Paste_From_TC_Sheet_To_SI_File(target_wb.sheets[target_sheet_name_3y], 'A8:IC106', 'A8', source_wb.sheets[target_sheet_name_3y])
        utils_helper.get_Copy_Paste_From_TC_Sheet_To_SI_File(target_wb.sheets[target_sheet_name_f], 'A8:IC106', 'A8', source_wb.sheets[target_sheet_name_f])
        print('Start export MUT file  ')
        mustPay = 'Resource/MUT_Output/' + 'ILPC0' + str(item[0]) + '.xlsx'
        full = 'Resource/MUT_Output/' + 'ILPF0' + str(item[0]) + '.xlsx'
        mustPay_Target_wb = app.books.open(mustPay)
        full_Target_wb = app.books.open(full)
        sheet_Control = source_wb.sheets[CONTROL_SHEET]
        utils_helper.get_Copy_Paste_MUT_Value(source_wb.sheets[source_Sheet_Mustpay], sheet_Control, MUSTPAY_COPY,
                                              FULL_MUSTPAY_PASTE, mustPay_Target_wb.sheets[MUT_SHEET_NAME])
        utils_helper.get_Copy_Paste_MUT_Value(source_wb.sheets[source_Sheet_Full], sheet_Control, FULL_COPY,
                                              FULL_MUSTPAY_PASTE, full_Target_wb.sheets[MUT_SHEET_NAME])
        # utils_helper.get_Remove_Column_Xw(mustPay_Target_wb.sheets[MUT_SHEET_NAME])
        # utils_helper.get_Remove_Column_Xw(full_Target_wb.sheets[MUT_SHEET_NAME])
        mustPay_Target_wb.save()
        mustPay_Target_wb.close()
        full_Target_wb.save()
        full_Target_wb.close()
        utils_helper.get_Remove_Column_Pandas(mustPay, MUT_SHEET_NAME)
        utils_helper.get_Remove_Column_Pandas(full, MUT_SHEET_NAME)
        print(f"Close file {testcase}")
        target_wb.save()
        target_wb.close()
        end_time = time.perf_counter()
        print(datetime.now())
        elapsed_time = end_time - start_time
        elapsed_time_index = f'D{i}'
        utils_helper.get_Input_Value(TC_SOURCE, TESTCASE_SHEET, elapsed_time_index, elapsed_time)
        print(f"Execution took:{start_time} - {end_time} {elapsed_time:.6f} seconds")
        TC_SOURCE.save()
        i = i + 1
        app.quit()


get_Values_From_Interim_ILP(
    source_path=SOURCE_FILE_ILP_2026,
    # target_path=TARGET_FILE,
    source_sheet_name=SOURCE_SHEET,
    target_sheet_name_3y=TARGET_SHEET_3Y,
    target_sheet_name_f=TARGET_SHEET_FULL,
    map_SI=map_SI_ILP2026,
    source_Sheet_Mustpay=SOURCE_SHEET_MUT_MUSTPAY,
    source_Sheet_Full=SOURCE_SHEET_MUT_FULL
)
