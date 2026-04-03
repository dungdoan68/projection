import time
import os
from datetime import datetime

import sys
print(sys.path.append(os.path.realpath("src")))
import utils_helper, map_SI_ILP2026
from data import data_source
import xlwings as xw

# SOURCE_WB = xw.Book(data_source.SOURCE_FILE_ILP_2026_PUBLIC_SI)
# TC_SOURCE_WB = xw.Book(data_source.TESTCASE_FILE)
LIST_TCS = utils_helper.get_List_Tcs()
dir_path = os.path.dirname(os.path.realpath(__file__))
cwd = os.getcwd()


def get_Result_Projection_ILP_By_List(source_sheet_name, target_sheet_name_3y, target_sheet_name_f, map_SI):
    for item in LIST_TCS:
        # app = xw.App(visible=False)
        start_time = time.perf_counter()
        get_Result_Projection_ILP_By_Case(source_sheet_name, target_sheet_name_3y, target_sheet_name_f,
            map_SI, item[0])
        elapsed_time = time.perf_counter() - start_time
        print(f"Execution took: {elapsed_time:.6f} seconds")
        # app.quit()


def get_Print_Public_SI_By_List(source_SI_wb, source_sheet_name, target_sheet_name_3y, target_sheet_name_f, map_SI):
    LIST_TCS1 = utils_helper.get_List_Tcs()
    for item in LIST_TCS1:
        start_time = time.perf_counter()
        get_Print_Public_SI_By_Case(source_SI_wb, source_sheet_name, target_sheet_name_3y, target_sheet_name_f, map_SI,
            item[0])
        elapsed_time = time.perf_counter() - start_time
        print(f"Execution took: {elapsed_time:.6f} seconds")
        # lastR = utils_helper.get_Last_Row(data_source.TESTCASE_FILE, data_source.TESTCASE_RESULT_SHEET)
        # utils_helper.get_Input_Value(source_SI_wb, data_source.TESTCASE_RESULT_SHEET, f'C{lastR}', item[0])
        # utils_helper.get_Input_Value(source_SI_wb, data_source.TESTCASE_RESULT_SHEET, f'D{lastR}', f'{elapsed_time:.6f}')
        source_SI_wb.save()


def get_Export_SI_File_To_UAT_File_By_List(source_SI_wb, source_sheet_name, target_sheet_name_3y, target_sheet_name_f, map_SI):
    TC_SOURCE_WB
    for item in LIST_TCS:
        # xw.App(visible=False)
        elapsed_time = 0
        start_time = time.perf_counter()
        print(datetime.now())
        tc_wb = get_Copy_From_SI_File_To_UAT_File(source_SI_wb, source_sheet_name, target_sheet_name_3y, target_sheet_name_f,
            map_SI, item[0])
        elapsed_time = time.perf_counter() - start_time
        print(f"Execution took: {elapsed_time:.6f} seconds")
        tc_wb.close()


def get_Result_Projection_ILP_By_Case(source_SI_wb, source_sheet_name, target_sheet_name_3y, target_sheet_name_f, map_SI, testCase):
    start_time = time.perf_counter()
    get_Copy_From_SI_File_To_UAT_File(source_SI_wb, source_sheet_name, target_sheet_name_3y, target_sheet_name_f, map_SI, testCase)
    result = get_Compare_PDF_SI_By_Case(source_SI_wb, target_sheet_name_3y, target_sheet_name_f, testCase)
    end_time = time.perf_counter()
    print(datetime.now())
    elapsed_time = end_time - start_time
    # lastR = utils_helper.get_Last_Row(data_source.TESTCASE_FILE, data_source.TESTCASE_RESULT_SHEET)
    # utils_helper.get_Input_Value(TC_SOURCE_WB, data_source.TESTCASE_RESULT_SHEET, f'B{lastR}', testCase)
    # utils_helper.get_Input_Value(TC_SOURCE_WB, data_source.TESTCASE_RESULT_SHEET, f'C{lastR}', result)
    # utils_helper.get_Input_Value(TC_SOURCE_WB, data_source.TESTCASE_RESULT_SHEET, f'D{lastR}', f'{elapsed_time:.6f}')
    print(f"Execution took:{start_time} - {end_time} {elapsed_time:.6f} seconds")
    # TC_SOURCE.save()
    return result


def get_Compare_PDF_SI_By_Case(source_SI_wb, target_sheet_name_3y, target_sheet_name_f, testCase):
    get_Copy_Data_TC_To_SI_File(source_SI_wb, target_sheet_name_3y, target_sheet_name_f, testCase)
    print('Start compare SI')
    source_SI_wb.api.RefreshAll()
    result_Sheet = source_SI_wb.sheets['Sum_Validation']
    value = utils_helper.get_Cell_Value(source_SI_wb, result_Sheet, 'I6')
    print(value)
    return value


def get_Copy_Data_TC_To_SI_File(source_SI_wb, target_sheet_name_3y, target_sheet_name_f, testCase):
    testCase = str(testCase)
    print(testCase)
    utils_helper.get_Input_Value(source_SI_wb, data_source.POL_INFO_SHEET, data_source.CELL_INPUT_TESTCASE, testCase)
    target_wb = xw.Book(data_source.UAT_FOLDER + testCase + '.xlsb')
    print('Start copy from UAT to SI file')
    utils_helper.get_Copy_Paste_From_TC_Sheet_To_SI_File(target_wb.sheets[target_sheet_name_3y], 'A1:II106', 'A1',
        source_SI_wb.sheets[target_sheet_name_3y])
    utils_helper.get_Copy_Paste_From_TC_Sheet_To_SI_File(target_wb.sheets[target_sheet_name_f], 'A1:II106', 'A1',
        source_SI_wb.sheets[target_sheet_name_f])
    target_wb.save()
    target_wb.close()


def get_Copy_From_SI_File_To_UAT_File(source_SI_wb, source_sheet_name, target_sheet_name_3y, target_sheet_name_f, map_SI, testCase):
    testCase = str(testCase)
    # SOURCE_WB = xw.Book(data_source.SOURCE_FILE_ILP_2026_PUBLIC_SI)
    source_sheet = source_SI_wb.sheets[source_sheet_name]
    print(data_source.ORIGINAL_UAT_FILE)
    uat_Tc = utils_helper.get_Copy_New_TC_File(data_source.ORIGINAL_UAT_FILE, data_source.UAT_FOLDER + testCase + ".xlsb")
    print(datetime.now())
    source_SI_wb.app.calculation = 'automatic'
    utils_helper.get_Input_Value(source_SI_wb, data_source.POL_INFO_SHEET, data_source.CELL_INPUT_TESTCASE, testCase)
    utils_helper.get_Input_Value(source_SI_wb, data_source.POL_INFO_SHEET, data_source.CELL_Prem_Term, 3)
    utils_helper.get_Input_Value(source_SI_wb, data_source.POL_INFO_SHEET, data_source.CELL_Scenario_Int, 'L')
    # uat_Tc1 = data_source.UAT_FOLDER + testCase + ".xlsb"
    # print(uat_Tc1)
    target_wb = xw.Book(uat_Tc)
    target_sheet_3y = target_wb.sheets[target_sheet_name_3y]
    utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_COI_Benefit_Low, target_sheet_3y)  # COPY PAGE I - PREMIUM && COPY PAGE II - BENEFIT
    utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_PW_BAV_PW_TAV_Value, target_sheet_3y)  # Copy PW BAV and PW TAV for Low Interest Rate
    utils_helper.get_Copy_Paste_Fund(source_SI_wb, map_SI.get_Fund_List, map_SI.Low_Fund_Source, map_SI.Low_Fund_Target)  # Copy low fund to another space
    utils_helper.get_Input_Value(source_SI_wb, data_source.POL_INFO_SHEET, data_source.CELL_Scenario_Int, 'H')
    utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_COI_Benefit_High, target_sheet_3y)  # COPY PAGE I - COI  &&  COPY PAGE II - BENEFIT
    utils_helper.get_Clear_Content(target_sheet_3y, 'S:S')
    utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Premium_Table, target_sheet_3y)  # PREM TABLE
    utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Benefit_High, target_sheet_3y)  # BENEFIT HIGH TABLE
    utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Benefit_Low, target_sheet_3y)  # BENEFIT LOW TABLE
    utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Fund_Table, target_sheet_3y)  # 'FUND TABLE
    utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Fund_High, target_sheet_3y)
    utils_helper.get_Input_Value(source_SI_wb, data_source.POL_INFO_SHEET, data_source.CELL_Scenario_Int, 'L')
    utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Fund_Low, target_sheet_3y)  # 'Copy Low Fund interest
    ####################################################
    target_sheet_full = target_wb.sheets[target_sheet_name_f]
    flexible_Term = utils_helper.get_Cell_Value(source_SI_wb, data_source.POL_INFO_SHEET, data_source.CELL_FLEXIBLE_IPL26)
    utils_helper.get_Input_Value(source_SI_wb, data_source.POL_INFO_SHEET, data_source.CELL_Prem_Term, flexible_Term)
    utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_COI_Benefit_Low, target_sheet_full)  # COPY PAGE I - PREMIUM && COPY PAGE II - BENEFIT
    utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Benefit_Low, target_sheet_full)  # BENEFIT LOW TABLE
    utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_PW_BAV_PW_TAV_Value, target_sheet_full)  # Copy PW BAV and PW TAV for Low Interest Rate
    utils_helper.get_Copy_Paste_Fund(source_SI_wb, map_SI.get_Fund_List, map_SI.Low_Fund_Source, map_SI.Low_Fund_Target)  # Copy low fund to another space
    utils_helper.get_Input_Value(source_SI_wb, data_source.POL_INFO_SHEET, data_source.CELL_Scenario_Int, 'H')
    utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_COI_Benefit_High, target_sheet_full)  # COPY PAGE I - COI && COPY PAGE II - BENEFIT
    utils_helper.get_Clear_Content(target_sheet_full, 'S:S')
    utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Premium_Table, target_sheet_full)  # PREM TABLE
    utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Benefit_High, target_sheet_full)  # BENEFIT HIGH TABLE
    utils_helper.get_Input_Value(source_SI_wb, data_source.POL_INFO_SHEET, data_source.CELL_Scenario_Int, 'L')
    utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Withdrawal_Low, target_sheet_full)  # Withdrawal LOW
    utils_helper.get_Clear_Content(target_sheet_full, 'AI:AI')
    utils_helper.get_Input_Value(source_SI_wb, data_source.POL_INFO_SHEET, data_source.CELL_Scenario_Int, 'H')
    utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Withdrawal_High, target_sheet_full)
    utils_helper.get_Clear_Content(target_sheet_full, 'AY:AY')
    utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Fund_Table, target_sheet_full)  # 'FUND TABLE BA:BG
    utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Fund_High, target_sheet_full)  # Copy High Fund interest
    utils_helper.get_Input_Value(source_SI_wb, data_source.POL_INFO_SHEET, data_source.CELL_Scenario_Int, 'L')
    utils_helper.get_Copy_Paste_Value(source_sheet, map_SI.get_Fund_Low, target_sheet_full)  # Copy Low Fund interest
    utils_helper.get_Input_Value(source_SI_wb, data_source.POL_INFO_SHEET, data_source.CELL_Scenario_Int, 'H')
    target_wb.save()
    return target_wb


def get_Print_Public_SI_By_Case(source_SI_wb, source_sheet_name, target_sheet_name_3y, target_sheet_name_f, map_SI, testCase):
    utils_helper.get_Input_Value(source_SI_wb, data_source.SI_INFO_SHEET, data_source.CELL_INPUT_TESTCASE_SI, testCase)
    get_Copy_From_SI_File_To_UAT_File(source_SI_wb, source_sheet_name, target_sheet_name_3y, target_sheet_name_f, map_SI, testCase)
    get_Copy_Data_TC_To_SI_File(source_SI_wb, target_sheet_name_3y, target_sheet_name_f, testCase)
    printer_Sheet = source_SI_wb.sheets["Printer"]
    info_Sheet = source_SI_wb.sheets[data_source.SI_INFO_SHEET]
    fileName = info_Sheet.range("C19").value
    # printer_Sheet.page_setup.print_area = "$A$1:$R$1008"
    # printer_Sheet.page_setup.LeftMargin = 0
    # printer_Sheet.page_setup.RightMargin = 0
    # printer_Sheet.page_setup.TopMargin = 0
    # printer_Sheet.page_setup.BottomMargin = 0
    # printer_Sheet.page_setup.HeaderMargin = 0
    # printer_Sheet.page_setup.FooterMargin = 0
    # printer_Sheet.page_setup.PrintHeadings = False
    # printer_Sheet.page_setup.PrintGridlines = False
    # printer_Sheet.page_setup.CenterHorizontally = True
    # printer_Sheet.page_setup.CenterVertically = True
    # printer_Sheet.page_setup.PaperSize = 2 #xlPaperA4
    # printer_Sheet.page_setup.ScaleWithDocHeaderFooter = True
    # printer_Sheet.page_setup.AlignMarginsHeaderFooter = True
    # pdf_path = os.path.join(data_source.PUBLIC_SI_FOLDER + fileName + ".pdf")
    pdf_path = os.path.join(data_source.PUBLIC_SI_FOLDER, fileName + ".pdf")
    print(pdf_path)
    if os.path.exists(pdf_path):
        print(pdf_path + ' is existed')
        os.remove(pdf_path)
        print(pdf_path + ' is removed')
    print("Start Print PDF file")
    printer_Sheet.api.ExportAsFixedFormat(0, pdf_path)
    # try:
    #     printer_Sheet.api.ExportAsFixedFormat(0, pdf_path)
    # except Exception as e:
    #     print(f"An error occurred: {e}")
    # finally:
    #     source_wb.save()


