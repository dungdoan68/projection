import os

dir_path = os.path.dirname(os.path.abspath(__file__))
cwd = os.getcwd()


SOURCE_FILE_ILP_2026_PUBLIC_SI = os.path.join(dir_path, 'SI file/2026ILP Illustration Spreadsheet V6.6_PublicSI_v7_2003.xlsb')
TESTCASE_FILE = os.path.join(dir_path,'testCases.xlsx')
SOURCE_PUBLIC_SI_ILP_2026 = 'Resource'
TARGET_FILE = "../2.xlsb"
SOURCE_SHEET = 'Interim'
TARGET_SHEET_3Y = 'Final_3yrs_Prem'
TARGET_SHEET_FULL = 'Final_flex_Prem'
ORIGINAL_MUT_FILE = 'MUT_Sample - Copy.xlsx'
ORIGINAL_UAT_FILE = os.path.join(dir_path , "Original_ILP2026.xlsb")
TESTCASE_SHEET = 'TC'
SOURCE_SHEET_MUT_MUSTPAY = 'Output MUT_3yrs_Prem'
SOURCE_SHEET_MUT_FULL = 'Output MUT_flex_Prem'
CELL_FLEXIBLE_IPL26 = 'H58'
MUT_SHEET_NAME = 'BaseAndRider'
UAT_FOLDER = os.path.join(dir_path , "UAT_Output\\")
# PUBLIC_SI_FOLDER = 'C:/Users/doananh/Documents/Public_SI_Folder/'
PUBLIC_SI_FOLDER = ""
#COPY MUT DATA OUTPUT
CELL_MUSTPAY_COPY_IPL26 = 'C16'
CELL_FULL_MUSTPAY_PASTE = 'A1'
CELL_FULL_COPY_IPL26 = 'C11'
CONTROL_SHEET = 'Control'
TESTCASE_RESULT_SHEET = 'Result'
CELL_INPUT_TESTCASE = 'B3'
CELL_INPUT_TESTCASE_SI = 'C34'
CELL_Scenario_Int = 'B56'
POL_INFO_SHEET = 'Pol_Info'
SI_INFO_SHEET = 'Info'
CELL_Prem_Term = 'H56'
POWERQUERY_SHEET='Input_PowerQuery'
CELL_LINK_POWERQUERY='D2'