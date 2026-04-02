from pathlib import Path

import streamlit as st
import os
import xlwings as xw
import tkinter as tk
from tkinter import filedialog
import sys
import os.path
sys.path.append(str(Path(__file__).parent / "src"))
import ilp2026, utils_helper, map_SI_ILP2026
from data import data_source

UL = 'UL_Copay0.xlsb'
Term = 'TermBase_Copay0.xlsm'
ILP_2026 = '2026ILP Illustration Spreadsheet V6.6_PublicSI_v7_2003.xlsb'
ILP_Package = 'RUV11_V11_Copay0.xlsb'

st.set_page_config(page_title="VERIFY PROJECTION", layout='wide')
st.title("VERIFY PROJECTION")
st.write("RUN SCRIPT AUTOMATE")


# download = "[![Click me](app/static/cat.png)](https://streamlit.io)"
# st.markdown(download)


def radio_change_callback():
    """Updates session state or performs other actions on change."""
    st.session_state.message = f"Option changed to: {st.session_state.selected_option}"
    # You can also change the value of other widgets here via st.session_state


# Initialize session state message
if 'message' not in st.session_state:
    st.session_state.message = "Initial state"

# # st.write(f"Current choice: **{st.session_state.message}**")
#
#
st.radio("How to get SI file", options=['Upload SI File',
                                        'Select SI File'], key="selected_option", on_change=radio_change_callback)

st.write(f"Your choice is: **{st.session_state.selected_option}**")


def get_File_SI():
    col_SI, col2 = st.columns(2)
    if st.session_state.selected_option=="Upload SI File":
        with col_SI:
            upload_File = st.file_uploader("Upload SI file", type=["xlsb", "xlsx"], accept_multiple_files=False)
            if upload_File!=None:
                temp_file_path = f"temp_{upload_File.name}"
                with open(temp_file_path, "wb") as f:
                    f.write(upload_File.getbuffer())
                return temp_file_path
    elif st.session_state.selected_option=="Select SI File":
        with col_SI:
            folder_path = os.path.join(data_source.dir_path, "SI file")
            selected_SI_File = st.selectbox("Select SI to run", [UL, ILP_2026, ILP_Package,
                                                                 Term], index=None, placeholder="Select file to run...", )
            if selected_SI_File!=None:
                file_path = os.path.join(folder_path, selected_SI_File)
                return file_path


source_wb = get_File_SI()


def get_Open_SI_File():
    if source_wb!=None:
        return xw.Book(source_wb)


selected_Type_to_Test = st.selectbox("Type run: ", ["PWS_By_List", "Compare_PDF_SI_By_Case_List",
                                                    "PWS_By_Case",
                                                    "Compare_PDF_SI_By_Case"], index=None, placeholder="Select file to run...", )


def select_Dir_Save_File():
    root = tk.Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    if st.button('Save file as'):
        folder_selected = filedialog.askdirectory(master=root)
        st.write("File save as: " + folder_selected)
        return os.path.abspath(folder_selected)


def select_Dir_File_Projection_PDF():
    root = tk.Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    if st.button('PDF file'):
        folder_selected = filedialog.askdirectory(master=root)
        st.write("Folder content PDF file: " + folder_selected)
        return os.path.abspath(folder_selected)


def get_Type_To_Run():
    match selected_Type_to_Test:
        case 'PWS_By_Case':
            input_Tc = st.text_input("Input your testcase")
            path = select_Dir_Save_File()
            return input_Tc, path
        case 'Compare_PDF_SI_By_Case':
            input_Tc = st.text_input("Input your testcase")
            path = select_Dir_File_Projection_PDF()
            return input_Tc, path
        case 'PWS_By_List':
            upload_TC_File = st.file_uploader("Upload TC file", type=["xlsb", "xlsx"], accept_multiple_files=False)
            path = select_Dir_Save_File()
            return upload_TC_File, path
        case 'Compare_PDF_SI_By_Case_List':
            upload_TC_File = st.file_uploader("Upload TC file", type=["xlsb", "xlsx"], accept_multiple_files=False)
            path = select_Dir_File_Projection_PDF()
            return upload_TC_File, path


testCase = get_Type_To_Run()


def on_button_click_Run_Public_SI_By_Case():
    st.session_state.clicked = True
    st.write("Current testcase: " + testCase[0])
    data_source.PUBLIC_SI_FOLDER = testCase[1]
    ilp2026.get_Print_Public_SI_By_Case(
        get_Open_SI_File(),
        data_source.SOURCE_SHEET,
        data_source.TARGET_SHEET_3Y,
        data_source.TARGET_SHEET_FULL,
        map_SI_ILP2026,
        testCase[0]
    )


def on_button_click_Run_Public_SI_By_List():
    st.session_state.clicked = True
    data_source.PUBLIC_SI_FOLDER = select_Dir_Save_File()
    data_source.TESTCASE_FILE = testCase
    LIST_TCS = utils_helper.get_List_Tcs()
    ilp2026.get_Print_Public_SI_By_List(
        get_Open_SI_File(),
        data_source.SOURCE_SHEET,
        data_source.TARGET_SHEET_3Y,
        data_source.TARGET_SHEET_FULL,
        map_SI_ILP2026
    )


def on_button_click_Run_Compare_SI_By_Case():
    st.session_state.clicked = True
    data_source.PUBLIC_SI_FOLDER = testCase[0]
    st.write("Current testcase: " + testCase)
    source_wb_SI = get_Open_SI_File()
    utils_helper.get_Input_Value(source_wb_SI, data_source.POWERQUERY_SHEET, data_source.CELL_LINK_POWERQUERY,
        testCase[1])
    ilp2026.get_Result_Projection_ILP_By_Case(
        source_wb_SI,
        data_source.SOURCE_SHEET,
        data_source.TARGET_SHEET_3Y,
        data_source.TARGET_SHEET_FULL,
        map_SI_ILP2026,
        testCase[0]
    )


def on_Click_Button():
    match selected_Type_to_Test:
        case 'PWS_By_Case':
            on_button_click_Run_Public_SI_By_Case()
        case 'Compare_PDF_SI_By_Case':
            on_button_click_Run_Compare_SI_By_Case()
        case 'PWS_By_List':
            on_button_click_Run_Public_SI_By_List()
        case 'Compare_PDF_SI_By_Case_List':
            st.write("Compare_PDF_SI_By_Case_List")


st.button('Trigger', on_click=on_Click_Button)
