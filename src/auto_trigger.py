import time
import logging

import create_OutpurMUT_ILP2026

import watchdog.events

TESTCASE_NO = 'B3'
Scenario_Int = 'B56'
POL_INFO = 'Pol_Info'
Prem_Term = 'H56'
SOURCE_FILE_ILP_2026 = '../Resource/SI file/2026ILP Illustration Spreadsheet V6.5.1.xlsb'
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


def verify_File_Change():
    logging.basicConfig(level=logging.INFO,
                        format='%(asctime)s - %(message)s',
                        datefmt='%Y-%m-%d %H:%M:%S')


class Handler(watchdog.events.PatternMatchingEventHandler):
    def __init__(self):
        watchdog.events.PatternMatchingEventHandler.__init__(self, patterns=['*.pdf'], ignore_directories=True,
                                                             case_sensitive=False)

    def on_created(self, event):
        print("Watchdog received created event - % s." % event.src_path)

    def on_modified(self, event):
        print("Watchdog received Modified event - % s." % event.src_path)
        print("START GET FILE")
        create_OutpurMUT_ILP2026.get_Values_From_Interim_ILP(
            source_path=SOURCE_FILE_ILP_2026,
            # target_path=TARGET_FILE,
            source_sheet_name=SOURCE_SHEET,
            target_sheet_name_3y=TARGET_SHEET_3Y,
            target_sheet_name_f=TARGET_SHEET_FULL,
            map_SI=map_SI_ILP2026,
            source_Sheet_Mustpay=SOURCE_SHEET_MUT_MUSTPAY,
            source_Sheet_Full=SOURCE_SHEET_MUT_FULL
        )
        print("DONE!!!!!")


if __name__ == "__main__":
    src_path = r"C:\Users\doananh\OneDrive - Manulife\Desktop\Auto FW\pricing\output\Projection\Template_UAT_riders"
    event_handler = Handler()
    observer = watchdog.observers.Observer()
    observer.schedule(event_handler, path=src_path, recursive=True)
    observer.start()
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
