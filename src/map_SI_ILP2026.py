COI_Low_Copy = 'Q8:R106'
COI_High_Copy = 'R8:S106'
Benefit_Low_Copy = 'X8:AH106'
Benefit_High_Copy = 'AR8:BB106'
Prem_Benefit_Copy = 'T8:W106'
Prem_Copy = 'A8:P106'

F1_Low_Copy = 'BM8:BX106'
F2_Low_Copy = 'CM8:CX106'
F3_Low_Copy = 'DM8:DX106'
F4_Low_Copy = 'EM8:EX106'
F5_Low_Copy = 'FM8:FX106'
F6_Low_Copy = 'GM8:GX106'
F7_Low_Copy = 'HM8:HX106'
F1_High_Copy = 'BZ8:CK106'
F2_High_Copy = 'CZ8:DK106'
F3_High_Copy = 'DZ8:EK106'
F4_High_Copy = 'EZ8:FK106'
F5_High_Copy = 'FZ8:GK106'
F6_High_Copy = 'GZ8:HK106'
F7_High_Copy = 'HZ8:IK106'
FUND_COPY = 'BM1:IK3'

COI_Low_Paste_1 = 'Q8'
COI_High_Paste_1 = 'R8'
Benefit_Low_Paste_1 = 'X8'
Benefit_High_Paste_1 = 'AN8'
Prem_Paste_1 = 'A8'
Prem_Benefit_Low_Paste_1 = 'T8'
Prem_Benefit_High_Paste_1 = 'AJ8'
FUND_PASTE = 'BA2'
F1_Low_Paste_1 = 'BK8'
F2_Low_Paste_1 = 'CK8'
F3_Low_Paste_1 = 'DK8'
F4_Low_Paste_1 = 'EK8'
F5_Low_Paste_1 = 'FK8'
F6_Low_Paste_1 = 'GK8'
F7_Low_Paste_1 = 'HK8'
F1_High_Paste_1 = 'BX8'
F2_High_Paste_1 = 'CX8'
F3_High_Paste_1 = 'DX8'
F4_High_Paste_1 = 'EX8'
F5_High_Paste_1 = 'FX8'
F6_High_Paste_1 = 'GX8'
F7_High_Paste_1 = 'HX8'

FUND_TABLE_COPY = 'BE8:BK106'
FUND_TABLE_PASTE = 'BA8'
Low_Fund_Source = 'E3:BG1205'
Low_Fund_Target = 'BH3'
PW_BAV_PW_TAV_COPY = 'BJ8:BK106'
PW_BAV_PW_TAV_PASTE = 'BH8'
WITHDRAWAL_LOW_COPY = 'AH8:AI106'
WITHDRAWAL_LOW_PASTE = 'AH8'
WITHDRAWAL_HIGH_COPY = 'BB8:BC106'
WITHDRAWAL_HIGH_PASTE = 'AX8'

get_Key_Values = {
    # key         source    target
    'COI_Low': [COI_Low_Copy, COI_Low_Paste_1],
    'COI_High': [COI_High_Copy, COI_High_Paste_1],
    'Benefit_Low': [Benefit_Low_Copy, Benefit_Low_Paste_1],
    'Benefit_High': [Benefit_High_Copy, Benefit_High_Paste_1],
    'Prem_Copy': [Prem_Copy, Prem_Paste_1],
    'Prem_Benefit_Low': [Prem_Benefit_Copy, Prem_Benefit_Low_Paste_1],
    'Prem_Benefit_High': [Prem_Benefit_Copy, Prem_Benefit_High_Paste_1],
    'FUND': [FUND_COPY, FUND_PASTE],
    'F1_Low': [F1_Low_Copy, F1_Low_Paste_1],
    'F2_Low': [F2_Low_Copy, F2_Low_Paste_1],
    'F3_Low': [F3_Low_Copy, F3_Low_Paste_1],
    'F4_Low': [F4_Low_Copy, F4_Low_Paste_1],
    'F5_Low': [F5_Low_Copy, F5_Low_Paste_1],
    'F6_Low': [F6_Low_Copy, F6_Low_Paste_1],
    'F7_Low': [F7_Low_Copy, F7_Low_Paste_1],
    'F1_High': [F1_High_Copy, F1_High_Paste_1],
    'F2_High': [F2_High_Copy, F2_High_Paste_1],
    'F3_High': [F3_High_Copy, F3_High_Paste_1],
    'F4_High': [F4_High_Copy, F4_High_Paste_1],
    'F5_High': [F5_High_Copy, F5_High_Paste_1],
    'F6_High': [F6_High_Copy, F6_High_Paste_1],
    'F7_High': [F7_High_Copy, F7_High_Paste_1]
}
get_Fund_Low = {
    # key         source    target
    # 'FUND': [FUND_COPY, FUND_PASTE],
    'F1_Low': [F1_Low_Copy, F1_Low_Paste_1],
    'F2_Low': [F2_Low_Copy, F2_Low_Paste_1],
    'F3_Low': [F3_Low_Copy, F3_Low_Paste_1],
    'F4_Low': [F4_Low_Copy, F4_Low_Paste_1],
    'F5_Low': [F5_Low_Copy, F5_Low_Paste_1],
    'F6_Low': [F6_Low_Copy, F6_Low_Paste_1],
    'F7_Low': [F7_Low_Copy, F7_Low_Paste_1],
}
get_Fund_High = {
    # key         source    target
    'F1_High': [F1_High_Copy, F1_High_Paste_1],
    'F2_High': [F2_High_Copy, F2_High_Paste_1],
    'F3_High': [F3_High_Copy, F3_High_Paste_1],
    'F4_High': [F4_High_Copy, F4_High_Paste_1],
    'F5_High': [F5_High_Copy, F5_High_Paste_1],
    'F6_High': [F6_High_Copy, F6_High_Paste_1],
    'F7_High': [F7_High_Copy, F7_High_Paste_1]
}

get_Benefit_High = {
    # key         source    target
    'Prem_Benefit_High': [Prem_Benefit_Copy, Prem_Benefit_High_Paste_1],
}

get_Benefit_Low = {
    # key         source    target
    'Prem_Benefit_Low': [Prem_Benefit_Copy, Prem_Benefit_Low_Paste_1],
}

get_Premium_Table = {
    # key         source    target
    'Prem_Copy': [Prem_Copy, Prem_Paste_1],
}

get_Fund_List = {
    'Money_Market': ['MM'],
    'Fixed_Income': ['FI'],
    'Diversified': ['DI'],
    'Balance': ['BA'],
    'Growth': ['GR'],
    'Target_Fund_2035': ['35'],
    'Target_Fund_2040': ['40'],
    'Target_Fund_2045': ['45'],
    'Aggressive': ['AG']
}

get_COI_Benefit_Low = {
    # key         source    target
    'COI_Low': [COI_Low_Copy, COI_Low_Paste_1],
    'Benefit_Low': [Benefit_Low_Copy, Benefit_Low_Paste_1],
}

get_COI_Benefit_High = {
    # key         source    target
    'COI_High': [COI_High_Copy, COI_High_Paste_1],
    'Benefit_High': [Benefit_High_Copy, Benefit_High_Paste_1],
}

get_Fund_Table = {
    'FUND_TABLE': [FUND_TABLE_COPY, FUND_TABLE_PASTE]
}

get_PW_BAV_PW_TAV_Value = {
    'PW_BAV_PW_TAV': [PW_BAV_PW_TAV_COPY, PW_BAV_PW_TAV_PASTE]
}
get_Withdrawal_Low = {
    'WITHDRAWAL_LOW': [WITHDRAWAL_LOW_COPY, WITHDRAWAL_LOW_PASTE]
}

get_Withdrawal_High = {
    'WITHDRAWAL_HIGH': [WITHDRAWAL_HIGH_COPY, WITHDRAWAL_HIGH_PASTE]
}

get_Remove_Redundant = {
    'UDR': ['E:E', 'F:F'],
    'ROI': ['G:G', 'H:H'],
    'GB': ['K:K', 'K:K'],
    'Regular_Special_Loyalty': ['AG:AG', 'AJ:AJ'],
    'Load_UDR': ['AI:AI', 'AI:AI'],
    'UDR_Allocated-Base_Allocated_Premium ': ['AO:AO', 'AP:AP'],
}
