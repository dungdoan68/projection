COI_Low_Copy = 'Q8:Q106'
COI_High_Copy = 'R8:R106'
Benefit_Low_Copy = 'X8:AF106'
Benefit_High_Copy = 'AP8:AX106'
Prem_Benefit_Copy_Low = 'T8:W106'
Prem_Benefit_Copy_High = 'T8:W106'
Prem_Copy = 'A8:P106'

F1_Low_Copy = 'BI8:BT106'
F2_Low_Copy = 'CI8:CT106'
F3_Low_Copy = 'DI8:DT106'
F4_Low_Copy = 'EI8:ET106'
F5_Low_Copy = 'FI8:FT106'
F6_Low_Copy = 'GI8:GT106'
F7_Low_Copy = 'HI8:HT106'
F1_High_Copy = 'BV8:CG106'
F2_High_Copy = 'CV8:DG106'
F3_High_Copy = 'DV8:EG106'
F4_High_Copy = 'EV8:FG106'
F5_High_Copy = 'FV8:GG106'
F6_High_Copy = 'GV8:HG106'
F7_High_Copy = 'HV8:IG106'
FUND_COPY = 'BI2:IG2'

COI_Low_Paste_1 = 'Q8'
COI_High_Paste_1 = 'R8'
Benefit_Low_Paste_1 = 'X8'
Benefit_High_Paste_1 = 'AL8'
Prem_Paste_1 = 'A8'
Prem_Benefit_Low_Paste_1 = 'T8'
Prem_Benefit_High_Paste_1 = 'AH8'
FUND_PASTE = 'BE2'
F1_Low_Paste_1 = 'BE8'
F2_Low_Paste_1 = 'CE8'
F3_Low_Paste_1 = 'DE8'
F4_Low_Paste_1 = 'EE8'
F5_Low_Paste_1 = 'FE8'
F6_Low_Paste_1 = 'GE8'
F7_Low_Paste_1 = 'HE8'
F1_High_Paste_1 = 'BR8'
F2_High_Paste_1 = 'CR8'
F3_High_Paste_1 = 'DR8'
F4_High_Paste_1 = 'ER8'
F5_High_Paste_1 = 'FR8'
F6_High_Paste_1 = 'GR8'
F7_High_Paste_1 = 'HR8'


FUND_TABLE_COPY = 'BA8:BG106'
FUND_TABLE_PASTE = 'AW8'
Low_Fund_Source = 'E3:BF1205'
Low_Fund_Target = 'BG3'
PW_BAV_PW_TAV_COPY = 'BJ8:BK106'
PW_BAV_PW_TAV_PASTE = 'BH8'


get_Key_Values = {
    # key         source    target
    'COI_Low': [COI_Low_Copy, COI_Low_Paste_1],
    'COI_High': [COI_High_Copy, COI_High_Paste_1],
    'Benefit_Low': [Benefit_Low_Copy, Benefit_Low_Paste_1],
    'Benefit_High': [Benefit_High_Copy, Benefit_High_Paste_1],
    'Prem_Copy': [Prem_Copy, Prem_Paste_1],
    'Prem_Benefit_Low': [Prem_Benefit_Copy_Low, Prem_Benefit_Low_Paste_1],
    'Prem_Benefit_High': [Prem_Benefit_Copy_High, Prem_Benefit_High_Paste_1],
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
get_Key_Values_Low = {
    # key         source    target
    # 'COI_Low': [COI_Low_Copy, COI_Low_Paste_1],
    # 'Benefit_Low': [Benefit_Low_Copy, Benefit_Low_Paste_1],
    # 'Prem_Copy': [Prem_Copy, Prem_Paste_1],
    'Prem_Benefit_Low': [Prem_Benefit_Copy_Low, Prem_Benefit_Low_Paste_1],
    'FUND': [FUND_COPY, FUND_PASTE],
    'F1_Low': [F1_Low_Copy, F1_Low_Paste_1],
    'F2_Low': [F2_Low_Copy, F2_Low_Paste_1],
    'F3_Low': [F3_Low_Copy, F3_Low_Paste_1],
    'F4_Low': [F4_Low_Copy, F4_Low_Paste_1],
    'F5_Low': [F5_Low_Copy, F5_Low_Paste_1],
    'F6_Low': [F6_Low_Copy, F6_Low_Paste_1],
    'F7_Low': [F7_Low_Copy, F7_Low_Paste_1],
}
get_Key_Values_High = {
    # key         source    target
    # 'COI_High': [COI_High_Copy, COI_High_Paste_1],
    # 'Benefit_High': [Benefit_High_Copy, Benefit_High_Paste_1],
    'Prem_Copy': [Prem_Copy, Prem_Paste_1],
    'Prem_Benefit_High': [Prem_Benefit_Copy_High, Prem_Benefit_High_Paste_1],
    'F1_High': [F1_High_Copy, F1_High_Paste_1],
    'F2_High': [F2_High_Copy, F2_High_Paste_1],
    'F3_High': [F3_High_Copy, F3_High_Paste_1],
    'F4_High': [F4_High_Copy, F4_High_Paste_1],
    'F5_High': [F5_High_Copy, F5_High_Paste_1],
    'F6_High': [F6_High_Copy, F6_High_Paste_1],
    'F7_High': [F7_High_Copy, F7_High_Paste_1]
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


