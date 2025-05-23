import pandas as pd
from data_import.utils import derive_rtu_address_and_protocol_from_po_rtu_name, convert_control_id_to_generic_control_id

def clean_controls_test(df: pd.DataFrame, eterra_rtu_map: pd.DataFrame) -> pd.DataFrame:
    """Clean the controls test dataframe."""
    # | Original Column         | New Column
    # |-------------------------|------------
    # | RTU                     | 
    # | Protocol                | 
    # | RTU_Ctrl                | 
    # | Test_Command            | 
    # | control_time            | 
    # | control_address         | AutoTestAddress
    # | control_status          | AutoTestStatus
    # | control_result          | AutoTestResult
    # | event_status            | 
    # | component_alias         | AutoTestAlias
    # | control_attribute       | AutoTestAttribute
    # | telecontrol_action      | AutoTestAction
    # | template_alias          | 
    # | template_name           | 
    # | pre_control_value       | 
    # | expected_value          | 
    # | found_value             | 
    # | number_of_events        | 
    # | event1_message          | 
    # | event1_text             | 
    # | event1_class            | 
    # | event1_zone             | 
    # | event1_substation       | 
    # | event2_message          | 
    # | event2_text             | 
    # | event2_class            | 
    # | event2_zone             | 
    # | event2_substation       | 
    # | event3_message          | 
    # | event3_text             | 
    # | event3_class            | 
    # | event3_zone             | 
    # | event3_substation       | 

    df.rename(columns={
        'control_address': 'AutoTestAddress',
        'control_status': 'AutoTestStatus',
        'control_result': 'AutoTestResult',
        'component_alias': 'AutoTestAlias',
        'control_attribute': 'AutoTestAttribute',
        'telecontrol_action': 'AutoTestAction'
    }, inplace=True)

    def convert_po_rtu_eterra_rtu_name(row):
        if row['RTU'] == '':
            return None
        else:
            return row['RTU'].replace('_RTU', '')
        
    def convert_poweron_word_to_generic_word(row):
        if row['Word'] == '':
            return None
        else:
            return str(int(row['Word'])+1) # +1 because the word is 1-based in eterra
        
    # decompose the AutoTestAddress into Card, Word, and CtrlId
    df['Card'] = df['AutoTestAddress'].str.split(':').str[0]
    df['Word'] = df['AutoTestAddress'].str.split(':').str[1]
    df['Word'] = df.apply(convert_poweron_word_to_generic_word, axis=1)
    df['CtrlId'] = df['AutoTestAddress'].str.split(':').str[2]
    df['GenericType'] = "C"
    # get the rtu_address and protocol from the RTU and the eterra_rtu_map dataframe
    df[['RTUAddress', 'Protocol']] = df.apply(lambda row: pd.Series(derive_rtu_address_and_protocol_from_po_rtu_name(row, eterra_rtu_map)), axis=1)
    df['RTU'] = df.apply(convert_po_rtu_eterra_rtu_name, axis=1)
    df['CtrlId'] = df.apply(lambda row: convert_control_id_to_generic_control_id(row['CtrlId'], row['GenericType']), axis=1)
    df['GenericPointAddress'] = '[(' + df['RTU'].astype(str) + ':' + df['RTUAddress'].astype(str) + '):' + df['Card'].astype(str) + ':' + df['Word'].astype(str) + '-' + df['CtrlId'].astype(str) + ' C]'

    # Only return the columns we need
    # We will only keep the columns in the New Column section
    columns_to_keep = [
        'AutoTestAddress',
        'AutoTestStatus',
        'AutoTestResult',
        'AutoTestAlias',
        'AutoTestAttribute',
        'AutoTestAction',
        'GenericPointAddress'
    ]
    return df
