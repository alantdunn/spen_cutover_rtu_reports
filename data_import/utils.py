import pandas as pd
from typing import List, Dict, Optional
from pathlib import Path

def filter_data_by_rtu(df: pd.DataFrame, rtu_name: str) -> pd.DataFrame:
    """Filter dataframe by RTU name."""
    return df[df['RTU'] == rtu_name]

def filter_data_by_substation(df: pd.DataFrame, substation: str) -> pd.DataFrame:
    """Filter dataframe by substation."""
    return df[df['Sub'] == substation]



# RTUId = (rtu:rtu_address)

# GenericType = SD | DD | A | C

# Example GenericPointAddress
# mk2a format: [(RTUId):card:word- Generic_Type]
# iec101 format: [(RTUId):CASDU:IAO- Generic_Type]
# [(ANDE3:33018):2000:100- DD]
# [(AREC:141):109:4- SD]
# [(ARIE3:33053):312:203- A]
# [(AREC:141):252:6-1 C]

def ignore_habbde_point(row):
    # check if 'PointName' key exists
    if 'PointName' in row:
        if "SPURIOUS ALARM" in row['PointName']:
            return True
    if 'DeviceId' in row:
        if "SPURIOUS" in row['DeviceId']:
            return True
    if 'DeviceType' in row and 'DeviceId' in row:
        if row['DeviceType'] == "UNUSED" and row['DeviceId'] == "SPURIOUS":
            return True
    return False

def get_controllable_for_taps(pointid: str):
    if pointid == 'TCP':
        return '1'
    else:
        return '0'


def derive_rtu_address_and_protocol_from_po_rtu_name(row, eterra_rtu_map: pd.DataFrame):
    # get the eTerra ERU name by removing the text _RTU from the PO_RTU column
    eterra_rtu_name = row['RTU'].replace('_RTU', '')
    # get the rtu_address and protocol from the eterra_rtu_map dataframe
    eterra_rtu_map_row = eterra_rtu_map[eterra_rtu_map['RTU'] == eterra_rtu_name]
    if eterra_rtu_map_row.empty:
        return None, None
    return eterra_rtu_map_row['RTUAddress'].values[0], eterra_rtu_map_row['Protocol'].values[0]

def convert_control_id_to_generic_control_id(control_id, generic_type):
    if generic_type == 'SETPOINT':
        return "2"
    else:
        if control_id == '1':
            return "1"
        else:
            return "0"

def derive_generic_address_for_poweron_export(row):
    # Handle ControlId - convert nan/None/empty to empty string
    if pd.isna(row['ControlId']) or row['ControlId'] == None or row['ControlId'] == '':
        CtrlText = ''
    else:

        CtrlText = convert_control_id_to_generic_control_id(row['ControlId'], row['GenericType'])

    if row['Protocol'] == 'IEC60870-101':
        return pd.Series({
            'GenericPointAddress': f"[{row['RTUId']}:{row['CASDU']}:{str(row['IOA'])}-{CtrlText} {row['GenericType']}]"
        })
    else:
        return pd.Series({
            'GenericPointAddress': f"[{row['RTUId']}:{row['Card']}:{str(row['Offset'])}-{CtrlText} {row['GenericType']}]"
        })
    
def split_ioa(row):
    """
    Split the IOA address (int) into IOA1 and IOA2 (bottom 2 bytes).
    
    Required input columns:
    - IOA: int - IOA address
    
    Returns:
    - tuple: IOA1 and IOA2
    """
    # check IOA column exists and can be converted to int
    try:
        if row['IOA'] is None:
            return None, None
        if row['IOA'] == '':
            return None, None
        
        IOA_int = int(row['IOA'])
    except:
        print (f" :heavy_exclamation_mark: Error: IOA is not an integer: r{row['RTU']}:c{row['Card']}:w{row['Word']}")
        return None, None
    
    return IOA_int >> 16, IOA_int & 0xFFFF

def combine_ioa(ioa1, ioa2):
    """
    Combine the IOA1 and IOA2 into a single IOA value.
    """
    return (ioa1 << 16) | ioa2

def compute_offset(row):
    """
    Compute the offset value for a row based on protocol and record type.
    
    Required input columns:
    - Protocol: str - Protocol type (e.g. 'IEC60870-101')
    - Word: str - Mk2a Word or IEC IOA address 
    - recordType: str - Type of record ('SD', 'DD', 'A', 'A2')
    - rtu: str - RTU identifier (used for error messages)
    - card: str - Card identifier (used for error messages) 
    - size: str - Size value (used for error messages)
    
    Returns:
    - int/float: Computed offset value
    - None: If there are any conversion errors or invalid record types
    """
    # for IEC rtu's just put the IOA into the offset column
    if row['Protocol'] == 'IEC60870-101':
        return str(int(row['Word']))
    
    try:
        word = int(row['Word'])
    except:
        print (f" :heavy_exclamation_mark: Error: word is not an integer: r{row['PO_RTU']}:c{row['Card']}:w{row['Word']}:b{row['Shift']}:s{row['Size']}")
        return None

    try:
        shift = int(row['Shift'])
    except:
        print (f" :heavy_exclamation_mark: Error: shift is not an integer: r{row['PO_RTU']}:c{row['Card']}:w{row['Word']}:b{row['Shift']}:s{row['Size']}")
        return None

    # convert the word to an int and add 1 to account for the fact that the word is 1 based in eTerra
    word_int = int(word)
    offset = 0

    if row['POType'] == 'DI':
        offset = str(int((word_int * 8) + shift))
    elif row['POType'] == 'DD':
        offset = str(int(((word_int * 8) + shift) / 2))
    elif row['POType'] == 'A1':
        offset = str(int(word_int))
    elif row['POType'] == 'A2':
        offset = str(int(word_int))
    else:
        offset = str(int(word_int))
    
    # convert to int and add 1 to account for the fact that the word is 1 based in eTerra
    return str(int(offset) + 1)
    









