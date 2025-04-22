import pandas as pd
from typing import List, Dict, Optional
from pathlib import Path

def filter_data_by_rtu(df: pd.DataFrame, rtu_id: str) -> pd.DataFrame:
    """Filter dataframe by RTU ID."""
    return df[df['RTU_ID'] == rtu_id]

def filter_data_by_substation(df: pd.DataFrame, substation: str) -> pd.DataFrame:
    """Filter dataframe by substation."""
    return df[df['Substation'] == substation]

def create_points_section(df: pd.DataFrame) -> pd.DataFrame:
    """Create the points section of the report."""
    points = []
    
    for _, row in df.iterrows():
        point = {
            'Type': 'Input',
            'SD/DD': row.get('SD_DD', ''),
            'SCADA Address': f"{row.get('RTU')}-{row.get('Card')}-{row.get('Word')}-{row.get('Size')}",
            'eTerra Key': row.get('eTerra_Key', ''),
            'PowerOn Alias': row.get('PowerOn_Alias', ''),
            'ICCP Flag': row.get('ICCP_Flag', ''),
            'Habdde Match Status': row.get('Match_Status', ''),
            'PowerOn Config Health Status': row.get('Config_Status', ''),
            'Alarm Match Status': row.get('Alarm_Status', ''),
            'Control Zone Status': row.get('Control_Zone', ''),
            'Controllable Flag': row.get('Controllable', ''),
            'Circuit Suggestion': row.get('Circuit', '')
        }
        points.append(point)
        
        # Add controls if they exist
        if row.get('Has_Controls', False):
            control = {
                'Type': 'Control',
                'DO': row.get('DO', ''),
                'SCADA Address': f"{row.get('RTU')}-{row.get('Card')}-{row.get('Word')}-{row.get('ControlId')}",
                'eTerra Key': row.get('Control_eTerra_Key', ''),
                'PowerOn Alias': row.get('Control_PowerOn_Alias', ''),
                'Auto Test Status': row.get('Auto_Test_Status', ''),
                'Manual Test Status': row.get('Manual_Test_Status', ''),
                'Manual Test Comments': row.get('Manual_Test_Comments', '')
            }
            points.append(control)
    
    return pd.DataFrame(points)

def create_analogs_section(df: pd.DataFrame) -> pd.DataFrame:
    """Create the analogs section of the report."""
    analogs = []
    
    for _, row in df.iterrows():
        analog = {
            'Type': 'Analog',
            'A': row.get('A', ''),
            'SCADA Address': f"{row.get('RTU')}-{row.get('Card')}-{row.get('Word')}-{row.get('Size')}",
            'eTerra Key': row.get('eTerra_Key', ''),
            'PowerOn Alias': row.get('PowerOn_Alias', ''),
            'ICCP Flag': row.get('ICCP_Flag', ''),
            'Habdde Match Status': row.get('Match_Status', ''),
            'PowerOn Config Health Status': row.get('Config_Status', ''),
            'Analog Value Match Status': row.get('Value_Match_Status', '')
        }
        analogs.append(analog)
        
        # Add controls if they exist
        if row.get('Has_Controls', False):
            control = {
                'Type': 'Control',
                'AO': row.get('AO', ''),
                'SCADA Address': f"{row.get('RTU')}-{row.get('Card')}-{row.get('Word')}-{row.get('ControlId')}",
                'eTerra Key': row.get('Control_eTerra_Key', ''),
                'PowerOn Alias': row.get('Control_PowerOn_Alias', ''),
                'Auto Test Status': row.get('Auto_Test_Status', ''),
                'Manual Test Status': row.get('Manual_Test_Status', ''),
                'Manual Test Comments': row.get('Manual_Test_Comments', '')
            }
            analogs.append(control)
    
    return pd.DataFrame(analogs)

def save_report(df: pd.DataFrame, output_path: Path):
    """Save the report to an Excel file."""
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='RTU Report', index=False)
        
        # Apply formatting
        worksheet = writer.sheets['RTU Report']
        for column in worksheet.columns:
            max_length = 0
            column = [cell for cell in column]
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column[0].column_letter].width = adjusted_width 

def derive_rtu_address_and_protocol_from_po_rtu_name(row, eterra_rtu_map: pd.DataFrame):
    # get the eTerra ERU name by removing the text _RTU from the PO_RTU column
    eterra_rtu_name = row['RTU'].replace('_RTU', '')
    # get the rtu_address and protocol from the eterra_rtu_map dataframe
    eterra_rtu_map_row = eterra_rtu_map[eterra_rtu_map['RTU'] == eterra_rtu_name]
    if eterra_rtu_map_row.empty:
        return None, None
    return eterra_rtu_map_row['RTUAddress'].values[0], eterra_rtu_map_row['Protocol'].values[0]


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
        return row['Word']
    
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

    if row['POType'] == 'DI':
        return (word * 8) + shift
    elif row['POType'] == 'DD':
        return (((word * 8) + shift) / 2)
    elif row['POType'] == 'A1':
        return word 
    elif row['POType'] == 'A2':
        return word
    
    return word
    




# RTUId = (rtu:rtu_address)

# GenericType = SD | DD | A | C

# Example GenericPointAddress
# mk2a format: [(RTUId):card:word- Generic_Type]
# iec101 format: [(RTUId):CASDU:IAO- Generic_Type]
# [(ANDE3:33018):2000:100- DD]
# [(AREC:141):109:4- SD]
# [(ARIE3:33053):312:203- A]
# [(AREC:141):252:6-1 C]

def derive_addresses_for_habdde_export(row):

    if row['GenericType'] not in ['CTRL', 'SETPOINT']:
        CtrlText = ""
    elif row['CtrlFunc'] == '' or row['CtrlFunc'] is None:
        CtrlText = ""
    else:
        CtrlText =  row['CtrlFunc']

    if row['Protocol'] == 'MK2A':
        return pd.Series({
            'CASDU': None,
            'IOA': None, 
            'IOA1': None,
            'IOA2': None,
            'GenericPointAddress': f"[{row['RTUId']}:{row['Card']}:{row['Word']}-{CtrlText} {row['GenericType']}]"
        })
    else:
        # Convert Word to int
        try:
            ioa1 = int(row['Card'])
            ioa2 = int(row['Word'])
            ioa = combine_ioa(ioa1, ioa2)
        except:
            print (f" :heavy_exclamation_mark: Error: Word is not an integer: r{row['RTU']}:c{row['Card']}:w{row['Word']} ({row['GenericType']})")
            return pd.Series({
                'CASDU': row['CASDU'],
                'IOA': None, 
                'IOA1': None,
                'IOA2': None,
                'GenericPointAddress': None})
        
        return pd.Series({
            'CASDU': row['CASDU'],
            'IOA': str(ioa),
            'IOA1': str(ioa1),
            'IOA2': str(ioa2),
            'GenericPointAddress': f"[{row['RTUId']}:{row['CASDU']}:{str(ioa)}-{CtrlText} {row['GenericType']}]"
        })
    
def derive_generic_address_for_poweron_export(row):
    # Handle ControlId - convert nan/None/empty to empty string
    if pd.isna(row['ControlId']) or row['ControlId'] == None or row['ControlId'] == '':
        CtrlText = ''
    else:
        CtrlText = str(row['ControlId'])

    if row['Protocol'] == 'IEC60870-101':
        return pd.Series({
            'GenericPointAddress': f"[{row['RTUId']}:{row['CASDU']}:{str(row['IOA'])}-{CtrlText} {row['GenericType']}]"
        })
    else:
        return pd.Series({
            'GenericPointAddress': f"[{row['RTUId']}:{row['Card']}:{str(row['Word'])}-{CtrlText} {row['GenericType']}]"
        })

def clean_eterra_point_export(df: pd.DataFrame) -> pd.DataFrame:
    """Clean the eTerra point export dataframe."""
    # | Original Column     | New Column
    # |---------------------|------------
    # | Enabled             | 
    # | RowNumber           | 
    # | eTerraKey           | eTerraKey
    # | sub                 | Sub
    # | devtyp              | DeviceType
    # | device_id           | DeviceId
    # | device_name         | DeviceName
    # | point_id            | PointId
    # | point_name          | PointName
    # | fep                 | 
    # | area                | eTerraZone
    # | site2               | 
    # | rtu                 | RTU
    # | address1            | CASDU
    # | rtu_address         | RTUAddress
    # | card                | Card and IOA1
    # | phyadr              | Word and IOA2
    # | concat_conect       | Size: 1 if concat_conect is 0 else 2
    # | Three               | 
    # | itpnd               | 
    # | ztpnd               | 
    # | pnttyp              | eTerraPtyType
    # | catpnt              | 
    # | text00              | 
    # | itpnd.1             | 
    # | ztpnd.1             | 
    # | sinvt               | Inverted
    # | protocol            | Protocol
    # | ctrlable            | Controllable
    # | arg                 | 
    # |                     | GenericType : derived from Size (pnttyp) to SD or DD
    # |                     | IOA : IOA1 << 16 + IOA2
    # |                     | RTUId : (RTU:RTUAddress)
    # |                     | GenericPointAddress : [(RTUId):card:word- Generic_Type] or [(RTUId):CASDU:IAO- Generic_Type]
    # |                     | eTerraAlias : Sub/DeviceType/DeviceId/PointId

    # rename the columns to the new column names using the mapping in the New Column section - skip the columns that are not in the New Column section
    df.rename(columns={
        'sub': 'Sub',
        'devtyp': 'DeviceType',
        'device_id': 'DeviceId',
        'device_name': 'DeviceName',
        'point_id': 'PointId',
        'point_name': 'PointName',
        'area': 'eTerraZone',
        'rtu': 'RTU',
        'address1': 'CASDU',
        'rtu_address': 'RTUAddress',
        'card': 'Card',
        'phyadr': 'Word',
        'pnttyp': 'eTerraPtyType',
        'sinvt': 'Inverted',
        'protocol': 'Protocol',
        'ctrlable': 'Controllable',
    }, inplace=True)

    # Derive the columns we need from the columns we have
    df['eTerraAlias'] = df['Sub'] + '/' + df['DeviceType'] + '/' + df['DeviceId'] + '/' + df['PointId']
    df['RTUId'] = '(' + df['RTU'] + ':' + df['RTUAddress'].astype(str) + ')'

    # Convert Size from 1->2 and 0->1
    # Convert concat_conect to int
    df['concat_conect'] = df['concat_conect'].astype(int)
    df['Size'] = df['concat_conect'].apply(lambda x: 2 if x == 1 else 1 if x == 0 else x)

    def derive_generic_type(row):
        if row['Size'] == 1:
            return 'SD'
        elif row['Size'] == 2:
            return 'DD'
        else:
            return None

    df['GenericType'] = df.apply(derive_generic_type, axis=1)

    df[['CASDU', 'IOA', 'IOA1', 'IOA2', 'GenericPointAddress']] = df.apply(derive_addresses_for_habdde_export, axis=1)

    # strip the eTerraKey of any leading or trailing whitespace
    df['eTerraKey'] = df['eTerraKey'].str.strip()

    # Only return the columns we need
    # We will only keep the columns in the New Column section
    columns_to_keep = [
        'eTerraKey',
        'eTerraAlias',
        'Sub',
        'DeviceType',
        'DeviceId',
        'DeviceName',
        'PointId',
        'PointName',
        'eTerraZone',
        'RTU',
        'RTUAddress',
        'Card',
        'Word',
        'CASDU',
        'IOA',
        'IOA1',
        'IOA2',
        'Size',
        'Inverted',
        'Protocol',
        'Controllable',
        'eTerraPtyType',
        'RTUId',
        'GenericPointAddress',
        'GenericType'
    ]  
    df = df[columns_to_keep]
    return df

def clean_eterra_analog_export(df: pd.DataFrame) -> pd.DataFrame:
    """Clean the eTerra analog export dataframe."""
    # | Original Column         | New Column
    # |-------------------------|------------
    # | eTerraKey               | eTerraKey
    # | sub                     | Sub
    # | devtyp                  | DeviceType
    # | device_id               | DeviceId
    # | device_name             | DeviceName
    # | analog_id               | PointId
    # | lo_reas                 | 
    # | hi_reas                 | 
    # | fep                     | 
    # | area                    | eTerraZone
    # | site2                   | 
    # | rtu                     | RTU
    # | address1                | CASDU
    # | rtu_address             | RTUAddress
    # | card                    | Card
    # | word                    | Word
    # | rawhigh                 | RawHigh
    # | rawlow                  | RawLow
    # | enghigh                 | EngHigh
    # | englow                  | EngLow
    # | itpnd                   | eTerraPtyType
    # | DIS                     | 
    # | disprior                | 
    # | protocol                | Protocol
    # | loreas                  | LoReas
    # | hireas                  | HiReas
    # | clmpdbnd                | ClmpDbnd
    # | pospolar                | PosPolar
    # | negpolar                | NegPolar
    # | negate                  | Negate
    # |                         | GenericType : A
    # |                         | IOA1 : IOA >> 8
    # |                         | IOA2 : IOA & 0xFF
    # |                         | RTUId : (RTU:RTUAddress)
    # |                         | GenericPointAddress : [(RTUId):card:word- Generic_Type] or [(RTUId):CASDU:IAO- Generic_Type]
    # |                         | eTerraAlias : Sub/DeviceType/DeviceId/PointId

    # rename the columns to the new column names using the mapping in the New Column section - skip the columns that are not in the New Column section
    df.rename(columns={
        'sub': 'Sub',
        'devtyp': 'DeviceType',
        'device_id': 'DeviceId',
        'device_name': 'DeviceName',
        'analog_id': 'PointId',
        'lo_reas': 'LoReas',
        'hi_reas': 'HiReas',
        'area': 'eTerraZone',
        'rtu': 'RTU',
        'address1': 'CASDU',
        'rtu_address': 'RTUAddress',
        'card': 'Card',
        'word': 'Word',
        'rawhigh': 'RawHigh',
        'rawlow': 'RawLow',
        'enghigh': 'EngHigh',
        'englow': 'EngLow',
        'protocol': 'Protocol',
        'clmpdbnd': 'ClmpDbnd',
        'pospolar': 'PosPolar',
        'negpolar': 'NegPolar',
        'negate': 'Negate',
    }, inplace=True)

    # Derive the columns we need from the columns we have
    df['eTerraAlias'] = df['Sub'] + '/' + df['DeviceType'] + '/' + df['DeviceId'] + '/' + df['PointId']
    df['RTUId'] = '(' + df['RTU'] + ':' + df['RTUAddress'].astype(str) + ')'

    # Set the GenericType to A for all rows
    df['GenericType'] = 'A'

    # Derive the GenericPointAddress from the RTUId, Card, Word, and GenericType
    df[['CASDU', 'IOA', 'IOA1', 'IOA2', 'GenericPointAddress']] = df.apply(derive_addresses_for_habdde_export, axis=1)

    # strip the eTerraKey of any leading or trailing whitespace
    df['eTerraKey'] = df['eTerraKey'].str.strip()

    # Create a dummy Controllable field to make columns match digital columns
    df['Controllable'] = '0'

    # Only return the columns we need
    # We will only keep the columns in the New Column section
    columns_to_keep = [
        'eTerraKey',
        'eTerraAlias',
        'Sub',
        'DeviceType',
        'DeviceId',
        'DeviceName',
        'PointId',
        'LoReas',
        'HiReas',
        'eTerraZone',
        'RTU',
        'RTUAddress',
        'Card',
        'Word',
        'RawHigh',
        'RawLow',
        'EngHigh',
        'EngLow',
        'Protocol',
        'ClmpDbnd',
        'PosPolar',
        'NegPolar',
        'Negate',
        'RTUId',
        'GenericPointAddress',
        'GenericType',
        'CASDU',
        'IOA',
        'IOA1',
        'IOA2',
        'Controllable' # Dummy field to make columns match digital columns
    ]
    
    df = df[columns_to_keep]
    return df

def clean_eterra_control_export(df: pd.DataFrame) -> pd.DataFrame:
    """Clean the eTerra control export dataframe."""
    # | Original Column         | New Column
    # |-------------------------|------------
    # | Enabled                 | 
    # | RowNumber               | 
    # | eTerraKey               | eTerraKey
    # | sub                     | Sub
    # | devtyp                  | DeviceType
    # | device_id               | DeviceId
    # | device_name             | DeviceName
    # | point_id                | PointId
    # | control_id              | ControlId
    # | rtu                     | RTU
    # | rtu_address             | RTUAddress
    # | card                    | Card and IOA1
    # | phyadr                  | Word and IOA2
    # | mdlparm1                | Parm1
    # | mdlparm2                | Parm2
    # | mdlparm3                | Parm3
    # | ctrlfunc                | CtrlFunc
    # | protocol                | Protocol
    # | address_id              | 
    # | address                 | CASDU
    # |                         | GenericType : CTRL
    # |                         | IOA : IOA1 << 16 + IOA2
    # |                         | RTUId : (RTU:RTUAddress)
    # |                         | GenericPointAddress : [(RTUId):card:word- Generic_Type] or [(RTUId):CASDU:IAO- Generic_Type]
    # |                         | eTerraAlias : Sub/DeviceType/DeviceId/PointId

    # rename the columns to the new column names using the mapping in the New Column section - skip the columns that are not in the New Column section
    df.rename(columns={
        'sub': 'Sub',
        'devtyp': 'DeviceType',
        'device_id': 'DeviceId',
        'device_name': 'DeviceName',
        'point_id': 'PointId',
        'control_id': 'ControlId',
        'rtu': 'RTU',
        'rtu_address': 'RTUAddress',
        'card': 'Card',
        'phyadr': 'Word',
        'mdlparm1': 'Parm1',
        'mdlparm2': 'Parm2',
        'mdlparm3': 'Parm3',
        'protocol': 'Protocol',
        'ctrlfunc': 'CtrlFunc',
        'address': 'CASDU',
    }, inplace=True)

    # Derive the columns we need from the columns we have
    df['eTerraAlias'] = df['Sub'] + '/' + df['DeviceType'] + '/' + df['DeviceId'] + '/' + df['PointId']
    df['RTUId'] = '(' + df['RTU'] + ':' + df['RTUAddress'].astype(str) + ')'

    # Set the GenericType to CTRL for all rows
    df['GenericType'] = 'CTRL'

    # Derive the GenericPointAddress from the RTUId, Card, Word, and GenericType
    df[['CASDU', 'IOA', 'IOA1', 'IOA2', 'GenericPointAddress']] = df.apply(derive_addresses_for_habdde_export, axis=1)

    # strip the eTerraKey of any leading or trailing whitespace
    df['eTerraKey'] = df['eTerraKey'].str.strip()

    # Only return the columns we need
    # We will only keep the columns in the New Column section
    columns_to_keep = [
        'eTerraKey',
        'eTerraAlias',
        'Sub',
        'DeviceType',
        'DeviceId',
        'DeviceName',
        'PointId',
        'ControlId',
        'RTU',
        'RTUAddress',
        'Card',
        'Word',
        'Parm1',
        'Parm2',
        'Parm3',
        'CtrlFunc',
        'Protocol',
        'RTUId',
        'GenericPointAddress',
        'GenericType',
        'CASDU',
        'IOA',
        'IOA1',
        'IOA2'
    ]
    
    df = df[columns_to_keep]
    return df

def clean_eterra_setpoint_control_export(df: pd.DataFrame) -> pd.DataFrame:
    """Clean the eTerra setpoint control export dataframe."""
    # These are only for 2 IEC RTUs, so don't need MK2A variants for address derivation
    # | Column Name        | Description
    # |--------------------|-------------
    # | Enabled            | 
    # | RowNumber          |    
    # | eTerraKey          | eTerraKey
    # | sub                | Sub
    # | devtyp             | DeviceType
    # | device_id          | DeviceId
    # | device_name        | DeviceName
    # | analog_id          | PointId
    # | rtu                | RTU
    # | rtu_address        | RTUAddress
    # | address1           | CASDU
    # | card               | IOA1
    # | phyadr             | IOA2
    # | mdlparm1           | 
    # | mdlparm2           | CtrlFunc
    # | mdlparm3           | 
    # | mdlparm4           | 
    # | mdlparm5           | 
    # | protocol           | Protocol
    # | address_id         | 
    # | enghigh            | EngHigh
    # | englow             | EngLow
    # |                    | GenericType : C
    # |                    | IOA : IOA1 << 16 + IOA2
    # |                    | RTUId : (RTU:RTUAddress)
    # |                    | GenericPointAddress : [(RTUId):CASDU:IAO- Generic_Type]
    # |                    | eTerraAlias : Sub/DeviceType/DeviceId/PointId

    # rename the columns to the new column names using the mapping in the New Column section - skip the columns that are not in the New Column section
    df.rename(columns={
        'sub': 'Sub',
        'devtyp': 'DeviceType',
        'device_id': 'DeviceId',
        'device_name': 'DeviceName',
        'analog_id': 'PointId',
        'rtu': 'RTU',
        'rtu_address': 'RTUAddress',
        'address1': 'CASDU',
        'card': 'IOA1',
        'phyadr': 'IOA2',
        'protocol': 'Protocol',
        'mdlparm2': 'CtrlFunc',
        'enghigh': 'EngHigh',
        'englow': 'EngLow',
    }, inplace=True)

    # Set the GenericType to C for all rows
    df['GenericType'] = 'SETPOINT'

    # Derive the GenericPointAddress from the RTUId, Card, Word, and GenericType
    df['RTUId'] = '(' + df['RTU'] + ':' + df['RTUAddress'].astype(str) + ')'
    df['eTerraAlias'] = df['Sub'] + '/' + df['DeviceType'] + '/' + df['DeviceId'] + '/' + df['PointId']

    # Set the card and word for compatibility with common functions
    df['Card'] = df['IOA1']
    df['Word'] = df['IOA2']

    # Derive the GenericPointAddress from the RTUId, CASDU, IOA1, IOA2, and GenericType
    df[['CASDU', 'IOA', 'IOA1', 'IOA2', 'GenericPointAddress']] = df.apply(derive_addresses_for_habdde_export, axis=1)

    # strip the eTerraKey of any leading or trailing whitespace
    df['eTerraKey'] = df['eTerraKey'].str.strip()
    
    # Only return the columns we need
    # We will only keep the columns in the New Column section
    columns_to_keep = [
        'eTerraKey',
        'eTerraAlias',
        'Sub',
        'DeviceType',
        'DeviceId',
        'DeviceName',
        'PointId',
        'RTU',
        'RTUAddress',
        'CASDU',
        'IOA1',
        'IOA2',
        'CtrlFunc',
        'EngHigh',
        'EngLow',
        'Protocol',
        'GenericPointAddress',
        'GenericType'
    ]
    
    df = df[columns_to_keep]
    return df


def clean_habdde_compare(df: pd.DataFrame) -> pd.DataFrame:
    """Clean the habdde compare dataframe."""
    # | Original Column         | New Column
    # |-------------------------|------------
    # | lineNo                  | 
    # | RTUid                   | 
    # | matched_status          | HbddeCompareStatus
    # | GenericPointAddress     | GenericPointAddress
    # | Generic_Type            | 
    # | Protocol                | 
    # | Description             | 
    # | Key                     | HabCompKey
    # | ID                      | 
    # | Key2                    | 
    # | RTUname                 | 
    # | RTUnum                  | 
    # | FUNC                    | 
    # | CASDU                   | 
    # | Area                    | 
    # | ID_SUBSTN               | 
    # | ID_DEVTYP               | 
    # | ID_DEVICE               | 
    # | Name_device             | 
    # | Bit                     | 
    # | ID_POINT                | 
    # | PointType               | 
    # | Priority                | 
    # | StateOn                 | 
    # | NormalState             | 
    # | SingleOrDouble          | 
    # | Site                    | 
    # | Controlable             | 
    # | PART                    | 
    # | ID_ANALOG               | 
    # | HighRawValue            | 
    # | LowRawValue             | 
    # | HighEngValue            | 
    # | LowEngValue             | 
    # | ScaledValue             | 
    # | ID_POINT_CTRL           | 
    # | ID_CTRL                 | 
    # | PhyadrRelay             | 
    # | MdlParm1Relay           | 
    # | MdlParm2Relay           | 
    # | MdlParm3Relay           | 
    # | ctrlfunc                | 
    # | IdAnalog                | 
    # | PhyAdrAnout             | 
    # | MdlParm1                | 
    # | MdlParm2                | 
    # | MdlParm3                | 
    # | MdlParm4                | 
    # | MdlParm5                | 
    # | IdAdrs                  | 
    # | PO PointAddress         | 
    # | PO RecordType           | 
    # | PO Description          | 
    # | PO Component Name       | 
    # | PO StateAlarmText       | 
    # | PO ComponentAlias       | 
    # | PO AttributeName        | 
    # | PO Action               | 
    # | PO ConfigHealth         | 
    # | PO ControlValue         | 
    # | PO EterraDevID          | 
    # | PO EterraDevType        | 
    # | PO EterraPointID        | 
    # | PO EterraPointName      | 
    # | PO EterraSubstation     | 
    # | PO VerifyAttribute      | 
    # | PO VerifyValue          | 
    # | PO Protocol             | 
    
    # rename the columns to the new column names using the mapping in the New Column section - skip the columns that are not in the New Column section
    df.rename(columns={
        'matched_status': 'HbddeCompareStatus',
        'GenericPointAddress': 'GenericPointAddress',
        'Key': 'HabCompKey'
    }, inplace=True)

    # Only return the columns we need
    # We will only keep the columns in the New Column section
    columns_to_keep = [
        'HbddeCompareStatus',
        'GenericPointAddress',
        'HabCompKey'
    ]

    df = df[columns_to_keep]
    return df

def clean_all_rtus(df: pd.DataFrame) -> pd.DataFrame:
    """Clean the all rtus dataframe."""
    # | Original Column         | New Column
    # |-------------------------|------------
    # | Protocol                | 
    # | RTU                     | PO_RTU
    # | RTU Address             | 
    # | addr1                   | Card or CASDU
    # | addr2                   | Word or IOA
    # | comp_alias              | POAlias
    # | comp_name               | 
    # | config_extra_info       | ConfigInfo
    # | config_health           | ConfigHealth
    # | control_attr            | 
    # | control_row             | 
    # | control_type            | 
    # | control_val             | ControlId
    # | desc                    | PODescription
    # | eng_max                 | 
    # | eng_min                 | 
    # | eterra_dev_id           | 
    # | eterra_dev_type         | 
    # | eterra_point_id         | 
    # | eterra_point_name       | 
    # | eterra_sub              | 
    # | raw_max                 | 
    # | raw_min                 | 
    # | recordType              | POType
    # | scan_attr               | 
    # | scan_row                | ScanInputRow
    # | shift                   | Shift
    # | siref1                  | ScanInputRef
    # | size                    | Size
    # | state_alarm_text        | 
    # | symbol_menu             | Menu
    # | symbol_name             | Symbol
    # | telecontrol_action      | TC Action
    # | verify_attribute        | 
    # | verify_value            | 
    # |                         | IOA1
    # |                         | IOA2
    # |                         | Offset
    # |                         | GenericType
    # |                         | GenericPointAddress
    # |                         | eTerraAlias

    # rename the columns to the new column names using the mapping in the New Column section - skip the columns that are not in the New Column section
    df.rename(columns={
        'Protocol': 'Protocol',
        'RTU': 'PO_RTU',
        'RTU Address': 'RTUAddress',
        'eterra_sub': 'Sub',
        'eterra_dev_type': 'DeviceType',
        'eterra_dev_id': 'DeviceId',
        'eterra_point_id': 'PointId',
        'addr1': 'Card',
        'addr2': 'Word',
        'comp_alias': 'POAlias',
        'comp_name': 'POName',
        'control_val': 'ControlId',
        'config_extra_info': 'ConfigInfo',
        'config_health': 'ConfigHealth',
        'desc': 'PODescription',
        'recordType': 'POType',
        'scan_row': 'ScanInputRow',
        'shift': 'Shift',
        'siref1': 'ScanInputRef',
        'size': 'Size',
        'symbol_menu': 'Menu',
        'symbol_name': 'Symbol',
        'telecontrol_action': 'TC Action'
    }, inplace=True)

    def derive_generic_type_from_po_type(po_type):
        if po_type in ['A1', 'A2', 'A4']:
            return 'A'
        elif po_type == 'DI':
            return 'SD'
        elif po_type == 'DD':
            return 'DD'
        elif po_type == 'DO':
            return 'C'
        elif po_type == 'AO':
            return 'SETPOINT'
        else:
            return 'Unknown'
        
    # Derive the GenericType from the POType
    df['GenericType'] = df['POType'].apply(derive_generic_type_from_po_type)

    # Derive the GenericPointAddress from the RTUId, Card, Word, and GenericType
    df['RTU'] = df['PO_RTU'].str.replace('_RTU', '')
    df['RTUId'] = '(' + df['RTU'] + ':' + df['RTUAddress'].astype(str) + ')'
    df['eTerraAlias'] = df['Sub'] + '/' + df['DeviceType'] + '/' + df['DeviceId'] + '/' + df['PointId']

    # Set the card and word for compatibility with common functions
    df['CASDU'] = df['Card']
    df['IOA'] = df['Word']
    df['IOA1', 'IOA2]'] = df.apply(split_ioa, axis=1)
    df['Offset'] = df.apply(compute_offset, axis=1)

    # Derive the GenericPointAddress from the RTUId, CASDU, IOA1, IOA2, and GenericType
    df[['GenericPointAddress']] = df.apply(derive_generic_address_for_poweron_export, axis=1)


    # Change a few columns to unique names before we return this dataframe
    df.rename(columns={
        'Protocol': 'PO_Protocol',
        'Card': 'PO_Card',
        'Word': 'PO_Word',
        'GenericType': 'PO_GenericType',
        'eTerraAlias': 'PO_eTerraAlias'
    }, inplace=True)

    # Only return the columns we need
    # We will only keep the columns in the New Column section
    columns_to_keep = [
        'PO_Protocol',
        'PO_RTU',
        'PO_Card',
        'PO_Word',
        'POAlias',
        'POName',
        'ConfigInfo',
        'ConfigHealth',
        'PODescription',
        'POType',
        'ScanInputRow',
        'Shift',
        'ScanInputRef',
        'Size',
        'Menu',
        'Symbol',
        'TC Action',
        'PO_GenericType',
        'GenericPointAddress',
        'PO_eTerraAlias'
    ]

    df = df[columns_to_keep]
    return df

def clean_compare_alarms(df: pd.DataFrame) -> pd.DataFrame:
    """Clean the compare alarms dataframe."""
    # | Original Column         | New Column
    # |-------------------------|------------
    # | RTU_Name                | CompAlarmRTU
    # | RTU_Address             | CompAlarmRTUAddress
    # | eTerra Alias            | CompAlarmEterraAlias
    # | PO Alias                | CompAlarmPOAlias
    # | Type                    | CompAlarmType
    # | Card                    | CompAlarmCard
    # | Offset                  | CompAlarmOffset
    # | Value                   | CompAlarmValue
    # | eTerraSubstation        | CompAlarmeTerraSubstation
    # | eTerraAlarmMessage      | CompAlarmeTerraAlarmMessage
    # | eTerraAlarmZone         | CompAlarmeTerraAlarmZone
    # | eTerraStatus            | CompAlarmeTerraStatus
    # | POSubstation            | CompAlarmPOsubstation
    # | POAlarmMessage          | CompAlarmPOAlarmMessage
    # | POAlarmZone             | CompAlarmPOAlarmZone
    # | POAlarmValue            | CompAlarmPOAlarmValue
    # | POAlarmRef              | CompAlarmPOAlarmRef
    # | POStatus                | CompAlarmPOStatus
    # | EventCategory           | 
    # | DevID                   | 
    # | PointID                 | 
    # | AlarmType               | 
    # | etoken1                 | eToken1
    # | etoken2                 | eToken2
    # | etoken3                 | eToken3
    # | etoken4                 | eToken4
    # | etoken5                 | eToken5
    # | ptoken1                 | pToken1
    # | ptoken2                 | pToken2
    # | ptoken3                 | pToken3
    # | ptoken4                 | pToken4
    # | ptoken5                 | pToken5
    # | T1Match                 | T1Match
    # | T2Match                 | T2Match
    # | T3Match                 | T3Match
    # | T4Match                 | T4Match
    # | T5Match                 | T5Match
    # | new_match               | CompAlarmNewMatch
    # | MatchScore              | CompAlarmMatchScore
    # | AlarmMessageMatch       | AlarmMessageMatch
    # | AlarmZoneMatch          | AlarmZoneMatch
    # | NumControls             | 
    # | DCB                     | 
    # | 314                     | 
    # | SC1E                    | 
    # | SC2E                    | 

    df.rename(columns={
        'RTU_Name': 'CompAlarmRTU',
        'RTU_Address': 'CompAlarmRTUAddress',
        'eTerra Alias': 'CompAlarmEterraAlias',
        'PO Alias': 'CompAlarmPOAlias',
        'Type': 'CompAlarmType',
        'Card': 'CompAlarmCard',
        'Offset': 'CompAlarmOffset',
        'Value': 'CompAlarmValue',
        'eTerraSubstation': 'CompAlarmeTerraSubstation',
        'eTerraAlarmMessage': 'CompAlarmeTerraAlarmMessage',
        'eTerraAlarmZone': 'CompAlarmeTerraAlarmZone',
        'eTerraStatus': 'CompAlarmeTerraStatus',
        'POSubstation': 'CompAlarmPOsubstation',
        'POAlarmMessage': 'CompAlarmPOAlarmMessage',
        'POAlarmZone': 'CompAlarmPOAlarmZone',
        'POAlarmValue': 'CompAlarmPOAlarmValue',
        'POAlarmRef': 'CompAlarmPOAlarmRef',
        'POStatus': 'CompAlarmPOStatus',
        'etoken1': 'eToken1',
        'etoken2': 'eToken2',
        'etoken3': 'eToken3',
        'etoken4': 'eToken4',
        'etoken5': 'eToken5',
        'ptoken1': 'pToken1',
        'ptoken2': 'pToken2',
        'ptoken3': 'pToken3',
        'ptoken4': 'pToken4',
        'ptoken5': 'pToken5',
        'T1Match': 'T1Match',
        'T2Match': 'T2Match',
        'T3Match': 'T3Match',
        'T4Match': 'T4Match',
        'T5Match': 'T5Match',
        'new_match': 'CompAlarmNewMatch',
        'MatchScore': 'CompAlarmMatchScore',
        'AlarmMessageMatch': 'CompAlarmAlarmMessageMatch',
        'AlarmZoneMatch': 'CompAlarmAlarmZoneMatch'
    }, inplace=True)

    # Only return the columns we need
    # We will only keep the columns in the New Column section
    columns_to_keep = [
        'CompAlarmRTU',
        'CompAlarmRTUAddress',
        'CompAlarmEterraAlias',
        'CompAlarmPOAlias',
        'CompAlarmType',
        'CompAlarmCard',
        'CompAlarmOffset',
        'CompAlarmValue',
        'CompAlarmeTerraSubstation',
        'CompAlarmeTerraAlarmMessage',
        'CompAlarmeTerraAlarmZone',
        'CompAlarmeTerraStatus',
        'CompAlarmPOsubstation',
        'CompAlarmPOAlarmMessage',
        'CompAlarmPOAlarmZone',
        'CompAlarmPOAlarmValue',
        'CompAlarmPOAlarmRef',
        'CompAlarmPOStatus',
        'eToken1',
        'eToken2',
        'eToken3',
        'eToken4',
        'eToken5',
        'pToken1',
        'pToken2',
        'pToken3',
        'pToken4',
        'pToken5',
        'T1Match',
        'T2Match',
        'T3Match',
        'T4Match',
        'T5Match',
        'CompAlarmNewMatch',
        'CompAlarmMatchScore',
        'CompAlarmAlarmMessageMatch',
        'CompAlarmAlarmZoneMatch'
    ]

    df = df[columns_to_keep]
    return df

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

    # decompose the AutoTestAddress into Card, Word, and CtrlId
    df['Card'] = df['AutoTestAddress'].str.split(':').str[0]
    df['Word'] = df['AutoTestAddress'].str.split(':').str[1]
    df['CtrlId'] = df['AutoTestAddress'].str.split(':').str[2]
    # get the rtu_address and protocol from the RTU and the eterra_rtu_map dataframe
    df[['RTUAddress', 'Protocol']] = df.apply(lambda row: pd.Series(derive_rtu_address_and_protocol_from_po_rtu_name(row, eterra_rtu_map)), axis=1)
    df['GenericPointAddress'] = '[(' + df['RTUAddress'].astype(str) + ':' + df['Protocol'].astype(str) + '):' + df['Card'].astype(str) + ':' + df['Word'].astype(str) + '-' + df['CtrlId'].astype(str) + ' C]'

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

def clean_manual_commissioning(df: pd.DataFrame) -> pd.DataFrame:
    """Clean the manual commissioning dataframe."""
    # | Original Column         | New Column
    # |-------------------------|------------
    # | testset                 | CommissioningTestset
    # | testdate                | CommissioningTestdate
    # | user                    | CommissioningUser
    # | control_address         | CommissioningControlAddress
    # | test_name               | CommissioningTestName
    # | result                  | CommissioningResult
    # | comments                | CommissioningComments
    # | RTUname                 | CommissioningRTUname
    # | voltage_group           | CommissioningVoltageGroup
    # | test_area               | CommissioningTestArea
    # | alias                   | CommissioningAlias

    df.rename(columns={
        'testset': 'CommissioningTestset',
        'testdate': 'CommissioningTestdate',
        'user': 'CommissioningUser',
        'control_address': 'CommissioningControlAddress',
        'test_name': 'CommissioningTestName',
        'result': 'CommissioningResult',
        'comments': 'CommissioningComments',
        'RTUname': 'CommissioningRTUname',
        'voltage_group': 'CommissioningVoltageGroup',
        'test_area': 'CommissioningTestArea',
        'alias': 'CommissioningAlias'
    }, inplace=True)


    return df

def add_control_info_to_eterra_export(eterra_export: pd.DataFrame, eterra_control_export: pd.DataFrame, eterra_setpoint_control_export: pd.DataFrame) -> pd.DataFrame:
    """Add control info to the eterra export dataframe."""

    # For each row in the eterra export dataframe, add the control info from the eterra control and eterra setpoint control dataframes
    # For digitals (SD or DD), there are between 0 and 2 controls per point in the eterra_control_export
    # For analogs (A), there are between 0 and 1 controls per point in the eterra_setpoint_control_export

    '''
    For Digitals we want to add the following info into Ctrl1Addr, Ctrl1Name, Ctrl2Addr, Ctrl2Name:
     - GenericPointAddress -> Ctrl1Addr or Ctrl2Addr
     - ControlId -> Ctrl1Name or Ctrl2Name
     For Analogs we want to add the following info into Ctrl1Addr, Ctrl1Name:
     - GenericPointAddress -> Ctrl1Addr 
     - "SETPNT" -> Ctrl1Name
    '''

    def get_control_info(row):
        # Get the control info from the eterra control and eterra setpoint control dataframes
        control_info = eterra_control_export[eterra_control_export['eTerraAlias'] == row['eTerraAlias']]
        return control_info 
    
    def get_setpoint_control_info(row):
        # Get the control info from the eterra setpoint control dataframe
        control_info = eterra_setpoint_control_export[eterra_setpoint_control_export['eTerraAlias'] == row['eTerraAlias']]
        return control_info

    # For each row in the eterra export dataframe, add the control info from the eterra control and eterra setpoint control dataframes
    for _, row in eterra_export.iterrows():

        # Initialize the Ctrl1Addr and Ctrl1Name columns
        eterra_export.at[_, 'Ctrl1Addr'] = ''
        eterra_export.at[_, 'Ctrl1Name'] = ''
        eterra_export.at[_, 'Ctrl2Addr'] = ''
        eterra_export.at[_, 'Ctrl2Name'] = ''

        if row['Controllable'] == '1':
            # Get the control info from the eterra control and eterra setpoint control dataframes
            control_info = get_control_info(row)

            # check how many controls are in the control_info dataframe
            if control_info.shape[0] > 0:
                # Add the first control info to the eterra export dataframe
                eterra_export.at[_, 'Ctrl1Addr'] = control_info.iloc[0]['GenericPointAddress']
                eterra_export.at[_, 'Ctrl1Name'] = control_info.iloc[0]['ControlId']
            if control_info.shape[0] > 1:
                # Add the second control info to the eterra export dataframe
                eterra_export.at[_, 'Ctrl2Addr'] = control_info.iloc[1]['GenericPointAddress']
                eterra_export.at[_, 'Ctrl2Name'] = control_info.iloc[1]['ControlId']

        if row['GenericType'] == 'A':
            # Get the control info from the eterra setpoint control dataframe
            control_info = get_setpoint_control_info(row)
            # Add the control info to the eterra setpoint control dataframe
            if control_info.shape[0] > 0:
                eterra_export.at[_, 'Ctrl1Addr'] = control_info.iloc[0]['GenericPointAddress']
                eterra_export.at[_, 'Ctrl1Name'] = 'SETPOINT'

    return eterra_export