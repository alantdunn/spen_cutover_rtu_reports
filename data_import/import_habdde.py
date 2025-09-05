
from data_import.utils import combine_ioa, ignore_habbde_point, convert_control_id_to_generic_control_id, get_controllable_for_taps
from pylib3i.habdde import read_habdde_tab_into_df, read_habdde_point_tab_into_df
import pandas as pd

def import_habdde_export_point_tab(file_path: str, debug_dir: str) -> pd.DataFrame:
    """Import the HABDDE export file."""
    
    eterra_point_export = read_habdde_point_tab_into_df(habdde_file=file_path,keep_cols=None,remove_dummy_points=False, try_to_use_sql_cache=True)
    if eterra_point_export is None:
        raise ValueError("Failed to read POINT tab from eTerra export")
    eterra_point_export = clean_eterra_point_export(eterra_point_export)
    eterra_point_export = eterra_point_export[~eterra_point_export.apply(ignore_habbde_point, axis=1)]

    if debug_dir:
        eterra_point_export.to_csv(f"{debug_dir}/eterra_point_export.csv", index=False)

    return eterra_point_export

def import_habdde_export_analog_tab(file_path: str, debug_dir: str) -> pd.DataFrame:
    """Import the HABDDE export file."""
    eterra_analog_export = read_habdde_tab_into_df(file_path, 'ANALOG', keep_cols=None, remove_dummy_points=True, try_to_use_sql_cache=True)
    if eterra_analog_export is None:
        raise ValueError("Failed to read ANALOG tab from eTerra export")
    eterra_analog_export = clean_eterra_analog_export(eterra_analog_export)
    eterra_analog_export = eterra_analog_export[~eterra_analog_export.apply(ignore_habbde_point, axis=1)]

    if debug_dir:
        eterra_analog_export.to_csv(f"{debug_dir}/eterra_analog_export.csv", index=False)

    return eterra_analog_export

def import_habdde_export_control_tab(file_path: str, debug_dir: str) -> pd.DataFrame:
    """Import the HABDDE export file."""
    eterra_control_export = read_habdde_tab_into_df(file_path, 'CTRL', keep_cols=None, remove_dummy_points=True, try_to_use_sql_cache=True)
    if eterra_control_export is None:
        raise ValueError("Failed to read CTRL tab from eTerra export")
    eterra_control_export = clean_eterra_control_export(eterra_control_export)
    eterra_control_export = eterra_control_export[~eterra_control_export.apply(ignore_habbde_point, axis=1)]
    if debug_dir:
        eterra_control_export.to_csv(f"{debug_dir}/eterra_control_export.csv", index=False)

    return eterra_control_export

def import_habdde_export_setpoint_control_tab(file_path: str, debug_dir: str) -> pd.DataFrame:
    """Import the HABDDE export file."""
    eterra_setpoint_control_export = read_habdde_tab_into_df(file_path, 'SETPNT', keep_cols=None, remove_dummy_points=True, try_to_use_sql_cache=True)
    if eterra_setpoint_control_export is None:
        raise ValueError("Failed to read SETPNT tab from eTerra export")
    
    eterra_setpoint_control_export = clean_eterra_setpoint_control_export(eterra_setpoint_control_export)
    eterra_setpoint_control_export = eterra_setpoint_control_export[~eterra_setpoint_control_export.apply(ignore_habbde_point, axis=1)]
    if debug_dir:
        eterra_setpoint_control_export.to_csv(f"{debug_dir}/eterra_setpoint_control_export.csv", index=False)

    return eterra_setpoint_control_export


def derive_rtu_addresses_and_protocols_from_eterra_export(eterra_point_export: pd.DataFrame, debug_dir: str) -> pd.DataFrame:
    """Derive the RTU addresses and protocols from the eTerra export."""
    eterra_rtu_map = eterra_point_export[['RTU', 'RTUAddress', 'Protocol']].drop_duplicates()
    if debug_dir:
        eterra_rtu_map.to_csv(f"{debug_dir}/eterra_rtu_map.csv", index=False)
    return eterra_rtu_map


def derive_addresses_for_habdde_export(row):

    if row['GenericType'] not in ['CTRL', 'SETPOINT']:
        CtrlText = ""
        TypeText = row['GenericType']
    else:
        TypeText = "C"
        if row['CtrlFunc'] == '' or row['CtrlFunc'] is None:
            CtrlText = ""
        else:
            CtrlText =  convert_control_id_to_generic_control_id(row['CtrlFunc'], row['GenericType'])

    if row['Protocol'] == 'MK2A':
        return pd.Series({
            'CASDU': None,
            'IOA': None, 
            'IOA1': None,
            'IOA2': None,
            'GenericPointAddress': f"[{row['RTUId']}:{row['Card']}:{row['Word']}-{CtrlText} {TypeText}]"
        })
    else:

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
            'GenericPointAddress': f"[{row['RTUId']}:{row['CASDU']}:{str(ioa)}-{CtrlText} {TypeText}]"
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
    # | values              | OPTIONAL - MAY NOT BE PRESENT
    # | xdis                | OPTIONAL - MAY NOT BE PRESENT
    # | sdis                | OPTIONAL - MAY NOT BE PRESENT
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
        # look for Dummy Rows - Card and CASDU will be empty or Nan
        if pd.isna(row['Card']) and pd.isna(row['CASDU']):
            return 'DUMMY'
        elif row['Size'] == 1:
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
    # Only keep columns that exist in the dataframe
    available_columns = [
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
        'GenericType',
        'IGNORE_RTU',
        'IGNORE_POINT',
        'OLD_DATA',
        'GridIncomer',
        'eTerra Alias',
        'ICCP_POINTNAME',
        'ICCP->PO',
        'ICCP_ALIAS',
        'PowerOn Alias',
        'PowerOn Alias Exists',
        'PowerOn Alias Linked to SCADA',
        'values',
        'xdis', 
        'sdis'
    ]

    # Only keep columns that exist in df to avoid KeyError
    columns_to_keep = []
    for col in available_columns:
        if col in df.columns:
            columns_to_keep.append(col)

    # Ensure we have at least some columns before filtering
    if len(columns_to_keep) > 0:
        df = df[columns_to_keep]
    else:
        print("Warning: No matching columns found in dataframe")
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
        'itpnd': 'eTerraPtyType',
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

    # Create a dummy Controllable field to make columns match digital columns, and set the Taps to controllable
    df['Controllable'] = '0'
    df['Controllable'] = df.apply(lambda row: get_controllable_for_taps(row['PointId']), axis=1)

    # Only return the columns we need
    # We will only keep the columns in the New Column section
    available_columns = [
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
        'eTerraPtyType',
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
        ,
        'IGNORE_RTU',
        'IGNORE_POINT',
        'OLD_DATA',
        'GridIncomer',
        'eTerra Alias',
        'ICCP_POINTNAME',
        'ICCP->PO',
        'ICCP_ALIAS',
        'PowerOn Alias',
        'PowerOn Alias Exists',
        'PowerOn Alias Linked to SCADA'
    ]
    columns_to_keep = [col for col in available_columns if col in df.columns]
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

def add_control_info_to_eterra_export(eterra_export: pd.DataFrame, eterra_control_export: pd.DataFrame, eterra_setpoint_control_export: pd.DataFrame, all_rtus: pd.DataFrame, controls_test: pd.DataFrame, manual_commissioning: pd.DataFrame) -> pd.DataFrame:
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
        # if the point id is TCP then edit the eTerraAlias to swap TCP for TAP
        if row['PointId'] == 'TCP':
            row['eTerraAlias'] = row['eTerraAlias'].replace('TCP', 'TAP')

        control_info = eterra_control_export[eterra_control_export['eTerraAlias'] == row['eTerraAlias']]

        # if the point id is TCP then edit the eTerraAlias to swap back to TCP
        if row['PointId'] == 'TCP':
            control_info.loc[:,'eTerraAlias'] = control_info['eTerraAlias'].replace('TAP', 'TCP')

        return control_info 
    
    def get_setpoint_control_info(row):
        # Get the control info from the eterra setpoint control dataframe
        control_info = eterra_setpoint_control_export[eterra_setpoint_control_export['eTerraAlias'] == row['eTerraAlias']]
        return control_info
    
    def get_control_po_config(row, all_rtus: pd.DataFrame):
        pass

    # For each row in the eterra export dataframe, add the control info from the eterra control and eterra setpoint control dataframes
    for _, row in eterra_export.iterrows():

        # Initialize the Ctrl1Addr and Ctrl1Name columns
        eterra_export.at[_, 'Ctrl1Addr'] = ''
        eterra_export.at[_, 'Ctrl1Name'] = ''
        eterra_export.at[_, 'Ctrl1SecureBit'] = ''
        eterra_export.at[_, 'Ctrl1SyncChannel'] = ''
        eterra_export.at[_, 'Ctrl1IECSingleDouble'] = ''
        eterra_export.at[_, 'Ctrl1Func'] = ''
        eterra_export.at[_, 'Ctrl2Addr'] = ''
        eterra_export.at[_, 'Ctrl2Name'] = ''
        eterra_export.at[_, 'Ctrl2SecureBit'] = ''
        eterra_export.at[_, 'Ctrl2SyncChannel'] = ''
        eterra_export.at[_, 'Ctrl2IECSingleDouble'] = ''
        eterra_export.at[_, 'Ctrl2Func'] = ''

        if row['Controllable'] == '1':
            # Get the control info from the eterra control and eterra setpoint control dataframes
            control_info = get_control_info(row)

            # check how many controls are in the control_info dataframe
            if control_info.shape[0] > 0:
                # Add the first control info to the eterra export dataframe
                eterra_export.at[_, 'Ctrl1Addr'] = control_info.iloc[0]['GenericPointAddress']
                eterra_export.at[_, 'Ctrl1Name'] = control_info.iloc[0]['ControlId']
                eterra_export.at[_, 'Ctrl1Func'] = control_info.iloc[0]['CtrlFunc']
                # for Mk2a controls, the secure bit is in mdlparm1, the sync channel is in mdlparm2, for IEC controls, the IEC single/double is in mdlparm2
                if control_info.iloc[0]['Protocol'] == 'MK2A':
                    eterra_export.at[_, 'Ctrl1SecureBit'] = control_info.iloc[0]['Parm1']
                    eterra_export.at[_, 'Ctrl1SyncChannel'] = control_info.iloc[0]['Parm2']
                else:
                    eterra_export.at[_, 'Ctrl1IECSingleDouble'] = control_info.iloc[0]['Parm2']
                    # if the control is a CLOSE control, look for a matching SC1E point and set the Ctrl1SyncChannel to 1
                    if control_info.iloc[0]['ControlId'] == 'CLOSE':
                        sc1e_point = eterra_export[eterra_export['eTerraAlias'] == row['eTerraAlias'].replace('SWDD', 'SC1E')]
                        if sc1e_point.shape[0] > 0:
                            eterra_export.at[_, 'Ctrl1SyncChannel'] = 1

            if control_info.shape[0] > 1:
                # Add the second control info to the eterra export dataframe
                eterra_export.at[_, 'Ctrl2Addr'] = control_info.iloc[1]['GenericPointAddress']
                eterra_export.at[_, 'Ctrl2Name'] = control_info.iloc[1]['ControlId']
                eterra_export.at[_, 'Ctrl2Func'] = control_info.iloc[1]['CtrlFunc']
                if control_info.iloc[1]['Protocol'] == 'MK2A':
                    eterra_export.at[_, 'Ctrl2SecureBit'] = control_info.iloc[1]['Parm1']
                    eterra_export.at[_, 'Ctrl2SyncChannel'] = control_info.iloc[1]['Parm2']
                else:
                    eterra_export.at[_, 'Ctrl2IECSingleDouble'] = control_info.iloc[1]['Parm2']
                    # if the control is a CLOSE control, look for a matching SC1E point and set the Ctrl2SyncChannel to 1
                    if control_info.iloc[1]['ControlId'] == 'CLOSE':
                        sc1e_point = eterra_export[eterra_export['eTerraAlias'] == row['eTerraAlias'].replace('SWDD', 'SC1E')]
                        if sc1e_point.shape[0] > 0:
                            eterra_export.at[_, 'Ctrl2SyncChannel'] = 1

        if row['GenericType'] == 'A':
            # Get the control info from the eterra setpoint control dataframe
            control_info = get_setpoint_control_info(row)
            # Add the control info to the eterra setpoint control dataframe
            if control_info.shape[0] > 0:
                eterra_export.at[_, 'Ctrl1Addr'] = control_info.iloc[0]['GenericPointAddress']
                eterra_export.at[_, 'Ctrl1Name'] = 'SETPOINT'

    return eterra_export

def set_grid_incomer_flag_based_on_eterra_alias(eterra_export: pd.DataFrame) -> pd.DataFrame:
    """Set the GridIncomer flag based on the eTerra Alias.
    
    To consider a component being part of Grid Incomers the following conditions have to be met:

    the second part of the eTerra Alias to include one of the following text: 
    '033_CB', '025_CB', '011_CB','033_IS'
    the third part of the eTerra Alias to include one of the following text: 
    'GRID', 'SG', 'GS', 'T4', '2L5', '3L5'
    """

    # strip the eTerraAlias of any trailing spaces
    eterra_export['eTerraAlias'] = eterra_export['eTerraAlias'].str.strip()

    # create a mask for the points that are grid incomers
    grid_incomer_mask = eterra_export['eTerraAlias'].str.contains(r'^.*033_CB.*GRID.*$') | \
                        eterra_export['eTerraAlias'].str.contains(r'^.*025_CB.*GRID.*$') | \
                        eterra_export['eTerraAlias'].str.contains(r'^.*011_CB.*GRID.*$') | \
                        eterra_export['eTerraAlias'].str.contains(r'^.*033_IS.*GRID.*$') | \
                        eterra_export['eTerraAlias'].str.contains(r'^.*033_CB.*SG.*$') | \
                        eterra_export['eTerraAlias'].str.contains(r'^.*025_CB.*SG.*$') | \
                        eterra_export['eTerraAlias'].str.contains(r'^.*011_CB.*SG.*$') | \
                        eterra_export['eTerraAlias'].str.contains(r'^.*033_IS.*SG.*$') | \
                        eterra_export['eTerraAlias'].str.contains(r'^.*033_IS.*GS.*$') | \
                        eterra_export['eTerraAlias'].str.contains(r'^.*025_CB.*GS.*$') | \
                        eterra_export['eTerraAlias'].str.contains(r'^.*011_CB.*GS.*$') | \
                        eterra_export['eTerraAlias'].str.contains(r'^.*033_IS.*GS.*$') | \
                        eterra_export['eTerraAlias'].str.contains(r'^.*033_IS.*T4.*$') | \
                        eterra_export['eTerraAlias'].str.contains(r'^.*033_CB.*2L5.*$') | \
                        eterra_export['eTerraAlias'].str.contains(r'^.*033_CB.*3L5.*$')

    # Set the GridIncomer flag based on the mask
    eterra_export['GridIncomer'] = grid_incomer_mask

    return eterra_export
