import pandas as pd
from data_import.utils import (
    derive_generic_address_for_poweron_export,
    split_ioa,
    compute_offset
)

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
    # | interpretation          | POInterpretation
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
    # | user_tag                | UserTag
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
        'interpretation': 'POInterpretation',
        'shift': 'Shift',
        'siref1': 'ScanInputRef',
        'size': 'Size',
        'symbol_menu': 'Menu',
        'symbol_name': 'Symbol',
        'telecontrol_action': 'TC Action',
        'user_tag': 'UserTag'
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
    df[['IOA1', 'IOA2']] = df.apply(split_ioa, axis=1, result_type='expand')
    df['Offset'] = df.apply(compute_offset, axis=1)

    # Derive the GenericPointAddress from the RTUId, CASDU, IOA1, IOA2, and GenericType
    df[['GenericPointAddress']] = df.apply(derive_generic_address_for_poweron_export, axis=1)

    # Change a few columns to unique names before we return this dataframe
    df.rename(columns={
        'Protocol': 'PO_Protocol',
        'Card': 'PO_Card',
        'Word': 'PO_Word',
        'IOA1': 'PO_IOA1',
        'IOA2': 'PO_IOA2',
        'Offset': 'PO_Offset',
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
        'PO_IOA1',
        'PO_IOA2',
        'PO_Offset',
        'POAlias',
        'POName',
        'ConfigInfo',
        'ConfigHealth',
        'PODescription',
        'POType',
        'ScanInputRow',
        'Shift',
        'ScanInputRef',
        'UserTag',
        'Size',
        'POInterpretation',
        'Menu',
        'Symbol',
        'TC Action',
        'PO_GenericType',
        'GenericPointAddress',
        'PO_eTerraAlias'
    ]

    # convert PO_Card to an int then a string
    try:
        df['PO_Card'] = df['PO_Card'].astype(int).astype(str)
    except:
        pass

    # convert PO_Word to an int then a string
    try:
        df['PO_Word'] = df['PO_Word'].astype(int).astype(str)
    except:
        pass

    # convert Shift to an int then a string
    try:
        df['Shift'] = df['Shift'].astype(int).astype(str)
    except:
        pass

    # convert Size to an int then a string
    try:
        df['Size'] = df['Size'].astype(int).astype(str)
    except:
        pass

    df = df[columns_to_keep]

    # We have had some issues with the eTerra source data that can result in duplciate rows in the all_rtus.csv file. 

    # Start by removing the CUMW RTU - it has duplciates in dataload 14
    df = df[df['PO_RTU'] != 'CUMW_RTU']

    # Look for any duplcicate addresses based on PO_RTU, PO_Card, PO_Word where the PO_Protocol is IEC60870-101
    # if we find any, first print out the duplciates, then order them by POType and keep the last one (thinking here is that we wnat to drop the analog which would be the first one)

    # Get indices of non_control rows with IEC60870-101 protocol
    df_iec = df[(df['PO_Protocol'] == 'IEC60870-101') & (df['PO_GenericType'] != 'C')]
    dup_indices = df_iec[df_iec.duplicated(subset=['PO_RTU', 'PO_Card', 'PO_Word'], keep=False)].index
    
    if len(dup_indices) > 0:
        print("🔍 Found duplicates in the all_rtus.csv file")
        print(df.loc[dup_indices])

        # Sort duplicates by POType and get indices to drop
        print("🔍 Sorting duplicates by POType")
        df_sorted = df.loc[dup_indices].sort_values(by=['POType'])
        indices_to_drop = df_sorted.duplicated(subset=['PO_RTU', 'PO_Card', 'PO_Word'], keep='last')
        indices_to_drop = df_sorted[indices_to_drop].index
        
        # Drop duplicates from main dataframe
        print(f"🔍 Dropping duplicates from main dataframe with indices: {indices_to_drop}")
        df = df.drop(indices_to_drop)

    # Check for duplicate GenericPointAddress values
    dup_gpa = df[df.duplicated(subset=['GenericPointAddress'], keep=False)]
    if len(dup_gpa) > 0:
        print("\n⚠️  WARNING: Duplicate GenericPointAddress values found:")
        print(dup_gpa[['GenericPointAddress', 'PO_RTU', 'PO_Card', 'PO_Word', 'POType']])
        
        # Group by GenericPointAddress and show the duplicated values
        duplicates = dup_gpa.groupby('GenericPointAddress')['GenericPointAddress'].count()
        print("\nDuplicate GenericPointAddress values:")
        for addr, count in duplicates.items():
            print(f"'{addr}' appears {count} times")
        
        print("\nThis will cause issues when using GenericPointAddress as an index.")
        print("Please check the source data and resolve duplicates.")
        print("Continuing with processing but merge operations may fail...\n\n")

        # Ask the user if they want to continue
        user_input = input("Do you want to continue anyway? (y/n): (recommend you do not, as this will likely crash when using GenericPointAddress as an index)")
        if user_input.lower() != 'y':
            print("Exiting...")
            exit()


    return df
