import pandas as pd

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
    # | DCB                     | IsDCB
    # | 314                     | Is314
    # | SC1E                    | IsSC1E
    # | SC2E                    | IsSC2E
    # | TemplateAlias           | 
    # | TemplateName            | 
    # | TemplateType            | 
    # | State Index             | 
    

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
        'AlarmZoneMatch': 'CompAlarmAlarmZoneMatch',
        'TemplateAlias': 'CompAlarmTemplateAlias',
        'TemplateName': 'CompAlarmTemplateName',
        'TemplateType': 'CompAlarmTemplateType',
        'StateIndex': 'CompAlarmStateIndex',
        'DCB': 'IsDCB',
        '314': 'Is314',
        'SC1E': 'IsSC1E',
        'SC2E': 'IsSC2E'
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
        'CompAlarmAlarmZoneMatch',
        'CompAlarmTemplateAlias',
        'CompAlarmTemplateName',
        'CompAlarmTemplateType',
        'CompAlarmStateIndex',
        'IsDCB',
        'Is314',
        'IsSC1E',
        'IsSC2E'
    ]

    df = df[columns_to_keep]
    return df
