import pandas as pd

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