import pandas as pd

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
