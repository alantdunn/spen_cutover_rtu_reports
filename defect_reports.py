import pandas as pd


def check_report_results(merged_data: pd.DataFrame, report_name: str):
    """
    Check the results of a defect report.
    """
    print(f"Check: {report_name} has {merged_data[merged_data[report_name] == True].shape[0]} rows")


def defect_report1(merged_data: pd.DataFrame) -> pd.DataFrame:
    """
    Generate a defect report for missing analog components in PowerOn.
    
    Args:
        merged_data (pd.DataFrame): The merged data from the RTU report generator.
        
    Returns:
        pd.DataFrame: A dataframe with the defect report.
    """
    print("Generating defect report 1 ...")
    # check that required columns exist
    required_cols = ['GenericType', 'PowerOn Alias Exists']
    missing_cols = [col for col in required_cols if col not in merged_data.columns]
    if missing_cols:
        raise ValueError(f"Missing required columns: {', '.join(missing_cols)}")

    merged_data['Report1'] =    (merged_data['GenericType'] == 'A') & \
                                (~merged_data['PowerOn Alias Exists']) & \
                                (~merged_data['IGNORE_RTU']) & \
                                (~merged_data['IGNORE_POINT']) & \
                                (~merged_data['OLD_DATA'])

    check_report_results(merged_data, 'Report1')
    return merged_data

def defect_report2(merged_data: pd.DataFrame) -> pd.DataFrame:
    """
    Generate a defect report for missing digital inputs in PowerOn.
    
    Args:
        merged_data (pd.DataFrame): The merged data from the RTU report generator.

    Returns:
        pd.DataFrame: A dataframe with the defect report.
    """
    print("Generating defect report 2 ...")
    # check that required columns exist
    required_cols = ['GenericType', 'PowerOn Alias Exists']
    missing_cols = [col for col in required_cols if col not in merged_data.columns]
    if missing_cols:
        raise ValueError(f"Missing required columns: {', '.join(missing_cols)}")
    
    

    merged_data['Report2'] =    ((merged_data['GenericType'] == 'SD') | \
                                (merged_data['GenericType'] == 'DD')) & \
                                (~merged_data['PowerOn Alias Exists']) & \
                                (~merged_data['IGNORE_RTU']) & \
                                (~merged_data['IGNORE_POINT']) & \
                                (~merged_data['OLD_DATA'])
    check_report_results(merged_data, 'Report2')
    return merged_data



def defect_report3(merged_data: pd.DataFrame) -> pd.DataFrame:
    """
    Generate a defect report for missing controllable points in PowerOn.
    
    Args:
        merged_data (pd.DataFrame): The merged data from the RTU report generator.

    Returns:
        pd.DataFrame: A dataframe with the defect report.
    """
    print("Generating defect report 3 ...")
    # check that required columns exist
    required_cols = ['GenericType', 'PowerOn Alias Exists']
    missing_cols = [col for col in required_cols if col not in merged_data.columns]
    if missing_cols:
        raise ValueError(f"Missing required columns: {', '.join(missing_cols)}")

    merged_data['Report3'] =    ((merged_data['Ctrl1Addr'].notna() | merged_data['Ctrl2Addr'].notna()) & \
                                (merged_data['Controllable'] == '1') & \
                                (merged_data['RTUId'] != '(€€€€€€€€:)') & \
                                (~merged_data['PowerOn Alias Exists']) & \
                                (~merged_data['IGNORE_RTU']) & \
                                (~merged_data['IGNORE_POINT']) & \
                                (~merged_data['OLD_DATA']))
    check_report_results(merged_data, 'Report3')
    return merged_data



def defect_report4(merged_data: pd.DataFrame) -> pd.DataFrame:
    """
    Generate a defect report for missing telecontrol actions in PowerOn.
    
    Args:
        merged_data (pd.DataFrame): The merged data from the RTU report generator.

    Returns:
        pd.DataFrame: A dataframe with the defect report.
    """
    print("Generating defect report 4 ...")
    # check that required columns exist
    required_cols = ['GenericType', 'Controllable', 'PowerOn Alias Exists']
    missing_cols = [col for col in required_cols if col not in merged_data.columns]
    if missing_cols:
        raise ValueError(f"Missing required columns: {', '.join(missing_cols)}")

    merged_data['Report4'] =    (merged_data['Controllable'] == '1') & \
                                ((merged_data['Ctrl1Addr'].notna()  & ~merged_data['Ctrl1TelecontrolAction'].isna()) | \
                                (merged_data['Ctrl2Addr'].notna()  & ~merged_data['Ctrl2TelecontrolAction'].isna())) & \
                                (merged_data['DeviceType'] != 'RTU') & \
                                (~merged_data['IGNORE_RTU']) & \
                                (~merged_data['IGNORE_POINT']) & \
                                (~merged_data['OLD_DATA'])
    check_report_results(merged_data, 'Report4')
    return merged_data


def defect_report5(merged_data: pd.DataFrame) -> pd.DataFrame:
    """
    Generate a defect report for Components Missing Alarm Reference.
    
    Args:
        merged_data (pd.DataFrame): The merged data from the RTU report generator.

    Returns:
        pd.DataFrame: A dataframe with the defect report.
    """
    print("Generating defect report 5 ...")
    # criteria is 
    # not old or ignore data
    # GenericTye is SD or DD
    # CompAlarmPOAlarmRef is 0 or null
    # any <num> of 0-3 for Alarm<num>_eTerraMessage is not null and  Alarm<num>_POMessage is null

    # debug some data
    # debugdata = merged_data[
    #                             ((merged_data['GenericType'] == 'SD') | \
    #                             (merged_data['GenericType'] == 'DD')) \
    #                             & \
    #                             (merged_data['PowerOn Alias Exists']) \
    #                             & \
    #                             (merged_data['PowerOn Alias Linked to SCADA'] == 2) \
    #                             & \
    #                             ((merged_data['CompAlarmPOAlarmRef'] == 0) | \
    #                             (merged_data['CompAlarmPOAlarmRef'].isna())) \
    #                             & \
    #                             ((merged_data['Alarm0_MessageMatch'] == 0) | \
    #                             (merged_data['Alarm1_MessageMatch'] == 0) | \
    #                             (merged_data['Alarm2_MessageMatch'] == 0) | \
    #                             (merged_data['Alarm3_MessageMatch'] == 0)) \
    #                             & \
    #                             (merged_data['ConfigHealth'] == 'GOOD') \
    #                             & \
    #                             (~merged_data['IGNORE_RTU']) & \
    #                             (~merged_data['IGNORE_POINT']) & \
    #                             (~merged_data['OLD_DATA'])
    # ]
    # print(debugdata[['eTerraAlias', 'CompAlarmPOAlarmRef','ConfigHealth', 'PowerOn Alias Linked to SCADA']].head(10))
    # exit(0)

    merged_data['Report5'] =   ((merged_data['GenericType'] == 'SD') | \
                                (merged_data['GenericType'] == 'DD')) \
                                & \
                                (merged_data['PowerOn Alias Exists']) \
                                & \
                                (merged_data['PowerOn Alias Linked to SCADA'] == 2) \
                                & \
                                ((merged_data['CompAlarmPOAlarmRef'] == 0) | \
                                (merged_data['CompAlarmPOAlarmRef'].isna())) \
                                & \
                                ((merged_data['Alarm0_MessageMatch'] == 0) | \
                                (merged_data['Alarm1_MessageMatch'] == 0) | \
                                (merged_data['Alarm2_MessageMatch'] == 0) | \
                                (merged_data['Alarm3_MessageMatch'] == 0)) \
                                & \
                                (merged_data['ConfigHealth'] == 'GOOD') \
                                & \
                                (~merged_data['IGNORE_RTU']) & \
                                (~merged_data['IGNORE_POINT']) & \
                                (~merged_data['OLD_DATA'])
    
    check_report_results(merged_data, 'Report5')

    return merged_data


def defect_report6(merged_data: pd.DataFrame) -> pd.DataFrame:
    """
    Generate a defect report for components that missed the match but have tested ok.
    
    criteria is 
    - Ctrl1Addr exists and Ctrl1MatchStatus is notinPO AND CtrlTest1Result is OK , or
    - Ctrl2Addr exists and Ctrl2MatchStatus is notinPO AND CtrlTest2Result is OK
    - not old or ignored data

    Args:
        merged_data (pd.DataFrame): The merged data from the RTU report generator.

    Returns:
        pd.DataFrame: A dataframe with the defect report.
    """
    print("Generating defect report 6 ...")

    merged_data['Report6'] =   (~merged_data['PowerOn Alias Exists']) & \
                                (merged_data['Ctrl1Addr'].notna() & \
                                (merged_data['Ctrl1MatchStatus'] == 'notinPO') & \
                                (merged_data['Ctrl1TestResult'] == 'OK')) | \
                                (merged_data['Ctrl2Addr'].notna() & \
                                (merged_data['Ctrl2MatchStatus'] == 'notinPO') & \
                                (merged_data['Ctrl2TestResult'] == 'OK')) & \
                                (~merged_data['IGNORE_RTU']) & \
                                (~merged_data['IGNORE_POINT']) & \
                                (~merged_data['OLD_DATA'])
    check_report_results(merged_data, 'Report6')
    return merged_data


def defect_report7(merged_data: pd.DataFrame) -> pd.DataFrame:
    """
    Generate a defect report for Controls that are not linked correctly.

    Criteria
    - Points are not old or ignored
    - GenericType is SD or DD
    - PowerOn Alias Exists
    - Either
        - Ctrl1Name is not empty AND CtrlConfigHealth is null , or
        - Ctrl2Name is not empty AND CtrlConfigHealth is null
    - DeviceType is not RTU

    Args:
        merged_data (pd.DataFrame): The merged data from the RTU report generator.
    """
    print("Generating defect report 7 ...")

    # debug some data - get the first 10 rows where Ctrl1Name is not null and Ctrl1ConfigHealth is null
    # debugdata = merged_data[
    #     (merged_data['Ctrl1Name'] != '') & \
    #     (merged_data['Ctrl1ConfigHealth'].isnull()) & \
    #     (~merged_data['OLD_DATA']) & \
    #     (~merged_data['IGNORE_POINT']) & \
    #     (~merged_data['IGNORE_RTU'])
    #     ]
    # debugdata = debugdata[['eTerraAlias', 'Ctrl1Name',  'Ctrl1ConfigHealth', 'Ctrl1Addr']].head(10)
    # print(debugdata)
    # # print the column types
    # print(debugdata.dtypes)
    # # add some columsn to describe the data
    # debugdata['Ctrl1NameNotNa'] = debugdata['Ctrl1Name'].notna()
    # debugdata['Ctrl1NameNotEmpty'] = debugdata['Ctrl1Name'] != ''
    # debugdata['Ctrl1ConfigHealthNotNa'] = debugdata['Ctrl1ConfigHealth'].notna()
    # debugdata['Ctrl1ConfigHealthIsNull'] = debugdata['Ctrl1ConfigHealth'].isnull()

    # print(debugdata)
    # exit (0)
    merged_data['Report7'] =   (~merged_data['IGNORE_RTU']) & \
                                (~merged_data['IGNORE_POINT']) & \
                                (~merged_data['OLD_DATA']) & \
                                ((merged_data['GenericType'] == 'SD') | \
                                (merged_data['GenericType'] == 'DD')) \
                                & \
                                (merged_data['PowerOn Alias Exists']) \
                                & \
                                (((merged_data['Ctrl1Name'] != '') & 
                                (merged_data['Ctrl1ConfigHealth'].isnull())) \
                                    | \
                                ((merged_data['Ctrl2Name'] != '') & \
                                (merged_data['Ctrl2ConfigHealth'].isnull()))) \
                                & \
                                (merged_data['DeviceType'] != 'RTU')
    check_report_results(merged_data, 'Report7')
    return merged_data

def defect_report8(merged_data: pd.DataFrame) -> pd.DataFrame:
    """
    Generate a defect report for points that are marked as controllable but have no Ctrl1Name or Ctrl2Name.
    
    Args:
        merged_data (pd.DataFrame): The merged data from the RTU report generator.
    """
    print("Generating defect report 8 ...")
    merged_data['Report8'] =   (merged_data['Controllable'] == 1) & \
                                ((merged_data['Ctrl1Name'].isnull() | merged_data['Ctrl2Name'].isnull())) & \
                                (~merged_data['IGNORE_RTU']) & \
                                (~merged_data['IGNORE_POINT']) & \
                                (~merged_data['OLD_DATA'])
    check_report_results(merged_data, 'Report8')
    return merged_data

def defect_report9(merged_data: pd.DataFrame) -> pd.DataFrame:
    """
    Generate a defect report for tbd.
    
    Args:
        merged_data (pd.DataFrame): The merged data from the RTU report generator.
    """
    print("Generating defect report 9 ...")

    merged_data['Report9'] = False
    check_report_results(merged_data, 'Report9')
    return merged_data

def defect_report10(merged_data: pd.DataFrame) -> pd.DataFrame:
    """
    Generate a defect report for tbd.
    
    Args:
        merged_data (pd.DataFrame): The merged data from the RTU report generator.
    """ 
    print("Generating defect report 10 ...")
    merged_data['Report10'] = False
    check_report_results(merged_data, 'Report10')
    return merged_data