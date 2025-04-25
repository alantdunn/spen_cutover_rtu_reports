import pandas as pd

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
    
    # Convert PowerOn Alias Exists to boolean if not already
    merged_data = merged_data.reset_index(drop=True)

    # Ensure boolean conversion works correctly
    merged_data['PowerOn Alias Exists'] = merged_data['PowerOn Alias Exists'].astype(int).astype(bool)

    merged_data['Report1'] = (merged_data['GenericType'] == 'A') & (~merged_data['PowerOn Alias Exists'])


    return merged_data


def defect_report2(merged_data: pd.DataFrame) -> pd.DataFrame:
    """
    Generate a defect report for missing controllable points in PowerOn.
    
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

    # TODO: Come back here - need to look at ctrl1 and ctrl2 for unklinked status


    return merged_data

def defect_report3(merged_data: pd.DataFrame) -> pd.DataFrame:
    """
    Generate a defect report for missing digital inputs in PowerOn.
    
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
    
    # Check for any duplicate named columns in merged_data and print out the duplicates
    duplicate_cols = merged_data.columns.duplicated()
    if duplicate_cols.any():
        print(f"Duplicate column names found: {merged_data.columns[duplicate_cols]}")
    
    # Convert PowerOn Alias Exists to boolean if not already
    merged_data = merged_data.reset_index(drop=True)

    # Ensure boolean conversion works correctly
    merged_data['PowerOn Alias Exists'] = merged_data['PowerOn Alias Exists'].astype(int).astype(bool)

    merged_data['Report3'] = (merged_data['GenericType'] == 'SD') & (~merged_data['PowerOn Alias Exists'])
    return merged_data


def defect_report4(merged_data: pd.DataFrame) -> pd.DataFrame:
    """
    Generate a defect report for missing telecontrol actions in PowerOn.
    
    Args:
        merged_data (pd.DataFrame): The merged data from the RTU report generator.

    Returns:
        pd.DataFrame: A dataframe with the defect report.
    """
    merged_data['Report4'] = merged_data['GenericType'] == 'C'  & ~merged_data['PowerOn Alias Exists']
    return merged_data


def defect_report5(merged_data: pd.DataFrame) -> pd.DataFrame:
    """
    Generate a defect report for missing items in PowerOn that are in eTerra.
    
    Args:
        merged_data (pd.DataFrame): The merged data from the RTU report generator.

    Returns:
        pd.DataFrame: A dataframe with the defect report.
    """

    merged_data['Report5'] = merged_data['GenericType'] == 'A'  & ~merged_data['PowerOn Alias Exists']
    return merged_data


def defect_report6(merged_data: pd.DataFrame) -> pd.DataFrame:
    """
    Generate a defect report for components missing alarm references in PowerOn.
    
    Args:
        merged_data (pd.DataFrame): The merged data from the RTU report generator.

    Returns:
        pd.DataFrame: A dataframe with the defect report.
    """
    merged_data['Report6'] = merged_data['GenericType'] == 'A'  & ~merged_data['PowerOn Alias Exists']
    return merged_data


