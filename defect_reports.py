import pandas as pd


def check_report_results(merged_data: pd.DataFrame, report_name: str):
    """
    Check the results of a defect report.
    """
    print(f"Check: {report_name} has {merged_data[merged_data[report_name] == True].shape[0]} rows")


def debug_report_data(merged_data: pd.DataFrame, report_criteria: dict, report_name: str):
    """Debug helper to print details when a report returns 0 rows"""
    print(f"\nDebugging {report_name}...")
    
    # Start with full dataset
    debug_data = merged_data.copy()
    rows = len(debug_data)
    
    # Apply each criteria and print remaining rows
    for criteria_name, criteria in report_criteria['criteria'].items():
        debug_data = debug_data[criteria]
        remaining = len(debug_data)
        filtered = rows - remaining
        print(f"After {criteria_name}: {remaining} rows (filtered {filtered})")
        rows = remaining

    if rows == 0:
        print("No rows matched all criteria\n")
    else:
        print(f"{rows} rows matched all criteria\n")
        
    return debug_data

def generate_defect_report(merged_data: pd.DataFrame, report_config: dict, debug: bool = False) -> pd.DataFrame:
    """
    Generic defect report generator that applies configured criteria
    
    Args:
        merged_data: The merged dataframe to analyze
        report_config: Dictionary containing report configuration
        debug: Whether to print debug info when 0 rows match
        
    Returns:
        DataFrame with report column added
    """
    report_name = report_config['name']
    print(f"Generating {report_name}...")

    # Validate required columns exist
    if 'required_columns' in report_config:
        missing_cols = [col for col in report_config['required_columns'] 
                       if col not in merged_data.columns]
        if missing_cols:
            raise ValueError(f"Missing required columns: {', '.join(missing_cols)}")

    # Combine all criteria with & operator
    final_criteria = None
    for criteria in report_config['criteria'].values():
        if final_criteria is None:
            final_criteria = criteria
        else:
            final_criteria = final_criteria & criteria

    # Add report column
    merged_data[report_name] = final_criteria

    # Check results
    matching_rows = merged_data[merged_data[report_name] == True].shape[0]
    print(f"Check: {report_name} has {matching_rows} rows")

    # Debug if requested and no matches found
    if debug and matching_rows == 0:
        debug_report_data(merged_data, report_config, report_name)

    return merged_data

# Report configurations
REPORT_CONFIGS = {
    'Report1': {
        'name': 'Report1',
        'description': 'Missing analog components in PowerOn',
        'required_columns': ['GenericType', 'PowerOn Alias Exists'],
        'criteria': {
            'type': merged_data['GenericType'] == 'A',
            'missing_alias': ~merged_data['PowerOn Alias Exists'],
            'not_ignored': ~merged_data['IGNORE_RTU'] & ~merged_data['IGNORE_POINT'] & ~merged_data['OLD_DATA']
        }
    },
    'Report2': {
        'name': 'Report2', 
        'description': 'Missing digital inputs in PowerOn',
        'required_columns': ['GenericType', 'PowerOn Alias Exists'],
        'criteria': {
            'type': (merged_data['GenericType'] == 'SD') | (merged_data['GenericType'] == 'DD'),
            'missing_alias': ~merged_data['PowerOn Alias Exists'],
            'not_ignored': ~merged_data['IGNORE_RTU'] & ~merged_data['IGNORE_POINT'] & ~merged_data['OLD_DATA']
        }
    },
    'Report3': {
        'name': 'Report3',
        'description': 'Missing controllable points in PowerOn',
        'required_columns': ['GenericType', 'PowerOn Alias Exists'],
        'criteria': {
            'has_controls': (merged_data['Ctrl1Addr'].notna() | merged_data['Ctrl2Addr'].notna()),
            'controllable': merged_data['Controllable'] == '1',
            'valid_rtu': merged_data['RTUId'] != '(€€€€€€€€:)',
            'missing_alias': ~merged_data['PowerOn Alias Exists'],
            'not_ignored': ~merged_data['IGNORE_RTU'] & ~merged_data['IGNORE_POINT'] & ~merged_data['OLD_DATA']
        }
    },
    # Add remaining report configs following same pattern...
}

# Generate all reports
def generate_all_defect_reports(merged_data: pd.DataFrame, debug: bool = False) -> pd.DataFrame:
    """Generate all configured defect reports"""
    for report_config in REPORT_CONFIGS.values():
        merged_data = generate_defect_report(merged_data, report_config, debug)
    return merged_data