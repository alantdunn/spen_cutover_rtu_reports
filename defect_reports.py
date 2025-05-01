import pandas as pd


# Dictionary defining all report configurations
REPORT_CONFIGS = {
    'Report1': {
        'name': 'Missing Analog Components',
        'debug': False,
        'required_cols': ['GenericType', 'PowerOn Alias Exists'],
        'criteria': [
            ('GenericType', '==', 'A'),
            ('PowerOn Alias Exists', '==', False),
            ('IGNORE_RTU', '==', False),
            ('IGNORE_POINT', '==', False), 
            ('OLD_DATA', '==', False)
        ],
        'combine_with': 'and'
    },
    'Report2': {
        'name': 'Missing Digital Components',
        'debug': False,
        'required_cols': ['GenericType', 'PowerOn Alias Exists'],
        'criteria': [
            ('GenericType', 'in', ['SD', 'DD']),
            ('PowerOn Alias Exists', '==', False),
            ('IGNORE_RTU', '==', False),
            ('IGNORE_POINT', '==', False),
            ('OLD_DATA', '==', False)
        ],
        'combine_with': 'and'
    },
    'Report3': {
        'name': 'Missing Controllable Components',
        'debug': False,
        'required_cols': ['GenericType', 'PowerOn Alias Exists'],
        'criteria': [
            ('Ctrl1Addr,Ctrl2Addr', 'any_notna'),
            ('Controllable', '==', '1'),
            ('RTUId', '!=', '(€€€€€€€€:)'),
            ('PowerOn Alias Exists', '==', False),
            ('IGNORE_RTU', '==', False),
            ('IGNORE_POINT', '==', False),
            ('OLD_DATA', '==', False)
        ],
        'combine_with': 'and'
    },
    'Report4': {
        'name': 'Components Missing Telecontrol Actions',
        'debug': False,
        'required_cols': ['GenericType', 'Controllable', 'PowerOn Alias Exists'],
        'criteria_groups': [
            {
                'criteria': [
                    ('Controllable', '==', '1'),
                    ('DeviceType', '!=', 'RTU'),
                    ('IGNORE_RTU', '==', False),
                    ('IGNORE_POINT', '==', False),
                    ('OLD_DATA', '==', False)
                ],
                'combine_with': 'and'
            },
            {
                'criteria': [
                    ('Ctrl1Addr,Ctrl1TelecontrolAction', 'notna_pair'),
                    ('Ctrl2Addr,Ctrl2TelecontrolAction', 'notna_pair')
                ],
                'combine_with': 'or'
            }
        ],
        'combine_groups_with': 'and'
    },
    'Report5': {
        'name': 'Components Missing Alarm Reference',
        'debug': False,
        'required_cols': ['GenericType', 'PowerOn Alias Exists', 'PowerOn Alias Linked to SCADA'],
        'criteria': [
            ('GenericType', 'in', ['SD', 'DD']),
            ('DeviceType', '!=', 'RTU'),
            ('PowerOn Alias Exists', '==', True),
            ('PowerOn Alias Linked to SCADA', '==', 2),
            ('CompAlarmPOAlarmRef', 'isnull_or_zero'),
            ('Alarm0_MessageMatch,Alarm1_MessageMatch,Alarm2_MessageMatch,Alarm3_MessageMatch','no_zeros'),
            ('ConfigHealth', '==', 'GOOD'),
            ('IGNORE_RTU', '==', False),
            ('IGNORE_POINT', '==', False),
            ('OLD_DATA', '==', False)
        ],
        'combine_with': 'and'
    },
    'Report6': {
        'name': 'Controls not in PO but tested ok',
        'debug': False,
        'required_cols': ['PowerOn Alias Exists'],
        'criteria_groups': [
            {
                'criteria': [
                    ('PowerOn Alias Exists', '==', False),
                    ('IGNORE_RTU', '==', False), 
                    ('IGNORE_POINT', '==', False),
                    ('OLD_DATA', '==', False)
                ],
                'combine_with': 'and'
            },
            {
                'criteria_groups': [
                    {
                        'criteria': [
                            ('Ctrl1Addr', 'notna'),
                            ('Ctrl1MatchStatus', '==', 'notinPO'),
                            ('Ctrl1TestResult', '==', 'OK')
                        ],
                        'combine_with': 'and'
                    },
                    {
                        'criteria': [
                            ('Ctrl2Addr', 'notna'),
                            ('Ctrl2MatchStatus', '==', 'notinPO'),
                            ('Ctrl2TestResult', '==', 'OK')
                        ],
                        'combine_with': 'and'
                    }
                ],
                'combine_with': 'or'
            }
        ],
        'combine_groups_with': 'and'
    },
    'Report7': {
        'name': 'Controls Not Linked',
        'debug': False,
        'required_cols': ['GenericType', 'PowerOn Alias Exists'],
        'criteria_groups': [
            {
                'criteria': [
                    ('GenericType', 'in', ['SD', 'DD']),
                    ('PowerOn Alias Exists', '==', True),
                    ('DeviceType', '!=', 'RTU'),
                    ('IGNORE_RTU', '==', False),
                    ('IGNORE_POINT', '==', False),
                    ('OLD_DATA', '==', False)
                ],
                'combine_with': 'and'
            },
            {
            'criteria_groups': [
                    {
                        'criteria': [
                            ('Ctrl1Name','!=', ''),
                            ('Ctrl1ConfigHealth','isna_or_blank'),
                        ],
                        'combine_with': 'and'
                    },
                    {
                        'criteria': [
                            ('Ctrl2Name','!=', ''),
                            ('Ctrl2ConfigHealth','isna_or_blank')
                        ],
                        'combine_with': 'and'
                    }
                ],
                'combine_with': 'or'
            }
        ],
        'combine_with': 'and'
    },
    'Report8': {
        'name': 'Ctrl-able eTerra Points with no Controls',
        'debug': False,
        'required_cols': ['Controllable'],
        'criteria': [
            ('Controllable', '==', '1'),
            ('Ctrl1Name', 'isna_or_blank'),
            ('Ctrl2Name', 'isna_or_blank'),
            ('IGNORE_RTU', '==', False),
            ('IGNORE_POINT', '==', False),
            ('OLD_DATA', '==', False)
        ],
        'combine_with': 'and'
    },
    'Report9': {
        'name': 'Alarm Mismatch Manual Actions',
        'debug': False,
        'required_cols': ['AlarmMismatchComment'],
        'criteria': [
            ('AlarmMismatchComment', '!=', ''),
            ('IGNORE_RTU', '==', False),
            ('IGNORE_POINT', '==', False),
            ('OLD_DATA', '==', False)
            ],
        'combine_with': 'and'
    },
    'Report10': {
        'name': 'TBD10', 
        'debug': False,
        'required_cols': [],
        'criteria': [('_', 'always_false', None)],
        'combine_with': 'and'
    }
}

def debug_report_criteria(df: pd.DataFrame, report_name: str, report_config: dict, debug: bool = False) -> None:
    """Debug helper to print intermediate results when criteria returns 0 rows"""
    if not debug:
        return
        
    print(f"\nDebugging {report_name}...")
    result = df.copy()
    result[report_name] = True
    
    if 'criteria_groups' in report_config:
        for group in report_config['criteria_groups']:
            for cols, op, val in group['criteria']:
                prev_count = len(result[result[report_name] == True])
                result = generate_defect_report(result, {'name': report_name, 'criteria': [(cols, op, val)], 'combine_with': group['combine_with']})
                print(f"After {cols} {op} {val}: {prev_count} -> {len(result[result[report_name] == True])} rows")
                if len(result[result[report_name] == True]) == 0:
                    break
    else:
        for cols, op, val in report_config['criteria']:
            prev_count = len(result[result[report_name] == True])
            result = generate_defect_report(result, {'name': report_name, 'criteria': [(cols, op, val)], 'combine_with': report_config.get('combine_with', 'and')})
            print(f"After {cols} {op} {val}: {prev_count} -> {len(result[result[report_name] == True])} rows")
            if len(result[result[report_name] == True]) == 0:
                break

def generate_defect_report(df: pd.DataFrame, report_name: str, report_config: dict, debug: bool = False) -> pd.DataFrame:
    """
    Generic report generator that applies criteria from config.
    Returns the full input DataFrame with a new boolean column named report_name.
    The column will be True for rows matching all criteria, False otherwise.
    """
    # Start with a copy of the input DataFrame to avoid modifying the original
    result = df
    
    # For AND operations, we start with True and filter down
    # For OR operations, we start with False and build up
    # This is because:
    # - AND: A AND B AND C means start True, then remove non-matches 
    # - OR: A OR B OR C means start False, then add matches
    initial_value = True if report_config.get('combine_with', 'and') == 'and' else False

    if report_config['debug']:
        debug = True
        print(f"\nDebugging {report_name}...")
        print(f"Initial value: {initial_value}")
        print(f"Report config: {report_config}")
        print

    
    if 'criteria_groups' in report_config:
        # Initialize report column
        result[report_name] = initial_value

        if debug:   
            print("\n================================================")
            print(f"{report_name}: Initial result {initial_value}: {result[report_name].sum()} rows")
            print("================================================\n")
        
        for group_idx, group in enumerate(report_config['criteria_groups']):
            if debug:
                print(f"\nGroup {group_idx + 1} (combine with {group.get('combine_with', 'and')})")
                
            # Same logic for group initialization
            group_result = pd.Series(True if group.get('combine_with') == 'and' else False, index=result.index)
            
            if 'criteria_groups' in group:
                # Handle nested groups
                for subgroup_idx, subgroup in enumerate(group['criteria_groups']):
                    if debug:
                        print(f"  Subgroup {subgroup_idx + 1} (combine with {subgroup.get('combine_with', 'and')})")
                        
                    subgroup_result = pd.Series(True if subgroup.get('combine_with') == 'and' else False, index=result.index)
                    
                    for criteria_idx, criteria in enumerate(subgroup['criteria']):
                        if len(criteria) == 2:
                            cols, op = criteria
                            val = None
                        else:
                            cols, op, val = criteria

                        if debug:
                            print(f"    Criteria {criteria_idx + 1}: {cols} {op} {val}")
                            
                        criteria_result = evaluate_criteria(result, cols, op, val)
                        prev_count = subgroup_result.sum()
                        
                        if subgroup.get('combine_with') == 'or':
                            subgroup_result |= criteria_result
                        else:  # default to 'and'
                            subgroup_result &= criteria_result
                            
                        if debug:
                            print(f"      Rows matching: {criteria_result.sum()}")
                            print(f"      After combining: {subgroup_result.sum()} ({prev_count} -> {subgroup_result.sum()})")
                            
                    prev_count = group_result.sum()
                    if group.get('combine_with') == 'or':
                        group_result |= subgroup_result
                    else:  # default to 'and'
                        group_result &= subgroup_result
                        
                    if debug:
                        print(f"    Subgroup result: {subgroup_result.sum()} rows")
                        print(f"    After combining with group: {group_result.sum()} ({prev_count} -> {group_result.sum()})")
            else:
                # Handle regular criteria
                for criteria_idx, criteria in enumerate(group['criteria']):
                    if len(criteria) == 2:
                        cols, op = criteria
                        val = None
                    else:
                        cols, op, val = criteria

                    if debug:
                        print(f"  Criteria {criteria_idx + 1}: {cols} {op} {val}")
                        
                    criteria_result = evaluate_criteria(result, cols, op, val)
                    prev_count = group_result.sum()
                    
                    if group.get('combine_with') == 'or':
                        group_result |= criteria_result
                    else:  # default to 'and'
                        group_result &= criteria_result
                        
                    if debug:
                        print(f"    Rows matching: {criteria_result.sum()}")
                        print(f"    After combining: {group_result.sum()} ({prev_count} -> {group_result.sum()})")
            
            prev_count = result[report_name].sum()
            if report_config.get('combine_groups_with') == 'or':
                result[report_name] |= group_result
            else:  # default to 'and'
                result[report_name] &= group_result
                
            if debug:
                print(f"  Group result: {group_result.sum()} rows")
                print(f"  After combining with final result: {result[report_name].sum()} ({prev_count} -> {result[report_name].sum()})")
    else:
        # Initialize report column if it doesn't exist
        if report_name not in result.columns:
            result[report_name] = initial_value

        for criteria_idx, criteria in enumerate(report_config['criteria']):
            if len(criteria) == 2:
                cols, op = criteria
                val = None
            else:
                cols, op, val = criteria

            if debug:
                print(f"\nCriteria {criteria_idx + 1}: {cols} {op} {val}")
                
            criteria_result = evaluate_criteria(result, cols, op, val)
            prev_count = result[report_name].sum()
            
            if report_config.get('combine_with') == 'or':
                result[report_name] |= criteria_result
            else:  # default to 'and'
                result[report_name] &= criteria_result
                
            if debug:
                print(f"  Rows matching: {criteria_result.sum()}")
                print(f"  After combining: {result[report_name].sum()} ({prev_count} -> {result[report_name].sum()})")

    if debug:
        if len(result[result[report_name] == True]) == 0:
            print("\nNo rows left after applying all criteria")
        print(f"\nFinal result: {result[report_name].sum()} rows")

    return result

def evaluate_criteria(df: pd.DataFrame, cols: str, op: str, val: any) -> pd.Series:
    """Evaluate a single criteria and return the result"""
    if op == '==':
        return (df[cols] == val)
    elif op == '!=':
        return (df[cols] != val)
    elif op == 'in':
        return df[cols].isin(val)
    elif op == 'notna_pair':
        col_pairs = cols.split('|')
        result = pd.Series(True, index=df.index)
        for pair in col_pairs:
            pair_cols = pair.split(',')
            result &= df[pair_cols].notna().all(axis=1)
        return result
    elif op == 'notna':
        return df[cols].notna()
    elif op == 'any_notna':
        col_list = cols.split(',')
        return df[col_list].notna().any(axis=1)
    elif op == 'all_null':
        col_list = cols.split(',')
    elif op == 'isna_or_blank':
        return df[cols].isna() | (df[cols] == '')
    elif op == 'isnull_or_zero':
        return (df[cols].isna() | (df[cols] == 0))
    elif op == 'any_zero':
        col_list = cols.split(',')
        return (df[col_list] == 0).any(axis=1)
    elif op == 'no_zeros': # all columns must be non-zero - there may be 1 or no values / nulls
        col_list = cols.split(',')
        return (df[col_list] != 0).all(axis=1)
    elif op == 'paired_notna':
        col_pairs = cols.split('|')
        result = pd.Series(True, index=df.index)
        for pair in col_pairs:
            pair_cols = pair.split(',')
            result &= df[pair_cols].notna().all(axis=1)
        return result
    elif op == 'notna_pair':
        col_pairs = cols.split('|')
        result = pd.Series(True, index=df.index)
        for pair in col_pairs:
            pair_cols = pair.split(',')
            result &= df[pair_cols].notna().all(axis=1)
        return result
    elif op == 'ctrl_test_ok':
        col_groups = cols.split('|')
        result = pd.Series(True, index=df.index)
        for group in col_groups:
            group_cols = group.split(',')
            result &= (
                df[group_cols[0]].notna() & 
                (df[group_cols[1]] == 'GOOD') &
                (df[group_cols[2]] == 'OK')
            )
        return result
    elif op == 'notinpo_test_ok':
        col_groups = cols.split('|')
        result = pd.Series(True, index=df.index)
        for group in col_groups:
            group_cols = group.split(',')
            result &= (
                df[group_cols[0]].notna() & 
                (df[group_cols[1]] != 'OK') &
                (df[group_cols[2]] != 'OK')
            )
        return result
    elif op == 'name_without_config':
        col_pairs = cols.split('|')
        result = pd.Series(True, index=df.index)
        for pair in col_pairs:
            name_col, config_col = pair.split(',')
            result &= (
                df[name_col].notna() & 
                (df[config_col].isna() | (df[config_col] != 'GOOD'))
            )
        return result
    elif op == 'always_false':
        return pd.Series(False, index=df.index)
    
    raise ValueError(f"Unknown operator: {op}")

def generate_defect_report_by_name(df: pd.DataFrame, report_name: str, debug: bool = False) -> pd.DataFrame:
    """Generate a defect report by name"""
    if report_name not in REPORT_CONFIGS:
        raise ValueError(f"Unknown report name: {report_name}")
    print(f"Generating report: {report_name} ... ", end='')
    updated_df = generate_defect_report(df, report_name, REPORT_CONFIGS[report_name], debug)
    print(f"{updated_df[updated_df[report_name] == True].shape[0]} matching rows. ({REPORT_CONFIGS[report_name]['name']})")
    return updated_df