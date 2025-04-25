import pandas as pd
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter
import openpyxl.styles


def create_points_section(df: pd.DataFrame) -> pd.DataFrame:
    """Create the points section of the report."""
    points = []
    
    for _, row in df.iterrows():
        
        if row.get('GenericType') in ['SD', 'DD']:
            point = {
                'Type': row.get('GenericType', ''),
                'SCADA Address': row.get('GenericPointAddress', ''),
                'eTerra Key': row.get('eTerraKey', ''),
                'PowerOn Alias': row.get('POAlias', ''),
                'ICCP Flag': row.get('ICCPFlag', ''),
                'Habdde Match Status': row.get('HbddeCompareStatus', ''),
                'PowerOn Config Health Status': row.get('ConfigHealth', ''),
                'Control Zone Status': row.get('CompAlarmAlarmZoneMatch', ''),
            }
            # Get the Ctrl Info
            if row.get('Controllable') == '1':
                if row.get('Ctrl1Addr','') != '':
                    point['Ctrl1Addr'] = row.get('Ctrl1Addr','')
                    point['Ctrl1Name'] = row.get('Ctrl1Name','')
                if row.get('Ctrl2Addr','') != '':
                    point['Ctrl2Addr'] = row.get('Ctrl2Addr','')
                    point['Ctrl2Name'] = row.get('Ctrl2Name','')
            else:
                point['Ctrl1Addr'] = ''
                point['Ctrl1Name'] = ''
                point['Ctrl2Addr'] = ''
                point['Ctrl2Name'] = ''

            # Get the Alarm Info
            if row.get('CompAlarmEterraAlias','') != '':
                point['CompAlarmeTerraAlarmZone'] = row.get('CompAlarmeTerraAlarmZone','')
                point['CompAlarmeTerraStatus'] = row.get('CompAlarmeTerraStatus','')
                point['CompAlarmPOsubstation'] = row.get('CompAlarmPOsubstation','')
                point['CompAlarmPOAlarmZone'] = row.get('CompAlarmPOAlarmZone','')
                point['CompAlarmPOAlarmRef'] = row.get('CompAlarmPOAlarmRef','')
                point['CompAlarmPOStatus'] = row.get('CompAlarmPOStatus','')
                point['CompAlarmAlarmZoneMatch'] = row.get('CompAlarmAlarmZoneMatch','')
                point['Alarm0_MessageMatch'] = row.get('Alarm0_MessageMatch','')
                point['Alarm1_MessageMatch'] = row.get('Alarm1_MessageMatch','')
                point['Alarm2_MessageMatch'] = row.get('Alarm2_MessageMatch','')
                point['Alarm3_MessageMatch'] = row.get('Alarm3_MessageMatch','')

            # Add the Report1 flag
            point['Report1'] = row.get('Report1', '')
            point['Report2'] = row.get('Report2', '')
            point['Report3'] = row.get('Report3', '')


            points.append(point)

    return pd.DataFrame(points)


def save_reports(reports: list, output_path: Path):
    """Save the report to an Excel file."""
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for report in reports:
            df = report['Content']
            df.to_excel(writer, sheet_name=report['RTU'], index=False)
        
        # Apply formatting
        worksheet = writer.sheets[report['RTU']]
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



def generate_defect_report_in_excel(df: pd.DataFrame, output_path: Path):
    report_columns = [
        {'ColName': 'GenericPointAddress', 'ColWidth': 25, 'ColFill': None},
        {'ColName': 'RTU', 'ColWidth': 7, 'ColFill': None}, 
        {'ColName': 'Sub', 'ColWidth': 7, 'ColFill': None},
        {'ColName': 'eTerraKey', 'ColWidth': 17, 'ColFill': None},
        {'ColName': 'eTerraAlias', 'ColWidth': 35, 'ColFill': None},
        {'ColName': 'GridIncomer', 'ColWidth': 10, 'ColFill': None},
        {'ColName': 'ICCP->PO', 'ColWidth': 7, 'ColFill': None},
        {'ColName': 'ICCP_ALIAS', 'ColWidth': 27, 'ColFill': None},
        {'ColName': 'PowerOn Alias', 'ColWidth': 35, 'ColFill': None},
        {'ColName': 'PowerOn Alias Exists', 'ColWidth': 7, 'ColFill': None},
        {'ColName': 'PowerOn Alias Linked to SCADA', 'ColWidth': 7, 'ColFill': None},
        {'ColName': 'Report1', 'ColWidth': 8, 'ColFill': None},
        {'ColName': 'Report2', 'ColWidth': 8, 'ColFill': None},
        {'ColName': 'Report3', 'ColWidth': 8, 'ColFill': None},
        {'ColName': 'Report4', 'ColWidth': 8, 'ColFill': None},
        {'ColName': 'Report5', 'ColWidth': 8, 'ColFill': None},
        {'ColName': 'Report6', 'ColWidth': 8, 'ColFill': None},
        {'ColName': 'Review Status', 'ColWidth': 12, 'ColFill': 'FFFFE0'},
        {'ColName': 'Comments', 'ColWidth': 60, 'ColFill': 'FFFFE0'}
    ]
    report_fields = [col['ColName'] for col in report_columns]

    # create a new dataframe with the report fields and the new columns
    report_df = pd.DataFrame(columns=report_fields)

    # First ensure all required columns exist in merged_data
    for col in report_fields:
        if col not in df.columns:
            df[col] = ''  # Add empty column if missing
            
    # Now we can safely select and concat
    report_df = pd.concat([report_df, df[report_fields]], ignore_index=True)

    # save the report dataframe to an xlsx file with formatting
    writer = pd.ExcelWriter(output_path / f"defect_report_all.xlsx", engine='openpyxl')
    report_df.to_excel(writer, index=False)
    
    # Get the worksheet
    worksheet = writer.sheets['Sheet1']
    
    # Add filters to row 1
    worksheet.auto_filter.ref = worksheet.dimensions
    
    # Format header row
    for cell in worksheet[1]:
        cell.font = openpyxl.styles.Font(bold=True)
        cell.fill = openpyxl.styles.PatternFill(start_color='B8CCE4', end_color='B8CCE4', fill_type='solid')
        
    # Freeze top row
    worksheet.freeze_panes = worksheet['F2']
    
    # Apply the column widths
    for idx, col in enumerate(report_columns, 1):
        worksheet.column_dimensions[get_column_letter(idx)].width = col['ColWidth']

    # Apply the fill colors and set all borders to be black and 1pt thick
    for idx, col in enumerate(report_columns, 1):

            for row in range(2, len(report_df) + 2):
                cell = worksheet[f"{get_column_letter(idx)}{row}"]
                if col['ColFill']:
                    cell.fill = openpyxl.styles.PatternFill(start_color=col['ColFill'], end_color=col['ColFill'], fill_type='solid')
                cell.border = openpyxl.styles.Border(left=openpyxl.styles.Side(style='thin', color='000000'),
                                                    right=openpyxl.styles.Side(style='thin', color='000000'),
                                                    top=openpyxl.styles.Side(style='thin', color='000000'),
                                                    bottom=openpyxl.styles.Side(style='thin', color='000000'))

    # Rename report columns to be more readable
    rename_cols = {
        'Report1': 'Missing Analog Components',
        'Report2': 'Missing Controllable Points', 
        'Report3': 'Missing Digital Inputs',
        'Report4': 'Components Missing Telecontrol Actions'
    }
    for idx, col in enumerate(report_columns, 1):
        if col['ColName'] in rename_cols:
            worksheet[f"{get_column_letter(idx)}1"].value = rename_cols[col['ColName']]

    writer.close()
    
    print(f"Defect report generated successfully: {output_path / f'defect_report_all.xlsx'}")

    return 

