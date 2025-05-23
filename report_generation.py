import pandas as pd
from pathlib import Path
from rich import print

import openpyxl
from openpyxl.utils import get_column_letter
import openpyxl.styles
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Border, Side, PatternFill, Font, NamedStyle, GradientFill, Alignment

# Define some openpyxl styles for the report

# Define a style for the header row
header_style = NamedStyle(name="header")
header_style.font = Font(bold=True, size=11)
bd = Side(style='thick', color="000000")
header_style.border = Border(left=bd, top=bd, right=bd, bottom=bd)
header_style.fill = PatternFill(start_color='B8CCE4', end_color='B8CCE4', fill_type='solid')

# Define a style for the navigation row
navigation_style = NamedStyle(name="navigation")
navigation_style.font = Font(color='0000AA', size=14)

# Define a style for the traffic light cells
traffic_light_style = NamedStyle(name="traffic_light")
traffic_light_style.font = Font(color='000000')
traffic_light_style.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
traffic_light_style.border = Border(left=Side(style='thick', color='000000'),
                                right=Side(style='thick', color='000000'),
                                top=Side(style='thick', color='000000'),
                                bottom=Side(style='thick', color='000000'))



def setup_navigation_row(ws, navigation_row):
    for cell in ws[navigation_row]:
        cell.style = 'navigation'

    # Insert a Hyperlink to the tab "Overview" in column A
    ws.cell(row=navigation_row, column=1).value = "Go to Overview"
    ws.cell(row=navigation_row, column=1).hyperlink = "Overview!A1"

def apply_header_style(ws, header_row):
    for cell in ws[header_row]:
        cell.style = 'header'




def apply_conditional_formatting(ws, report_columns, start_row=2, debug=False):
    for idx, col in enumerate(report_columns, start=1):
        col_letter = ws.cell(row=1, column=idx).column_letter
        cf_type = col.get('ConditionalFormatting')

        if not cf_type:
            continue  # Skip if no conditional formatting needed

        last_row = ws.max_row
        range_ref = f"{col_letter}{start_row}:{col_letter}{last_row}"  # all rows from start_row to last row

        if cf_type == 'ZeroOne':
            if debug:
                print(f"ZeroOne: {range_ref}")
            # Zero should have no formatting
            ws.conditional_formatting.add(range_ref,
                CellIsRule(operator='equal', formula=['0'], fill=PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')))
            # One should have a light orange background with dark orange text   
            ws.conditional_formatting.add(range_ref,
                CellIsRule(operator='equal', formula=['1'], fill=PatternFill(start_color='FFD966', end_color='FFD966', fill_type='solid')))

        elif cf_type == 'ZeroTwo':
            if debug:
                print(f"ZeroTwo: {range_ref}")
            # Zero should have a light red background with dark red text
            ws.conditional_formatting.add(range_ref,
                CellIsRule(operator='equal', formula=['0'], fill=PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')))
            # Two should have a light green background with dark green text
            ws.conditional_formatting.add(range_ref,
                CellIsRule(operator='equal', formula=['2'], fill=PatternFill(start_color='A9D08E', end_color='A9D08E', fill_type='solid')))

        elif cf_type == 'TrueFalse':
            if debug:
                print(f"TrueFalse: {range_ref}")
            # True should have a light green background with dark green text
            ws.conditional_formatting.add(range_ref,
                CellIsRule(operator='equal', formula=['TRUE'], fill=PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')))
            # False should have a light red background with dark red text
            ws.conditional_formatting.add(range_ref,
                CellIsRule(operator='equal', formula=['FALSE'], fill=PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')))

        elif cf_type == 'GoodBadNA':
            if debug:
                print(f"GoodBadNA: {range_ref}")
            # To start set a 2 px white border around the cell
            ws.conditional_formatting.add(range_ref,
                CellIsRule(operator='equal', formula=['""'], fill=PatternFill(start_color='FFFFFF', end_color='FFFFFF', 
                                                                fill_type='solid'), border=Border(
                                                                        left=Side(style='thin', color='000000'), 
                                                                        right=Side(style='thin', color='000000'), 
                                                                        top=Side(style='thin', color='000000'), 
                                                                        bottom=Side(style='thin', color='000000'))))

            # GOOD should have a green background  green text
            ws.conditional_formatting.add(range_ref,
                CellIsRule(operator='equal', formula=['1'], fill=PatternFill(start_color='63EF45', end_color='63EF45', fill_type='solid'), font=Font(color='63EF45')))
            # BAD should have a red background red text
            ws.conditional_formatting.add(range_ref,
                CellIsRule(operator='equal', formula=['0'], fill=PatternFill(start_color='C0504D', end_color='C0504D', fill_type='solid'), font=Font(color='C0504D')))
            # blank cells should have no formatting
            ws.conditional_formatting.add(range_ref,
                CellIsRule(operator='equal', formula=['""'], fill=PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid'), font=Font(color='FFFFFF')))
            
        elif cf_type == 'XBlank':
            if debug:
                print(f"XBlank: {range_ref}")
            # X should have a light orange background with dark orange text
            ws.conditional_formatting.add(range_ref,
                FormulaRule(formula=[f'EXACT({col_letter}{start_row},"X")'], 
                            fill=PatternFill(start_color='FFD966', end_color='FFD966', fill_type='solid')))
            # Blank should have no formatting
            ws.conditional_formatting.add(range_ref,
                FormulaRule(formula=[f'ISBLANK({col_letter}{start_row})'], 
                            fill=PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')))

        elif cf_type == 'Bold':
            if debug:
                print(f"Bold: {range_ref}")
            ws.conditional_formatting.add(range_ref,
                FormulaRule(formula=[f'LEN(TRIM({col_letter}{start_row}))>0'], 
                            font=Font(bold=True)))

        elif cf_type == 'Italic':
            if debug:
                print(f"Italic: {range_ref}")
            ws.conditional_formatting.add(range_ref,
                FormulaRule(formula=[f'LEN(TRIM({col_letter}{start_row}))>0'], 
                            font=Font(italic=True)))
            



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

            # Add the Report flags
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

    print("="*80)
    print("Generating Defect Report in Excel")
    print("="*80)

    # Derive the alarm status cols (for alarm number 0..3)
    for i in range(4):
        # df[f'Alarm{i}'] = df[f'Alarm{i}_MessageMatch'].apply(lambda x: '1' if x == 'TRUE' else '0' if x == 'FALSE' else '')
        df[f'Alarm{i}'] = df[f'Alarm{i}_MessageMatch'].apply(lambda x: 1 if x == True else 0 if x == False else '')
        #df[f'Alarm{i}'] = df[f'Alarm{i}_MessageMatch']

    # Derive the ctrl status cols (for ctrl number 1..2)
    for i in range(1,3):
        df[f'Ctrl{i}'] = df.apply(lambda row: 1 if row[f'Ctrl{i}TestResult'] == 'OK' else 0 if row[f'Ctrl{i}TestResult'] == 'Fail' else '', axis=1)
        df[f'Ctrl{i}V'] = df.apply(lambda row: 1 if row[f'Ctrl{i}VisualCheckResult'] == 'OK' else 0 if row[f'Ctrl{i}VisualCheckResult'] == 'Fail' else '', axis=1)
        df[f'Ctrl{i}C'] = df.apply(lambda row: 1 if row[f'Ctrl{i}ControlSentResult'] == 'OK' else 0 if row[f'Ctrl{i}ControlSentResult'] == 'Fail' else '', axis=1)

    # Sort out a few bespoke columns
    # # first get a Type column that also flags the dummy rows are DUMMY, Get the value of GenericType unless the RTUId = '(€€€€€€€€:)'
    # df['Type'] = df.apply(lambda row: 'DUMMY' if row['RTUId'] == '(€€€€€€€€:)' else row['GenericType'], axis=1)
    # # now make an ignore column that is TRUE if any of IGNORE_RTU, IGNORE_POINT, OLD_DATA are TRUE
    # df['Ignore'] = df.apply(lambda row: True if (row['IGNORE_RTU'] == True or row['IGNORE_POINT'] == True or row['OLD_DATA'] == True ) else False, axis=1)
    # # We want to add a new column 'RTUComms' to the df that is True if the DeviceType is 'RTU and the eTerraAlias does not contain 'LDC'
    # df['RTUComms'] = df.apply(lambda row: True if row['DeviceType'] == 'RTU' and 'LDC' not in row['eTerraAlias'] else False, axis=1)

    report_columns = [
        {'dfCol': 'GenericPointAddress',            'ColName': 'GenericPointAddress',           'ColWidth': 25,     'Align': 'left',    'ColFill': None,        'ConditionalFormatting': None,      'Hidden': False},
        {'dfCol': 'Type',                           'ColName': 'Type',                          'ColWidth': 3,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': None,      'Hidden': False},
        {'dfCol': 'RTU',                            'ColName': 'RTU',                           'ColWidth': 7,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': None,      'Hidden': False}, 
        {'dfCol': 'Sub',                            'ColName': 'Sub',                           'ColWidth': 7,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': None,      'Hidden': False},
        {'dfCol': 'Ignore',                         'ColName': 'Ignore',                        'ColWidth': 7,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': None,      'Hidden': False},
        {'dfCol': 'eTerraKey',                      'ColName': 'eTerraKey',                     'ColWidth': 17,     'Align': 'left',    'ColFill': None,        'ConditionalFormatting': None,      'Hidden': False},
        {'dfCol': 'eTerraAlias',                    'ColName': 'eTerraAlias',                   'ColWidth': 35,     'Align': 'left',    'ColFill': None,        'ConditionalFormatting': None,      'Hidden': False},
        {'dfCol': 'GridIncomer',                    'ColName': 'GridIncomer',                   'ColWidth': 10,     'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'ZeroOne',      'Hidden': False},
        {'dfCol': 'RTUComms',                       'ColName': 'RTUComms',                      'ColWidth': 7,     'Align': 'center',  'ColFill': None,         'ConditionalFormatting': None,      'Hidden': False},
        {'dfCol': 'PointId',                        'ColName': 'PointId',                       'ColWidth': 5,     'Align': 'center',  'ColFill': None,         'ConditionalFormatting': 'None',      'Hidden': False},
        {'dfCol': 'ICCP->PO',                       'ColName': 'ICCP->PO',                      'ColWidth': 7,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'XBlank',      'Hidden': False},
        {'dfCol': 'ICCP_ALIAS',                     'ColName': 'ICCP_ALIAS',                    'ColWidth': 27,     'Align': 'left',    'ColFill': None,        'ConditionalFormatting': None,      'Hidden': False},
        {'dfCol': 'PowerOn Alias',                  'ColName': 'PowerOn Alias',                 'ColWidth': 35,     'Align': 'left',    'ColFill': None,        'ConditionalFormatting': None,      'Hidden': False},
        {'dfCol': 'CompAlarmTemplateAlias',         'ColName': 'Template'                  ,    'ColWidth': 30,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'None',      'Hidden': False},
        {'dfCol': 'PowerOn Alias Exists',           'ColName': 'PowerOn Alias Exists',          'ColWidth': 7,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'TrueFalse',      'Hidden': False},
        {'dfCol': 'PowerOn Alias Linked to SCADA',  'ColName': 'PowerOn Alias Linked to SCADA', 'ColWidth': 7,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'ZeroTwo',      'Hidden': False},
        {'dfCol': 'CompAlarmTemplateType',          'ColName': 'TemplateType',                  'ColWidth': 3,     'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'None',      'Hidden': False},
        {'dfCol': 'CompAlarmStateIndex',            'ColName': 'StateIndex',                    'ColWidth': 3,     'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'None',      'Hidden': False},
        {'dfCol': 'Alarm0',                         'ColName': 'Alarm0',                        'ColWidth': 4,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'GoodBadNA',      'Hidden': False},
        {'dfCol': 'Alarm1',                         'ColName': 'Alarm1',                        'ColWidth': 4,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'GoodBadNA',      'Hidden': False},
        {'dfCol': 'Alarm2',                         'ColName': 'Alarm2',                        'ColWidth': 4,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'GoodBadNA',      'Hidden': False},
        {'dfCol': 'Alarm3',                         'ColName': 'Alarm3',                        'ColWidth': 4,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'GoodBadNA',      'Hidden': False},
        {'dfCol': 'Controllable',                   'ColName': 'Controllable',                  'ColWidth': 10,     'Align': 'center',  'ColFill': None,        'ConditionalFormatting': None,      'Hidden': False},
        {'dfCol': 'Ctrl1',                          'ColName': 'Ctrl1',                         'ColWidth': 4,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'GoodBadNA',      'Hidden': False},
        {'dfCol': 'Ctrl1V',                         'ColName': 'Ctrl1V',                        'ColWidth': 1,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'GoodBadNA',      'Hidden': True},
        {'dfCol': 'Ctrl1C',                         'ColName': 'Ctrl1C',                        'ColWidth': 1,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'GoodBadNA',      'Hidden': True},
        {'dfCol': 'Ctrl2',                          'ColName': 'Ctrl2',                         'ColWidth': 4,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'GoodBadNA',      'Hidden': False},
        {'dfCol': 'Ctrl2V',                         'ColName': 'Ctrl2V',                        'ColWidth': 1,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'GoodBadNA',      'Hidden': True},
        {'dfCol': 'Ctrl2C',                         'ColName': 'Ctrl2C',                        'ColWidth': 1,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'GoodBadNA',      'Hidden': True},
        {'dfCol': 'Ctrl1Name',                      'ColName': 'Ctrl1Name',                     'ColWidth': 10,     'Align': 'left',    'ColFill': None,        'ConditionalFormatting': 'Bold',      'Hidden': False},
        {'dfCol': 'Ctrl1Comments',                  'ColName': 'Ctrl1Comments',                 'ColWidth': 10,     'Align': 'left',    'ColFill': None,        'ConditionalFormatting': 'Italic',      'Hidden': True},
        {'dfCol': 'Ctrl1ConfigHealth',              'ColName': 'Ctrl1ConfigHealth',             'ColWidth': 10,     'Align': 'left',    'ColFill': None,        'ConditionalFormatting': 'Italic',      'Hidden': False},
        {'dfCol': 'Ctrl2Name',                      'ColName': 'Ctrl2Name',                     'ColWidth': 10,     'Align': 'left',    'ColFill': None,        'ConditionalFormatting': 'Bold',      'Hidden': False},
        {'dfCol': 'Ctrl2Comments',                  'ColName': 'Ctrl2Comments',                 'ColWidth': 10,     'Align': 'left',    'ColFill': None,        'ConditionalFormatting': 'Italic',      'Hidden': True},
        {'dfCol': 'Ctrl2ConfigHealth',              'ColName': 'Ctrl2ConfigHealth',             'ColWidth': 10,     'Align': 'left',    'ColFill': None,        'ConditionalFormatting': 'Italic',      'Hidden': False},
        {'dfCol': 'AlarmMismatchComment',           'ColName': 'AlarmMismatchComment',          'ColWidth': 10,     'Align': 'left',    'ColFill': None,        'ConditionalFormatting': None,      'Hidden': False},
        {'dfCol': 'AlarmMismatchTemplateAlias',     'ColName': 'AlarmMismatchTemplateAlias',    'ColWidth': 10,     'Align': 'left',    'ColFill': None,        'ConditionalFormatting': None,      'Hidden': False},
        {'dfCol': 'Report1',            'ColName': 'Missing Analog Components',                 'ColWidth': 8,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'TrueFalse',      'Hidden': False},
        {'dfCol': 'Report2',            'ColName': 'Missing Digital Components',                'ColWidth': 8,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'TrueFalse',      'Hidden': False},
        {'dfCol': 'Report3',            'ColName': 'Missing Controllable Components',           'ColWidth': 8,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'TrueFalse',      'Hidden': False},
        {'dfCol': 'Report4',            'ColName': 'Components Missing Telecontrol Actions',    'ColWidth': 8,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'TrueFalse',      'Hidden': False},
        {'dfCol': 'Report5',            'ColName': 'Components Missing Alarm Reference',        'ColWidth': 8,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'TrueFalse',      'Hidden': False},
        {'dfCol': 'Report6',            'ColName': 'Controls not in PO but tested ok',          'ColWidth': 8,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'TrueFalse',      'Hidden': False},
        {'dfCol': 'Report7',            'ColName': 'Controls Not Linked',                       'ColWidth': 8,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'TrueFalse',      'Hidden': False},
        {'dfCol': 'Report8',            'ColName': 'Ctrl-able eTerra Points with no Controls',  'ColWidth': 8,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'TrueFalse',      'Hidden': False},
        {'dfCol': 'Report9',            'ColName': 'Alarm Mismatch Manual Actions',             'ColWidth': 8,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'TrueFalse',      'Hidden': False},
        {'dfCol': 'Report10',           'ColName': 'RESET w/ CtrlFunc 0',                       'ColWidth': 8,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'TrueFalse',      'Hidden': False},
        {'dfCol': 'Report11',           'ColName': 'SWDD with LAMP symbol',                     'ColWidth': 8,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'TrueFalse',      'Hidden': False},
        {'dfCol': 'Report12',           'ColName': 'Missing from DLPoint',                      'ColWidth': 8,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'TrueFalse',      'Hidden': False},
        {'dfCol': 'Report13',           'ColName': 'DD symbol should be SD',                    'ColWidth': 8,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'TrueFalse',      'Hidden': False},
        {'dfCol': 'Report14',           'ColName': 'SD symbol should be DD',                    'ColWidth': 8,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'TrueFalse',      'Hidden': False},
        {'dfCol': 'ReportANY',          'ColName': 'Any Defect',                                'ColWidth': 8,      'Align': 'center',  'ColFill': None,        'ConditionalFormatting': 'TrueFalse',      'Hidden': False},
        {'dfCol': 'Review Status',                  'ColName': 'Review Status',                 'ColWidth': 12,     'Align': 'left',    'ColFill': 'FFFFE0',    'ConditionalFormatting': None,      'Hidden': False},
        {'dfCol': 'Comments',                       'ColName': 'Comments',                      'ColWidth': 60,     'Align': 'left',    'ColFill': 'FFFFE0',    'ConditionalFormatting': None,      'Hidden': False}
    ]
    
    for idx, col in enumerate(report_columns, 1):
        if col['dfCol'] not in df.columns:
            print(f"Column {col['dfCol']} not found in df ... adding empty column")
            df[col['dfCol']] = ''  # Add empty column if missing
        # if the ColName is different from the dfCol, then rename the column
        if col['ColName'] != col['dfCol']:
            df = df.rename(columns={col['dfCol']: col['ColName']})

    report_fields = [col['ColName'] for col in report_columns]

    # create a new dataframe with the report fields and the new columns
    report_df = pd.DataFrame(columns=report_fields)

    # First ensure all required columns exist in merged_data
    for col in report_fields:
        if col not in df.columns:
            print(f"Column {col} not found in df ... adding empty column")
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

    for idx, col in enumerate(report_columns, 1):

        # Rename report columns to be more readable going from dfCol to ColName 
        if col['dfCol'] != col['ColName']:
            worksheet[f"{get_column_letter(idx)}1"].value = col['ColName']

        # Setup column widths
        worksheet.column_dimensions[get_column_letter(idx)].width = col['ColWidth']

        # apply the alignment
        for row in range(1, len(report_df) + 2):
            worksheet[f"{get_column_letter(idx)}{row}"].alignment = openpyxl.styles.Alignment(horizontal=col['Align'])

        # Apply the fill colors and set all borders to be black and 1pt thick
        for row in range(2, len(report_df) + 2):
                cell = worksheet[f"{get_column_letter(idx)}{row}"]
                if col['ColFill']:
                    cell.fill = openpyxl.styles.PatternFill(start_color=col['ColFill'], end_color=col['ColFill'], fill_type='solid')
                cell.border = openpyxl.styles.Border(left=openpyxl.styles.Side(style='thin', color='000000'),
                                                    right=openpyxl.styles.Side(style='thin', color='000000'),
                                                    top=openpyxl.styles.Side(style='thin', color='000000'),
                                                    bottom=openpyxl.styles.Side(style='thin', color='000000'))

    apply_conditional_formatting(worksheet, report_columns)

    # Hide the columns that are hidden
    for idx, col in enumerate(report_columns, 1):
        if col['Hidden']:
            worksheet.column_dimensions[get_column_letter(idx)].hidden = True

    writer.close()
    
    print(f"Defect report generated successfully: {output_path / f'defect_report_all.xlsx'}")

    return 

