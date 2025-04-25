import pandas as pd
from pathlib import Path




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