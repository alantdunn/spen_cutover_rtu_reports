#!/usr/bin/env python3

import os
import sys
import pandas as pd
import sqlite3
from pathlib import Path
from typing import List, Dict, Optional
from data_import.utils import (
    filter_data_by_rtu,
    filter_data_by_substation,
)
from data_import.import_habdde import (
    import_habdde_export_point_tab,
    import_habdde_export_analog_tab,
    derive_rtu_addresses_and_protocols_from_eterra_export,
    import_habdde_export_control_tab,
    import_habdde_export_setpoint_control_tab,
    add_control_info_to_eterra_export
)
from data_import.import_poweron_rtu_report import clean_all_rtus
from data_import.import_alarm_compare import clean_compare_alarms
from data_import.import_controls_auto_test_report import clean_controls_test
from data_import.import_manual_commissioning_data import clean_manual_commissioning
from data_import.import_habdde_compare import clean_habdde_compare
from report_generation import (
    create_points_section,
    save_reports
)
from defect_reports import (
    defect_report1,
    defect_report2,
    defect_report3,
    defect_report4,
    defect_report5,
    defect_report6
)
import configparser



CONFIG_FILE = 'rtu_reports.ini'
DEFAULT_DATA_DIR = 'rtu_report_data'
DEFAULT_CONFIG_DIR = 'rtu_report_config'
DEFAULT_OUTPUT_DIR = 'reports'

def load_config(config_path=CONFIG_FILE):
    config = configparser.ConfigParser()
    config.read(config_path)
    return config

class RTUReportGenerator:
    def __init__(self, config_path=CONFIG_FILE, data_dir=""):
        self.config = load_config(config_path)
        if not self.config.has_section('Paths'):
            raise ValueError("Config file missing required 'Paths' section")
        
        if data_dir == "":
            self.data_dir = Path(self.config['Paths']['data_dir'])
        else:
            self.data_dir = Path(data_dir)

        if not self.data_dir.exists():
            raise ValueError(f"Data directory not found: {self.data_dir}")
        
        # Create debug directory if specified in config
        if 'debug_dir' in self.config['Paths']:
            self.debug_dir = Path(self.config['Paths']['debug_dir'])
            self.debug_dir.mkdir(exist_ok=True)
            print(f"Debug directory: {self.debug_dir}")
        else:
            self.debug_dir = None
            print("Debug directory not specified in config")

        # Create output directory if specified in config
        if 'output_dir' in self.config['Paths']:
            self.output_dir = Path(self.config['Paths']['output_dir'])
            self.output_dir.mkdir(exist_ok=True)
            print(f"Output directory: {self.output_dir}")
        else:
            self.output_dir = Path(DEFAULT_OUTPUT_DIR)
        
        # Default file names that will be overridden by config
        self.required_files = {
            'eterra_export': "habdde_eTerra_export.xlsx",
            'habdde_compare': "habdde_compare_report.xlsx", 
            'all_rtus': "all_rtus.csv",
            'controls_test': "controls_test.xlsx",
            'iccp_compare': "eterra_poweron_iccp_compare_report.xlsx",
            'compare_alarms': "Comparison_eTerra_dl11_all_po_dl11_4.xlsx",
            'controls_db': "controls.db"
        }
            
        if Path(config_path).exists():
            config = configparser.ConfigParser()
            config.read(config_path)
            if 'Files' in config:
                for key in self.required_files:
                    if key in config['Files']:
                        self.required_files[key] = config['Files'][key]
        
        # Initialize dataframes
        self.eterra_point_export = None
        self.eterra_analog_export = None
        self.eterra_control_export = None
        self.eterra_setpoint_control_export = None
        self.eterra_rtu_map = None
        self.eterra_export = None
        self.habdde_compare = None
        self.all_rtus = None
        self.controls_test = None
        self.iccp_compare = None
        self.compare_alarms = None
        self.manual_commissioning = None

    def validate_data_files(self) -> bool:
        """Check if all required files exist in the data directory."""
        missing_files = []
        for file_key, filename in self.required_files.items():
            if not (self.data_dir / filename).exists():
                missing_files.append(filename)
        
        if missing_files:
            print(f"Error: Missing required files: {', '.join(missing_files)}")
            return False
        return True

    def load_data(self, rtu_name: Optional[str] = None, substation: Optional[str] = None):
        """Load all source data into dataframes."""
        try:
            print(f"Loading eTerra export from {self.data_dir / self.required_files['eterra_export']}")
            self.eterra_point_export = import_habdde_export_point_tab(self.data_dir / self.required_files['eterra_export'], self.debug_dir)

            # Create a map of RTU addresses and protocols from the eTerra export
            self.eterra_rtu_map = derive_rtu_addresses_and_protocols_from_eterra_export(self.eterra_point_export, self.debug_dir)

            print(f"Loading analog export from {self.data_dir / self.required_files['eterra_export']}")
            self.eterra_analog_export = import_habdde_export_analog_tab(self.data_dir / self.required_files['eterra_export'], self.debug_dir)
            
            print(f"Loading control export from {self.data_dir / self.required_files['eterra_export']}")
            self.eterra_control_export = import_habdde_export_control_tab(self.data_dir / self.required_files['eterra_export'], self.debug_dir)

            print(f"Loading setpoint control export from {self.data_dir / self.required_files['eterra_export']}")
            self.eterra_setpoint_control_export = import_habdde_export_setpoint_control_tab(self.data_dir / self.required_files['eterra_export'], self.debug_dir)

            # Get just the common columns from point and analog and concatenate them together, sort by GenericPointAddress
            common_columns = [  'GenericPointAddress', 'CASDU', 'Protocol', 'RTU', 'Card',
                                'RTUAddress', 'RTUId', 'IOA2', 'IOA1', 'IOA', 'PointId', 
                                'GenericType', 'DeviceType', 'DeviceName', 'DeviceId', 
                                'Sub', 'Word', 'eTerraKey', 'eTerraAlias', 'Controllable']
            potentially_common_columns = ['IGNORE_RTU',
                                'IGNORE_POINT',
                                'OLD_DATA',
                                'GridIncomer',
                                'eTerra Alias',
                                'ICCP_POINTNAME',
                                'ICCP->PO',
                                'ICCP_ALIAS',
                                'PowerOn Alias',
                                'PowerOn Alias Exists',
                                'PowerOn Alias Linked to SCADA']
            # add the potentially common columns to the common columns if they exist in the point and analog exports
            common_columns.extend([col for col in potentially_common_columns if col in self.eterra_point_export.columns])
            common_columns.extend([col for col in potentially_common_columns if col in self.eterra_analog_export.columns])

            print("Combining point and analog exports...")
            eterra_points_common_cols = self.eterra_point_export[common_columns]
            eterra_analogs_common_cols = self.eterra_analog_export[common_columns]
            eterra_export = pd.concat([eterra_points_common_cols, eterra_analogs_common_cols], ignore_index=True)
            eterra_export = eterra_export.sort_values(by='GenericPointAddress')
            if self.debug_dir:
                eterra_export.to_csv(f"{self.debug_dir}/eterra_export.csv", index=False)

            if rtu_name:
                eterra_export = eterra_export[eterra_export['RTU'] == rtu_name]
            elif substation:
                eterra_export = eterra_export[eterra_export['Sub'] == substation]

            print("Adding control info to eTerra export...")
            self.eterra_export = add_control_info_to_eterra_export(eterra_export, self.eterra_control_export, self.eterra_setpoint_control_export)
            if self.debug_dir:
                self.eterra_export.to_csv(f"{self.debug_dir}/eterra_export_with_control_info.csv", index=False)

            print(f"Loading habdde compare from {self.data_dir / self.required_files['habdde_compare']}")
            self.habdde_compare = pd.read_csv(self.data_dir / self.required_files['habdde_compare'], low_memory=False)
            self.habdde_compare = clean_habdde_compare(self.habdde_compare)
            if self.debug_dir:
                self.habdde_compare.to_csv(f"{self.debug_dir}/habdde_compare.csv", index=False)

            print(f"Loading all RTUs from {self.data_dir / self.required_files['all_rtus']}")
            self.all_rtus = pd.read_csv(self.data_dir / self.required_files['all_rtus'], low_memory=False)
            self.all_rtus = clean_all_rtus(self.all_rtus)
            if self.debug_dir:
                self.all_rtus.to_csv(f"{self.debug_dir}/all_rtus.csv", index=False)

            print(f"Loading controls test from {self.data_dir / self.required_files['controls_test']}")
            self.controls_test = pd.read_csv(self.data_dir / self.required_files['controls_test'])
            self.controls_test = clean_controls_test(self.controls_test, self.eterra_rtu_map)
            if self.debug_dir:
                self.controls_test.to_csv(f"{self.debug_dir}/controls_test.csv", index=False)

            print(f"Loading compare alarms from {self.data_dir / self.required_files['compare_alarms']}")
            self.compare_alarms = pd.read_excel(self.data_dir / self.required_files['compare_alarms'], sheet_name='Event Detail')
            self.compare_alarms = clean_compare_alarms(self.compare_alarms)
            if self.debug_dir:
                self.compare_alarms.to_csv(f"{self.debug_dir}/compare_alarms.csv", index=False)

            print(f"Loading manual commissioning from {self.data_dir / self.required_files['controls_db']}")
            conn = sqlite3.connect(self.data_dir / self.required_files['controls_db'])
            self.manual_commissioning = pd.read_sql_query("SELECT * FROM test_results", conn)
            conn.close()
            self.manual_commissioning = clean_manual_commissioning(self.manual_commissioning)
            if self.debug_dir:
                self.manual_commissioning.to_csv(f"{self.debug_dir}/manual_commissioning.csv", index=False)

        except Exception as e:
            print(f"Error loading data: {str(e)}")
            print(f"Error type: {type(e)}")
            import traceback
            print(traceback.format_exc())
            sys.exit(1)

    def merge_data(self) -> pd.DataFrame:
        """Merge all data sources into a single dataframe."""

        # Merge eTerra export with habdde compare
        print(f"Merging eTerra export with habdde compare on {self.eterra_export.shape[0]} rows")
        merged = pd.merge(
            self.eterra_export,
            self.habdde_compare,
            on=['GenericPointAddress'],
            how='left'
        )
        print(f"Merged eTerra export with habdde compare on {merged.shape[0]} rows")
        
        print("Merging with all RTUs ... ")
        # Merge with all RTUs
        merged = pd.merge(
            merged,
            self.all_rtus,
            on=['GenericPointAddress'],
            how='left'
        )
        print(f"Merged with all RTUs on {merged.shape[0]} rows")
        
        # Merge with ICCP compare
        # merged = pd.merge(
        #     merged,
        #     self.iccp_compare,
        #     on=['RTU', 'Card', 'Word'],
        #     how='left'
        # )

        print("Merge report columns:")
        print(merged.columns)

        
        # Merge with compare report - this needs to be done in 2 parts
        # 1. First there are some columsn that we want that are associate with the point - to get these we need to get a subset fo columns then de-duplicate before the merge
        # 2. Then we want to add a small set of fields for each associated alarm.
        # There are either 2 or 4 alarms per point
        # SD - 2 alarms for 0 and 1
        # DD - 4 alarms for 0, 1, 2, 3

        #1.a) get just the useful and point related columns
        # Get the columns that exist in the merged dataframe
        available_columns = self.compare_alarms.columns.tolist()
        
        # Define the columns we want if they exist
        desired_columns = [
            'CompAlarmEterraAlias',
            'CompAlarmPOAlias', 
            'CompAlarmeTerraAlarmZone',
            'CompAlarmeTerraStatus',
            'CompAlarmPOsubstation',
            'CompAlarmPOAlarmZone', 
            'CompAlarmPOAlarmRef',
            'CompAlarmPOStatus',
            'CompAlarmAlarmZoneMatch'
        ]
        
        # Only include columns that exist in the merged dataframe
        point_related_columns = [col for col in desired_columns if col in available_columns]
        
        if not point_related_columns:
            print("Warning: No matching columns found for point_related_columns")
            point_related_df = self.compare_alarms.copy()
        else:
            point_related_df = self.compare_alarms[point_related_columns]
            point_related_df = point_related_df.drop_duplicates()

        #1.b) merge the point related df with the compare alarms df
        merged = pd.merge(
            merged,
            point_related_df,
            left_on=['eTerraAlias'],
            right_on=['CompAlarmEterraAlias'],
            how='left'
        )

        #1.c de-duplicate the merged df
        merged = merged.drop_duplicates()
        #1.d) add the alarm related columns into Alarm<value>_eTerraMessage and Alarm<value>_POMessage, and Alarm<value>_MessageMatch
        # Initialize empty columns for all possible alarm values
        for value in range(4):
            merged[f'Alarm{value}_eTerraMessage'] = None
            merged[f'Alarm{value}_POMessage'] = None 
            merged[f'Alarm{value}_MessageMatch'] = None

        # For each row in merged, find matching rows in compare_alarms and populate corresponding alarm columns
        for idx, row in merged.iterrows():
            matching_alarms = self.compare_alarms[
                self.compare_alarms['CompAlarmEterraAlias'] == row['eTerraAlias']
            ]
            
            # For each matching alarm row, populate the corresponding alarm columns based on CompAlarmValue
            for _, alarm_row in matching_alarms.iterrows():
                value = alarm_row['CompAlarmValue']
                merged.at[idx, f'Alarm{value}_eTerraMessage'] = alarm_row['CompAlarmeTerraAlarmMessage']
                merged.at[idx, f'Alarm{value}_POMessage'] = alarm_row['CompAlarmPOAlarmMessage']
                merged.at[idx, f'Alarm{value}_MessageMatch'] = alarm_row['CompAlarmAlarmMessageMatch']
        
        # Control information needs to be joined differently as only a few key fields are requried for each associated control

        # Merge with controls test

        # Merge with manual commissioning

        if self.debug_dir:
            merged.to_csv(f"{self.debug_dir}/merged.csv", index=False)
        
        return merged
    
    def add_issue_report_flags(self, merged_data: pd.DataFrame) -> pd.DataFrame:
        """Add issue report flags to the merged data."""
        # Add a flag for each issue type
        # 1. Missing Analog components in PowerOn
        # 2. Missing Controllable Points in PowerOn
        # 3. Missing Digital Inputs in PowerOn
        # 4. Components Missing Telecontrol Actions in PowerOn
        # 5. Items missing from PowerOn that are in eTerra
        #6. Components missing alarm references in Poweron
        merged_data = defect_report1(merged_data)
        
        return merged_data

    def generate_report(self, rtu_name: Optional[str] = None, substation: Optional[str] = None):
        """Generate report for specified RTU or substation."""
        if not self.validate_data_files():
            sys.exit(1)
            
        self.load_data(rtu_name, substation)
        # self.debug_print_dataframes()
        merged_data = self.merge_data()
        merged_data = self.add_issue_report_flags(merged_data)
        
        # Filter data based on criteria
        # if rtu_name:
        #     filtered_data = filter_data_by_rtu(merged_data, rtu_name)
        # elif substation:
        #     filtered_data = filter_data_by_substation(merged_data, substation)
        # else:
        #     filtered_data = merged_data
        filtered_data = merged_data
        # Get list of RTUs in the filtered data
        rtus = filtered_data['RTU'].unique()
        print(f"{len(rtus)} RTUs in the filtered data for the report")

        # We will create a report for each RTU in the filtered data
        reports = []
        for rtu in rtus:
            rtu_data = filter_data_by_rtu(filtered_data, rtu)
            # Create report sections
            points_section = create_points_section(rtu_data)

            # Combine sections
            report_content = pd.concat([points_section], ignore_index=True)
            report = {'RTU': rtu, 'Content': report_content}
            reports.append(report)
        
        # Save report
        output_path = self.output_dir / f"rtu_report_{rtu_name or substation or 'all'}.xlsx"
        save_reports(reports, output_path)
        print(f"Report generated successfully: {output_path}")

    
    def debug_print_dataframes(self):
        try:
            # print the row counts and columns in each dataframe in a readable format
            print("eterra_point_export row count:", self.eterra_point_export.shape[0])
            print("eterra_point_export columns:")
            print(self.eterra_point_export.columns)
            print("eterra_analog_export row count:", self.eterra_analog_export.shape[0])
            print("eterra_analog_export columns:")
            print(self.eterra_analog_export.columns)
            print("eterra_control_export row count:", self.eterra_control_export.shape[0])
            print("eterra_control_export columns:")
            print(self.eterra_control_export.columns)
            print("eterra_setpoint_control_export row count:", self.eterra_setpoint_control_export.shape[0])
            print("eterra_setpoint_control_export columns:")
            print(self.eterra_setpoint_control_export.columns)
            print("eterra_export row count:", self.eterra_export.shape[0])
            print("eterra_export columns:")
            print(self.eterra_export.columns)
            print("habdde_compare row count:", self.habdde_compare.shape[0])
            print("habdde_compare columns:")
            print(self.habdde_compare.columns)
            print("all_rtus row count:", self.all_rtus.shape[0])
            print("all_rtus columns:")
            print(self.all_rtus.columns)
            print("controls_test row count:", self.controls_test.shape[0])
            print("controls_test columns:")
            print(self.controls_test.columns)
            print("compare_alarms row count:", self.compare_alarms.shape[0])
            print("compare_alarms columns:")
            print(self.compare_alarms.columns)
            print("manual_commissioning row count:", self.manual_commissioning.shape[0])
            print("manual_commissioning columns:")
            print(self.manual_commissioning.columns)
        except Exception as e:
            print(f"Error: {str(e)}")
            sys.exit(1)

def main():
    import argparse
    
    parser = argparse.ArgumentParser(description="Generate RTU reports from various source files")
    parser.add_argument("--rtu", help="Generate report for specific RTU")
    parser.add_argument("--substation", help="Generate report for specific substation")
    parser.add_argument("--data-dir", default=DEFAULT_DATA_DIR, help="Directory containing source files")
    parser.add_argument("--config-dir", default=DEFAULT_CONFIG_DIR, help="Directory containing config files")
    
    args = parser.parse_args()
    
    generator = RTUReportGenerator(args.config_dir + '/' + CONFIG_FILE, args.data_dir)
    generator.generate_report(args.rtu, args.substation)

if __name__ == "__main__":
    main() 