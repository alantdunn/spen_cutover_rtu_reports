#!/usr/bin/env python3

import os
import sys
import pandas as pd
import sqlite3
from pathlib import Path
from typing import List, Dict, Optional
from utils import (
    filter_data_by_rtu,
    filter_data_by_substation,
    create_points_section,
    create_analogs_section,
    save_report,
    clean_eterra_point_export,
    clean_eterra_analog_export,
    clean_eterra_control_export,
    clean_habdde_compare,
    clean_all_rtus,
    clean_controls_test,
    clean_compare_alarms,
    clean_manual_commissioning,
    clean_eterra_setpoint_control_export,
    add_control_info_to_eterra_export
)
import configparser

from pylib3i.habdde import read_habdde_tab_into_df


CONFIG_FILE = 'rtu_reports.ini'
DEFAULT_DATA_DIR = 'rtu_report_data'
DEFAULT_CONFIG_DIR = 'rtu_report_config'

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

    def load_data(self):
        """Load all source data into dataframes."""
        try:
            print(f"Loading eTerra export from {self.data_dir / self.required_files['eterra_export']}")
            self.eterra_point_export = read_habdde_tab_into_df(self.data_dir / self.required_files['eterra_export'], 'POINT')
            if self.eterra_point_export is None:
                raise ValueError("Failed to read POINT tab from eTerra export")
            self.eterra_point_export = clean_eterra_point_export(self.eterra_point_export)
            if self.debug_dir:
                self.eterra_point_export.to_csv(f"{self.debug_dir}/eterra_point_export.csv", index=False)

            # Create a map of RTU addresses and protocols from the eTerra export
            self.eterra_rtu_map = self.eterra_point_export[['RTU', 'RTUAddress', 'Protocol']].drop_duplicates()
            if self.debug_dir:
                self.eterra_rtu_map.to_csv(f"{self.debug_dir}/eterra_rtu_map.csv", index=False)

            print(f"Loading analog export from {self.data_dir / self.required_files['eterra_export']}")
            self.eterra_analog_export = read_habdde_tab_into_df(self.data_dir / self.required_files['eterra_export'], 'ANALOG')
            if self.eterra_analog_export is None:
                raise ValueError("Failed to read ANALOG tab from eTerra export")
            self.eterra_analog_export = clean_eterra_analog_export(self.eterra_analog_export)
            if self.debug_dir:
                self.eterra_analog_export.to_csv(f"{self.debug_dir}/eterra_analog_export.csv", index=False)
            
            print(f"Loading control export from {self.data_dir / self.required_files['eterra_export']}")
            self.eterra_control_export = read_habdde_tab_into_df(self.data_dir / self.required_files['eterra_export'], 'CTRL')
            if self.eterra_control_export is None:
                raise ValueError("Failed to read CTRL tab from eTerra export")
            self.eterra_control_export = clean_eterra_control_export(self.eterra_control_export)
            if self.debug_dir:
                self.eterra_control_export.to_csv(f"{self.debug_dir}/eterra_control_export.csv", index=False)

            print(f"Loading setpoint control export from {self.data_dir / self.required_files['eterra_export']}")
            self.eterra_setpoint_control_export = read_habdde_tab_into_df(self.data_dir / self.required_files['eterra_export'], 'SETPNT')
            if self.eterra_setpoint_control_export is None:
                raise ValueError("Failed to read SETPNT tab from eTerra export")
            self.eterra_setpoint_control_export = clean_eterra_setpoint_control_export(self.eterra_setpoint_control_export)
            if self.debug_dir:
                self.eterra_setpoint_control_export.to_csv(f"{self.debug_dir}/eterra_setpoint_control_export.csv", index=False)

            # Get just the common columns from point and analog and concatenate them together, sort by GenericPointAddress
            common_columns = [  'GenericPointAddress', 'CASDU', 'Protocol', 'RTU', 'Card',
                                'RTUAddress', 'RTUId', 'IOA2', 'IOA1', 'IOA', 'PointId', 
                                'GenericType', 'DeviceType', 'DeviceName', 'DeviceId', 
                                'Sub', 'Word', 'eTerraKey', 'eTerraAlias', 'Controllable']
            
            print("Combining point and analog exports...")
            eterra_points_common_cols = self.eterra_point_export[common_columns]
            eterra_analogs_common_cols = self.eterra_analog_export[common_columns]
            eterra_export = pd.concat([eterra_points_common_cols, eterra_analogs_common_cols], ignore_index=True)
            eterra_export = eterra_export.sort_values(by='GenericPointAddress')
            if self.debug_dir:
                eterra_export.to_csv(f"{self.debug_dir}/eterra_export.csv", index=False)

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

        
        # Merge with compare report
        merged = pd.merge(
            merged,
            self.compare_alarms,
            left_on=['eTerraAlias'],
            right_on=['CompAlarmEterraAlias'],
            how='left'
        )


        # Control information needs to be joined differently as only a few key fields are requried for each associated control

        # Merge with controls test

        # Merge with manual commissioning

        if self.debug_dir:
            merged.to_csv(f"{self.debug_dir}/merged.csv", index=False)
        
        return merged

    def generate_report(self, rtu_id: Optional[str] = None, substation: Optional[str] = None):
        """Generate report for specified RTU or substation."""
        if not self.validate_data_files():
            sys.exit(1)
            
        self.load_data()
        self.debug_print_dataframes()
        merged_data = self.merge_data()
        
        # Filter data based on criteria
        if rtu_id:
            filtered_data = filter_data_by_rtu(merged_data, rtu_id)
        elif substation:
            filtered_data = filter_data_by_substation(merged_data, substation)
        else:
            filtered_data = merged_data
        
        # Create report sections
        points_section = create_points_section(filtered_data)
        analogs_section = create_analogs_section(filtered_data)
        
        # Combine sections
        report = pd.concat([points_section, analogs_section], ignore_index=True)
        
        # Save report
        output_path = self.data_dir / f"rtu_report_{rtu_id or substation or 'all'}.xlsx"
        save_report(report, output_path)
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