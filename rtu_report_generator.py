#!/usr/bin/env python3

import os
import sys
import pandas as pd
import sqlite3
from rich import print
from rich.progress import Progress

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
    save_reports,
    generate_defect_report_in_excel
)
from defect_reports import (
    generate_defect_report_by_name
)
from local_query.po_query import check_if_component_alias_exists_in_poweron, checkIfComponentAliasInScanPointComponents
import configparser

from openpyxl.utils import get_column_letter
from pylib3i.habdde import remove_dummy_points_from_df, get_dummy_points_from_df, read_habdde_card_tab_into_df

CONFIG_FILE = 'rtu_reports.ini'
DEFAULT_DATA_DIR = 'rtu_report_data'
DEFAULT_CONFIG_DIR = 'rtu_report_config'
DEFAULT_OUTPUT_DIR = 'reports'

def load_config(config_path=CONFIG_FILE):
    config = configparser.ConfigParser()
    config.read(config_path)
    return config

# def add_control_stats_to_each_row(row):
#     # Count how many controls are in the row - one for each of Ctrl1Name, Ctrl2Name that are not empty
#     control_count = 0
#     control_missing_links = 0
#     controls_commissioned = 0
#     for ctrlnum in range(1,3):
#         ctrl_name = f'Ctrl{ctrlnum}Name'
#         if row[ctrl_name] != '':
#             control_count += 1
#         if row[f'Ctrl{ctrlnum}ConfigHealth'] == '' or row[f'Ctrl{ctrlnum}ConfigHealth'].isna():
#             control_missing_links += 1
#         if row[f'Ctrl{ctrlnum}Commissioned'] == 'OK':
#             controls_commissioned += 1

#     # Add the control stats to the row
#     row['Num Controls'] = control_count
#     row['Num Controls Missing Links'] = control_missing_links
#     row['Num Controls Commissioned'] = controls_commissioned
#     row['Percent Controls Commissioned'] = (controls_commissioned / control_count) * 100 if control_count > 0 else 0
#     return row

# def add_alarm_stats_to_each_row(row):
#     # Count how many alarms are in the row - one for each of Alarm1Name, Alarm2Name that are not empty
#     alarm_count = 0
#     alarms_matched = 0
#     for alarmnum in range(0,4):
#         alarm_name = f'Alarm{alarmnum}_eTerraMessage'
#         if row[alarm_name] != '':
#             alarm_count += 1
#         if row[f'Alarm{alarmnum}_MessageMatch'] == True:
#             alarms_matched += 1

#     # Add the alarm stats to the row
#     row['Num Alarms'] = alarm_count
#     row['Num Alarms Matched'] = alarms_matched
#     row['Percent Alarms Matched'] = (alarms_matched / alarm_count) * 100 if alarm_count > 0 else 0
#     return row

class RTUReportGenerator:
    def __init__(self, config_path=CONFIG_FILE, data_dir="", write_cache=False, read_cache=False):
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
            'controls_db': "controls.db",
            'alarm_mismatch_manual_actions': "AlarmMismatchManualActions.xlsx"
        }
            
        if Path(config_path).exists():
            config = configparser.ConfigParser()
            config.read(config_path)
            if 'Files' in config:
                for key in self.required_files:
                    if key in config['Files']:
                        self.required_files[key] = config['Files'][key]

        # set the data cache db
        if write_cache or read_cache:
            # get the data cache directory from the config file
            self.data_cache_dir = Path(self.config['Paths']['data_cache_dir'])

            if self.data_cache_dir is None or self.data_cache_dir == "":
                print("Error: Data cache directory is not specified in the config file")
                sys.exit(1)

            if not self.data_cache_dir.exists():
                print(f"Error: Data cache directory does not exist: {self.data_cache_dir}")
                sys.exit(1)

            self.data_cache_db = self.data_cache_dir / 'data_cache.db'
        else:
            self.data_cache_db = None
            write_cache = False
            read_cache = False

        print(f"write_cache: {write_cache} read_cache: {read_cache}")
        print(f"data_cache_db: {self.data_cache_db}")

        
        # Initialize dataframes
        self.eterra_full_point_export = None
        self.eterra_point_export = None
        self.eterra_dummy_point_export = None
        self.eterra_analog_export = None
        self.eterra_control_export = None
        self.eterra_setpoint_control_export = None
        self.eterra_rtu_map = None
        self.eterra_export = None
        self.eterra_card_tab = None
        self.habdde_compare = None
        self.all_rtus = None
        self.controls_test = None
        self.iccp_compare = None
        self.compare_alarms = None
        self.manual_commissioning = None
        self.merged_data = None
        self.alarm_mismatch_manual_actions = None

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
    
    def write_data_cache(self):
        """Write the data cache to the database."""
        # write the merged data to the database
        if self.merged_data is None:
            print("No merged data to write to cache")
            return
        print(f"Writing data cache to {self.data_cache_db}")

        conn = sqlite3.connect(self.data_cache_db)
        self.merged_data.to_sql('merged_data', conn, if_exists='replace', index=False)
        conn.close()

    def read_data_cache(self, rtu_name: Optional[str] = None, substation: Optional[str] = None):
        """Read the data cache from the database."""
        # read the merged data from the database
        print(f" :mag_right: Reading data cache from {self.data_cache_db}")
        conn = sqlite3.connect(self.data_cache_db)
        self.merged_data = pd.read_sql_query('SELECT * FROM merged_data', conn)
        conn.close()
        print(f"Read {self.merged_data.shape[0]} rows from data cache")
        # filter the data using the rtu_name and substation if they are provided
        if rtu_name:
            self.merged_data = self.merged_data[self.merged_data['RTU'] == rtu_name]
        if substation:
            self.merged_data = self.merged_data[self.merged_data['Sub'] == substation]
        print(f"Filtered data to {self.merged_data.shape[0]} rows")

    def load_eterra_export(self):
        print(f" :arrow_right: Loading eTerra export from {self.data_dir / self.required_files['eterra_export']}")
        self.eterra_full_point_export = import_habdde_export_point_tab(self.data_dir / self.required_files['eterra_export'], self.debug_dir)
        self.eterra_point_export = remove_dummy_points_from_df(self.eterra_full_point_export)
        self.eterra_dummy_point_export = get_dummy_points_from_df(self.eterra_full_point_export)
        if self.debug_dir:
            self.eterra_dummy_point_export.to_csv(f"{self.debug_dir}/eterra_dummy_point_export.csv", index=False)
        # Create a map of RTU addresses and protocols from the eTerra export
        self.eterra_rtu_map = derive_rtu_addresses_and_protocols_from_eterra_export(self.eterra_point_export, self.debug_dir)

        print(f" :arrow_right: Loading analog export from {self.data_dir / self.required_files['eterra_export']}")
        self.eterra_analog_export = import_habdde_export_analog_tab(self.data_dir / self.required_files['eterra_export'], self.debug_dir)
        
        print(f" :arrow_right: Loading control export from {self.data_dir / self.required_files['eterra_export']}")
        self.eterra_control_export = import_habdde_export_control_tab(self.data_dir / self.required_files['eterra_export'], self.debug_dir)

    def add_no_input_controls(self):
        # Look for any controls that are not in the point export, then look for these as dummy points
        print(f" :arrow_right: Looking for controls that are not in the point export ... ", end="")
        no_input_controls = self.eterra_control_export[~self.eterra_control_export['eTerraAlias'].isin(self.eterra_point_export['eTerraAlias'])]
        # remove any rows that have PointId = "TAP" - these are dealt with through the TAP/TPC connection in the control load
        no_input_controls = no_input_controls[no_input_controls['PointId'] != "TAP"]
        print(f"found {no_input_controls.shape[0]} controls.")
        if self.debug_dir:
            no_input_controls.to_csv(f"{self.debug_dir}/no_input_controls.csv", index=False)

        print(f" :arrow_right: Looking for dummy points for no input controls ... ", end="")
        no_input_dummy_points = self.eterra_dummy_point_export[self.eterra_dummy_point_export['eTerraAlias'].isin(no_input_controls['eTerraAlias'])]
        print(f"found {no_input_dummy_points.shape[0]} dummy points.")
        if self.debug_dir:
            no_input_dummy_points.to_csv(f"{self.debug_dir}/no_input_dummy_points.csv", index=False)

        # make a copy of no_input_controls so we can create a vesion without eTerraAlias duplicates
        no_input_controls_deduped = no_input_controls.drop_duplicates(subset=['eTerraAlias'])
        # in the no_input_dummy_points dataframe, set the RTU to the RTU value from the corresponding row in no_input_controls_deduped
        no_input_dummy_points['RTU'] = no_input_dummy_points['eTerraAlias'].map(no_input_controls_deduped.set_index('eTerraAlias')['RTU'])
        # set the PowerOn Alias to the eTerraAlias
        no_input_dummy_points['PowerOn Alias'] = no_input_dummy_points['eTerraAlias']
        # get the PowerOn Alias Exists by querying
        no_input_dummy_points['PowerOn Alias Exists'] = no_input_dummy_points['PowerOn Alias'].apply(check_if_component_alias_exists_in_poweron)
        

        # Add the no_input_dummy_points to the self.eterra_point_export dataframe
        self.eterra_point_export = pd.concat([self.eterra_point_export, no_input_dummy_points], ignore_index=True)



        # # Get the list of no_input_controls that are not in the dummy points by comparing the original eTerraAlias values
        # alias_of_no_input_controls = no_input_controls['eTerraAlias'].tolist()
        # alias_of_no_input_dummy_points = no_input_dummy_points['eTerraAlias'].tolist()
        # no_input_controls_not_dummy_points = [alias for alias in alias_of_no_input_controls if alias not in alias_of_no_input_dummy_points]
        # print(f"found {len(no_input_controls_not_dummy_points)} controls that are not in the dummy points.")
        # if self.debug_dir:
        #     pd.DataFrame(no_input_controls_not_dummy_points, columns=['eTerraAlias']).to_csv(f"{self.debug_dir}/no_input_controls_not_dummy_points.csv", index=False)

        # # get the duplicates from no_input_controls
        # no_input_controls_duplicates = no_input_controls[no_input_controls.duplicated(subset=['eTerraAlias'])]
        # print(f"found {no_input_controls_duplicates.shape[0]} duplicate controls.")
        # if self.debug_dir:
        #     no_input_controls_duplicates.to_csv(f"{self.debug_dir}/no_input_controls_duplicates.csv", index=False)

    def load_eterra_setpoint_control_export(self):
        print(f" :arrow_right: Loading setpoint control export from {self.data_dir / self.required_files['eterra_export']}")
        self.eterra_setpoint_control_export = import_habdde_export_setpoint_control_tab(self.data_dir / self.required_files['eterra_export'], self.debug_dir)

    def load_eterra_card_tab(self):
        print(f" :arrow_right: Loading card tab from {self.data_dir / self.required_files['eterra_export']}")
        self.eterra_card_tab = read_habdde_card_tab_into_df(self.data_dir / self.required_files['eterra_export'], self.debug_dir)

    def create_base_eterra_export_by_combining_point_and_analog_exports(self):
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
        # TODO: consider if the analog df might have different columns - I dont' think it should
        #common_columns.extend([col for col in potentially_common_columns if col in self.eterra_analog_export.columns])

        print("Combining point and analog exports...")
        eterra_points_common_cols = self.eterra_point_export[common_columns]
        eterra_analogs_common_cols = self.eterra_analog_export[common_columns]
        self.eterra_export = pd.concat([eterra_points_common_cols, eterra_analogs_common_cols], ignore_index=True)
        self.eterra_export = self.eterra_export.sort_values(by='GenericPointAddress')
        if self.debug_dir:
            self.eterra_export.to_csv(f"{self.debug_dir}/eterra_export.csv", index=False)

    ''' ********** filter_eterra_export_by_rtu_name_or_substation ********** '''
    def filter_eterra_export_by_rtu_name_or_substation(self, rtu_name: Optional[str] = None, substation: Optional[str] = None):
        # Filter by RTU name if provided - because we build the whole report off this list this is the only filter we need
        if rtu_name:
            print(f"Filtering by RTU name: {rtu_name}")
            self.eterra_export = self.eterra_export[self.eterra_export['RTU'] == rtu_name]
        elif substation:
            print(f"Filtering by substation: {substation}")
            self.eterra_export = self.eterra_export[self.eterra_export['Sub'] == substation]

    ''' ********** load_habdde_compare ********** '''
    def load_habdde_compare(self):
        print(f" :arrow_right: Loading habdde compare from {self.data_dir / self.required_files['habdde_compare']}")
        self.habdde_compare = pd.read_csv(self.data_dir / self.required_files['habdde_compare'], low_memory=False)
        self.habdde_compare = clean_habdde_compare(self.habdde_compare)
        if self.debug_dir:
            self.habdde_compare.to_csv(f"{self.debug_dir}/habdde_compare.csv", index=False)

    ''' ********** load_poweron_data ********** '''
    def load_poweron_data(self):
        print(f" :arrow_right: Loading poweron data from {self.data_dir / self.required_files['all_rtus']}")
        self.all_rtus = pd.read_csv(self.data_dir / self.required_files['all_rtus'], low_memory=False)
        self.all_rtus = clean_all_rtus(self.all_rtus)
        if self.debug_dir:
            self.all_rtus.to_csv(f"{self.debug_dir}/all_rtus.csv", index=False)

    ''' ********** load_controls_auto_test_results ********** '''
    def load_controls_auto_test_results(self):
        print(f" :arrow_right: Loading controls auto test results from {self.data_dir / self.required_files['controls_test']}")
        self.controls_test = pd.read_csv(self.data_dir / self.required_files['controls_test'])
        self.controls_test = clean_controls_test(self.controls_test, self.eterra_rtu_map)
        if self.debug_dir:
            self.controls_test.to_csv(f"{self.debug_dir}/controls_test.csv", index=False)

    ''' ********** load_compare_alarms ********** '''
    def load_compare_alarms(self):
        print(f" :arrow_right: Loading compare alarms from {self.data_dir / self.required_files['compare_alarms']}")
        self.compare_alarms = pd.read_excel(self.data_dir / self.required_files['compare_alarms'], sheet_name='Event Detail')
        self.compare_alarms = clean_compare_alarms(self.compare_alarms)
        if self.debug_dir:
            self.compare_alarms.to_csv(f"{self.debug_dir}/compare_alarms.csv", index=False)

    ''' ********** load_manual_commissioning_results ********** '''
    def load_manual_commissioning_results(self):
        print(f" :arrow_right: Loading manual commissioning results from {self.data_dir / self.required_files['controls_db']}")
        conn = sqlite3.connect(self.data_dir / self.required_files['controls_db'])
        self.manual_commissioning = pd.read_sql_query("SELECT * FROM test_results", conn)
        conn.close()
        self.manual_commissioning = clean_manual_commissioning(self.manual_commissioning)
        if self.debug_dir:
            self.manual_commissioning.to_csv(f"{self.debug_dir}/manual_commissioning.csv", index=False)

    ''' ********** add_control_info_to_input_rows_in_eterra_export ********** '''
    def add_control_info_to_input_rows_in_eterra_export(self):
        print("Adding control info to input rows in eTerra export...")
        self.eterra_export = add_control_info_to_eterra_export(self.eterra_export, self.eterra_control_export, self.eterra_setpoint_control_export, self.all_rtus, self.controls_test, self.manual_commissioning)
        if self.debug_dir:
            self.eterra_export.to_csv(f"{self.debug_dir}/eterra_export_with_control_info.csv", index=False)


    ''' ********** load_alarm_mismatch_manual_actions ********** '''
    def load_alarm_mismatch_manual_actions(self):
        # if the alarm_mismatch_manual_actions file exists, load it
        if (self.data_dir / self.required_files['alarm_mismatch_manual_actions']).exists():
            print(f" :arrow_right: Loading alarm mismatch manual actions from {self.data_dir / self.required_files['alarm_mismatch_manual_actions']}")
            self.alarm_mismatch_manual_actions = pd.read_excel(self.data_dir / self.required_files['alarm_mismatch_manual_actions'], sheet_name='Sheet1')
            if self.debug_dir:
                self.alarm_mismatch_manual_actions.to_csv(f"{self.debug_dir}/alarm_mismatch_manual_actions.csv", index=False)
        else:
            print(f"Warning: alarm mismatch manual actions file does not exist: {self.data_dir / self.required_files['alarm_mismatch_manual_actions']}")



    ''' ********** load_data ********** '''
    def load_data(self, rtu_name: Optional[str] = None, substation: Optional[str] = None):
        """Load all source data into dataframes."""
        try:
            self.load_eterra_export() # creates eterra_full_point_export, eterra_point_export, eterra_dummy_point_export, eterra_analog_export, eterra_control_export
            self.load_eterra_setpoint_control_export() # creates eterra_setpoint_control_export
            self.add_no_input_controls() # creates no_input_controls, no_input_dummy_points, no_input_controls_not_dummy_points
            self.create_base_eterra_export_by_combining_point_and_analog_exports() # creates eterra_export - a dataframe with all the point and analog data and only the common columns
            self.filter_eterra_export_by_rtu_name_or_substation(rtu_name, substation) # creates eterra_export_filtered - a filtered version of eterra_export that is filtered by rtu_name or substation
            self.load_habdde_compare()
            self.load_poweron_data()
            self.load_controls_auto_test_results()
            self.load_compare_alarms()
            self.load_manual_commissioning_results()
            self.add_control_info_to_input_rows_in_eterra_export()
            self.load_alarm_mismatch_manual_actions()

        except Exception as e:
            print(f"Error loading data: {str(e)}")
            print(f"Error type: {type(e)}")
            import traceback
            print(traceback.format_exc())
            sys.exit(1)

    def merge_data(self) -> pd.DataFrame:
        """Merge all data sources into a single dataframe."""

        merged = self.merge_eterra_export_with_habdde_compare()
        merged = self.merge_all_rtus_data(merged)
        merged = self.merge_iccp_compare_data(merged)
        merged = self.merge_compare_alarms_data(merged)
        merged = self.merge_control_data(merged)
        merged = self.merge_alarm_mismatch_manual_actions(merged)
        merged = self.add_derived_columns(merged)

        if self.debug_dir:
            print(f" :arrow_right: Writing merged data to {self.debug_dir}/merged.csv")
            merged.to_csv(f"{self.debug_dir}/merged.csv", index=False)
        
        return merged
    
    #####################################################################
    #                                                                    
    #                      Add Derived Columns                           
    #                                                                    
    #  - Adds Type column that flags dummy rows as 'DUMMY'               
    #  - Creates Ignore column based on IGNORE flags and OLD_DATA        
    #  - Adds RTUComms column for RTU devices without 'LDC'              
    #                                                                    
    #####################################################################
    def add_derived_columns(self, merged: pd.DataFrame) -> pd.DataFrame:
        print(f" ✅ Adding derived columns to merged data...")
        # first get a Type column that also flags the dummy rows are DUMMY, Get the value of GenericType unless the RTUId = '(€€€€€€€€:)'
        merged['Type'] = merged.apply(lambda row: 'DUMMY' if row['RTUId'] == '(€€€€€€€€:)' else row['GenericType'], axis=1)
        # now make an ignore column that is TRUE if any of IGNORE_RTU, IGNORE_POINT, OLD_DATA are TRUE
        merged['Ignore'] = merged.apply(lambda row: True if (row['IGNORE_RTU'] == True or row['IGNORE_POINT'] == True or row['OLD_DATA'] == True ) else False, axis=1)
        # We want to add a new column 'RTUComms' to the df that is True if the DeviceType is 'RTU and the eTerraAlias does not contain 'LDC'
        merged['RTUComms'] = merged.apply(lambda row: True if row['DeviceType'] == 'RTU' and 'LDC' not in row['eTerraAlias'] else False, axis=1)

        # Create 2 new columsn eTerraAliasExistsInPO, eTerraAliasLinkedToSCADA
        # For every value in the eTerraAlias column, set the value of these new fields: 1/0 if that comp exists in PowerOn, 1/0 if that comp is scada linked in PowerOn
        # we will use 2 functions to do this
        def get_poweron_alias_exists(alias):
            if alias is None or alias == "":
                return 0
            exists = check_if_component_alias_exists_in_poweron(alias)
            return 1 if exists is not None and exists else 0
        
        def get_poweron_alias_linked_to_scada(alias):
            if alias is None or alias == "":
                return 0
            linked = checkIfComponentAliasInScanPointComponents(alias)
            return  1 if linked is not None and linked else 0
        
        # 1. get_poweron_alias_exists(eTerraAlias) - returns 1/0 if the eTerraAlias exists in PowerOn
        # 2. get_poweron_alias_linked_to_scada(eTerraAlias) - returns 1/0 if the eTerraAlias is scada linked in PowerOn
        merged['eTerraAliasExistsInPO'] = merged.apply(lambda row: get_poweron_alias_exists(row['eTerraAlias']), axis=1)
        merged['eTerraAliasLinkedToSCADA'] = merged.apply(lambda row: get_poweron_alias_linked_to_scada(row['eTerraAlias']), axis=1)

        # And do the same for ICCP Alias
        merged['ICCPAliasExists'] = merged.apply(lambda row: get_poweron_alias_exists(row['ICCP_Alias']), axis=1)
        merged['ICCPAliasLinkedToSCADA'] = merged.apply(lambda row: get_poweron_alias_linked_to_scada(row['ICCP_Alias']), axis=1)

        return merged


    def merge_eterra_export_with_habdde_compare(self) -> pd.DataFrame:
        # Merge eTerra export with habdde compare
        print(f" 🧠 Merging eTerra export with habdde compare on {self.eterra_export.shape[0]} rows")
        merged = pd.merge(
            self.eterra_export,
            self.habdde_compare,
            on=['GenericPointAddress'],
            how='left'
        )
        # drop the HabCompKey column
        merged = merged.drop(columns=['HabCompKey'])
        print(f" ✅ Merged eTerra export with habdde compare on {merged.shape[0]} rows")
        return merged

    
    def merge_all_rtus_data(self, merged: pd.DataFrame) -> pd.DataFrame:
        """Merge with all RTUs data."""
        print(" 🧠 Merging with all RTUs ... ")
        # Merge with all RTUs
        merged = pd.merge(
            merged,
            self.all_rtus,
            on=['GenericPointAddress'],
            how='left'
        )
        print(f" ✅ Merged with all RTUs on {merged.shape[0]} rows")

        # remove the TC Action column - we don't want this for input points but we'll query for it later when we add the control info
        merged = merged.drop(columns=['TC Action'])

        return merged
    
    def merge_iccp_compare_data(self, merged: pd.DataFrame) -> pd.DataFrame:
        # Merge with ICCP compare
        # merged = pd.merge(
        #     merged,
        #     self.iccp_compare,
        #     on=['RTU', 'Card', 'Word'],
        #     how='left'
        # )

        # print("Merge report columns:")
        # print(merged.columns)
        pass
        return merged

    def merge_compare_alarms_data(self, merged: pd.DataFrame) -> pd.DataFrame:
        print(" 🧠 Merging with alarm compare report ... ")
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
            'CompAlarmAlarmZoneMatch',
            'CompAlarmTemplateAlias',
            'CompAlarmTemplateType',
            'CompAlarmStateIndex'
        ]
        
        # Only include columns that exist in the merged dataframe
        point_related_columns = [col for col in desired_columns if col in available_columns]
        
        if not point_related_columns:
            print("Warning: No matching columns found for point_related_columns")
            point_related_df = self.compare_alarms.copy()
        else:
            # Sort by CompAlarmPOStatus so 'Matched' comes first
            point_related_df = self.compare_alarms[point_related_columns].sort_values(
                by=['CompAlarmEterraAlias', 'CompAlarmPOStatus'],
                ascending=[True, False]  # False puts 'Matched' first
            )
            # Keep first row for each eTerraAlias (which will be 'Matched' if exists)
            point_related_df = point_related_df.groupby('CompAlarmEterraAlias').first().reset_index()

        print(f"  🧠 Merging with Component level information for {point_related_df.shape[0]} rows")
        #1.b) merge the point related df with the compare alarms df
        merged = pd.merge(
            merged,
            point_related_df,
            left_on=['eTerraAlias'],
            right_on=['CompAlarmEterraAlias'],
            how='left'
        )
        print(f"  ✅ Merged with Component level information into {merged.shape[0]} rows")

        # Debug the merge - we are getting duplicates in the merged df
        print(f"    Merged df has {merged.shape[0]} rows")
        number_of_duplicates = merged.duplicated().sum()
        print(f"    Number of duplicates: {number_of_duplicates}")


        #1.c de-duplicate the merged df
        print(f"    De-duplicating merged data...")
        merged = merged.drop_duplicates()

        print(f"    After de-duplication, merged df has {merged.shape[0]} rows")
        #if we had duplciates, exit
        if number_of_duplicates > 0:
            print(f" ❌ Duplicates were found, exiting")
            sys.exit(1)

        print(f"  🧠 Initializing alarm related columns...")
        #1.d) add the alarm related columns into Alarm<value>_eTerraMessage and Alarm<value>_POMessage, and Alarm<value>_MessageMatch
        # Initialize empty columns for all possible alarm values
        for value in range(4):
            merged[f'Alarm{value}_eTerraMessage'] = None
            merged[f'Alarm{value}_POMessage'] = None 
            merged[f'Alarm{value}_MessageMatch'] = None

        # For each row in merged, find matching rows in compare_alarms and populate corresponding alarm columns
        # print(f"   Adding alarm compare related columns to merged data...")
        with Progress() as progress:
            task = progress.add_task("    Adding alarm compare related columns to merged data...", total=merged.shape[0])
            for idx, row in merged.iterrows():
                matching_alarms = self.compare_alarms[
                    self.compare_alarms['CompAlarmEterraAlias'] == row['eTerraAlias']
                ]
                
                # For each matching alarm row, populate the corresponding alarm columns based on CompAlarmValue
                num_alarms = 0
                num_alarms_matched = 0
                for _, alarm_row in matching_alarms.iterrows():
                    value = alarm_row['CompAlarmValue']
                    merged.at[idx, f'Alarm{value}_eTerraMessage'] = alarm_row['CompAlarmeTerraAlarmMessage']
                    merged.at[idx, f'Alarm{value}_POMessage'] = alarm_row['CompAlarmPOAlarmMessage']
                    merged.at[idx, f'Alarm{value}_MessageMatch'] = alarm_row['CompAlarmAlarmMessageMatch']
                    num_alarms += 1
                    if alarm_row[f'CompAlarmAlarmMessageMatch'] == '1':
                        num_alarms_matched += 1
                    
                merged.at[idx, 'NumAlarms'] = num_alarms
                merged.at[idx, 'NumAlarmsMatched'] = num_alarms_matched
                merged.at[idx, 'PercentAlarmsMatched'] = num_alarms_matched / num_alarms if num_alarms > 0 else 0
                progress.update(task, advance=1)

        print(f"  ✅ Added alarm compare related columns to merged data on {merged.shape[0]} rows")
        return merged
    
    def merge_alarm_mismatch_manual_actions(self, merged: pd.DataFrame) -> pd.DataFrame:
        # Merge with alarm mismatch manual actions
        print(" 🧠 Merging with alarm mismatch manual actions ... ")
        if self.alarm_mismatch_manual_actions is not None:

            # rename the columns to have better names
            self.alarm_mismatch_manual_actions.rename(columns={
                'eTerra Alias': 'eTerraAlias',
                'Comments on missmatch': 'AlarmMismatchComment',
                'TemplateAlias': 'AlarmMismatchTemplateAlias'
            }, inplace=True)

            merged = pd.merge(
                merged,
                self.alarm_mismatch_manual_actions,
                on=['eTerraAlias'],
                how='left'
            )

            # set any na values to '' for the 2 columns that are added
            merged['AlarmMismatchComment'] = merged['AlarmMismatchComment'].fillna('')
            merged['AlarmMismatchTemplateAlias'] = merged['AlarmMismatchTemplateAlias'].fillna('')

        print(f" ✅ Merged with alarm mismatch manual actions on {merged.shape[0]} rows")
        return merged
    
    def merge_control_data(self, merged: pd.DataFrame) -> pd.DataFrame:
        print(" 🧠 Adding control info to merged data...")
        # Control information needs to be joined differently as only a few key fields are requried for each associated control
        # For each control we need to get the following:
        # 1. match status from habdde compare
        # 2. config health from poweron
        # 3. auto test status from controls test
        # 4. test result from manual commissioning
        # 5. telecontrol action from poweron


        # Get the columns that exist in the merged dataframe
        available_columns = self.eterra_export.columns.tolist()

        # Add the columns we need to the merged dataframe, but insert them after the CtrlNAddr column
        for ctrl_num in [1, 2]:
            merged.insert(merged.columns.get_loc(f'Ctrl{ctrl_num}Addr') + 1, f'Ctrl{ctrl_num}MatchStatus', None)
            merged.insert(merged.columns.get_loc(f'Ctrl{ctrl_num}MatchStatus') + 1, f'Ctrl{ctrl_num}ConfigHealth', None)
            merged.insert(merged.columns.get_loc(f'Ctrl{ctrl_num}ConfigHealth') + 1, f'Ctrl{ctrl_num}AutoTestStatus', None)
            merged.insert(merged.columns.get_loc(f'Ctrl{ctrl_num}AutoTestStatus') + 1, f'Ctrl{ctrl_num}TestResult', None)
            merged.insert(merged.columns.get_loc(f'Ctrl{ctrl_num}TestResult') + 1, f'Ctrl{ctrl_num}TelecontrolAction', None)
            merged.insert(merged.columns.get_loc(f'Ctrl{ctrl_num}TelecontrolAction') + 1, f'Ctrl{ctrl_num}VisualCheckResult', None)
            merged.insert(merged.columns.get_loc(f'Ctrl{ctrl_num}VisualCheckResult') + 1, f'Ctrl{ctrl_num}ControlSentResult', None)
            merged.insert(merged.columns.get_loc(f'Ctrl{ctrl_num}ControlSentResult') + 1, f'Ctrl{ctrl_num}Comments', None)
        
        # Go through every row that has at least one control
        # Pre-process the lookups into dictionaries for faster access
        habdde_compare_dict = self.habdde_compare.set_index('GenericPointAddress').to_dict('index')
        poweron_dict = self.all_rtus.set_index('GenericPointAddress').to_dict('index')
        controls_test_dict = self.controls_test.set_index('GenericPointAddress').to_dict('index')
        
        # Create dictionaries for manual commissioning lookups
        manual_commission_dict = {}
        visual_check_dict = {}
        control_sent_dict = {}
        
        for _, row in self.manual_commissioning.iterrows():
            addr = row['CommissioningControlAddress']
            test = row['CommissioningTestName']
            if test == 'Action Verified':
                manual_commission_dict[addr] = row
            elif test == 'Visual Check':
                visual_check_dict[addr] = row
            elif test == 'Control Sent':
                control_sent_dict[addr] = row

        # Process all rows at once using vectorized operations where possible
        controllable_rows = merged[merged['Controllable'] == '1']
        
        for idx, row in controllable_rows.iterrows():
            num_controls = 0
            num_controls_matched = 0
            num_controls_config_good = 0
            num_controls_commission_ok = 0
            num_controls_all_commission_ok = 0

            for ctrl_num in [1, 2]:
                ctrl_addr = row[f'Ctrl{ctrl_num}Addr']
                if ctrl_addr != '':
                    num_controls += 1
                    num_controls_matched += 1

                    # Lookup data from dictionaries
                    habdde_compare_info = habdde_compare_dict.get(ctrl_addr, {})
                    poweron_info = poweron_dict.get(ctrl_addr, {})
                    controls_test_info = controls_test_dict.get(ctrl_addr, {})
                    manual_commission_info = manual_commission_dict.get(ctrl_addr, {})
                    visual_check_info = visual_check_dict.get(ctrl_addr, {})
                    control_sent_info = control_sent_dict.get(ctrl_addr, {})

                    if poweron_info and poweron_info.get('ConfigHealth') == 'GOOD':
                        num_controls_config_good += 1

                    # Check manual commissioning results
                    if all(info.get('CommissioningResult') == 'OK' for info in 
                          [visual_check_info, control_sent_info, manual_commission_info]):
                        num_controls_all_commission_ok += 1

                    if manual_commission_info.get('CommissioningResult') == 'OK':
                        num_controls_commission_ok += 1

                    # Update control columns
                    merged.at[idx, f'Ctrl{ctrl_num}MatchStatus'] = habdde_compare_info.get('HbddeCompareStatus')
                    merged.at[idx, f'Ctrl{ctrl_num}ConfigHealth'] = poweron_info.get('ConfigHealth')
                    merged.at[idx, f'Ctrl{ctrl_num}AutoTestStatus'] = controls_test_info.get('AutoTestResult')
                    merged.at[idx, f'Ctrl{ctrl_num}TestResult'] = manual_commission_info.get('CommissioningResult')
                    merged.at[idx, f'Ctrl{ctrl_num}VisualCheckResult'] = visual_check_info.get('CommissioningResult')
                    merged.at[idx, f'Ctrl{ctrl_num}ControlSentResult'] = control_sent_info.get('CommissioningResult')
                    merged.at[idx, f'Ctrl{ctrl_num}TelecontrolAction'] = str(poweron_info.get('TC Action', ''))
                    
                    # Combine comments
                    comments = ' '.join(filter(None, [
                        visual_check_info.get('CommissioningComments', ''),
                        control_sent_info.get('CommissioningComments', ''),
                        manual_commission_info.get('CommissioningComments', '')
                    ])).strip()
                    merged.at[idx, f'Ctrl{ctrl_num}Comments'] = comments

            # Update summary columns
            merged.at[idx, 'NumControls'] = num_controls
            merged.at[idx, 'NumControlsMatched'] = num_controls_matched
            merged.at[idx, 'NumControlsConfigGood'] = num_controls_config_good
            merged.at[idx, 'NumControlsCommissionOk'] = num_controls_commission_ok
            merged.at[idx, 'NumControlsAllCommissionOk'] = num_controls_all_commission_ok
            merged.at[idx, 'PercentControlsMatched'] = num_controls_matched / num_controls if num_controls > 0 else 0
            merged.at[idx, 'PercentControlsConfigGood'] = num_controls_config_good / num_controls if num_controls > 0 else 0
            merged.at[idx, 'PercentControlsCommissionOk'] = num_controls_commission_ok / num_controls if num_controls > 0 else 0
            merged.at[idx, 'PercentControlsAllCommissionOk'] = num_controls_all_commission_ok / num_controls if num_controls > 0 else 0

        print(f" ✅ Added control info to merged data on {merged.shape[0]} rows")
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

        print(" 🧠 Adding issue report flags to the merged data...")
        # Convert some columns to boolean if not already
        merged_data = merged_data.reset_index(drop=True)

        # Handle NaN values before converting to bool
        merged_data['PowerOn Alias Exists'] = merged_data['PowerOn Alias Exists'].fillna(0).astype(int).astype(bool)
        merged_data['IGNORE_RTU'] = merged_data['IGNORE_RTU'].fillna(0).astype(int).astype(bool) 
        merged_data['IGNORE_POINT'] = merged_data['IGNORE_POINT'].fillna(0).astype(int).astype(bool)
        merged_data['OLD_DATA'] = merged_data['OLD_DATA'].fillna(0).astype(int).astype(bool)

        # HACK - remove RTU MICR4 from the data
        merged_data = merged_data[merged_data['RTU'] != 'MICR4']
        print(f"Removed RTU MICR4 from the data, now have {merged_data.shape[0]} rows")

        reports_list = [
            'Report1',
            'Report2',
            'Report3',
            'Report4',
            'Report5',
            'Report6',
            'Report7',
            'Report8',
            'Report9',
            'Report10',
            'Report11',
            'Report12',
            'Report13',
            'Report14',
            'ReportANY'
        ]

        for report in reports_list:
            merged_data = generate_defect_report_by_name(merged_data, report)
        print(f" ✅ Added issue report flags to the merged data on {merged_data.shape[0]} rows")
        return merged_data
    
    def generate_mk2a_card_report(self):
        """Generate the MK2A card report."""
        # First get a list of the unique RTU Name/Card pairs for all MK2A cards
        mk2a_cards = self.merged_data[self.merged_data['Protocol'] == 'MK2A'].copy()
        mk2a_cards = mk2a_cards[['RTU', 'Card']].drop_duplicates()
        # print(f"Found {mk2a_cards.shape[0]} unique RTU/Card pairs for MK2A cards")

        # Now check that every RTU/Card pair has at least one corresponding record in the merged_data dataframe, PO section, that is a good config health
        # print("="*80)
        # print(f"Checking {mk2a_cards.shape[0]} RTU/Card pairs for MK2A cards using point data")
        # for index, row in mk2a_cards.iterrows():
        #     rtu = row['RTU']
        #     card = row['Card']
            
        #     if card in ['251', '252', '253', '254' '119', '']:
        #         # ignore these cards as they are internal to the RTU
        #         continue
        #     valid_rows_for_card = self.merged_data[(self.merged_data['PO_RTU'] == (rtu + '_RTU')) & (self.merged_data['PO_Card'] == card) & (self.merged_data['ConfigHealth'] == 'GOOD')]

        #     if valid_rows_for_card.shape[0] == 0:
        #         print(f"RTU/Card pair {rtu}/{card} has no matching record in PowerOn with a good config health")

        # print("="*80)
        # print("")

        print("="*80)
        print(f"Checking {mk2a_cards.shape[0]} RTU/Card pairs for MK2A cards using card tab")
        # Load the eTerra card tab
        result_list = []
        self.eterra_card_tab = read_habdde_card_tab_into_df(self.data_dir / self.required_files['eterra_export'])
        # go through each card in the eTerra card tab, and check if a corresponding record exists in the merged_data dataframe, PO section, that is a good config health
        for index, row in self.eterra_card_tab.iterrows():
            rtu = row['rtu']
            # if rtu != 'CURR':
            #     continue
            protocol = row['protocol']
            if protocol != 'MK2A':
                continue
            card_type = row['CASDU']
            if card_type != '1':
                continue
            card = row['card']
            if card in ['250', '251', '252', '253', '254', '119', '']:
                continue
            valid_rows_for_card = self.merged_data[(self.merged_data['PO_RTU'] == (rtu + '_RTU')) & (self.merged_data['PO_Card'] == card) & (self.merged_data['ConfigHealth'] == 'GOOD')]
            if valid_rows_for_card.shape[0] == 0:
                print(f"RTU/Card pair {rtu}/{card} has no matching record in PowerOn with a good config health")
                result_list.append(f"{rtu},{card}")
        print("="*80)
        print("")

        # output the result list to a csv file, with header RTU, Card
        with open(self.output_dir / 'mk2a_missing_card_report.csv', 'w') as f:
            f.write('RTU,Card\n')
            for result in result_list:
                f.write(f"{result}\n")
        
    def generate_statistics(self, merged_data: pd.DataFrame):
        """Generate statistics for the merged data."""
        print("Generating statistics for the merged data...")
        # split the data by type of points in the merged data
        SD_points = merged_data[merged_data['GenericType'] == 'SD']
        DD_points = merged_data[merged_data['GenericType'] == 'DD']
        A_points = merged_data[merged_data['GenericType'] == 'A']
        DUMMY_points = merged_data[merged_data['RTUId'] == '(€€€€€€€€:)']

        # The number of controls is the summation of the nnumber of controls in each row
        Control_count = merged_data[merged_data['DeviceType'] != 'RTU']['NumControls'].sum()
        Simple_commissioned_count = merged_data[merged_data['DeviceType'] != 'RTU']['NumControlsCommissionOk'].sum()
        All_commissioned_count = merged_data[merged_data['DeviceType'] != 'RTU']['NumControlsAllCommissionOk'].sum()
        print(f"Total number of controls: {Control_count}")
        print(f"Total number of simple commissioned controls: {Simple_commissioned_count} ({(Simple_commissioned_count / Control_count * 100) if Control_count > 0 else 0.0:.1f}%)")
        print(f"Total number of all commissioned controls: {All_commissioned_count} ({(All_commissioned_count / Control_count * 100) if Control_count > 0 else 0.0:.1f}%)")

        num_points = merged_data.shape[0]
        num_SD_points = SD_points.shape[0]
        num_DD_points = DD_points.shape[0]
        num_A_points = A_points.shape[0]
        num_DUMMY_points = DUMMY_points.shape[0]

        # # Add control stats to each row
        # merged_data = merged_data.apply(add_control_stats_to_each_row, axis=1)
        # # Add alarm stats to each row
        # merged_data = merged_data.apply(add_alarm_stats_to_each_row, axis=1)
        
        print(f"{num_points} points in the merged data: {num_SD_points} SD, {num_DD_points} DD, {num_A_points} A, {num_DUMMY_points} DUMMY")



    ''' ================================
        Generate the report
        ================================ '''
    def generate_report(self, rtu_name: Optional[str] = None, substation: Optional[str] = None, write_cache: bool = False, read_cache: bool = False):
        """Generate report for specified RTU or substation."""
        if not self.validate_data_files():
            sys.exit(1)
        
        if read_cache:
            self.read_data_cache(rtu_name, substation)
        
        if not read_cache or self.merged_data.shape[0] == 0:
            self.load_data(rtu_name, substation)
            # self.debug_print_dataframes()
            self.merged_data = self.merge_data()
            
            if write_cache:
                self.write_data_cache()
        
        self.merged_data = self.add_issue_report_flags(self.merged_data)

        self.generate_defect_report(self.merged_data)

        # self.generate_mk2a_card_report()

        self.generate_statistics(self.merged_data)

        # Get list of RTUs in the filtered data
        rtus = self.merged_data['RTU'].unique()
        print(f"{len(rtus)} RTUs in the filtered data for the report")
        if (len(rtus) == 0):
            return False

        # We will create a report for each RTU in the filtered data
        reports = []
        for rtu in rtus:
            rtu_data = filter_data_by_rtu(self.merged_data, rtu)
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


    def generate_defect_report(self, merged_data: pd.DataFrame):
        """Generate a defect report for the merged data."""
        generate_defect_report_in_excel(merged_data, self.output_dir )
        return

    
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
    parser.add_argument("--writecache", action="store_true", help="Write the data cache to the database")
    parser.add_argument("--readcache", action="store_true", help="Read the data cache from the database")
    parser.add_argument("--data-dir", default=DEFAULT_DATA_DIR, help="Directory containing source files")
    parser.add_argument("--config-dir", default=DEFAULT_CONFIG_DIR, help="Directory containing config files")
    
    args = parser.parse_args()

    if args.writecache and args.readcache:
        print("Error: --writecache and --readcache cannot both be True")
        sys.exit(1)
    
    generator = RTUReportGenerator(args.config_dir + '/' + CONFIG_FILE, args.data_dir, args.writecache, args.readcache)
    generator.generate_report(args.rtu, args.substation, args.writecache, args.readcache)

if __name__ == "__main__":
    main() 