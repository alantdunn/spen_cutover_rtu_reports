# spen_cutover_rtu_reports

This project is desigend to take all the varouops outputs of the automated and manual checks for the SPEN SPT SCADA validation and create a spreadseet that has all the information in one place and with a simple intutitive summary of the RTUs status so that any issues cna ne easily fixed or sigend off as not required prior to go live.

We start by importing a set of reports

1. habdde eTerra export - this is the source list of data to be migrated.  There are a couple of exceptions where point types are not imoprted 1-1 (e.g. Load Reduction) and we will manually interpret these.
2. habdde compare report - this report compared the scada addresses loaded into PowerOn with the ones that existed in eTerra and gave a match report
3. all_rtus.csv - an export of the relevant scada data and linked components PowerOn
4. controls_test - the  resutls of the automated control tests on PowerOn
5. eterra_poweron_iccp_compare_report - the comaprison of the outputs recored from eterra and PowerOn using the same scada simulation (plus comparisons of pre and post load ICCP linked points in PowerOn )
6. manual commissioning results from the controls.db sqlite database

The utility will run in one of 3 modes

1. for a single RTU
2. for a single Substation (which may include more than 1 RTU)
3. for all RTUs

The source data for the selected RTU or RTUS is output into the target spreadsheet for reference and we  create a new report tab in teh following format

**Points**
Type:Input, SD/DD, SCADA Address (RTU-Card-Word-Size),eTerra Key, PowerOn Alias, ICCP Flag, Habdde Match Status, PowerOn Config Health Status, Alarm Match Status, Control Zone Status, Controllable Flag, Circuit Suggestion
[and , if the point has controls the controls are listed below the point]
Type:Control, DO, SCADA Address (RTU-Card-Word-ControlId),eTerra Key, PowerOn Alias, Auto Test status, Manual Test Status, Manual TestComments

**Controls with no feedback**
Type:Control, DO, SCADA Address (RTU-Card-Word-ControlId),eTerra Key, PowerOn Alias, Auto Test status, Manual Test Status, Manual Test Comments

**Analogs**
Type:Analog, A, SCADA Address (RTU-Card-Word-Size),eTerra Key, PowerOn Alias, ICCP Flag, Habdde Match Status, PowerOn Config Health Status, Analog Value Match Status
[and , if the analog has controls the controls are listed below the analog]
Type:Control, AO, SCADA Address (RTU-Card-Word-ControlId),eTerra Key, PowerOn Alias, Auto Test status, Manual Test Status, Manual Test Comments

## logic

First we check the command line parameters are valid and setup the criteria for the report
Check we can see all the reports we need
The spreadsheets are then imported into dataframes and the relvant data from the manual commissioning is read into a df as well
We then cut down the data to match our criteria
then we create a new excel file, put the relevant source data into  tabs then create our report in a new tab at the start of the workbook.
The report will be formatted to be clear and easy to read and filter the data as required.

## INSTRUCTIONS

**:warning:Check the database name in local_query/po_query.py:warning:**

`python rtu_report_generator.py --writecache`
`python rtu_report_generator.py --readcache`

`python rtu_report_generator.py --report-name defect_report --readcache`
`python rtu_report_generator.py --report-name mk2a_card_report --readcache`

`python rtu_report_generator.py --report-name defect_report --readcache`

*to do the excel defined one - check name of tab in ReportDefinitions.xlsx*

`python rtu_report_generator.py --report-name defect_report_all --readcache`

## Copy exisiting comments onto newly generated sheet

`cd utils`

`python utils/copy_comments_from_defect_report.py --oldfile rtu_report_data/from_spen/defect_report_all-05-MAY-2025-archive.xlsx --newfile reports/defect_report_all.xlsx`

`python utils/copy_comments_from_defect_report.py --oldfile rtu_report_data/from_spen/defect_report_all-14-MAY-2025-archive.xlsx --newfile reports/defect_report_all-22-MAY-2025.xlsx`

`python copy_comments_from_defect_report.py --oldfile ../rtu_report_data/from_spen/defect_report_all-26-MAY-2025-archive.xlsx --newfile ../reports/defect_report_all.xlsx`

/Users/alan/Documents/Projects/spen_cutover_rtu_reports/rtu_report_data/from_spen/defect_report_all-05-MAY-2025.xlsx
/Users/alan/Documents/Projects/spen_cutover_rtu_reports/rtu_report_data/from_spen/defect_report_all-05-MAY-2025-archive.xlsx
defect_report_all-14-MAY-2025-archive.xlsx
defect_report_all-22-MAY-2025.xlsx
