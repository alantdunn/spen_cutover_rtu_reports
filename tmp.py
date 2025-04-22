import pandas as pd
from tabulate import tabulate


# Define New Column mappings from the three tables

point_new_columns = ['eTerraKey', 'Sub', 'DeviceType', 'DeviceId', 'DeviceName', 'PointId',
       'PointName', 'eTerraZone', 'RTU', 'RTUAddress', 'Card', 'Word', 'CASDU',
       'IOA', 'IOA1', 'IOA2', 'Size', 'Inverted', 'Protocol', 'Controllable',
       'eTerraPtyType', 'RTUId', 'GenericPointAddress', 'GenericType']

analog_new_columns = ['eTerraKey', 'Sub', 'DeviceType', 'DeviceId', 'DeviceName', 'PointId',
       'LoReas', 'HiReas', 'eTerraZone', 'RTU', 'RTUAddress', 'Card', 'Word',
       'RawHigh', 'RawLow', 'EngHigh', 'EngLow', 'Protocol', 'ClmpDbnd',
       'PosPolar', 'NegPolar', 'Negate', 'RTUId', 'GenericPointAddress',
       'GenericType', 'CASDU', 'IOA', 'IOA1', 'IOA2']

control_new_columns = ['eTerraKey', 'Sub', 'DeviceType', 'DeviceId', 'DeviceName', 'PointId',
       'ControlId', 'RTU', 'RTUAddress', 'Card', 'Word', 'Parm1', 'Parm2',
       'Parm3', 'CtrlFunc', 'Protocol', 'RTUId', 'GenericPointAddress',
       'GenericType', 'CASDU', 'IOA', 'IOA1', 'IOA2']

# Create the superset of all new column names
all_columns = sorted(set(point_new_columns + analog_new_columns + control_new_columns))

# Create a DataFrame with indicators for each table
df = pd.DataFrame({
    'New Column': all_columns,
    'Point': ['Y' if col in point_new_columns else '' for col in all_columns],
    'Analog': ['Y' if col in analog_new_columns else '' for col in all_columns], 
    'Control': ['Y' if col in control_new_columns else '' for col in all_columns],
    'All': ['Y' if col in point_new_columns and col in analog_new_columns and col in control_new_columns else '' for col in all_columns]
})

# sort by All, Point, Analog, Control
df = df.sort_values(by='All', ascending=False)

print(tabulate(df, headers='keys', tablefmt='grid'))

# print the columns that are in all three tables as a comma separated list
print(','.join(df[df['All'] == 'Y']['New Column']))