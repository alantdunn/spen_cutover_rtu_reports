# We will read in  defect_report_all.xlsx and copy the coments to a new version of the same report
# we need the filenames to be explicity passed in as arguments
# we need to count how many are in the old one, check theres none in the new onw and ask teh user to confrim before making the update

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import os

default_sheet_name = 'Sheet1'


def read_report_df_and_wb(filename, sheet_name=default_sheet_name):
    # we want to read the values and fill colors for the following fields:
    # GenericPointAddress
    # eTerraAlias
    # Review Status
    # Comments

    df= pd.read_excel(filename)
    #df = df[['GenericPointAddress', 'eTerraAlias', 'Review Status', 'Comments']]

    # Load workbook and sheet using openpyxl
    wb = load_workbook(filename, data_only=True)

    return df, wb




    

def count_comments(df):
    return df['Comments'].count()

def count_review_status(df):
    return df['Review Status'].count()

def copy_values_and_fill_color(old_df, old_wb, new_df, new_wb, matchmethod):
    # we need to copy the values and fill color for the following fields, matching on the matchmethod field:
    # Review Status
    # Comments

    # we need to find the index of the matchmethod field in the old_df
    matchmethod_index = old_df.columns.get_loc(matchmethod)

    # we need to find the index of the Review Status field in the old_df
    review_status_index = old_df.columns.get_loc('Review Status')

    # we need to find the index of the Comments field in the old_df
    comments_index = old_df.columns.get_loc('Comments')

    # we need to find the index of the matchmethod field in the new_df
    matchmethod_index = new_df.columns.get_loc(matchmethod)

    # we need to find the index of the Review Status field in the new_df
    review_status_index = new_df.columns.get_loc('Review Status')

    # we need to find the index of the Comments field in the new_df
    comments_index = new_df.columns.get_loc('Comments')

    # we need to iterate over the old_df and copy the values to the new_df
    for index, row in old_df.iterrows():
        # we need to find the row in the new_df that matches the matchmethod field
        new_row = new_df[new_df[matchmethod] == row[matchmethod]]

        # we need to copy the values to the new_df
        new_df.at[index, 'Review Status'] = row['Review Status']
        new_df.at[index, 'Comments'] = row['Comments']

    # now we need to copy the fill color to the new_df
    old_ws = old_wb[default_sheet_name]
    new_ws = new_wb[default_sheet_name]

    # Get the index of the matchmethod column
    matchmethod_index_old = old_ws.iter_rows(min_row=1, max_col=1, values_only=True)[0].index(matchmethod)
    matchmethod_index_new = new_ws.iter_rows(min_row=1, max_col=1, values_only=True)[0].index(matchmethod)

    for column_name in ['Review Status', 'Comments']:
        # Get the index of the column by looking at the header row (first row)
        column_index = old_ws.iter_rows(min_row=1, max_col=1, values_only=True)[0].index(column_name)

        # Iterate over the rows in the old_ws
        for row in old_ws.iter_rows(min_row=2, values_only=True):
            # Get the matchmethod value from the old_ws
            matchmethod_value = row[matchmethod_index_old]

            # Get the value from the old_ws
            value = row[column_index]
            
            # Copy the fill color to the new_ws
            new_ws.cell(row=row[0], column=column_index).fill = row[column_index]


    

def get_params():
    import argparse
    # we need the filenames to be explicity passed in as arguments

    parser = argparse.ArgumentParser(description="Copy cReview Status and Comments from defect report, values and fill color")
    parser.add_argument("--old_file", required=True, help="Path to the old defect report file")
    parser.add_argument("--new_file", required=True, help="Path to the new defect report file")
    parser.add_argument("--matchmethod", choices=["eTerraAlias", "GenericPointAddress"], default="GenericPointAddress", required=True, help="Method to match the rows between the two files")
    args = parser.parse_args()

    # Check if the files exist
    if not os.path.exists(args.old_file):
        print(f"Error: The file {args.old_file} does not exist.")
        return
    if not os.path.exists(args.new_file):
        print(f"Error: The file {args.new_file} does not exist.")
        return

    return args.old_file, args.new_file, args.matchmethod

def main():
    old_file, new_file, matchmethod = get_params()

    # we need to count how many are in the old one, check theres none in the new onw and ask teh user to confrim before making the update

    old_df, old_wb = read_report_df_and_wb(old_file)
    new_df, new_wb = read_report_df_and_wb(new_file)

    print(f"Opened old file {old_file} and new file {new_file}...")

    old_comments_count = count_comments(old_df)
    old_review_status_count = count_review_status(old_df)
    new_comments_count = count_comments(new_df)
    new_review_status_count = count_review_status(new_df)

    print(f"Old comments count: {old_comments_count}")
    print(f"Old review status count: {old_review_status_count}")
    print(f"New comments count: {new_comments_count}")
    print(f"New review status count: {new_review_status_count}")

    # Ask the user to confirm before making the update
    user_confirm = input("Are you sure you want to make the update? (y/n): ")
    if user_confirm != "y":
        print("Update cancelled.")
        return

    # Make the update for the values and fill color
    copy_values_and_fill_color(old_df, old_wb, new_df, new_wb, matchmethod)

    # Save the new file
    new_wb.save(new_file)

    print(f"Updated {new_file} with the values and fill color from {old_file}.")


if __name__ == "__main__":
    main() 

