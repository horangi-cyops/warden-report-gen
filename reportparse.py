from tokenize import group
import pandas as pd
import argparse


def gen_report(raw, group):
    reportdf = pd.DataFrame(columns=["Issue Title","Rating", "Score", "Affected Module", "Affected Resource", "Notes"])
    for index, row in group.iterrows(): # iterate over issues in the grouped sheet
        if row["Issue Title"].lower() != "not an issue": # check if not an issue based on manual review
            for ind, r in raw.iterrows(): #iterate over rules in the raw sheet
                if row['Affected Module'] == r['Rule Title']:
                    reportdf.loc[len(reportdf.index)] = ([row["Issue Title"], row["Severity"], row["Score"],row["Affected Module"], r["Resource"], r["Notes"]])
    return reportdf


def alt_gen_report_df(raw_df: pd.DataFrame, group_df: pd.DataFrame) -> pd.DataFrame:
    """Generates a report given a DataFrame of raw checks, and another
    DataFrame with manually grouped issues.
    """

    # Keeping only the columns we need
    raw_df = raw_df[["Rule Title", "Resource", "Notes", "Recommendation", "References"]]
    group_df = group_df[["Issue Title", "Severity", "Score", "Affected Module"]]

    # Removing rows in group_df that have "not an issue" in title
    # This is also case insensitive so that's nice
    group_df = group_df[~group_df['Issue Title'].str.contains("not an issue", case=False)]

    # merging the two dataframes
    merged_df = raw_df.merge(group_df, how='inner', right_on='Affected Module', left_on='Rule Title')

    # Renaming columns to keep them consistent and also removing duplicates
    merged_df = merged_df.rename(columns={'Resource':'Affected Resource', 'Severity': 'Rating'})
    
    print(merged_df.columns)
    # finally, keeping only the columns needed
    merged_df = merged_df[['Issue Title', 'Rating', 'Score', 'Affected Module', 'Affected Resource', 'Notes', 'Recommendation', 'References']]
    merged_df = merged_df.drop_duplicates()

    return merged_df


if __name__ == '__main__':
    aparse = argparse.ArgumentParser(description="Exports an excel instance report based on Warden raw output combined with manual issue categories", usage="\nreport-parse.py <excel-file>")
    aparse.add_argument("--excel_file", type=str, help="input the excel file here")
    aparse.add_argument("--raw", type=str, default="raw-checks", help="raw check spreadsheet name", required=False)
    aparse.add_argument("--grouped", type=str, default="Issue Grouping", help="grouped check spreadsheet name", required=False)
    args = aparse.parse_args()

    raw_df = pd.read_excel(args.excel_file, sheet_name=args.raw)
    group_df = pd.read_excel(args.excel_file, sheet_name=args.grouped)

    finaldf = alt_gen_report_df(raw_df, group_df)
    
    writer = pd.ExcelWriter("Report.xlsx", engine='xlsxwriter')
    finaldf.to_excel(writer, index=False, sheet_name="Instances")

    workbook = writer.book
    worksheet = writer.sheets['Instances']

    worksheet.set_zoom(90)

    worksheet.set_column('A:H', 25)
    worksheet.set_column('B:C', 10)

    # num_rows = len(finaldf.index)+1
    # colour_range = "B2:B{}".format(num_rows)


    # worksheet.conditional_format(colour_range, {'type': 'text',
    #                                     'criteria': '=',
    #                                     'value': 'Very High',
    #                                     'format': {'bg_color': '#CC001C'}})
    # worksheet.conditional_format(colour_range, {'type': 'text',
    #                                     'criteria': '=',
    #                                     'value': 'High',
    #                                     'format': {'bg_color': '#f4750e'}})
    # worksheet.conditional_format(colour_range, {'type': 'text',
    #                                     'criteria': '=',
    #                                     'value': 'Medium High',
    #                                     'format': {'bg_color': '#ffaa00'}})
    # worksheet.conditional_format(colour_range, {'type': 'text',
    #                                     'criteria': '=',
    #                                     'value': 'Medium',
    #                                     'format': {'bg_color': '#ffd23e'}})
    # worksheet.conditional_format(colour_range, {'type': 'text',
    #                                     'criteria': '=',
    #                                     'value': 'Low',
    #                                     'format': {'bg_color': '#3aa537'}})
    writer.close()

    print("Done!")