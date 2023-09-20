import argparse
import pandas as pd
import traceback
import logging
import sys
import datetime

# Add Logging
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)
# create file handler which logs even debug messages
fh = logging.FileHandler('transform_logs.log')
fh.setLevel(logging.DEBUG)
logger.addHandler(fh)

def transform(csv_list: list, output_report_path):
    def format_xlsx(pivot_table: pd.DataFrame, sheet: str = 'Отчет 1',
                    name: str = "pivot_table_gradient_colorscale.xlsx"):
        # Create an Excel writer and export the pivot table to an Excel file
        excel_file_path = name
        with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
            pivot_table.to_excel(writer, sheet_name=sheet, index=True)

            # Get the xlsxwriter workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets[sheet]

            # Get the dimensions of the pivot table
            max_row = len(pivot_table)
            max_col = len(pivot_table.columns)

            # Add a format for the header cells
            header_format = workbook.add_format(
                {'bold': True, 'text_wrap': True, 'valign': 'top', 'border': 1, 'bg_color': '#EFEFEF',
                 'align': 'center'})

            # Set the column width and format for the header
            for col_num, value in enumerate(pivot_table.columns.values):
                worksheet.write(0, col_num + 1, value, header_format)
                column_len = max(pivot_table[value].astype(str).str.len().max(), len(value)) + 2
                worksheet.set_column(col_num + 1, col_num + 1, column_len)

            # Apply gradient color scale to value cells
            # https://xlsxwriter.readthedocs.io/working_with_conditional_formats.html
            worksheet.conditional_format(1, 1, max_row, max_col, {
                'type': '3_color_scale',
                'min_color': '#A6D86E',  # Green
                'mid_color': '#FFFFFE',  # White (for NaN)
                'max_color': '#e85f5f',  # Red
                'min_type': 'num',
                'mid_type': 'num',
                'max_type': 'num'
            })
            print(f'file: {name} -- Transformed 0')
    try:
        # Concatenate all csv to a single biiig df
        df = pd.DataFrame()
        for i in csv_list:
            df_add = pd.read_csv(i, sep=';', header=0)
            df = pd.concat([df, df_add], ignore_index=True)
        # Fix мультидоговоры for RSB
        mask = df['№ п/п'].isna()
        df = df[~mask]
        # Convert the 'Длительность звонка' column to Timedelta
        df['Длительность звонка'] = pd.to_timedelta(df['Длительность звонка'])
        # 1. Filter the dataframe
        filtered_df = df[df['Результат автооценки'] != 100]  # !!!
        # filtered_df = df # !!!
        # 2. Create the pivot table
        pivot_table = filtered_df.pivot_table(index='Имя колл-листа',
                                            columns='Результат робота',
                                            values='Результат автооценки',
                                            aggfunc='size',
                                            fill_value=0)
        # Convert to xlsx and apply colorscale
        format_xlsx(pivot_table, name=output_report_path)
        print('Exit Code 0')
        return 0
    except ValueError or KeyError as e:
        logger.exception(f'\n{datetime.datetime.now()}\nExit Code 1 (Pandas Error): %s', e)
        print('Exit Code 1 (Pandas Error)')
        return 1
    except Exception as e:
        logger.exception(f'\n{datetime.datetime.now()}\nExit Code 2 (Unknown Error): %s', e)
        print('Exit Code 2 (Unknown Error)')
        return 2

if __name__ == '__main__':
    try:
        # Create a parser to handle command-line arguments
        parser = argparse.ArgumentParser(description='Process CSV files and create an Excel pivot table with color scaling.')

        # Add arguments for CSV list and output report path
        parser.add_argument('--csv_list', nargs='+', help='List of CSV file paths', required=True)
        parser.add_argument('--output_report_path', help='Path for the output Excel report', required=True)

        # Capture standard error (stderr) and redirect it to the logger
        sys.stderr = logging.StreamHandler(fh)

        # Parse the command-line arguments
        args = parser.parse_args()

        # Check if required arguments are missing
        if not args.csv_list or not args.output_report_path:
            logger.error(f'\n{datetime.datetime.now()}\nExit Code 3 (Input Error): One or both required arguments are missing')
            print('Exit Code 3 (Input Error): One or both required arguments are missing')
        else:
            # If required arguments are present, proceed with transformation
            transform(csv_list=args.csv_list, output_report_path=args.output_report_path)
    except Exception as ee:
        logger.exception(f'\n{datetime.datetime.now()}\nExit Code 4 (Script Error): %s', ee)

