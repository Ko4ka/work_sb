import argparse
import pandas as pd
import logging
import sys
import datetime
import io

# Add Logging
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)
# create file handler which logs even debug messages
fh = logging.FileHandler('transform_logs.log')
fh.setLevel(logging.DEBUG)
logger.addHandler(fh)

class StdErrLogger(io.StringIO):
    def write(self, message):
        # Log the error message
        logger.error(f'\n{datetime.datetime.now()}\nExit Code 3 (Input Error): %s', message)

def transform(csv_list: list, output_report_path):
    def format_xlsx(pivot_table: pd.DataFrame, sheet: str = 'Общий',
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
                {'bold': False, 'text_wrap': True, 'border': 1, 'bg_color': '#EFEFEF', 'align': 'center', 'valign': 'bottom', 'rotation': 270})

            # Apply gradient color scale to value cells
            # https://xlsxwriter.readthedocs.io/working_with_conditional_formats.html
            worksheet.conditional_format(1, 1, max_row, max_col, {
                'type': '3_color_scale',
                'min_color': '#A6D86E',  # Green
                'mid_color': '#FFFFFE',  # White (for NaN)
                'max_color': '#e85f5f',  # Red
                'min_type': 'percentile',
                'min_value': 0,
                'mid_type': 'percentile',
                'mid_value': 50,
                'max_type': 'percentile',
                'max_value': 100
            })
            print(f'file: {name} -- Transformed 0')
            white_fill_format = workbook.add_format({'bg_color': '#FFFFFF', 'border': 0, 'align': 'center'})
            worksheet.set_column(0, 0, 60, white_fill_format)
            worksheet.set_column(1, 100, 6, white_fill_format)
            worksheet.set_row(0, 50, header_format)
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
        logger.info('Script completed successfully at %s', datetime.datetime.now())
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
        logger.info('Script started at %s', datetime.datetime.now())
        # Capture the original stderr
        original_stderr = sys.stderr
        # Redirect stderr to the custom stream
        sys.stderr = stderr_logger = StdErrLogger()
        # Create a parser to handle command-line arguments
        parser = argparse.ArgumentParser(description='Process CSV files and create an Excel pivot table with color scaling.')
        # Add arguments for CSV list and output report path
        parser.add_argument('--csv_list', nargs='+', help='List of CSV file paths', required=True)
        parser.add_argument('--output_report_path', help='Path for the output Excel report', required=True)
        # Parse the command-line arguments
        args = parser.parse_args()
        # Check if required arguments are missing
        if not args.csv_list or not args.output_report_path:
            raise ValueError('One or both required arguments are missing')
        # If required arguments are present, proceed with transformation
        transform(csv_list=args.csv_list, output_report_path=args.output_report_path)
    except Exception as ee:
        logger.exception(f'\n{datetime.datetime.now()}\nExit Code 4 (Script Error): %s', ee)
    finally:
        # Restore the original stderr
        sys.stderr = original_stderr

