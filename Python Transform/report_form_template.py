import argparse
import pandas as pd
import logging
import datetime

# Add Logging
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)
fh = logging.FileHandler('transform_logs.log')
fh.setLevel(logging.DEBUG)
logger.addHandler(fh)
# Add name
NAME = 'report_form_1.py'

def transform(csv_list: list, output_report_path):
    '''Preprocess'''
    def prep_data(csv_list=csv_list):
        pass

    '''Pandas Code'''
    def create_pivot(df, rpc=False):
        pass

    '''Excel Code'''
    def format_xlsx(pivot_table: pd.DataFrame, sheet: str = 'Отчет общий',
                    name: str = "pivot_table_2_call_lists.xlsx", **kwargs):
        pass

if __name__ == '__main__':
    try:
        logger.info(f'{NAME} \nScript started at %s', datetime.datetime.now())
        parser = argparse.ArgumentParser(description='Process CSV files and create an Excel pivot table with color scaling.')
        # Add arguments for CSV list and output report path
        parser.add_argument('--csv_list', nargs='+', help='List of CSV file paths', required=False)
        parser.add_argument('--output_report_path', help='Path for the output Excel report', required=False)
        # Parse the command-line arguments
        args = parser.parse_args()
        # Check Input
        if not args.csv_list or not args.output_report_path:
            raise ValueError(f'One or both required arguments are missing')
        transform(csv_list=args.csv_list, output_report_path=args.output_report_path)
    except Exception as ee:
        logger.exception(f'{NAME} Exit Code 1 (Script Error): {datetime.datetime.now()} %s', ee)