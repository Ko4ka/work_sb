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
NAME = ''

def transform(csv_list: list, output_report_path):
    '''Preprocess'''
    def prep_data(csv_list=csv_list):
        return df

    '''Pandas Code'''
    def create_pivot(df, rpc=False):
        def run(df):
           return df
       # Create RPC-only frame
        rpc_df = df[df['Контактное лицо'] == 'Должник']
        pivot_all, percent, errors = run(df)
        del df
        pivot_rpc, _, _ = run(rpc_df)
        del rpc_df
        return pivot_all, rpc_df, percent, errors

    '''Excel Code'''
    def format_xlsx(pivot_all: pd.DataFrame,
                    pivot_rpc: pd.DataFrame,
                    name: str = "pivot_table_2_call_lists.xlsx", **kwargs):
        # Settings
        with pd.ExcelWriter(excel_file_path=output_report_path, engine='xlsxwriter') as writer:
            '''Function Start'''
            def create_sheet(pivot_table, sheet_name):
                pass
            # Create Sheets
            create_sheet(pivot_all, 'Все звонки')
            create_sheet(pivot_rpc, 'RPC')
            # Create Summary Sheet
        print(f'file: {name} -- Transformed 0')

    '''Run Script'''
    df = prep_data(csv_list=csv_list)
    df, rpc_df, percent, errors = create_pivot(df, rpc=False)
    format_xlsx(df, rpc_df,
                name=output_report_path,
                percent=percent,
                errors=errors)
    logger.info(f'%s {NAME}: exit code 0 (Success)', datetime.datetime.now())
    print('Exit Code 0')
    
if __name__ == '__main__':
    try:
        logger.info(f'%s {NAME}: script started', datetime.datetime.now())
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
        logger.exception(f'{datetime.datetime.now()} {NAME}: exit code 1: (Script Error)\n%s', ee)