import argparse
import pandas as pd
import logging
import datetime

# Add Logging
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)
fh = logging.FileHandler('transform_logs.log', encoding='utf-8')
fh.setLevel(logging.DEBUG)
logger.addHandler(fh)
# Add name
NAME = 'report_form_2c_ms.py'

def transform(csv_list: list, output_report_path):
    def construct_marker_matrix(csv_list):
        # Base DF
        df = pd.DataFrame()
        for i in csv_list:
            '''
            Take report files 1-by-1 and the merge then on external index from indices.py
            This will cut RAM cost 30 times (and make shit slower)
            '''
            # Merge 2 frames
            df_add = pd.read_csv(i, sep=';', encoding='utf-8',header=0)
            df = pd.concat([df, df_add], ignore_index=True)

        # Extracting suffixes from 'балл' and 'комментарий' columns
        ball_columns = [col for col in df.columns if col.endswith('балл')]
        comment_columns = [col for col in df.columns if col.endswith('комментарий')]
        suffixes = sorted(set(col.replace(', балл', '').replace(', комментарий', '') for col in ball_columns + comment_columns))

        # Creating a new ordered list of columns
        new_order = []
        for suffix in suffixes:
            ball_col = f'{suffix}, балл'
            comment_col = f'{suffix}, комментарий'  
            if ball_col in df.columns:
                new_order.append(ball_col)
            if comment_col in df.columns:
                new_order.append(comment_col)
            else:
                print(f'No matching комментарий column for {suffix}')
        # Adding columns that do not end with 'балл' or 'комментарий'
        new_order = [col for col in df.columns if col not in new_order] + new_order
        # Reordering the DataFrame columns
        df = df[new_order]

        return df

    '''Run Script'''
    df = construct_marker_matrix(csv_list)
    df.to_excel(output_report_path, 'Выгрузка')
    logger.info(f'%s {NAME}: exit code 0 (Success)', datetime.datetime.now())
    print('Exit Code 0')
    return 0
    
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