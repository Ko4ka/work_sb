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
NAME = 'standard_marker_matrix.py'

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

        # Forward fill NaN values in 'Маркер' column
        df.fillna(method='ffill', inplace=True)
        # Create a pivot table with 'Маркер' as columns and 'Маркер - количество совпадений' as values
        pivot_df = df.pivot_table(index=['№ п/п', 'ID звонка', 'Оператор', 'Дата звонка', 'Длительность звонка', 'Всего пауз, сек'], columns='Маркер', values='Маркер - количество совпадений').reset_index()
        # Reset the index and rename the columns
        pivot_df.columns.name = None  # Remove the columns' name
        # CREATED A MATRIX FORMAT
        pivot_df = pivot_df.fillna(0)
        pivot_df['Длительность звонка'] = pd.to_timedelta(pivot_df['Длительность звонка'])
        pivot_df['Дата'] = pd.to_datetime(pivot_df['Дата звонка'], format='%d.%m.%Y %H:%M:%S')
        pivot_df['Дата'] = pivot_df['Дата'].dt.strftime('%d.%m.%Y')
        pivot_df = pivot_df.reset_index(drop=True)
        return pivot_df
    
    def construct_pivot(df):
        # Свожу универсальную простыню с маркерами
        pivot_df = df.pivot_table(index='Оператор', aggfunc='sum')
        return pivot_df

    '''Run Script'''
    df = construct_pivot(construct_marker_matrix(csv_list))
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