import argparse
import pandas as pd
import logging
import datetime
from indices import form_1_indices as INDICES
from indices import form_1_colors as COLORS

# Add Logging
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)
fh = logging.FileHandler('transform_logs.log', encoding='utf-8')
fh.setLevel(logging.DEBUG)
logger.addHandler(fh)
# Add name
NAME = 'lsr_form_1_pivot.py'

def transform(csv_list: list, output_report_path):
    def construct_df(csv_list):
        '''
        Linear time:
        ~12 sec for 1 day
        ~6 min for 1 month
        '''
        def create_pivot(df):
            columns_to_keep = [
                'Длительность звонка',
                'Оператор',
                'Подразделение',
                'Лист автооценки',
                'Результат автооценки'
            ]
            df = df.loc[:, columns_to_keep]
            df = df[df['Результат автооценки'].notna()]
            df = df.reset_index(drop=True)
            # Create the pivot table
            pivot_table = pd.pivot_table(
                df,
                values=['Результат автооценки'],  # 'Результат автооценки' is used for both mean and count
                index=['Подразделение', 'Оператор'],  # First level: Подразделение, Second level: Оператор
                aggfunc={'Результат автооценки': ['count', 'mean']}
            )

            # Rename the columns to reflect 'mean' and 'count' clearly
            pivot_table.columns = ['Кол-во оцененных звонков', 'Средний результат автооценки']
            pivot_table['Средний результат автооценки'] = pivot_table['Средний результат автооценки'].round(2)

            # Sort the pivot table by 'Подразделение'
            pivot_table = pivot_table.sort_index(level='Подразделение')
            return pivot_table
        
        def create_plain_df(df):
            df_plain = df
            df_plain = df_plain[df_plain['Результат автооценки'].notna()]
            df_plain = df_plain.reset_index(drop=True)
            return df_plain

        # Assemble a full DF from fractions
        df = pd.DataFrame()
        for i in csv_list:
            '''
            Take report files 1-by-1 and the merge then on external index from indices.py
            This will cut RAM cost 30 times (and make shit slower)
            '''
            # Merge 2 frames
            df_add = pd.read_csv(i, sep=';', encoding='utf-8',header=0)
            df = pd.concat([df, df_add], ignore_index=True)

        # Construct 2 files
        df_plain = create_plain_df(df)
        df_pivot = create_pivot(df)
        del df
        return df_pivot, df_plain

    '''Excel Code'''
    def format_xlsx(df_pivot: pd.DataFrame,
                    df_plain: pd.DataFrame,
                    name: str = "pivot_autoscore.xlsx"):
        # Specify the Excel file path
        excel_file_path = name

        # Create a Pandas Excel writer using xlsxwriter as the engine
        with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
            # Write the DataFrame to the Excel file
            df_pivot.to_excel(writer, sheet_name='Свод по оценке')
            # Access the xlsxwriter workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets['Свод по оценке']
            white_fill_format = workbook.add_format({'text_wrap': True, 'bg_color': '#FFFFFF', 'border': 0, 'align':'center'})
            # Beautify formats
            worksheet.set_column(0, 1, 40, white_fill_format)
            worksheet.set_column(2, 2, 35, white_fill_format)
            worksheet.set_column(3, 3, 35, white_fill_format)
            worksheet.set_column(4, 100, 10, white_fill_format)
            df_plain.to_excel(writer, sheet_name='Оцененные звонки')
            workbook = writer.book
            worksheet = writer.sheets['Оцененные звонки']
            worksheet.set_column(0, 1, 5)
            worksheet.set_column(2, 8, 40)
        print(f'file: {name} -- Transformed 0')

    '''Run Script'''
    df_pivot, df_plain = construct_df(csv_list=csv_list)
    format_xlsx(df_pivot, df_plain,
                name=output_report_path)
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