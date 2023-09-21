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
NAME = 'report_form_3'

def transform(csv_list: list, output_report_path):
    '''Preprocess'''
    def prep_data(csv_list=csv_list):
        # Concatenate all csv to a single biiig df
        df = pd.DataFrame()
        for i in csv_list:
            df_add = pd.read_csv(i, sep=';', encoding='utf-8',header=0)
            df = pd.concat([df, df_add], ignore_index=True)
        # Fix мультидоговоры
        mask = df['№ п/п'].isna()
        df = df[~mask]
        # Step 1: Convert the 'Длительность звонка' column to Timedelta
        #df['Длительность звонка'] = pd.to_timedelta(df['Длительность звонка'])
        df['Ошибки'] = df['Результат автооценки'] != 100
        df['Дата'] = pd.to_datetime(df['Дата звонка'], format='%d.%m.%Y %H:%M:%S')
        df['Дата'] = df['Дата'].dt.strftime('%d.%m.%Y')
        df = df.reset_index(drop=True)
        return df

    '''Pandas Code'''
    def create_pivot(df, rpc=False):
        # GET HEADERS AND REMOVE IRRELEVANT ONES
        queries_list = list(df.columns)
        columns_to_remove = [
            '№ п/п',
            'ID звонка',
            'Имя колл-листа',
            'Результат робота',
            'Дата звонка',
            'Результат автооценки',
            'Ошибки',
            'Дата',
            'Поисковый запрос: Все звонки, балл'
        ]
        # Remove the specified columns
        queries_list = [col for col in queries_list if col not in columns_to_remove]

        # Все звонки fix
        df['Всего звонков по листу'] = df['Поисковый запрос: Все звонки, балл']
        queries_list.append('Всего звонков по листу')
        #df['Всего звонков по листу'] = df['Всего звонков по листу'] * -1

        for col in queries_list:
            df[col] = df[col] / df[col]

        queries_list.append('Ошибки')

        # First, melt the DataFrame to convert 'Запрос_1', 'Запрос_2', 'Запрос_3' into rows
        melted_df = pd.melt(df, id_vars=['Имя колл-листа', 'Дата'], value_vars = queries_list, var_name='Запрос', value_name='Ошибки шт.')
        melted_df.reset_index(drop=True)
        '''I could insert check here that will assign a block based on Имя кол-лист, but it will decrease the time'''
        del df
        # Now, create a pivot table to calculate sums
        pivot_table = melted_df.pivot_table(
            values='Ошибки шт.',
            index=['Имя колл-листа', 'Запрос'],
            columns='Дата',
            aggfunc='sum'
        )
        del melted_df
        # If you want to reset the index and have a cleaner view
        pivot_table.reset_index()
        return pivot_table

    '''Excel Code'''
    def format_xlsx(pivot_table: pd.DataFrame,
                    name: str = "pivot_table_2_call_lists.xlsx", **kwargs):
        # Specify the Excel file path
        excel_file_path = name
        # Create a Pandas Excel writer using xlsxwriter as the engine
        writer = pd.ExcelWriter(excel_file_path, engine='xlsxwriter')
        # Write the DataFrame to the Excel file
        pivot_table.to_excel(writer, sheet_name='Все звонки')
        # Access the xlsxwriter workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Все звонки']
        # Define a white fill format
        white_fill_format = workbook.add_format({'text_wrap': True, 'bg_color': '#FFFFFF', 'border': 0})
        # Define a white fill format
        white_fill_format = workbook.add_format({'text_wrap': True, 'bg_color': '#FFFFFF', 'border': 0})
        # Apply the white background to the entire worksheet
        worksheet.set_column(0, 0, 40, white_fill_format)
        worksheet.set_column(1, 1, 80, white_fill_format)
        worksheet.set_column(0, 100, 20, white_fill_format)
        # Save the Excel file
        writer.save()
        print(f'file: {name} -- Transformed 0')

    '''Run Script'''
    df = prep_data(csv_list=csv_list)
    df = create_pivot(df, rpc=False)
    format_xlsx(df,
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