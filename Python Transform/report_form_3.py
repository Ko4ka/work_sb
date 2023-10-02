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
NAME = 'report_form_3.py'

def transform(csv_list: list, output_report_path):
    def construct_df(csv_list):
        '''
        Linear time:
        ~9 sec for 1 day
        ~5 min for 1 month
        '''
        def create_pivot(df):
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
                'Поисковый запрос: Все звонки, балл',
                'Контактное лицо',
                'Длительность звонка',
                'Тип договора'
            ]
            # Remove the specified columns
            queries_list = [col for col in queries_list if col not in columns_to_remove]
            # Все звонки fix: для отображния стрроки 'Все звонки'
            df['Всего звонков по листу'] = df['Поисковый запрос: Все звонки, балл']
            queries_list.append('Всего звонков по листу')
            # Each column has -25/0, by dividing I get a 1/0 format
            for col in queries_list:
                df[col] = df[col].astype('Int8')  # Limit RAM usage
                df[col] = df[col] / df[col]
            queries_list.append('Ошибки')
            # INTENSIVE MELTING
            # First, melt the DataFrame to convert 'Запрос_1', 'Запрос_2', 'Запрос_3' into rows
            melted_df = pd.melt(df, id_vars=['Имя колл-листа', 'Дата'], value_vars = queries_list, var_name='Запрос', value_name='Ошибки шт.')
            melted_df.reset_index(drop=True)
            #del df
            # Now, create a pivot table to calculate sums
            pivot_table = melted_df.pivot_table(
                values='Ошибки шт.',
                index=['Имя колл-листа', 'Запрос'],
                columns='Дата',
                aggfunc='sum'
            )
            # del melted_df
            return pivot_table
        
        df_main = pd.DataFrame()
        df_rpc = pd.DataFrame()    
        for iteration, i in enumerate(csv_list):
            '''
            Take report files 1-by-1 and the merge then on external index from indices.py
            This will cut RAM cost 30 times (and make shit slower)
            '''
            # Merge 2 frames
            df = pd.read_csv(i, sep=';', encoding='utf-8',header=0)
            # Remove мультидоговоры for RSB
            mask = df['№ п/п'].isna()
            df = df[~mask]
            # Convert the 'Длительность звонка' column to Timedelta
            df['Длительность звонка'] = pd.to_timedelta(df['Длительность звонка'])
            df['Ошибки'] = df['Результат автооценки'] != 100
            # Fix Date
            df['Дата'] = pd.to_datetime(df['Дата звонка'], format='%d.%m.%Y %H:%M:%S')
            df['Дата'] = df['Дата'].dt.strftime('%d.%m.%Y')
            df = df.reset_index(drop=True)
            # Create RPC
            rpc_df = df[df['Контактное лицо'] == 'Должник']
            rpc_df = rpc_df.reset_index(drop=True)
            # Warn if dates != 1
            if len(df['Дата'].unique().tolist()) > 1:
                logger.warning('%s Warning: more than a single date in df...', datetime.datetime.now())
            # MEMORY MANAGEMENT: CONCAT TO INDEX AND DELETE
            main_pivot = create_pivot(df)
            df_main = pd.concat([df_main, main_pivot], axis=1)
            del main_pivot  # Save 10MB
            rpc_pivot = create_pivot(rpc_df)
            df_rpc = pd.concat([df_rpc, rpc_pivot], axis=1)
            del rpc_pivot  # Save 10MB
            # Log stage
            logger.info(f'%s {NAME}: iteration #{iteration} done...', datetime.datetime.now())

        return df_main, df_rpc

    '''Excel Code'''
    def format_xlsx(pivot_all: pd.DataFrame,
                    pivot_rpc: pd.DataFrame,
                    name: str = "pivot_table_2_call_lists.xlsx",
                    enable_filtering=True,
                    **kwargs):
        # Settings
        if enable_filtering:
            pivot_all = pivot_all.reset_index()
            pivot_rpc = pivot_rpc.reset_index()
        excel_file_path = name

        # Colors
        # Define a green data bar format
        green_databar_format = {
            'type': 'data_bar',
            'bar_color': '#63C384',  # Hex color code for green
        }

        # Define a red data bar format
        red_databar_format = {
            'type': 'data_bar',
            'bar_color': '#e85f5f',  # Hex color code for green
        }
        color_scale_rule_percent = {
                        'type': '2_color_scale',
                        'min_color': '#FFFFFF',  # White
                        'max_color': '#e85f5f',  # Red
                        'min_type': 'num',
                        'min_value': 0,
                        'max_type': 'percentile',
                        'max_value': 100
                        }
        
        with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
            '''Function Start'''
            def create_sheet(pivot_table, sheet_name):
                # Write the DataFrame to the Excel file
                pivot_table.to_excel(writer, sheet_name=sheet_name)
                # Access the xlsxwriter workbook and worksheet objects
                workbook = writer.book
                worksheet = writer.sheets[sheet_name]
                white_fill_format = workbook.add_format({'text_wrap': False, 'bg_color': '#FFFFFF', 'border': 0})
                table_fill_format = workbook.add_format({
                'text_wrap': False,
                'bg_color': '#FFFFFF',
                'border': 1,  # Standard border
                'border_color': '#BFBFBF'  # Line color
                })
                ref_format = workbook.add_format({'text_wrap': False, 'bold': True, 'align':'right', 'bg_color': '#FFFFFF', 'border': 0})
                worksheet.set_column(0, 0, 5, white_fill_format)
                worksheet.set_column(1, 1, 25, white_fill_format)
                worksheet.set_column(2, 2, 100, white_fill_format)
                worksheet.set_column(3, pivot_table.shape[1]+1, 10, table_fill_format)
                worksheet.set_column(pivot_table.shape[1]+1, 100, 10, white_fill_format)
                format_tmp_start = []
                format_tmp_stop = []
                for row_num, row in enumerate(pivot_table['Запрос']):
                    if "Всего звонков по листу" in row:
                        worksheet.conditional_format(f'D{row_num+2}:AZ{row_num+2}', green_databar_format)
                        worksheet.write(f'C{row_num+2}', 'Всего звонков по листу',ref_format)
                        format_tmp_start.append(row_num+4)
                        format_tmp_stop.append(row_num+1)  # Hack I append a list of previous querie
                    elif "Ошибки" in row:
                        worksheet.conditional_format(f'D{row_num+2}:AZ{row_num+2}', red_databar_format)
                        worksheet.write(f'C{row_num+2}', 'Ошибки', ref_format)
                # Prepare ranges for formatting
                format_tmp_start = format_tmp_start[:-1]
                format_tmp_stop = format_tmp_stop[1:]
                for num, i in enumerate(format_tmp_start):
                    worksheet.conditional_format(f'D{i}:AZ{format_tmp_stop[num]}', color_scale_rule_percent)
            # Create Sheets
            create_sheet(pivot_all, 'Все звонки')
            create_sheet(pivot_rpc, 'RPC')
            # Create Summary Sheet
        print(f'file: {name} -- Transformed 0')

    '''Run Script'''
    df_main, df_rpc = construct_df(csv_list=csv_list)
    # Sort columns
    df_main = df_main.sort_index(axis=1)
    df_rpc = df_rpc.sort_index(axis=1)
    format_xlsx(df_main.replace(0, pd.NA), df_rpc.replace(0, pd.NA),
                name=output_report_path,
                enable_filtering=True)
    
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
    except MemoryError as oom:
        logger.exception(f'{datetime.datetime.now()} {NAME}: exit code 3: (OOM Error)\n%s', oom)
    except Exception as ee:
        logger.exception(f'{datetime.datetime.now()} {NAME}: exit code 1: (Python Error)\n%s', ee)