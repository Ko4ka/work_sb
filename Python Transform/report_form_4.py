import argparse
import pandas as pd
import logging
import datetime
from indices import form_4_indices as INDICES
from indices import form_1_colors as COLORS

# Add Logging
logger = logging.getLogger()
logger.setLevel(logging.DEBUG)
fh = logging.FileHandler('transform_logs.log', encoding='utf-8')
fh.setLevel(logging.DEBUG)
logger.addHandler(fh)
# Add name
NAME = 'report_form_4.py'

def transform(csv_list: list, output_report_path):
    def construct_df(csv_list):
        '''
        Linear time:
        ~12 sec for 1 day
        ~6 min for 1 month
        '''
        def create_pivot(df):
            # Run Transforms for this day
            pivot_df_mistakes = df.pivot_table(index='Тип договора', columns='Дата', values='Ошибки', aggfunc='sum')
            pivot_df_mistakes = pivot_df_mistakes.fillna(0)
            pivot_df_mistakes = pivot_df_mistakes.replace(0.00, '')
            pivot_df_mistakes.columns = pd.to_datetime(pivot_df_mistakes.columns, format='%d.%m.%Y')  # Fix date Time
            pivot_df_mistakes = pivot_df_mistakes.sort_index(axis=1)  # Fix date Time
            tmp_pivot_df_mistakes = pivot_df_mistakes.copy()  # Fix %%
            pivot_df_mistakes.index = pivot_df_mistakes.index + ' (ошибки шт.)'
            # Create dynamic Calls count (2)
            pivot_df_calls = df.pivot_table(index='Тип договора', columns='Дата', values='Результат автооценки', aggfunc='count', fill_value=0)
            pivot_df_calls.columns = pd.to_datetime(pivot_df_calls.columns, format='%d.%m.%Y')  # Fix date Time
            pivot_df_calls = pivot_df_calls.sort_index(axis=1)  # Fix date Time
            tmp_pivot_df_calls = pivot_df_calls.copy()  # Fix %%
            pivot_df_calls.index = pivot_df_calls.index + ' (всего шт.)'
            # Create dynamic Mean Autoscore (3)
            pivot_df_mean = df.pivot_table(index='Тип договора', columns='Дата', values='Результат автооценки', aggfunc='mean', fill_value='')
            pivot_df_mean.columns = pd.to_datetime(pivot_df_mean.columns, format='%d.%m.%Y')  # Fix date Time
            pivot_df_mean = pivot_df_mean.sort_index(axis=1)  # Fix date Time
            pivot_df_mean.index = pivot_df_mean.index + ' (средняя АО)'
            # Create dynamic Error Percentage (4)
            pivot_df_mistakes_filled = tmp_pivot_df_mistakes.replace('', 0)
            pivot_df_error_rate = (pivot_df_mistakes_filled / tmp_pivot_df_calls).applymap(lambda x: x if not pd.isna(x) else '')
            pivot_df_error_rate.index = pivot_df_error_rate.index + ' (доля ошибок %)'
            # Create Mega-Pivot
            # Concatenate the pivot tables vertically along rows (axis=0)
            pivot_table = pd.concat([pivot_df_calls, pivot_df_mistakes, pivot_df_error_rate, pivot_df_mean], axis=0)
            pivot_table = pivot_table.sort_index()
            return pivot_table
        # Concatenate all csv to a single biiig df
        df_main = pd.DataFrame(index=INDICES)
        df_rpc = pd.DataFrame(index=INDICES)
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
        # Returns 2 complete pivots
        return df_main, df_rpc

    def construct_summary(df_main, df_rpc):
        '''
        Calculations for the summary sheet
        WEIGHTED BT DAY -> Simple
        '''
        def create_col(pivot, title):
            pivot = pivot.apply(pd.to_numeric, errors='coerce')
            calls = pivot[pivot.index.str.contains('(всего шт.)')].apply(pd.to_numeric, errors='coerce')
            calls = calls.sum(axis=1)
            error_rate = pivot[pivot.index.str.contains('(доля ошибок %)')].apply(pd.to_numeric, errors='coerce')
            error_rate = error_rate.mean(axis=1,skipna=True,numeric_only=True)
            errors = pivot[pivot.index.str.contains('(ошибки шт.)')].apply(pd.to_numeric, errors='coerce')
            errors = errors.sum(axis=1)
            score = pivot[pivot.index.str.contains('(средняя АО)')].apply(pd.to_numeric, errors='coerce')
            score = score.mean(axis=1,skipna=True, numeric_only=True)
            summary = pd.DataFrame(index=INDICES)
            summary[f'Свод: {title}'] = pd.concat([calls, error_rate, errors, score], axis=0)
            summary = summary.sort_index()
            return summary
        # Create Summary DF
        df_summary = pd.DataFrame()
        df_summary = pd.concat([create_col(df_main, 'все звонки'),
                                create_col(df_rpc, 'RPC') ], axis=1)
        # Returns Dataframe
        return df_summary

    '''Excel Code'''
    def format_xlsx(pivot_all: pd.DataFrame,
                    pivot_rpc: pd.DataFrame,
                    pivot_summary: pd.DataFrame,
                    name: str = "pivot_table_2_call_lists.xlsx", **kwargs):
        # Specify the Excel file path
        excel_file_path = name
        min_percent = kwargs['min percent']
        max_percent = kwargs['max percent']
        min_errors = kwargs['min errors']
        max_errors = kwargs['max errors']
        mid_divisor = kwargs['mid divisor']
        # Create a color scale conditional formatting rule
        color_scale_rule_percent = {
                        'type': '3_color_scale',
                        'min_color': '#A6D86E',  # Green
                        'mid_color': '#FCFAA0',  # White (for NaN)
                        'max_color': '#e85f5f',  # Red
                        'min_type': 'num',
                        'min_value': min_percent,
                        'mid_type': 'num',
                        'mid_value': (max_percent-min_percent)/mid_divisor,
                        'max_type': 'num',
                        'max_value': max_percent
                        }

        # Create a color scale conditional formatting rule
        color_scale_rule_errors = {
                        'type': '3_color_scale',
                        'min_color': '#A6D86E',  # Green
                        'mid_color': '#FCFAA0',  # White (for NaN)
                        'max_color': '#e85f5f',  # Red
                        'min_type': 'num',
                        'min_value': min_errors,
                        'mid_type': 'num',
                        'mid_value': (max_errors-min_errors)/mid_divisor,
                        'max_type': 'num',
                        'max_value': max_errors
                        }

        # Create a Pandas Excel writer using xlsxwriter as the engine
        with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
            '''Function Start'''
            def create_sheet(pivot_table, sheet_name):
                # Write the DataFrame to the Excel file
                pivot_table.to_excel(writer, sheet_name=sheet_name)

                # Access the xlsxwriter workbook and worksheet objects
                workbook = writer.book
                worksheet = writer.sheets[sheet_name]
                bold_format = workbook.add_format({'bold': True})
                worksheet.write('A1', 'Тип договора', bold_format)
                for row_num, row in enumerate(pivot_table.index):
                    if "(доля ошибок %)" in row:
                        format_range = 'B' + str(row_num + 2) + ':AZ' + str(row_num + 2)  # Adjust column range as needed
                        worksheet.conditional_format(format_range, color_scale_rule_percent)
                # Define a percentage format
                percentage_format = workbook.add_format({'num_format': '0.00%', 'bg_color': '#FFFFFF'})

                for row_num, row in enumerate(pivot_table.index):
                    if "(доля ошибок %)" in row:
                        worksheet.set_row(row_num+1, None, percentage_format)

                for row_num, row in enumerate(pivot_table.index):
                    if "(ошибки шт.)" in row:
                        format_range = 'B' + str(row_num + 2) + ':AZ' + str(row_num + 2)  # Adjust column range as needed
                        worksheet.conditional_format(format_range, color_scale_rule_errors)

                # Custom index format
                index_format = workbook.add_format(
                    {'bold': True, 'border': 1, 'bg_color': '#FFFFFF', 'align': 'left'})
                for row_num, row in enumerate(pivot_table.index):
                    worksheet.write(f'A{row_num+2}', row, index_format)

                # Define a white fill format
                white_fill_format = workbook.add_format({'bg_color': '#FFFFFF', 'border': 0})

                # Apply the white background to the entire worksheet
                worksheet.set_column(0, 0, 60, white_fill_format)
                worksheet.set_column(1, 100, 18, white_fill_format)
            # Create 2 sheets
            create_sheet(pivot_all, 'Все звонки')
            create_sheet(pivot_rpc, 'RPC')
            create_sheet(pivot_summary, 'Общий срез')
            '''ADD SUMMARY'''
            # With = Save
        print(f'file: {name} -- Transformed 0')

    '''Run Script'''
    df_main, df_rpc = construct_df(csv_list=csv_list)
    df_summary = construct_summary(df_main=df_main,
                                   df_rpc=df_rpc)
    # Sort Dates
    df_main.sort_index(axis=1, level=0, inplace=True)
    df_rpc.sort_index(axis=1, level=0, inplace=True)
    format_xlsx(df_main, df_rpc, df_summary,
                name=output_report_path, **COLORS)
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