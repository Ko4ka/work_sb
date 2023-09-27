import argparse
import pandas as pd
import logging
import datetime
from indices import form_2_indices as INDICES
from indices import form_2_colors as COLORS

# Add Logging
logger = logging.getLogger()
logger.setLevel(logging.NOTSET)
fh = logging.FileHandler('transform_logs.log', encoding='utf-8')
fh.setLevel(logging.NOTSET)
logger.addHandler(fh)
# Add name
NAME = 'report_form_2.py'

def transform(csv_list: list, output_report_path):
    def construct_df(csv_list):
        '''
        Linear time:
        ~9 sec for 1 day
        ~5 min for 1 month
        '''
        def create_pivot(df):
            # Create multiindex
            def create_multiindex(dataframe, sub_index:str):
                # Create MultiIndex
                multiindex = []
                for i, column in enumerate(dataframe):
                    multiindex.append((column, sub_index))
                dataframe.columns = pd.MultiIndex.from_tuples(multiindex)
                return dataframe
            # Create Pivot FUNC
            pivot_df_calls = df.pivot_table(index=['Имя колл-листа', 'Результат робота'], columns='Дата', values='Результат автооценки', aggfunc='count', fill_value='')
            pivot_df_calls.columns = pd.to_datetime(pivot_df_calls.columns, format='%d.%m.%Y')  # Fix date Time
            pivot_df_calls = pivot_df_calls.sort_index(axis=1)  # Fix date Time
            # Calculate number of errors
            pivot_df_errors = df.pivot_table(index=['Имя колл-листа', 'Результат робота'], columns='Дата', values='Ошибки', aggfunc='sum', fill_value='')
            pivot_df_errors.columns = pd.to_datetime(pivot_df_errors.columns, format='%d.%m.%Y')  # Fix date Time
            pivot_df_errors = pivot_df_errors.sort_index(axis=1)  # Fix date Time
            # Calculate mean autoscore
            pivot_df_mean = df.pivot_table(index=['Имя колл-листа', 'Результат робота'], columns='Дата', values='Результат автооценки', aggfunc='mean', fill_value='')
            pivot_df_mean.columns = pd.to_datetime(pivot_df_mean.columns, format='%d.%m.%Y')  # Fix date Time
            pivot_df_mean = pivot_df_mean.sort_index(axis=1)  # Fix date Time
            # Calculate error rate
            pivot_df_error_rate = (pivot_df_errors.replace("", pd.NA) / pivot_df_calls.replace("", pd.NA)).applymap(lambda x: x if not pd.isna(x) else '')
            pivot_df_error_rate.columns = pd.to_datetime(pivot_df_error_rate.columns, format='%d.%m.%Y')  # Fix date Time
            pivot_df_error_rate = pivot_df_error_rate.sort_index(axis=1)  # Fix date Time
            # Create MultiIndex
            pivot_df_calls = create_multiindex(pivot_df_calls, 'Зв.(шт.)')
            pivot_df_errors = create_multiindex(pivot_df_errors, 'Ошб.(шт.)')
            pivot_df_mean = create_multiindex(pivot_df_mean, 'Ср.АО')
            pivot_df_error_rate = create_multiindex(pivot_df_error_rate, 'Ошб.%')
            # Create a list of the DataFrames you want to merge
            #dfs_to_merge = [pivot_df_calls, pivot_df_errors, pivot_df_mean, pivot_df_error_rate]
            dfs_to_merge = [pivot_df_error_rate, pivot_df_errors, pivot_df_calls, pivot_df_mean]
            # Initialize an empty DataFrame with the same index as the original DataFrames
            merged_df = pd.DataFrame(index=pivot_df_calls.index)
            # Create Multiindex
            multi_index = []
            # Iterate through the DataFrames and concatenate their columns in the desired order
            for num, column in enumerate(pivot_df_calls.columns):
                    for dataframe in dfs_to_merge:
                            col_name = (dataframe.iloc[:, num].name[0], dataframe.iloc[:, num].name[1])
                            # Append the column name tuple to the list
                            multi_index.append(col_name)
                            merged_df[col_name] = dataframe.iloc[:, num]
            merged_df.columns = pd.MultiIndex.from_tuples(multi_index)
            # Returns pivot
            return merged_df
        # Create Base dfs for pivots
        multi_index = pd.MultiIndex.from_tuples(INDICES)
        multi_header = pd.MultiIndex.from_tuples([('tmp1','tmp2')])
        # Create your empty DataFrames with the MultiIndex
        df_main = pd.DataFrame(index=multi_index, columns=multi_header)
        df_rpc = pd.DataFrame(index=multi_index, columns=multi_header)
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
        # Remove TMP columns
        del df_main[('tmp1','tmp2')]
        del df_rpc[('tmp1','tmp2')]
        # Returns 2 complete pivots
        return df_main, df_rpc

    def construct_summary(df_main, df_rpc):
        '''
        Calculations for the summary sheet
        WEIGHTED BT DAY -> Simple
        '''
        def create_col(pivot, title):
            df = pivot.apply(pd.to_numeric, errors='coerce')
            errors_percent = df.loc[:, df.columns.get_level_values(1) == 'Ошб.%']
            errors_percent = errors_percent.mean(axis=1, skipna=True)
            errors_count = df.loc[:, df.columns.get_level_values(1) == 'Ошб.(шт.)']
            errors_count = errors_count.sum(axis=1, numeric_only=True)
            calls_count = df.loc[:, df.columns.get_level_values(1) == 'Зв.(шт.)']
            calls_count = calls_count.sum(axis=1, numeric_only=True)
            score = df.loc[:, df.columns.get_level_values(1) == 'Ср.АО']
            score = score.mean(axis=1, skipna=True)
            df = pd.DataFrame(index=pd.MultiIndex.from_tuples(df.index), columns=pd.MultiIndex.from_tuples([(title,'')]))
            # Mask Error Count
            mask = calls_count == 0
            df[(title, 'Ошб.%')] = errors_percent
            df[(title, 'Ошб.(шт.)')] = errors_count[~mask]  # pd.mean considers NA = 0
            df[(title, 'Зв.(шт.)')] = calls_count.replace(0,pd.NA)
            df[(title, 'Ср.АО')] = score
            del df[(title, '')]
            # Returns weighted sumary
            return df
            # Create Summary DF
        df_summary = pd.DataFrame()
        df_summary = pd.concat([create_col(df_main, 'Срез: все звонки'),
                                create_col(df_rpc, 'Срез: RPC') ], axis=1)
        # Returns Dataframe
        return df_summary

    '''Excel Code'''
    def format_xlsx(pivot_all: pd.DataFrame,
                    pivot_rpc: pd.DataFrame,
                    summary: pd.DataFrame,
                    name: str = "pivot_table_2_call_lists.xlsx",
                    enable_filtering=True,
                    **kwargs):
        # Reset index to enable filters
        if enable_filtering:
            pivot_all = pivot_all.reset_index()
            pivot_rpc = pivot_rpc.reset_index()
            summary = summary.reset_index()
        # Settings
        excel_file_path = name
        min_errors = kwargs['min errors']
        max_errors = kwargs['max errors']
        divisor = max_errors / kwargs['mid divisor']
        # Create a color scale conditional formatting rule
        color_scale_rule_errors = {
                        'type': '3_color_scale',
                        'min_color': '#A6D86E',  # Green
                        'mid_color': '#FCFAA0',  # White (for NaN)
                        'max_color': '#e85f5f',  # Red
                        'min_type': 'num',
                        'min_value': min_errors,
                        'mid_type': 'num',
                        'mid_value': divisor,
                        'max_type': 'num',
                        'max_value': max_errors
                        }
        with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
            '''Function Start'''
            def create_sheet(pivot_table, sheet_name):
                # Write the DataFrame to the Excel file
                pivot_table.to_excel(writer, sheet_name=sheet_name)

                # Access the xlsxwriter workbook and worksheet objects
                workbook = writer.book
                worksheet = writer.sheets[sheet_name]
                # Define a white fill format
                white_fill_format = workbook.add_format({'text_wrap': False,
                                                         'bg_color': '#FFFFFF',
                                                         'border': 0})
                white_index_format = workbook.add_format({'text_wrap': True,
                                                        'bg_color': '#FFFFFF',
                                                        'border': 0,
                                                        'bold': True
                                                        })
                bold_format = workbook.add_format({'bold': True,
                                                   'border': 1,
                                                   'valign': 'vcenter'})
                # Apply the white background to the entire worksheet
                worksheet.set_column(0, 0, 5, white_fill_format)
                worksheet.set_column(1, 1, 25, white_index_format)
                worksheet.set_column(2, 2, 30, white_index_format)
                worksheet.set_column(3, 100, 13, white_fill_format)
                worksheet.merge_range('A1:A2', '№', bold_format)
                worksheet.merge_range('B1:B2', 'Имя колл-листа', bold_format)
                worksheet.merge_range('C1:C2', 'Статус робота', bold_format)
                percentage_format = workbook.add_format({'num_format': '0.00%', 'bg_color': '#FFFFFF', 'border': 0})

                for i, j in enumerate(pivot_table.head()):
                    if i>2:
                        worksheet.set_column(i+1, i+1, 10, white_fill_format)
                    if j[1] == 'Ошб.%':
                        worksheet.set_column(i+1, i+1, None, percentage_format)
                        worksheet.conditional_format(2, i+1, 999, i+1, color_scale_rule_errors)
                # Autosave
            # Create Sheets
            create_sheet(pivot_all, 'Все звонки')
            create_sheet(pivot_rpc, 'RPC')
            create_sheet(summary, 'Общий срез')
            # Create Summary Sheet
        print(f'file: {name} -- Transformed 0')

    '''Run Script'''
    df_main, df_rpc = construct_df(csv_list=csv_list)
    df_summary = construct_summary(df_main, df_rpc)
    format_xlsx(df_main, df_rpc, df_summary,
                name=output_report_path,
                **COLORS)
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