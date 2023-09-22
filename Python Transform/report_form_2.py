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
NAME = 'report_form_2.py'

def transform(csv_list: list, output_report_path):
    '''Preprocess'''
    def prep_data(csv_list=csv_list):
        # Concatenate all csv to a single big df
        df = pd.DataFrame()
        for i in csv_list:
            df_add = pd.read_csv(i, sep=';', encoding='utf-8', header=0)
            df = pd.concat([df, df_add], ignore_index=True)
        # Fix мультидоговоры for RSB
        mask = df['№ п/п'].isna()
        df = df[~mask]
        # Convert the 'Длительность звонка' column to Timedelta
        df['Длительность звонка'] = pd.to_timedelta(df['Длительность звонка'])
        df['Ошибки'] = df['Результат автооценки'] != 100
        df['Дата'] = pd.to_datetime(df['Дата звонка'], format='%d.%m.%Y %H:%M:%S')
        df['Дата'] = df['Дата'].dt.strftime('%d.%m.%Y')
        df = df.reset_index(drop=True)
        return df

    '''Pandas Code'''
    def create_pivot(df):
        def create_summary_pivot(df):
            def create_col(df, title, main_index=None):
                '''
                df: df main
                title: RPC?
                main_index: None id it's the main df, any pivot from the main df if it is RPC
                '''
                header_label = title
                pivot_df_calls = df.pivot_table(index=['Имя колл-листа', 'Результат робота'], values='Результат автооценки', aggfunc='count', fill_value='')
                # Calculate number of errors
                pivot_df_errors = df.pivot_table(index=['Имя колл-листа', 'Результат робота'], values='Ошибки', aggfunc='sum', fill_value='')
                # Calculate mean autoscore
                pivot_df_mean = df.pivot_table(index=['Имя колл-листа', 'Результат робота'], values='Результат автооценки', aggfunc='mean', fill_value='')
                # Calculate error rate
                # Calculate error rate
                pivot_df_error_rate = pd.DataFrame(pivot_df_errors['Ошибки'] / pivot_df_calls['Результат автооценки'])
                pivot_df_error_rate = pivot_df_error_rate.where(pivot_df_errors['Ошибки'] != 0, other='')
                # Create a DataFrame with 'Общий' as the main header
                header = pd.MultiIndex.from_tuples([(f'Срез: {header_label}', 'Ошб.%'), (f'Срез: {header_label}', 'Ошб.(шт.)'), (f'Срез: {header_label}', 'Зв.(шт.)'), (f'Срез: {header_label}', 'Ср.АО')])
                # Handle RPC Case
                if main_index is pd.DataFrame:
                    summary = pd.DataFrame(columns=header, index=main_index.index)
                else:
                    summary = pd.DataFrame(columns=header, index=pivot_df_error_rate.index)
                # Assign your Series to the corresponding columns
                summary[(f'Срез: {header_label}', 'Ошб.%')] = pivot_df_error_rate
                summary[(f'Срез: {header_label}', 'Ошб.(шт.)')] = pivot_df_errors
                summary[(f'Срез: {header_label}', 'Зв.(шт.)')] = pivot_df_calls
                summary[(f'Срез: {header_label}', 'Ср.АО')] = pivot_df_mean
                return summary, pivot_df_error_rate
            # Create full summary
            final_summary = pd.DataFrame()
            full_summary, ref_index = create_col(df, 'все звонки', main_index=None)
            rpc_df = df[df['Контактное лицо'] == 'Должник']
            rpc_summary, _ = create_col(rpc_df, 'RPC', main_index=ref_index)
            final_summary = pd.concat([full_summary, rpc_summary], axis=1)
            return final_summary
        def create_multiindex(dataframe, sub_index:str):
            # Create MultiIndex
            multiindex = []
            for i, column in enumerate(dataframe):
                multiindex.append((column, sub_index))
            dataframe.columns = pd.MultiIndex.from_tuples(multiindex)
            return dataframe
        def run(df):
           # Calculate number of calls for each pair
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
            max_error_rate = pivot_df_error_rate.apply(pd.to_numeric, errors='coerce').max().max()
            min_error_rate = pivot_df_error_rate.apply(pd.to_numeric, errors='coerce').min().min()
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
                            #print(dataframe.iloc[:, num].name)
                            #merged_df[dataframe.iloc[:, num].]
                            col_name = (dataframe.iloc[:, num].name[0], dataframe.iloc[:, num].name[1])
                            # Append the column name tuple to the list
                            multi_index.append(col_name)
                            merged_df[col_name] = dataframe.iloc[:, num]
            merged_df.columns = pd.MultiIndex.from_tuples(multi_index)

            return merged_df, [max_error_rate, min_error_rate]
        # Create RPC-only frame
        rpc_df = df[df['Контактное лицо'] == 'Должник']
        summary = create_summary_pivot(df)
        pivot_all, errors = run(df)
        del df
        pivot_rpc, _= run(rpc_df)
        del rpc_df
        return pivot_all, pivot_rpc, summary, errors

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
        min_errors = kwargs['errors'][1]
        max_errors = kwargs['errors'][0]
        # Create a color scale conditional formatting rule
        color_scale_rule_errors = {
                        'type': '3_color_scale',
                        'min_color': '#A6D86E',  # Green
                        'mid_color': '#FCFAA0',  # White (for NaN)
                        'max_color': '#e85f5f',  # Red
                        'min_type': 'num',
                        'min_value': min_errors,
                        'mid_type': 'num',
                        'mid_value': (max_errors-min_errors)/4,
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
                # Apply the white background to the entire worksheet
                worksheet.set_column(0, 0, 5, white_fill_format)
                worksheet.set_column(1, 1, 25, white_index_format)
                worksheet.set_column(2, 2, 30, white_index_format)
                worksheet.set_column(3, 100, 13, white_fill_format)

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
    df = prep_data(csv_list=csv_list)
    df, rpc_df, summary, errors = create_pivot(df)
    format_xlsx(df, rpc_df, summary,
                name=output_report_path,
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