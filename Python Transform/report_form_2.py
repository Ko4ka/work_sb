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
        pivot_all, errors = run(df)
        del df
        pivot_rpc, _= run(rpc_df)
        del rpc_df
        return pivot_all, pivot_rpc, errors

    '''Excel Code'''
    def format_xlsx(pivot_all: pd.DataFrame,
                    pivot_rpc: pd.DataFrame,
                    name: str = "pivot_table_2_call_lists.xlsx", **kwargs):
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
                white_fill_format = workbook.add_format({'text_wrap': True, 'bg_color': '#FFFFFF', 'border': 0})
                # Apply the white background to the entire worksheet
                worksheet.set_column(0, 0, 25, white_fill_format)
                worksheet.set_column(1, 1, 35, white_fill_format)
                worksheet.set_column(2, 100, 10, white_fill_format)
                percentage_format = workbook.add_format({'num_format': '0.00%', 'bg_color': '#FFFFFF', 'border': 0})

                for i, j in enumerate(pivot_table.head()):
                    worksheet.set_column(i+2, i+2, 10, white_fill_format)
                    if j[1] == 'Ошб.%':
                        worksheet.conditional_format(2, i+2, 999, i+2, color_scale_rule_errors)
                        worksheet.set_column(i+2, i+2, None, percentage_format)

                      

            # Create Sheets
            create_sheet(pivot_all, 'Все звонки')
            create_sheet(pivot_rpc, 'RPC')
            # Create Summary Sheet
        print(f'file: {name} -- Transformed 0')

    '''Run Script'''
    df = prep_data(csv_list=csv_list)
    df, rpc_df, errors = create_pivot(df)
    format_xlsx(df, rpc_df,
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