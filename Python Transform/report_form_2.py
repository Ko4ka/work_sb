import argparse
import pandas as pd
import traceback

# Create a parser to handle command-line arguments
parser = argparse.ArgumentParser(description='Process CSV files and create an Excel pivot table with color scaling.')

# Add arguments for CSV list and output report path
parser.add_argument('--csv_list', nargs='+', help='List of CSV file paths', required=True)
parser.add_argument('--output_report_path', help='Path for the output Excel report', required=True)

# Parse the command-line arguments
args = parser.parse_args()

# Access the arguments using args.csv_list and args.output_report_path in your code

def transform(csv_list: list, output_report_path):
    def format_xlsx(pivot_table: pd.DataFrame, sheet: str = 'Общий',
                    name: str = "pivot_table_2_call_lists.xlsx", **kwargs):
        # Specify the Excel file path
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

        # Create a Pandas Excel writer using xlsxwriter as the engine
        writer = pd.ExcelWriter(excel_file_path, engine='xlsxwriter')

        # Write the DataFrame to the Excel file
        pivot_table.to_excel(writer, sheet_name=sheet)

        # Access the xlsxwriter workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Общий']
        # Define a white fill format
        white_fill_format = workbook.add_format({'text_wrap': True, 'bg_color': '#FFFFFF', 'border': 0})
        # Apply the white background to the entire worksheet
        worksheet.set_column(0, 0, 25, white_fill_format)
        worksheet.set_column(1, 1, 30, white_fill_format)
        worksheet.set_column(2, 100, 9, white_fill_format)

        percentage_format = workbook.add_format({'num_format': '0.00%', 'bg_color': '#FFFFFF', 'border': 0})

        for i, j in enumerate(pivot_table.head()):
            if j[1] == 'Ошб.%':
                worksheet.set_column(i+2, i+2, None, percentage_format)
                worksheet.conditional_format(0, i+2, 999, i+2, color_scale_rule_errors)

        # Save the Excel file
        writer.save()
        print(f'file: {name} -- Transformed 0')
    try:
        # Concatenate all csv to a single biiig df
        df = pd.DataFrame()
        for i in csv_list:
            df_add = pd.read_csv(i, sep=';', header=0)
            df = pd.concat([df, df_add], ignore_index=True)
        # Fix мультидоговоры for RSB
        mask = df['№ п/п'].isna()
        df = df[~mask]
        # Convert the 'Длительность звонка' column to Timedelta
        df['Длительность звонка'] = pd.to_timedelta(df['Длительность звонка'])
        df['Ошибки'] = df['Результат автооценки'] != 100
        df['Дата'] = df['Дата звонка'].str.split(' ').str[0]
        df = df.reset_index(drop=True)
        # 2. Create the pivot table
        def create_multiindex(dataframe, sub_index:str):
            # Create MultiIndex
            multiindex = []
            for i, column in enumerate(dataframe):
                multiindex.append((column, sub_index))
            dataframe.columns = pd.MultiIndex.from_tuples(multiindex)
            return dataframe

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
        dfs_to_merge = [pivot_df_calls, pivot_df_errors, pivot_df_mean, pivot_df_error_rate]

        # Initialize an empty DataFrame with the same index as the original DataFrames
        merged_df = pd.DataFrame(index=pivot_df_calls.index)

        '''CREATE MULTIINDEX'''
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
        # Create excel File
        format_xlsx(merged_df, name=output_report_path, errors=[max_error_rate, min_error_rate])
        print('Exit Code 0')
        return 0
    except ValueError or KeyError:
        traceback.print_exc()
        print('Exit Code 1 (Pandas Error)')
        return 1
    except Exception:
        traceback.print_exc()
        print('Exit Code 2 (Unknown Error)')
        return 2

if __name__ == '__main__':
    transform(csv_list=args.csv_list, output_report_path=args.output_report_path)

