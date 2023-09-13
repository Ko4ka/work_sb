import argparse
import pandas as pd
import traceback
from openpyxl import Workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import PatternFill
from openpyxl.styles import numbers
from openpyxl.styles import Alignment

# Create a parser to handle command-line arguments
parser = argparse.ArgumentParser(description='Process CSV files and create an Excel pivot table with color scaling.')

# Add arguments for CSV list and output report path
parser.add_argument('--csv_list', nargs='+', help='List of CSV file paths', required=True)
parser.add_argument('--output_report_path', help='Path for the output Excel report', required=True)

# Parse the command-line arguments
args = parser.parse_args()

# Access the arguments using args.csv_list and args.output_report_path in your code

def transform(csv_list: list, output_report_path):
    def format_xlsx(pivot_table: pd.DataFrame, sheet: str = 'Отчет 1',
                    name: str = "pivot_table_2_call_lists.xlsx", **kwargs):
        # Specify the Excel file path
        excel_file_path = name
        min_percent = kwargs['percent'][1]
        max_percent = kwargs['percent'][0]
        min_errors = kwargs['errors'][1]
        max_errors = kwargs['errors'][0]
        # Create a color scale conditional formatting rule
        color_scale_rule_percent = ColorScaleRule(
            start_type="num",
            start_value=min_percent,         # Set the minimum value to 0
            start_color="A6D86E",  # Green for minimum value
            mid_type="percentile",
            mid_value=75,          # Set the midpoint to 50%
            mid_color="FCFAA0",    # Yellow for mid-value
            end_type = "num",
            end_value=max_percent,           # Set the maximum value to 1
            end_color="e85f5f"     # Red for maximum value
        )

        # Create a color scale conditional formatting rule
        color_scale_rule_errors = ColorScaleRule(
            start_type="num",
            start_value=min_errors,         # Set the minimum value to 0
            start_color="A6D86E",  # Green for minimum value
            mid_type="percentile",
            mid_value=75,          # Set the midpoint to 50%
            mid_color="FCFAA0",    # Yellow for mid-value
            end_type = "num",
            end_value=max_errors,           # Set the maximum value to 1
            end_color="e85f5f"     # Red for maximum value
        )

        # Create a Pandas Excel writer using xlsxwriter as the engine
        writer = pd.ExcelWriter(excel_file_path, engine='openpyxl')

        # Write the DataFrame to the Excel file
        mega_pivot.to_excel(writer, sheet_name='Sheet1')

        # Access the xlsxwriter workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # Apply the color scale conditional formatting by row

        for row_num, row in enumerate(mega_pivot.index):
            if "(доля ошибок %)" in row:
                format_range = 'B' + str(row_num + 2) + ':AZ' + str(row_num + 2)  # Adjust column range as needed
                worksheet.conditional_formatting.add(format_range, color_scale_rule_percent)

        for row_num, row in enumerate(mega_pivot.index):
            if "(доля ошибок %)" in row:
                for col_num, col in enumerate(mega_pivot.columns):
                    cell = worksheet.cell(row=row_num + 2, column=col_num + 2)  # Adjust row and column numbers as needed
                    cell.number_format = '0.00%'  # Set number format for percentage columns

        for row_num, row in enumerate(mega_pivot.index):
            if "(ошибки шт.)" in row:
                format_range = 'B' + str(row_num + 2) + ':AZ' + str(row_num + 2)  # Adjust column range as needed
                worksheet.conditional_formatting.add(format_range, color_scale_rule_errors)

        # Set alignment for the entire column A (column index 1) to left
        for cell in worksheet['A']:
            cell.alignment = Alignment(horizontal='left')
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
        # Add numbers to index
        def update_index(dataframe):
            new_index = [f'{i if len(str(i)) > 1 else f"0{i}"} {row}' for i, row in enumerate(dataframe.index)]
            dataframe.index = new_index
            return dataframe
        
        # Create dynamic Mistakes count (1)
        pivot_df_mistakes = df.pivot_table(index='Имя колл-листа', columns='Дата', values='Ошибки', aggfunc='sum')
        pivot_df_mistakes = pivot_df_mistakes.fillna(0)
        pivot_df_mistakes = pivot_df_mistakes.replace(0.00, '')
        pivot_df_mistakes.columns = pd.to_datetime(pivot_df_mistakes.columns, format='%d.%m.%Y')  # Fix date Time
        pivot_df_mistakes = pivot_df_mistakes.sort_index(axis=1)  # Fix date Time
        tmp_pivot_df_mistakes = pivot_df_mistakes.copy()  # Fix %%
        pivot_df_mistakes.index = pivot_df_mistakes.index + ' (ошибки шт.)'
        pivot_df_mistakes = update_index(pivot_df_mistakes)
        # Create dynamic Calls count (2)
        pivot_df_calls = df.pivot_table(index='Имя колл-листа', columns='Дата', values='Результат автооценки', aggfunc='count', fill_value=0)
        pivot_df_calls.columns = pd.to_datetime(pivot_df_calls.columns, format='%d.%m.%Y')  # Fix date Time
        pivot_df_calls = pivot_df_calls.sort_index(axis=1)  # Fix date Time
        tmp_pivot_df_calls = pivot_df_calls.copy()  # Fix %%
        pivot_df_calls.index = pivot_df_calls.index + ' (всего шт.)'
        pivot_df_calls = update_index(pivot_df_calls)
        # Create dynamic Mean Autoscore (3)
        pivot_df_mean = df.pivot_table(index='Имя колл-листа', columns='Дата', values='Результат автооценки', aggfunc='mean', fill_value='')
        pivot_df_mean.columns = pd.to_datetime(pivot_df_mean.columns, format='%d.%m.%Y')  # Fix date Time
        pivot_df_mean = pivot_df_mean.sort_index(axis=1)  # Fix date Time
        pivot_df_mean.index = pivot_df_mean.index + ' (средняя АО)'
        pivot_df_mean = update_index(pivot_df_mean)
        # Create dynamic Error Percentage (4)
        pivot_df_mistakes_filled = tmp_pivot_df_mistakes.replace('', 0)
        pivot_df_error_rate = (pivot_df_mistakes_filled / tmp_pivot_df_calls).applymap(lambda x: x if not pd.isna(x) else '')
        pivot_df_error_rate.index = pivot_df_error_rate.index + ' (доля ошибок %)'
        pivot_df_error_rate = update_index(pivot_df_error_rate)
        # Create Mega-Pivot
        # Concatenate the pivot tables vertically along rows (axis=0)
        mega_pivot = pd.concat([pivot_df_mistakes, pivot_df_calls, pivot_df_mean, pivot_df_error_rate], axis=0)
        mega_pivot = mega_pivot.sort_index()
        error_rate_rows = mega_pivot[mega_pivot.index.str.contains("(доля ошибок %)")]
        errors_rows = mega_pivot[mega_pivot.index.str.contains("(ошибки шт.)")]
        # Find min/max percent
        max_percent = error_rate_rows.apply(pd.to_numeric, errors='coerce').max().max()
        min_percent = error_rate_rows.apply(pd.to_numeric, errors='coerce').min().min()
        # Find min/max mistakes
        max_errors = errors_rows.apply(pd.to_numeric, errors='coerce').max().max()
        min_errors = errors_rows.apply(pd.to_numeric, errors='coerce').min().min()

        # Create excel File
        format_xlsx(mega_pivot, name=output_report_path, percent=[max_percent, min_percent], errors=[max_errors, min_errors])
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

