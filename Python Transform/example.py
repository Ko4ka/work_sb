import pandas as pd
import traceback

import subprocess
import sys

def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

def transform(csv_list: list, output_report_path):
    # Format pd.DF as xlsx
    def format_xlsx(pivot_table: pd.DataFrame, sheet: str = 'Отчет 1',
                    name: str = "pivot_table_gradient_colorscale.xlsx"):
        # Create an Excel writer and export the pivot table to an Excel file
        excel_file_path = name
        with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
            pivot_table.to_excel(writer, sheet_name=sheet, index=True)

            # Get the xlsxwriter workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets[sheet]

            # Get the dimensions of the pivot table
            max_row = len(pivot_table)
            max_col = len(pivot_table.columns)

            # Add a format for the header cells
            header_format = workbook.add_format(
                {'bold': True, 'text_wrap': True, 'valign': 'top', 'border': 1, 'bg_color': '#EFEFEF',
                 'align': 'center'})

            # Set the column width and format for the header
            for col_num, value in enumerate(pivot_table.columns.values):
                worksheet.write(0, col_num + 1, value, header_format)
                column_len = max(pivot_table[value].astype(str).str.len().max(), len(value)) + 2
                worksheet.set_column(col_num + 1, col_num + 1, column_len)

            # Apply gradient color scale to value cells
            # https://xlsxwriter.readthedocs.io/working_with_conditional_formats.html
            worksheet.conditional_format(1, 1, max_row, max_col, {
                'type': '3_color_scale',
                'min_color': '#A6D86E',  # Green
                'mid_color': '#FFFFFE',  # White (for NaN)
                'max_color': '#e85f5f',  # Red
                'min_type': 'num',
                'mid_type': 'num',
                'max_type': 'num'
            })
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
        # 1. Filter the dataframe
        filtered_df = df[df['Результат автооценки'] != 100]  # !!!
        # filtered_df = df # !!!
        # 2. Create the pivot table
        pivot_table = filtered_df.pivot_table(index='Имя колл-листа',
                                            columns='Результат робота',
                                            values='Результат автооценки',
                                            aggfunc='size',
                                            fill_value=0)
        # Convert to xlsx and apply colorscale
        format_xlsx(pivot_table, name=output_report_path)
        print('Exit Code 0')
    except ValueError or KeyError:
        traceback.print_exc()
        print('Exit Code 1 (Pandas Error)')
    except Exception:
        traceback.print_exc()
        print('Exit Code 2 (Unknown Error)')



if __name__ == '__main__':
    install('xlsxwriter')
    transform(csv_list=['C:/Users/Alex/Work_python/work_sb/Python Transform/Reports/01-08.csv'],
              output_report_path="pivot_table_gradient_colorscale.xlsx")
