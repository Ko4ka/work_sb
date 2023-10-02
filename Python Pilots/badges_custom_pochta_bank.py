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
NAME = 'badges_custom_pochta_bank.py'

def transform(csv_list: list, output_report_path):
    def construct_marker_matrix(csv_list):
        # Base DF
        df = pd.DataFrame()
        for i in csv_list:
            '''
            Take report files 1-by-1 and the merge then on external index from indices.py
            This will cut RAM cost 30 times (and make shit slower)
            '''
            # Merge 2 frames
            df_add = pd.read_csv(i, sep=';', encoding='utf-8',header=0)
            df = pd.concat([df, df_add], ignore_index=True)

        # Forward fill NaN values in 'Маркер' column
        df.fillna(method='ffill', inplace=True)
        # Create a pivot table with 'Маркер' as columns and 'Маркер - количество совпадений' as values
        pivot_df = df.pivot_table(index=['№ п/п', 'ID звонка', 'Оператор', 'Дата звонка', 'Длительность звонка', 'Всего пауз, сек'], columns='Маркер', values='Маркер - количество совпадений').reset_index()
        # Reset the index and rename the columns
        pivot_df.columns.name = None  # Remove the columns' name
        # CREATED A MATRIX FORMAT
        pivot_df = pivot_df.fillna(0)
        pivot_df['Длительность звонка'] = pd.to_timedelta(pivot_df['Длительность звонка'])
        pivot_df['Дата'] = pd.to_datetime(pivot_df['Дата звонка'], format='%d.%m.%Y %H:%M:%S')
        pivot_df['Дата'] = pivot_df['Дата'].dt.strftime('%d.%m.%Y')
        pivot_df = pivot_df.reset_index(drop=True)
        return pivot_df
    
    def construct_dfs_custom(df):
        '''
        df = marker matrix
        return = a pivot table Operatos/each Marker count
        '''
        # Свожу универсальную простыню с маркерами
        pivot_df = df.pivot_table(index='Оператор', aggfunc='sum')
        # TABS
        # Blocks dfs
        blocks = [0, [], [], [], [], [], [], []]
        for i in df.columns:
            for j in range(1, 8):
                if str(j) in i[:1]: 
                    blocks[j].append(i)
        # Politeness
        products = [
            'ИИС+ПИФ',
            'ИСЖ',
            'Категория КК',
            'Моя карта ',
            'Мультибонус',
            'Вклады',
            'Вездедоход (КК)'
        ]
        # Warning
        negativity = [
        'Слова-паразиты',
        'Нарушение стандартов взаимодействия. Провоцирование конфликтов. ', 
        'Нежелательные выражения',
        'Нецензурная лексика']
        # Start
        politeness = [
        'Соблюдение стандартов взаимодействия. Вежливость+эмпатия в процессе обслуживания (Стандарты)',
        'Соблюдение стандартов взаимодействия. Положительные эмоции',
        'Соблюдение стандартов взаимодействия. Эмпатия.',
        'Эмпатия и Заинтересованность'
        ]
        # Time management
        total_time_df = df.pivot_table(index='Оператор', values='Длительность звонка', aggfunc='sum')
        # Convert timedelta to seconds and then to hours
        total_time_df['Длительность звонка (часы)'] = total_time_df['Длительность звонка'].dt.total_seconds() / 3600
        total_time_df.columns = ['Всего записи', 'Всего записи (часы)']
        total_wait_df = df.pivot_table(index='Оператор', values='Всего пауз, сек', aggfunc='sum') / 3600
        total_wait_df.columns = ['Всего пауз (часы)']
        # TAB 1 SUMMARY
        block_1 = pivot_df[blocks[1]].sum(axis=1)
        block_1.name = 'Блок 1: Установка контакта'
        block_2 = pivot_df[blocks[2]].sum(axis=1)
        block_2.name = 'Блок 2: Выявление и формирование потребности'
        block_3 = pivot_df[blocks[3]].sum(axis=1)
        block_3.name = 'Блок 3: Презентация основного продукта'
        block_4 = pivot_df[blocks[4]].sum(axis=1)
        block_4.name = 'Блок 4: Работа с возражениями'
        block_5 = pivot_df[blocks[5]].sum(axis=1)
        block_5.name = 'Блок 5: Завершение продажи'
        block_6 = pivot_df[blocks[6]].sum(axis=1)
        block_6.name = 'Блок 6: Кросс- продажа'
        block_7 = pivot_df[blocks[7]].sum(axis=1)
        block_7.name = 'Блок 7: Завершение встречи'
        script_df = pd.concat([block_1, block_2, block_3, block_4, block_5, block_6, block_7], axis=1)
        # TAB 2 ПРОДУКТЫ
        products_df = pivot_df[products]
        # TAB 2 НЕГАТИВ
        negative_df = pivot_df[negativity]
        # TAB 3 Вежливость
        politeness_df = pivot_df[politeness]
        # TAB 0 Свод
        summary_df = pd.concat([total_time_df, total_wait_df], axis=1)
        summary_df['К активности'] = 1 - (summary_df['Всего пауз (часы)'] / summary_df['Всего записи (часы)'])
        k_products = block_7 / summary_df['Всего записи (часы)']
        k_products.name = 'К кросс-продаж*'
        k_conformity = script_df.sum(axis=1) / summary_df['Всего записи (часы)']
        k_conformity.name = 'К соответсвия**'
        k_negativity = negative_df.drop(columns=['Слова-паразиты']).sum(axis=1) / summary_df['Всего записи (часы)']
        k_negativity.name = 'К негатива***'
        summary_df = pd.concat([summary_df, k_products, k_conformity, k_negativity], axis=1)
        del summary_df['Всего записи']

        return [summary_df, script_df, products_df, negative_df, politeness_df, pivot_df]

    '''Excel Code'''
    def format_xlsx(dataframes,
                    name: str = "pivot_table_2_call_lists.xlsx", **kwargs):
        color_scale_rule_percent = {
                        'type': '2_color_scale',
                        'min_color': '#FFFFFF',  # White
                        'max_color': '#8EA9DB',  # Red
                        'min_type': 'percentile',
                        'min_value': 0,
                        'max_type': 'percentile',
                        'max_value': 100
                        }
        excel_file_path = name
        # Create a Pandas Excel writer using xlsxwriter as the engine
        with pd.ExcelWriter(name, engine='xlsxwriter') as writer:
            def create_sheet(pivot_df, sheet_name, summary=False):
                # Write the DataFrame to the Excel file
                pivot_df.to_excel(writer, sheet_name=sheet_name)
                # Access the xlsxwriter workbook and worksheet objects
                workbook = writer.book
                worksheet = writer.sheets[sheet_name]
                white_fill_format = workbook.add_format({'text_wrap': False, 'bg_color': '#FFFFFF', 'border': 0})
                bold_format = workbook.add_format({'italic': True, 'text_wrap': False, 'bg_color': '#FFFFFF', 'border': 0})
                worksheet.set_column(0, 200, 20, white_fill_format)
                for i in range(1, pivot_df.shape[1]+1):
                    worksheet.conditional_format(1, i, 999, i, color_scale_rule_percent)
                if summary:
                    worksheet.write(f'A{pivot_df.shape[0]+3}', '* - кол-во кросс-продаж (согласно маркерам) в час', bold_format)
                    worksheet.write(f'A{pivot_df.shape[0]+4}', '** - кол-во ключевых фраз из методологических материалов (согласно маркерам) в час', bold_format)
                    worksheet.write(f'A{pivot_df.shape[0]+5}', '*** - кол-во маркеров негатива в час', bold_format)
            # Create Excel TABS
            create_sheet(dataframes[0], 'Сводный отчет', summary=True)        
            create_sheet(dataframes[1], 'Следование скрипту')
            create_sheet(dataframes[2], 'Продукты банка')
            create_sheet(dataframes[3], 'Негатив')
            create_sheet(dataframes[4], 'Вежливость')
            create_sheet(dataframes[5], 'Все маркеры')
        print(f'file: {name} -- Transformed 0')

    '''Run Script'''
    all_dfs = construct_dfs_custom(construct_marker_matrix(csv_list))
    format_xlsx(all_dfs, name=output_report_path)
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