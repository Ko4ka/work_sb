# Кастомные отчеты (Трансформеры)

# Dependencies
```
RUN pip install pandas
RUN pip install xlsxwriter
```
В принципе, можно использовать любой питон моложе 3.7.6. Проверено на 3.7.9 и 3.8

# Описание директории
1. transformer_config.json - конфиг файл, необходим для интеграции пользовательской части РА с трансформерами
```
[
  {
    "name": "НАЗВАНИЕ ФОРМЫ, КОТОРОЕ ВИДИТ ПОЛЬЗОВАТЕЛЬ",
    "script": "НАЗВАНИЕ .py СКРИПТА",
    "description": "ТЕКСТОВОЕ ОПИСАНИЕ ФОРМЫ, КОТОРОЕ ВИДИТ ПОЛЬЗОВАТЕЛЬ"
  },
  ...
  ]
```
2. transform_logs.log - файл для логов. Важный момент, т.к. python вызывается в контейнере через exec() бэка, ошибки формата ООМ или некоректно переданные аргументы логироваться не будут.

3. indices.py - кастомные индексы для pd.Dataframe, которые используются в скриптах

4. %name%.py - сами скрипты для трансформации

# Описание скриптов

1. Argparse - python парсер аргументов, переданных извне ([документация](https://docs.python.org/3/library/argparse.html)). Позволяет передать аргументы с бэка в формате `--csv_list` - список CSV файлов для трансформации + `--output_report_path` - путь для сохранения xlsx файла. Общий формат одного скрипта следующий:
```
import...
...
def transform(csv_list: list, output_report_path):
  def construct_df(csv_list):
      ...
  def format_xlsx(dfs, **kwargs):
      ...
      
  # Run Script
  df_main, df_rpc = construct_df(csv_list=csv_list)
  df_summary = construct_summary(dfs)
  # Sort columns
  df = df.sort_index(axis=1)
  format_xlsx(df, **COLORS)
  return 0

if __name__ == '__main__':
    try:
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

```

2. `transform(csv_list=args.csv_list, output_report_path=args.output_report_path)` - общая функция для трансформации файла, принимает аргументы, ничего не возвращает.

3. Внутри `transform()`, как правило 2 подфункции `construct_df(csv_list)` и `def format_xlsx():`. Первая строит сводную таблицу с помощью pandas, вторая раскрашивает и форматирует Excel-файл. В зависимости от задачи, содержимое функций меняется.

4. Большие данные. При работе с временными рядами за промежуток времени более недели, кол-во строк может достигнуть 1.5 млрд. Это много, для этого в ряде отчетов используется процессинг данных по дням:
```
df_main = pd.DataFrame()
for iteration, i in enumerate(csv_list):
            '''
            Take report files 1-by-1 and the merge then on external index from indices.py
            This will cut RAM cost 30 times (and make shit slower)
            '''
            # Merge 2 frames
            df = pd.read_csv(i, sep=';', encoding='utf-8',header=0)
            # Remove мультидоговоры for RSB
            ...
            # Warn if dates != 1
            if len(df['Дата'].unique().tolist()) > 1:
                logger.warning('%s Warning: more than a single date in df...', datetime.datetime.now())
            # MEMORY MANAGEMENT: CONCAT TO INDEX AND DELETE
            main_pivot = create_pivot(df)
            df_main = pd.concat([df_main, main_pivot], axis=1) # !!!
```
При таком подходе, необходимо отфильтровать сводную таблицу по дате перед форматированием в Excel:
```
df_main, df_rpc = construct_df(csv_list=csv_list)
    # Sort columns
    df_main = df_main.sort_index(axis=1)  # !!!
    format_xlsx(df_main.replace(0, pd.NA),
                name=output_report_path,
                enable_filtering=True)
```
Дополнительной оптимизацией может служить перекодировка соответствующих ячеек с `float64` на `Int8`:
```
for col in queries_list:
                df[col] = df[col].astype('Int8')
```
- `Int8`` а не `int8` , т.к. последний не поддерживает NaN и пустую строку ''

5. Скрипт логирует Exception-ы и другие события для информативности `logger.info(f'%s {NAME}: iteration #{iteration} done...', datetime.datetime.now())`

6. Дополнительно

- Работа с [xlsxwritter](https://xlsxwriter.readthedocs.io/index.html)
