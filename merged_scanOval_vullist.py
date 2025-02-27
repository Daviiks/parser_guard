import pandas as pd

def merge_excel_files(file1_path, file2_path, output_path, merge_column):
    """Объединяет два Excel-файла по указанному столбцу."""
    try:
        df1 = pd.read_excel(file1_path)
        df2 = pd.read_excel(file2_path)
        # Проверка на наличие столбца для объединения
        if merge_column not in df1.columns or merge_column not in df2.columns:
            raise ValueError(f"Столбец '{merge_column}' отсутствует в одном или обоих файлах.")
        # Объединение файлов по указанному столбцу (внешнее объединение, чтобы сохранить все строки)
        merged_df = pd.merge(df1, df2, on=merge_column, how='inner')
        merged_df.to_excel(output_path, index=False)
        print(f"Файлы успешно объединены в {output_path}")

    except FileNotFoundError:
        print("Один или оба файла не найдены.")
    except ValueError as e:
        print(f"Ошибка: {e}")
    except Exception as e:
        print(f"Произошла непредвиденная ошибка: {e}")

# Указываете свой путь и название файла.
file1_path = 'D:\\output1.xlsx'
file2_path = 'D:\\vullist1.xlsx'
output_path = 'D:\\merged_file.xlsx'
merge_column = 'Название'

merge_excel_files(file1_path, file2_path, output_path, merge_column)


