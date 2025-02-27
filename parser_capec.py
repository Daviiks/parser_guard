import pandas as pd
from bs4 import BeautifulSoup
import logging
from tqdm import tqdm
import requests
import time
import re
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler("D:\\parsing.log"),  # Логи в файл, путь свой указываете
        logging.StreamHandler()  # Логи в консоль, чтобы наглядно видеть что происходит
    ]
)

# Функция для извлечения номера страницы из URL
def extract_page_number(url):
    """Извлекает номер страницы из URL."""
    match = re.search(r'\d+', url)  # Ищем последовательность цифр в URL
    return match.group(0) if match else None

# Функция для извлечения таблиц и текста из div с определённым id
def extract_tables_and_div(url):
    """Извлекает таблицы и текст из div с id Likelihood_Of_Exploit."""
    try:
        # Извлечение номера страницы из URL
        page_number = extract_page_number(url)
        if not page_number:
            logging.warning(f"Номер страницы не найден в URL: {url}")
            return url, []

        # Загрузка HTML-кода страницы
        session = requests.Session()
        retries = Retry(total=5, backoff_factor=1, status_forcelist=[500, 502, 503, 504])
        session.mount("http://", HTTPAdapter(max_retries=retries))
        session.mount("https://", HTTPAdapter(max_retries=retries))
        response = session.get(url)
        response.raise_for_status()  # Проверка на ошибки HTTP

        # Загрузка таблиц с веб-страницы
        ds = pd.read_html(url)
        time.sleep(1)  # Задержка между запросами, можете убрать если не нужна

        # Ищем таблицу с определёнными заголовками
        target_headers = ["CAPEC-ID", "Attack Pattern Name"]
        found_tables = []

        for i, table in enumerate(ds):
            if all(header in table.columns for header in target_headers):
                found_tables.append(table)  # Таблица с CAPEC-ID и Attack Pattern Name

        return url, found_tables  # Возвращаем URL и таблицы
    except Exception as e:
        logging.error(f"Ошибка при обработке URL {url}: {e}")
        return url, []  # Возвращаем URL и пустой список таблиц

# Загрузка URL из Excel-файла
file_path = 'D:\\merged_file.xlsx' # Свой путь и тд
df = pd.read_excel(file_path)

# Проверка наличия столбца с URL
if 'NewUrl' not in df.columns:
    raise ValueError("Столбец 'NewUrl' не найден в файле.")

# Список для хранения всех найденных таблиц и текста
all_tables = []
visited_urls = set()
empty_urls = []

# Последовательная обработка URL
for url in tqdm(df['NewUrl'], desc="Обработка URL"):
    if url in visited_urls:
        logging.info(f"URL {url} уже обработан, пропускаем.")
        continue

    # Обработка URL
    url, found_tables = extract_tables_and_div(url)
    if found_tables:
        for table in found_tables:
            # Добавляем столбец с URL источника
            table['Source URL'] = url
        all_tables.extend(found_tables)
    else:
        empty_urls.append(url)  # Добавляем URL в список пустых
    
    visited_urls.add(url)  # Добавляем URL в множество обработанных

# Объединение всех найденных таблиц в один DataFrame
if all_tables:
    results_df = pd.concat(all_tables, ignore_index=True)  # Объединяем таблицы
    
    # Удаление дубликатов по столбцу CAPEC-ID
    if 'CAPEC-ID' in results_df.columns:
        results_df = results_df.drop_duplicates(subset=['CAPEC-ID'])
    else:
        logging.warning("Столбец 'CAPEC-ID' не найден в результатах.")
    
    logging.info("Найдены таблицы:")
    logging.info(results_df)
    # Сохранение всех данных в Excel
    output_path = 'D:\\parsed_results.xlsx' # Свой путь
    results_df.to_excel(output_path, index=False)
    logging.info(f"Все данные сохранены в файл: {output_path}")

else:
    logging.warning("Таблицы с указанными заголовками не найдены.")

# Сохранение списка пустых URL в файл, чтобы знать на каких страницах не было таблиц с CAPEC
if empty_urls:
    empty_urls_df = pd.DataFrame(empty_urls, columns=['Empty URLs'])
    empty_urls_path = 'D:\\empty_urls.xlsx' # Свой путь и название
    empty_urls_df.to_excel(empty_urls_path, index=False)
    logging.info(f"Список пустых URL сохранен в файл: {empty_urls_path}")
else:
    logging.info("Пустые URL не обнаружены.")