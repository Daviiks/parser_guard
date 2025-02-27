import pandas as pd
from bs4 import BeautifulSoup as bs
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
        logging.FileHandler("D:\\parsing.log"),  # Логи в файл. Путь свой
        logging.StreamHandler()  # Логи в консоль
    ]
)

def extract_page_number(url):
    """Извлекает номер страницы из URL."""
    match = re.search(r'\d+', url)  # Ищем последовательность цифр в URL
    return match.group(0) if match else None

def extract_tables_and_div(url):
    """Извлекает таблицы и текст из div с id Likelihood_Of_Attack."""
    try:
        # Извлечение номера страницы из URL
        page_number = extract_page_number(url)
        if not page_number:
            logging.warning(f"Номер страницы не найден в URL: {url}")
            return url, None  # Возвращаем URL и None, если номер страницы не найден

        # Загрузка HTML-кода страницы
        session = requests.Session()
        retries = Retry(total=5, backoff_factor=1, status_forcelist=[500, 502, 503, 504])
        session.mount("http://", HTTPAdapter(max_retries=retries))
        session.mount("https://", HTTPAdapter(max_retries=retries))
        response = session.get(url)
        response.raise_for_status() 
        soup = bs(response.text, 'lxml')

        # Формирование id с номером страницы
        capec_id = f'oc_{page_number}_Likelihood Of Attack'
        likelihood_div = soup.find('div', id=capec_id)
        likelihood_text = likelihood_div.get_text(strip=True) if likelihood_div else 'No chance'

        time.sleep(1)  # Задержка между запросами, можете убрать если не нужно
        return url, likelihood_text  # Возвращаем URL и текст из div
    except Exception as e:
        logging.error(f"Ошибка при обработке URL {url}: {e}")
        return url, None  # Возвращаем URL и None в случае ошибки

# Загрузка URL из Excel-файла
file_path = 'D:\\capec_url.xlsx' # свой путь и название
df = pd.read_excel(file_path)

# Проверка наличия столбца с URL
if 'UrlCapec' not in df.columns:
    raise ValueError("Столбец 'UrlCapec' не найден в файле.")

# Список для хранения данных
likelihood_data = []
visited_urls = set()

# Последовательная обработка URL
for url in tqdm(df['UrlCapec'], desc="Обработка URL"):
    if url in visited_urls:
        logging.info(f"URL {url} уже обработан, пропускаем.")
        continue

    # Обработка URL
    url, likelihood_text = extract_tables_and_div(url)

    # Добавляем значение Likelihood_Of_Attack (или None, если его нет)
    likelihood_data.append({'URL': url, 'Likelihood_Of_Attack': likelihood_text})
    visited_urls.add(url)  # Добавляем URL в множество обработанных

# Сохранение данных в Excel
if likelihood_data:
    likelihood_df = pd.DataFrame(likelihood_data)
    output_path_likelihood = 'D:\\likelihood_data_capec.xlsx' # Свой путь и название
    likelihood_df.to_excel(output_path_likelihood, index=False)
    logging.info(f"Данные из div сохранены в файл: {output_path_likelihood}")
else:
    logging.warning("Данные не найдены.")