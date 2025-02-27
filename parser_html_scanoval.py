from bs4 import BeautifulSoup #pip install openpyxl bs4 Если нужно
from openpyxl import Workbook

def html_to_excel(html_file, excel_file):
    """Извлекает данные из HTML-файла и записывает их в Excel-файл."""

    with open(html_file, 'r', encoding='utf-8') as f:
        html_content = f.read()

    soup = BeautifulSoup(html_content, 'lxml')

    table = soup.find('table', class_='vulnerabilitiesTbl')
    if table:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.append(['Название', 'Уровень опасности','Описание'])

        for row in table.find_all('tr'):
            row_data = []
            cells = row.find_all('td')
            for cell in cells:
                if cell.get('class') == ['bdu']:
                    row_data.append(cell.text.strip())
                elif cell.get('class') == ['risk riskTextColor']:
                    row_data.append(cell.text.strip())
                elif cell.get('class') == ['desc']:
                    row_data.append(cell.text.strip())
            if row_data: 
                worksheet.append(row_data)

        workbook.save(excel_file)
        print(f"Данные успешно записаны в {excel_file}")
    else:
        print("Таблица не найдена в HTML-файле.")

html_file = 'D:\ScanOval\ScanOval_Report_20_02_2025_new.html' #Прописываете свой путь, какой хотите
excel_file = 'output.xlsx' # И название файла на свое усмотрение
html_to_excel(html_file, excel_file)