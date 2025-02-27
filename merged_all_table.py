import pandas as pd

# Загрузка данных. Указываете свой путь, название файла, названия своих столбцов в excel и тд
merged_files = pd.read_excel('D:\\merged_file.xlsx')
capec_url = pd.read_excel('D:\\capec_url.xlsx')
likelihood_attack = pd.read_excel('D:\\likelihood_data_capec.xlsx')

# Извлечение номера CWE из ссылки в capec_url
capec_url['CWE'] = capec_url['Source URL'].str.extract(r'/(\d+)$').fillna('Unknown')

# Извлечение номера CAPEC из ссылки в likelihood_attack
likelihood_attack['CAPEC_ID'] = likelihood_attack['URL'].str.extract(r'/(\d+)$').fillna('Unknown')

# Приведение столбцов к единому формату
capec_url['CAPEC-ID'] = capec_url['CAPEC-ID'].str.extract(r'CAPEC-(\d+)').fillna('Unknown')
likelihood_attack['CAPEC_ID'] = likelihood_attack['CAPEC_ID'].astype(str)

# Объединение capec_url и likelihood_attack по CAPEC ID
capec_likelihood = pd.merge(capec_url, likelihood_attack, left_on='CAPEC-ID', right_on='CAPEC_ID', how='outer')

# Приведение столбца CWE к единому формату
merged_files['CWE'] = merged_files['Тип ошибки CWE'].str.extract(r'CWE-(\d+)').fillna('Unknown')
merged_files['CWE'] = merged_files['CWE'].astype(str)
capec_likelihood['CWE'] = capec_likelihood['CWE'].astype(str)

# Объединение с merged_files по CWE
merged_data = pd.merge(merged_files, capec_likelihood, left_on='CWE', right_on='CWE', how='outer')

# Группировка данных по Названию и CWE
def group_capec_by_likelihood(df):
    high = df[df['Likelihood_Of_Attack'] == 'High']['CAPEC_ID'].dropna().unique()
    medium = df[df['Likelihood_Of_Attack'] == 'Medium']['CAPEC_ID'].dropna().unique()
    low = df[df['Likelihood_Of_Attack'] == 'Low']['CAPEC_ID'].dropna().unique()
    no_chance = df[df['Likelihood_Of_Attack'] == 'No chance']['CAPEC_ID'].dropna().unique()
    
    return pd.Series({
        'CAPEC_High': ', '.join(map(str, high)),
        'CAPEC_Medium': ', '.join(map(str, medium)),
        'CAPEC_Low': ', '.join(map(str, low)),
        'CAPEC_No_Chance': ', '.join(map(str, no_chance))
    })

# Применение группировки
result = merged_data.groupby(['Название', 'CWE']).apply(group_capec_by_likelihood, include_groups=False).reset_index()

# Сохранение результата. Свой путь и название файла
result.to_excel('D:\\final_result.xlsx', index=False)
print('Файл успешно создан или нет, или да. Я хз')