import pandas as pd
import requests
import json
import logging

COMPANY_FIELD_INN = 'UF_CRM_1721163777924'
DUPLICATE_FIELD = 'UF_CRM_1721320084'

file_path = 'outputx.xlsx'


# Подготовить функцию на замену ИНН для определенной компании, у которой мы знаем ID
def update_company_inn(company_id, inn):
    global inn_field

    url = 'https://bitrix.aliton.ru/rest/2460/pbcohltiapum5nxf/crm.company.update.json'
    data = {
        'id': company_id,  # Указываем ID компании, которую нужно обновить
        'fields': {
            inn_field: inn  # Указываем поле, которое нужно обновить, и новое значение
        }
    }
    response = requests.post(url, json=data)
    return response.json()


# Подготовить функцию на получение определенной компании, у которой мы знаем ID
def get_company_info(company_id):
    url = 'https://bitrix.aliton.ru/rest/2460/vr0vhfbl4cs6qttf/crm.company.get.json'
    data = {
        'id': company_id,  # Указываем ID компании, которую нужно обновить
    }
    response = requests.post(url, json=data)

    # Проверяем, удалось ли выполнить запрос
    if response.status_code != 200:
        print(f"Ошибка при выполнении запроса: {response.status_code} для ID компании: {company_id}")
        return None

    result = response.json()

    # Проверяем, есть ли в ответе ошибка
    if 'error' in result:
        print(f"Ошибка: {result['error_description']}")
        return None

    return result


#  Подготовить функцию на обработку определенной компании, у которой мы знаем ID
def try_update_company_info(company_id, company_name, inn):
    global inn_field

    company_info = get_company_info(company_id)

    if company_info is None or 'result' not in company_info:
        print(f"Компания с ID {company_id} не найдена.")
        return

    company_data = company_info['result']

    if inn_field in company_data and company_data[inn_field]:
        print(f"Компания с ID {company_id} ({company_name}) уже имеет ИНН: {company_data[inn_field]}")
    else:
        update_result = update_company_inn(company_id, inn)
        if update_result is None or 'result' not in update_result:
            print(f"Не удалось обновить ИНН для компании с ID {company_id} ({company_name})")
        else:
            print(f"ИНН для компании с ID {company_id} ({company_name}) успешно обновлен на {inn}")


# Подготовить функцию на обработку всех компаний из таблички
def process_companies_from_excel(file_path):
    df = pd.read_excel(file_path)

    for index, row in df.iterrows():
        company_id = row['ID']
        company_name = row['Название компании']
        inn = row['ИНН']

        if pd.notna(inn) and pd.notna(company_id):
            if inn != '' and company_id != '':
                try_update_company_info(company_id, company_name, inn)




# Функция для поиска компании по ИНН
def search_company_by_inn(inn):
    search_url = 'https://bitrix.aliton.ru/rest/2460/qme0m0g41lulfzjj/crm.company.list.json'
    params = {
        'filter': {inn_field: inn},
        'select': ['ID']
    }
    response = requests.post(search_url, json=params)  # Изменен на POST с передачей JSON в теле запроса
    print(f"Поиск компании по ИНН {inn}: {response.status_code} {response.text}")  # Отладочный вывод

    if response.status_code == 200:
        result = response.json()
        if result and 'result' in result and len(result['result']) > 0:
            return result['result'][0]  # Возвращаем первую найденную компанию
    return None


# Функция для добавления компании из Excel
def company_add_from_excel(file_path):
    df = pd.read_excel(file_path)
    url = 'https://bitrix.aliton.ru/rest/2460/qme0m0g41lulfzjj/crm.company.add.json'

    for index, row in df.iterrows():
        company_name = row['Название компании']
        inn = str(row['ИНН']).split('.')[0].strip()  # Удаляем дробную часть и пробелы

        # Проверка существования компании по ИНН
        existing_company = search_company_by_inn(inn)

        if existing_company:
            print(f"Компания с ИНН {inn} уже существует. Пропускаем добавление.")
            continue

        # Данные для добавления новой компании
        data = {
            'fields': {
                inn_field: inn,
                "TITLE": company_name,
                responsible_field: '743'  # Добавляем ответственного
            }
        }
        response = requests.post(url, json=data)

        if response.status_code == 200:
            print(f"Компания {company_name} добавлена успешно.")
        else:
            print(f"Ошибка при добавлении компании {company_name}: {response.text}")





def get_all_companies():
    start = 0
    companies = []
    while True:
        print(f"Запрос компаний с позиции {start}...")
        response = requests.get(
            'https://bitrix.aliton.ru/rest/2460/h80q1pv568jdald1/crm.company.list.json',
            params={
                'start': start,
                'select': ['ID', COMPANY_FIELD_INN]
            }
        )
        result = response.json().get('result', [])
        companies.extend(result)
        print(f"Получено {len(result)} компаний.")
        if len(result) < 50:
            break
        start += 50
    print(f"Всего получено {len(companies)} компаний.")
    return companies

def mark_duplicates(companies):
    inn_counts = {}
    for company in companies:
        inn = company.get(COMPANY_FIELD_INN)
        if inn:
            if inn in inn_counts:
                inn_counts[inn].append(company['ID'])
            else:
                inn_counts[inn] = [company['ID']]

    duplicate_count = 0
    for inn, ids in inn_counts.items():
        if len(ids) > 1:
            duplicate_count += len(ids)
            print(f"Найден дубликат для ИНН {inn}: компании {ids}")
            for company_id in ids:
                update_company(company_id, {DUPLICATE_FIELD: 'Да'})
    print(f"Обработано {duplicate_count} компаний с дубликатами.")

def update_company(company_id, fields):
    print(f"Обновление компании {company_id}...")
    response = requests.post(
        'https://bitrix.aliton.ru/rest/2460/pbcohltiapum5nxf/crm.company.update.json',
        params={
            'ID': company_id,
            'fields': fields
        }
    )
    result = response.json()
    if result.get('result', False):
        print(f"Компания {company_id} успешно обновлена.")
    else:
        print(f"Ошибка при обновлении компании {company_id}: {result}")
    return result

def main():
    print("Начало процесса проверки дубликатов...")
    companies = get_all_companies()
    mark_duplicates(companies)
    print("Процесс завершен.")

if __name__ == '__main__':
    main()