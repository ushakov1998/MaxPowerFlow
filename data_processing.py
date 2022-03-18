import json
import csv


def csv_to_list(path: str) -> [dict]:
    """
    Парсим .сsv в список словарей
    path: str путь до файла .csv
    """

    dict_list = []

    with open(path, newline='') as csv_data:
        csv_dic = csv.DictReader(csv_data)

        # Создание пустого списка и добавление рядов
        for row in csv_dic:
            dict_list.append(row)

    return dict_list


def json_to_dic(path: str) -> dict:
    """
    Парсим .json в словарь
    path: str путь до файла .json
    """

    with open(path, "r") as json_data:
        dictionary = json.load(json_data)

    return dictionary
