import calculation
import data_processing
import create_files
import os

# Основные переменные
faults_lines = None
flowgate_lines = None
trajectory_nodes = None
p_fluctuations = 30

# Ввод файла flowgate.json с указанием директории
try:
    flowgate_lines = data_processing.json_to_dic(input("Укажите путь к "
                                                    "flowgate .json: "))
except:
    print("Ошибка! Неизвестная (директория,тип файла) или файл не содержится в данной директории")
    exit()

# Ввод файла faults.json с указанием директории
try:
    faults_lines = data_processing.json_to_dic(input("Укажите путь к "
                                                  "faults .json: "))
except:
    print("Ошибка! Неизвестная (директория/тип файла) или файл не содержится в данной директории")
    exit()

# Ввод файла vector.csv с указанием директории
try:
    trajectory_nodes = data_processing.csv_to_list(input("Укажите путь к "
                                                      "vector.csv: "))
except:
    print("Ошибка! Неизвестная (директория/тип файла) или файл не содержится в данной директории")
    exit()

# Создание файлов КС и траектории утяжеления RastrWin3
create_files.create_file_sch(flowgate_lines, 'flow_gate_name')
create_files.create_file_ut2(trajectory_nodes)

# Вычисление МДП по критериям
print(f"МДП по 1 критерию: "
      f"P(мдп1)= {calculation.criterion1(p_fluctuations)} МВт")
print(f"МДП по 2 критерию: "
      f"P(мдп2) = {calculation.criterion2(p_fluctuations)} МВт")
print(f"МДП по 3 критерию: "
      f"P(мдп3)= {calculation.criterion3(p_fluctuations, faults_lines)} МВт")
print(f"МДП по 4 критерию: "
      f"P(мдп4) = {calculation.criterion4(p_fluctuations, faults_lines)} МВт")
print(f"МДП по критерию 5.1: "
      f"P(мдп5.1) = {calculation.criterion5(p_fluctuations)} МВт")
print(f"МДП по критерию 5.2: "
      f"P(мдп5.2)= {calculation.criterion6(p_fluctuations, faults_lines)} МВт")

# Удаление временных файлов
for item in os.listdir('.'):
    if item.endswith(".sch") or item.endswith(".ut2"):
        os.remove(item)


