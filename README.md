# Max Power Flow
Автор: Ушаков А.В
## Назначение:
Данный скрипт предназначен для определения величины максимально допустимого перетока мощности (МДП) в контролируемом сечение (КС).

Определение МДП проводится согласно [стандарту АО "СО ЕЭС"](https://www.so-ups.ru/fileadmin/files/laws/standards/st_max_power_rules_004-2020.pdf) .

## Технические требования:
* [Python (32-bit)](https://www.python.org/downloads/windows/)
* [pywin32](https://pypi.org/project/pywin32/)
* [RastrWin3 (x86) v 2.0.0.5709 и выше](https://www.rastrwin.ru/rastr/)

## Файлы для работы со скриптом:
| Формат | Описание | 
:-------- |:-----:| 
-rg2 | Файл режима RastrWin3 | 
-rg2 (template)  | Шаблон файла режима RastrWin3 | 
faults.json  | Файл-описание нормативных возмущений |
outages.json  | Файл-описание КС |
vector.csv | Файл-описание траектории утяжеления |


