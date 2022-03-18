import win32com.client
rastrWin3 = win32com.client.Dispatch('Astra.Rastr')


def create_file_sch(flowgate_lines: dict, flowgate_name: str) -> None:
    """
    Создание файла .sch по шаблону RastrWin3 из словаря сечения
    flowgate_lines: словарь ветвей образующий сечение
    flowgate_name: str наименование сечения
    """

    # Создание пустого файла .sch (по шаблону RastrWin3)
    rastrWin3.Save('sech.sch', 'rst.patterns/сечения.sch')
    # Открытие созданного файла
    rastrWin3.Load(1, 'sech.sch', 'rst.patterns/сечения.sch')

    # Определение объектов RastrWin3
    flow_gate = rastrWin3.Tables('sechen')
    group_line = rastrWin3.Tables('grline')

    # Очистка строк в файле .sch
    flow_gate.DelRows()
    group_line.DelRows()

    # Создание сечения
    flow_gate.AddRow()
    flow_gate.Cols('ns').SetZ(0, 1)
    # Give a name for the flowgate
    flow_gate.Cols('name').SetZ(0, flowgate_name)
    flow_gate.Cols('sta').SetZ(0, 1)

    # Заполним список ЛЭП,образующих сечение
    for i, line in enumerate(flowgate_lines):
        group_line.AddRow()
        group_line.Cols('ns').SetZ(i, 1)

        # начало ЛЭП
        start_node = flowgate_lines[line]['ip']
        # конец ЛЭП
        end_node = flowgate_lines[line]['iq']

        group_line.Cols('ip').SetZ(i, start_node)
        group_line.Cols('iq').SetZ(i, end_node)

    # Сохранение файла .sch
    rastrWin3.Save('sech.sch', 'rst.patterns/сечения.sch')


def create_file_ut2(trajectory_nodes: list) -> None:
    """
    Создание файла .ut2 file
    из списка узлов траекторий утяжеления
    trajectory_nodes: список узлов,формирующий траекторию утяж.режима
    """

    # Создание пустого шаблона файла .ut2
    rastrWin3.Save('traj.ut2', 'rst.patterns/траектория утяжеления.ut2')
    # Открытие файла
    rastrWin3.Load(1, 'traj.ut2', 'rst.patterns/траектория утяжеления.ut2')

    # Определение объектов RastrWin3
    trajectory = rastrWin3.Tables('ut_node')

    # Очистка строк в файле.ut2
    trajectory.DelRows()

    # Наполнение траектории списком узлов
    # Для избежания повторений создаем пустой словарь
    # Узел может иметь Pg и Pn одновременно
    node_data = {}  # создание пустово словаря
    i = 0
    for node in trajectory_nodes:
        node_type = node['variable']  # Pg - ген / Pn - нагрузка
        node_number = node['node']
        power_change = float(node['value'])
        power_tg = node['tg']  # Тангенс нагрузки

        # Проверка на содержание узла в словаре
        if node_number not in node_data:
            # Создание пары номер узла - индекс
            node_data[node_number] = i
            i += 1
            # Заполняем строки в файле .ut2
            trajectory.AddRow()
            trajectory.Cols('ny').SetZ(node_data[node_number],
                                       node_number)
            trajectory.Cols(node_type).SetZ(node_data[node_number],
                                            power_change)
        else:
            # Находим существующую пару и добавляем в сущ.строку в файл.ut2
            trajectory.Cols(node_type).SetZ(node_data[node_number],
                                            power_change)

        # Добавляем коэф.нагрузки
        if trajectory.Cols('tg').Z(node_data[node_number]) == 0:
            trajectory.Cols('tg').SetZ(node_data[node_number], power_tg)

    # Сохранение файла .ut2
    rastrWin3.Save('traj.ut2', 'rst.patterns/траектория утяжеления.ut2')