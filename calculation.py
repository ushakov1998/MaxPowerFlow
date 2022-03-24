import win32com.client
import regime_config

rastr = win32com.client.Dispatch('Astra.Rastr')


def iterating() -> None:
    """
    Итерирование по таблице узлов с
    опеределением u_kr и u_min
    """
    # Определяем COM-путь к таблице узлов
    nodes = rastr.Tables('node')
    for i in range(nodes.Size):
        # Поиск узла нагрузки (1 - узел с нагрукой)
        if nodes.Cols('tip').Z(i) == 1:
            u_kr = nodes.Cols('uhom').Z(i) * 0.7  # Критический уровень напряжения
            u_min = u_kr * 1.15  # Допустимый уровень напряжения
            nodes.Cols('umin').SetZ(i, u_min)
            nodes.Cols('contr_v').SetZ(i, 1)


def criterion1(p_fluctuations: float) -> float:
    """
    Обеспечение 20% запаса по статической устойчивости
    в нормальной схеме
    p_fluctuations: амплитуда нерегулярных колебаний активной мощности
    """

    # Загрузка режима и выбор параметров утяжеления
    regime_config.load_clean_regime(rastr)
    regime_config.load_sech(rastr)
    regime_config.load_traj(rastr)
    regime_config.set_regime(rastr, 200, 1, 1, 1)

    # Пошаговое утяжеление режима
    regime_config.do_regime_weight(rastr)

    # МДП по 1 критерию
    mpf_1 = abs(
        rastr.Tables('sechen').Cols('psech').Z(0)) * 0.8 - p_fluctuations
    mpf_1 = round(mpf_1, 1)
    return mpf_1


def criterion2(p_fluctuations: float) -> float:
    """
    Обеспечение 15% коэффициента запаса по напряжению
    в узлах нагрузки в нормальной схеме
    p_fluctuations: амплитуда нерегулярных колебаний активной мощности
    """

    # Загрузка режима и выбор параметров утяжеления
    regime_config.load_clean_regime(rastr)
    regime_config.load_sech(rastr)
    regime_config.load_traj(rastr)
    regime_config.set_regime(rastr, 200, 1, 0, 1)

    # Итерирование по таблице узлов
    iterating()

    # Пошаговое утяжеление режима
    regime_config.do_regime_weight(rastr)

    # МДП по 2 критерию
    mpf_2 = abs(rastr.Tables('sechen').Cols('psech').Z(0)) - p_fluctuations
    mpf_2 = round(mpf_2, 1)
    return mpf_2


def criterion3(p_fluctuations: float, faults_lines: dict) -> float:
    """
    Обеспечение 8% коэффициента запаса статической
    устойчивости в послеаварийных режимах
    после нормативных возмущений
    p_fluctuations: амплитуда нерегулярных колебаний активной мощности
    faults_lines: словарь моделируемых возмущений
    """

    # Загрузка режима и выбор параметров утяжеления
    regime_config.load_clean_regime(rastr)
    regime_config.load_sech(rastr)
    regime_config.load_traj(rastr)
    regime_config.set_regime(rastr, 200, 1, 1, 1)

    # COM путь к таблице ветвей и таблице сечений
    branches = rastr.Tables('vetv')
    flowgate = rastr.Tables('sechen')

    # Список МДП для каждого возмущения
    mpf_3 = []

    # Итерирование по каждому возмущению
    for line in faults_lines:
        # Номер узла начала ветви
        node_start_branch = faults_lines[line]['ip']
        # Номер узла конца ветви
        node_end_branch = faults_lines[line]['iq']
        # Номер параллельности
        parallel_number = faults_lines[line]['np']
        # Состояние ветви (0 - вкл / 1 - выкл)
        branch_status = faults_lines[line]['sta']

        # Итерирование по каждой ветви в RastrWin3
        for i in range(branches.Size):

            # Поиск ветви с возмущением
            if (branches.Cols('ip').Z(i) == node_start_branch) and \
                    (branches.Cols('iq').Z(i) == node_end_branch) and \
                    (branches.Cols('np').Z(i) == parallel_number):

                # Запоминаем предыдущий статус ветви
                pr_branch_status = branches.Cols('sta').Z(i)
                # Производим возмущение
                branches.Cols('sta').SetZ(i, branch_status)

                # Производим утяжеление режима
                regime_config.do_regime_weight(rastr)

                # МДП по 3 критерию
                mpf = abs(flowgate.Cols('psech').Z(0))
                # Допустимый уровень напряжения для схемы
                mpf_acceptable = abs(flowgate.Cols('psech').Z(0)) * 0.92

                # Определение COM пути к коллекции режимов (шагов утяжеления)
                toggle = rastr.GetToggle()

                # Итеративный возврат к допустимому уровню МДП
                j = 1
                while mpf > mpf_acceptable:
                    toggle.MoveOnPosition(len(toggle.GetPositions()) - j)
                    mpf = abs(flowgate.Cols('psech').Z(0))
                    j += 1

                # Убираем возмущение
                branches.Cols('sta').SetZ(i, pr_branch_status)
                # Расчет режима
                rastr.rgm('p')

                # МДП по 3 критерию
                mpf = abs(
                    rastr.Tables('sechen').Cols('psech').Z(0)) - p_fluctuations
                mpf = round(mpf, 1)
                mpf_3.append(mpf)

                # Сброс для очистки режима
                toggle.MoveOnPosition(1)
                branches.Cols('sta').SetZ(i, pr_branch_status)
                break
    return min(mpf_3)


def criterion4(p_fluctuations: float, faults_lines: dict) -> float:
    """
    Обеспечение 10% коэффициента запаса по напряжению в узлах
    нагрузки в послеаварийных режимах после нормативных возмущений
    p_fluctuations: амплитуда нерегулярных колебаний активной мощности
    faults_lines: словарь моделируемых возмущений
    """

    # Загрузка режима и выбор параметров утяжеления
    regime_config.load_clean_regime(rastr)
    regime_config.load_sech(rastr)
    regime_config.load_traj(rastr)
    regime_config.set_regime(rastr, 200, 1, 0, 1)

    # Redefine the COM path to the RastrWin3 branch table
    branches = rastr.Tables('vetv')
    # Redefine the COM path to the RastrWin3 flowgate table
    flowgate = rastr.Tables('sechen')

    # Итерирование по таблице узлов
    iterating()

    # Лист МДП для каждого возмущения
    mpf_4 = []

    # Итерирование по каждому возмущению
    for line in faults_lines:
        # Номер узла начала ветви
        node_start_branch = faults_lines[line]['ip']
        # Номер узла конца ветви
        node_end_branch = faults_lines[line]['iq']
        # Номер параллельности
        parallel_number = faults_lines[line]['np']
        # Состояние ветви (0 - вкл / 1 - выкл)
        branch_status = faults_lines[line]['sta']

        # Итерирование по каждой ветви в RastrWin3
        for i in range(branches.Size):

            # Поиск ветви с возмущением
            if (branches.Cols('ip').Z(i) == node_start_branch) and \
                    (branches.Cols('iq').Z(i) == node_end_branch) and \
                    (branches.Cols('np').Z(i) == parallel_number):
                # Запоминаем предыдущий статус ветви
                pr_branch_status = branches.Cols('sta').Z(i)
                # Производим возмущение
                branches.Cols('sta').SetZ(i, branch_status)

                # Производим утяжеление режима
                regime_config.do_regime_weight(rastr)
                # Убираем возмущение
                branches.Cols('sta').SetZ(i, pr_branch_status)
                # Расчет режима
                rastr.rgm('p')

                # МДП по 4 критерию
                mpf = abs(
                    flowgate.Cols('psech').Z(0)) - p_fluctuations
                mpf = round(mpf, 1)
                mpf_4.append(mpf)

                # Сброс для очистки режима
                rastr.GetToggle().MoveOnPosition(1)
                branches.Cols('sta').SetZ(i, pr_branch_status)
                break
    return min(mpf_4)


def criterion5(p_fluctuations: float) -> float:
    """
    Обеспечение ДДТН линий электропередачи и
    электросетевого оборудования в нормальной схеме
    p_fluctuations: амплитуда нерегулярных колебаний активной мощности
    """

    # Загрузка режима и выбор параметров утяжеления
    regime_config.load_clean_regime(rastr)
    regime_config.load_sech(rastr)
    regime_config.load_traj(rastr)
    regime_config.set_regime(rastr, 200, 1, 1, 0)

    # Redefine the COM path to the RastrWin3 branch table
    branches = rastr.Tables('vetv')
    # Redefine the COM path to the RastrWin3 flowgate table
    flowgate = rastr.Tables('sechen')
    # Redefine the COM path to collection of regimes RastrWin3

    # Итерирование по каждой ветви в RastrWin3
    for i in range(branches.Size):
        branches.Cols('contr_i').SetZ(i, 1)
        branches.Cols('i_dop').SetZ(i, branches.Cols('i_dop_r').Z(i))

    # Пошаговое утяжеление режима
    regime_config.do_regime_weight(rastr)

    # МДП по 5 критерию
    mpf_5 = abs(flowgate.Cols('psech').Z(0)) - p_fluctuations
    mpf_5 = round(mpf_5, 1)
    return mpf_5


def criterion6(p_fluctuations: float, faults_lines: dict):
    """
    Обеспечение АДТН линий электропередачи и электросетевого
    оборудования в послеаварийных режимах
    после нормативных возмущений
    p_fluctuations: амплитуда нерегулярных колебаний активной мощности
    faults_lines: словарь моделируемых возмущений
    """

    # Загрузка режима и выбор параметров утяжеления
    regime_config.load_clean_regime(rastr)
    regime_config.load_sech(rastr)
    regime_config.load_traj(rastr)
    regime_config.set_regime(rastr, 200, 1, 1, 0)

    # Redefine the COM path to the RastrWin3 branch table
    branches = rastr.Tables('vetv')
    # Redefine the COM path to the RastrWin3 flowgate table
    flowgate = rastr.Tables('sechen')

    # Итерирование по каждой ветви в RastrWin3
    for j in range(branches.Size):
        branches.Cols('contr_i').SetZ(j, 1)
        branches.Cols('i_dop').SetZ(j, branches.Cols('i_dop_r_av').Z(j))

    # Список МДП для каждого возмущения
    mpf_6 = []

    # Итерирование по каждому возмущению
    for line in faults_lines:
        # Номер узла начала ветви
        node_start_branch = faults_lines[line]['ip']
        # Номер узла конца ветви
        node_end_branch = faults_lines[line]['iq']
        # Номер параллельности
        parallel_number = faults_lines[line]['np']
        # Состояние ветви (0 - вкл / 1 - выкл)
        branch_status = faults_lines[line]['sta']

        # Итерирование по каждой ветви в RastrWin3
        for i in range(branches.Size):
            # Поиск ветви с возмущением
            if (branches.Cols('ip').Z(i) == node_start_branch) and \
                    (branches.Cols('iq').Z(i) == node_end_branch) and \
                    (branches.Cols('np').Z(i) == parallel_number):
                # Запоминаем предыдущий статус ветви
                pr_branch_status = branches.Cols('sta').Z(i)
                # Производим возмущение
                branches.Cols('sta').SetZ(i, branch_status)

                # Производим утяжеление режима
                regime_config.do_regime_weight(rastr)

                # Убираем возмущение
                branches.Cols('sta').SetZ(i, pr_branch_status)
                # Расчет режима
                rastr.rgm('p')

                # МДП по 6 критерию
                mpf = abs(flowgate.Cols('psech').Z(0)) - p_fluctuations
                mpf = round(mpf, 1)
                mpf_6.append(mpf)

                # Сброс для очистки режима
                rastr.GetToggle().MoveOnPosition(1)
                branches.Cols('sta').SetZ(i, pr_branch_status)
                break
    return min(mpf_6)
