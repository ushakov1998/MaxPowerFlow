from win32com.client import Dispatch


def do_regime_weight(rastr: Dispatch) -> None:
    """
    Пошаговое утяжеление режима
    rastr: COM path to RastrWin3
    """
    if rastr.ut_utr('i') > 0:
        rastr.ut_utr('')


def load_clean_regime(rastr: Dispatch) -> None:
    """
    Загрузка чистого файла режима .rg2
    rastr: COM path to RastrWin3
    """
    rastr.Load(1, 'samples/regime.rg2', 'rst.patterns/режим.rg2')


def load_sech(rastr: Dispatch) -> None:
    """
    Загрузка чистого файла сечений .sch
    rastr: COM path to RastrWin3
    """
    rastr.Load(1, 'sech.sch', 'rst.patterns/сечения.sch')


def load_traj(rastr: Dispatch) -> None:
    """
    Загрузка файла утяжеления .ut2
    rastr: COM path to RastrWin3
    """
    rastr.Load(1, 'traj.ut2', 'rst.patterns/траектория утяжеления.ut2')


def set_regime(rastr: Dispatch,
               max_steps: int,
               full_control: int,
               disable_v: int,
               disable_i: int) -> None:
    """
    Установка параметром утяжеления
    rastr: COM path to RastrWin3
    max_steps: int - максимальное количество шагов утяжеления
    full_control: int (0 - вкл.контроля P,U,I; 1 - выкл.контроля P,U,I)
    disable_v: int (0 - вкл.контроля U; 1 - выкл.контроля U)
    disable_i: int (0 - вкл.контроля I; 1 - выкл.контроля I)
    """

    # Установка максимального количества шагов утяжеления
    rastr.Tables('ut_common').Cols('iter').SetZ(0, max_steps)
    # Контроль всех параметров утяжеления
    rastr.Tables('ut_common').Cols('enable_contr').SetZ(0, full_control)
    # Отключение контроля напряжения
    rastr.Tables('ut_common').Cols('dis_v_contr').SetZ(0, disable_v)
    # Отключение контроля тока
    rastr.Tables('ut_common').Cols('dis_i_contr').SetZ(0, disable_i)
    
