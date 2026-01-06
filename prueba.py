
from pywinauto import Application
app = Application(backend="uia").connect(title_re=r"UNICON\s+-\s+Módulo de PEDIDOS_DISTRIBUCION\s+-\s+AGREGADOS.*")
win = app.window(title_re=r"UNICON\s+-\s+Módulo de PEDIDOS_DISTRIBUCION\s+-\s+AGREGADOS.*")
win.print_control_identifiers()
