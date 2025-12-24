
import os
import time

# Fecha todas as instâncias do Excel (forçado)
os.system("taskkill /f /im excel.exe")

# Pequena pausa pra garantir que fechou tudo
time.sleep(3)