from Entities.CJI3 import CJI3
from Entities.FilesManipulation import Files
import traceback
from datetime import datetime
import os
from time import sleep
from getpass import getuser

def erro_log():
    path_log_error = "log_error"
    if not os.path.exists(path_log_error):
        os.makedirs(path_log_error)
    with open(f"{path_log_error}\\{datetime.now().strftime('%d-%m-%Y %H_%M_%S')}.txt", 'w')as _file:
        _file.write(traceback.format_exc())



if __name__ == "__main__":
    try:
        pass
        #CJI3().gerarRelatorio()
    except Exception as error:
        erro_log()
        sleep(1)
    for _ in range(5):
        try:
            bot = Files()
            bot.gerar_arquivos()
            bot.copiar_destino(f"C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\Janela da Engenharia Controle de Obras - Incorridos - SAP\\")
            break
        except Exception as error:
            print(traceback.format_exc())
            erro_log()
        sleep(1)
