from Entities.CJI3 import CJI3
from Entities.FilesManipulation import Files
import traceback
from datetime import datetime
import os

if __name__ == "__main__":
    try:
        CJI3().gerarRelatorio()
        Files().gerar_arquivos()
    except Exception as error:
        path_log_error = "log_error"
        if not os.path.exists(path_log_error):
            os.makedirs(path_log_error)
        with open(f"{path_log_error}\\{datetime.now().strftime('%d-%m-%Y %H_%M_%S')}.txt", 'w')as _file:
            _file.write(traceback.format_exc())
        error.with_traceback()