import os
import traceback

from Entities.CJI3 import CJI3
from Entities.FilesManipulation import Files
from Entities.credenciais.credenciais import Credential # type: ignore
from Entities.sharePointFolder import SharePointFolder # type: ignore
from datetime import datetime
from getpass import getuser
#from time import sleep
#from getpass import getuser

def erro_log():
    path_log_error = "log_error"
    if not os.path.exists(path_log_error):
        os.makedirs(path_log_error)
    with open(f"{path_log_error}\\{datetime.now().strftime('%d-%m-%Y %H_%M_%S')}.txt", 'w')as _file:
        _file.write(traceback.format_exc())



if __name__ == "__main__":
    date:datetime = datetime.now()
    crd:dict = Credential("credencialSAP").load()
    
    gerar_relatorios:bool = True
    manipular_relatorio:bool = True
    
    try:
        
        infor = SharePointFolder.infor_obras(path=f"C:/Users/renan.oliveira/PATRIMAR ENGENHARIA S A/Janela da Engenharia Controle de Obras - _Base de Dados - Geral/Informações de Obras.xlsx")
        
        if gerar_relatorios:
            botSAP: CJI3 = CJI3(date=date)
            botSAP.conectar(user=crd['user'], password=crd['password'])
            botSAP.gerar_relatorios_SAP(lista=infor, gerar_quantos=1)
        
        if manipular_relatorio:
            files_manipulation: Files = Files(date)
            files_manipulation.gerar_incorridos(infor=infor)
            files_manipulation.salvar_no_destino(destino=f"C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\Janela da Engenharia Controle de Obras - Incorridos - SAP\\")        
    
    
    except Exception as error:
        print(traceback.format_exc())
        erro_log()
