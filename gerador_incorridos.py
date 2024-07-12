import os
import traceback

from Entities.CJI3 import CJI3
from Entities.FilesManipulation import Files
from Entities.crenciais import Credential # type: ignore
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
    crd:dict = Credential('SAP_PRD').load()
    
    gerar_relatorios:bool = True
    manipular_relatorio:bool = True
    
    sharepoint_path:str = f"C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\Janela da Engenharia Controle de Obras - Incorridos - SAP"
    
    try:
        if not os.path.exists(sharepoint_path):
            raise FileNotFoundError(f"não foi possivel localizar a pasta do sharepoint '{sharepoint_path}'")
        
        infor_obras_path:str = os.path.join(sharepoint_path,"Informações de Obras.xlsx")
        if not os.path.exists(infor_obras_path):
            raise FileNotFoundError(f"não foi possivel localizar o arquivo do sharepoint '{infor_obras_path}'")
        
        descri_sap_path:str = os.path.join(sharepoint_path, "Descrição SAP.xlsx")
        if not os.path.exists(descri_sap_path):
            raise FileNotFoundError(f"não foi possivel localizar o arquivo do sharepoint '{descri_sap_path}'")
        
        infor = SharePointFolder.infor_obras(path=infor_obras_path)
        
        if gerar_relatorios:
            botSAP: CJI3 = CJI3(date=date)
            botSAP.conectar(user=crd['user'], password=crd['password'])
            botSAP.gerar_relatorios_SAP(lista=infor)
        
        if manipular_relatorio:
            files_manipulation: Files = Files(date, description_sap_tags_path=descri_sap_path)
            files_manipulation.gerar_incorridos(infor=infor)
            files_manipulation.salvar_no_destino(destino=sharepoint_path)        
    
    except Exception as error:
        print(traceback.format_exc())
        erro_log()
        path:str = os.path.join(os.getcwd(), "logs/")
        if not os.path.exists(path):
            os.makedirs(path)
        file_name = os.path.join(path, f"LogError_{datetime.now().strftime('%d%m%Y%H%M%Y')}.txt")
        with open(file_name, 'w', encoding='utf-8')as _file:
            _file.write(traceback.format_exc())
        raise error
                    
