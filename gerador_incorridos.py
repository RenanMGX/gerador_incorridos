import os
from Entities.CJI3 import CJI3
from Entities.FilesManipulation import Files
from Entities.dependencies.config import Config
from Entities.dependencies.logs import Logs, traceback
from Entities.sharePointFolder import SharePointFolder # type: ignore
from datetime import datetime
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
    
    gerar_relatorios:bool = True
    manipular_relatorio:bool = True
    
    sharepoint_path:str = Config()['paths']['sharepoint_path']
    sharepoint_incorridos_path:str = Config()['paths']['sharepoint_incorrido']
    
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
            botSAP.gerar_relatorios_SAP(lista=infor)#, gerar_quantos=2, numero_relatorios="10000")
        
        if manipular_relatorio:
            files_manipulation: Files = Files(date, description_sap_tags_path=descri_sap_path)
            files_manipulation.gerar_incorridos(infor=infor)
            #files_manipulation.salvar_no_destino(destino=r"C:\\Users\\renan.oliveira\Downloads") # <-------------- alterar depois 
            files_manipulation.salvar_no_destino(destino=sharepoint_path)
            files_manipulation.salvar_Incorridos(target=sharepoint_incorridos_path)

        Logs().register(status='Concluido', description="Automação finalizada com Sucesso!")
    except Exception as err:
        Logs().register(status='Error', description=str(err), exception=traceback.format_exc())
                    