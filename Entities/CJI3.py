import win32com.client
from datetime import datetime
import pandas as pd
import xlwings as xw # type: ignore
from time import sleep
import os
from getpass import getuser
import traceback
import subprocess
import json
import psutil
from credenciais.credenciais import Credenciais # type: ignore


speak:bool=False

dados_credenciais = Credenciais().read()

def add_bar(path: str) -> str:
    """adiciona barra no final da string

    Args:
        path (str): caminho

    Returns:
        str: caminho com as barras adicionadas
    """
    if path[-1] != "\\":
        path += "\\"
    return path

def mountDefaultPath(path: str) -> str:
    """cria uma pasta padrao no perfil do usuario para controle do script

    Args:
        path (str): pasta que vai ser criada

    Returns:
        str: caminho completo com a pasta criada
    """
    tempPath: str = add_bar(f"C:\\Users\\{getuser()}\\.bot_ti\\")
    if not os.path.exists(tempPath):
        os.mkdir(tempPath)
    tempPath += add_bar(path)
    if not os.path.exists(tempPath):
        os.mkdir(tempPath)
    return tempPath


class CJI3:
    def __init__(self, date:datetime=datetime.now(), path:str="CJI3") -> None:
        """methodo construtor da classe

        Args:
            date (datetime, optional): data para operar o script. Defaults to datetime.now().
            path (str, optional): nome da pasta onde os arquivos serão salvos. Defaults to "CJI3".
        """
        self.date: datetime = date
        self.dateSTR: str = self.date.strftime("%d.%m.%Y")
        self.initialDate: str = "01.01.2000"
        
        self.tempPath: str = mountDefaultPath(path)
        
    
    def conectar_sap(f):
        """Decorador

        Args:
            f (_type_): função que será decorada
        """
        def verificar_sap_aberto() -> bool:
            """verifica se o programa SAP está aberto

            Returns:
                bool: retorna True se o programa está aberto
            """
            for process in psutil.process_iter(['name']):
                if "saplogon" in process.name().lower():
                    return True
            return False
        def wrap(*args, **kwargs):
            """função que aplica a decoração

            Raises:
                TypeError: se o Kargs estiver faltando
                Exception: não conseguiu se conectar ao SAP

            Returns:
                _type_: retorna a propria função decorada
            """
            try:
                if not verificar_sap_aberto():

                    path_sap = r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe"
                    subprocess.Popen(path_sap)
                    sleep(5)
                    
                SapGuiAuto: win32com.client.CDispatch = win32com.client.GetObject("SAPGUI")# type: ignore
                application: win32com.client.CDispatch = SapGuiAuto.GetScriptingEngine# type: ignore
                connection = application.OpenConnection("S4P", True) # type: ignore
                session: win32com.client.CDispatch = connection.Children(0)# type: ignore
                session.findById("wnd[0]/usr/txtRSYST-BNAME").text = dados_credenciais["user"] # Usuario
                session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = dados_credenciais["pass"] # Senha
                session.findById("wnd[0]").sendVKey(0)
                    
                kwargs["session"] = session
                retorno = f(*args, **kwargs)
                return retorno
            except TypeError:
                raise TypeError("faltando o **Kargs na função principal")
            except Exception as error:
                raise Exception(f"não foi possivel conectar ao sap --> {type(error).__class__} -> {error}")
        return wrap
    
    def _listar_empreendimentos(self) -> list:
        """gera uma lista dos empreendimentos que serão carregados

        Raises:
            FileExistsError: caso não encontre o arquivo

        Returns:
            list: lista com os arquivos encontrador
        """
        path_base:str = f"C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\"
        path:str = [(path_base + x + '\\Informações de Obras.xlsx') for x in os.listdir(path_base) if 'Base de Dados - Geral' in x][0]        
        
        if os.path.exists(path):
            df = pd.read_excel(path)['Código da Obra']
            return df.unique().tolist()
        raise FileExistsError(f"arquivo não encontrado -> {path}")
    
    @conectar_sap  # type: ignore  
    def gerarRelatorio(self, *args, **kwargs) -> None:
        """gera os relatorios que serão salvas no caminho self.tempPath

        Raises:
            Exception: erro na execução de gerar relatorios
            FileNotFoundError: Não foi selecionado nenhum objeto com os critérios de seleção indicados.
        """
        agora:datetime = datetime.now()
        session: win32com.client.CDispatch = kwargs["session"]
        
        try:
            relatorios:list = [".po"]
            for empre in self._listar_empreendimentos():
                for relatorio in relatorios:
                    empreendimento:str = empre + relatorio
                    print(f"{datetime.now().strftime('%d/%m/%Y - %H:%M:%S')} {empreendimento} -> Iniciado")
                    try:
                        #executando CJI3
                        #session.findById("wnd[0]").maximize()
                        session.findById("wnd[0]/tbar[0]/okcd").text = ""
                        session.findById("wnd[0]/tbar[0]/okcd").text = "/nCJI3"
                        session.findById("wnd[0]").sendVKey(0)
                        try:
                            if not "Seleções gestão projetos (Outro perfil BD: ZPS000000001)" in session.findById("/app/con[0]/ses[0]/wnd[0]/usr/boxSEL_TEXT").Text:
                                raise Exception()
                        except:
                            session.findById("wnd[0]").sendVKey(4)
                            session.findById("wnd[2]/usr/lbl[6,19]").setFocus()
                            session.findById("wnd[2]").sendVKey(2)
                            session.findById("wnd[1]/tbar[0]/btn[0]").press()
                            session.findById("wnd[1]").sendVKey(4)
                            session.findById("wnd[2]/usr/lbl[14,14]").setFocus()
                            session.findById("wnd[2]").sendVKey(2)
                            session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        
                        #ronda de empresas
                        session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").text = empreendimento # empreendimento
                        session.findById("wnd[0]/usr/ctxtR_BUDAT-LOW").text = self.initialDate
                        session.findById("wnd[0]/usr/ctxtR_BUDAT-HIGH").text = self.dateSTR
                        session.findById("wnd[0]/usr/ctxtP_DISVAR").text = "/FABRICIO"
                        session.findById("wnd[0]/usr/btnBUT1").press()
                        session.findById("wnd[1]/usr/txtKAEP_SETT-MAXSEL").text = "999999999" # valor 999999999
                        session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        session.findById("wnd[0]/tbar[1]/btn[8]").press()
                        
                        if session.findById("wnd[0]/sbar").text == "Não foi selecionado nenhum objeto com os critérios de seleção indicados.":
                            raise FileNotFoundError("Não foi selecionado nenhum objeto com os critérios de seleção indicados.")
                        
                        #salvando Relatorio
                        file:str = self.tempPath + empreendimento.upper() + ".xlsx"
                        if os.path.exists(file):
                            try:
                                os.unlink(file)
                            except PermissionError:
                                app = xw.Book(file)
                                app.close()
                                os.unlink(file)                        
                        sleep(1)
                        session.findById("wnd[0]").sendVKey(43)
                        session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        session.findById("wnd[1]/usr/ctxtDY_PATH").text = self.tempPath
                        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = empreendimento.upper() + ".xlsx"
                        session.findById("wnd[1]/tbar[0]/btn[0]").press()
                        
                        sleep(7)
                        if os.path.exists(file):
                            app = xw.Book(file)
                            app.close()
                            
                        print(f"{datetime.now().strftime('%d/%m/%Y - %H:%M:%S')}            Finalizado!")
                    except Exception as error:
                        print(f"{datetime.now().strftime('%d/%m/%Y - %H:%M:%S')}            error -> {type(error)} -> {error}")
                        continue
                #break            
            tempo:str = f"tempo de execução: {datetime.now() - agora}"
            print(tempo) if speak else None
            with open("temp.txt", "w")as _file:
                _file.write(tempo)   
        
        finally:
            sleep(1)
            session.findById("wnd[0]").close()
            sleep(1)
            session.findById('wnd[1]/usr/btnSPOP-OPTION1').press()

if __name__ == "__main__":
    """como usar
    """
    speak=True
    try:
        bot: CJI3 = CJI3()
        print(bot.gerarRelatorio())
    except Exception:
        error = traceback.format_exc()
        print(error) if speak else None
        with open("temp.txt", "w")as _file:
            _file.write(error)               
    input()
    #print(bot.tempPath)
    