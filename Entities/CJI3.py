import win32com.client
from datetime import datetime
import pandas as pd
import xlwings as xw # type: ignore
from time import sleep
import os
from getpass import getuser
import traceback


speak:bool=False

def add_bar(path: str) -> str:
    if path[-1] != "\\":
        path += "\\"
    return path

def mountDefaultPath(path: str) -> str:
    tempPath: str = add_bar(f"C:\\Users\\{getuser()}\\.bot_ti\\")
    if not os.path.exists(tempPath):
        os.mkdir(tempPath)
    tempPath += add_bar(path)
    if not os.path.exists(tempPath):
        os.mkdir(tempPath)
    return tempPath


class CJI3:
    def __init__(self, date:datetime=datetime.now(), path:str="CJI3") -> None:
        self.date: datetime = date
        self.dateSTR: str = self.date.strftime("%d.%m.%Y")
        self.initialDate: str = "01.01.2000"
        
        self.tempPath: str = mountDefaultPath(path)
        for file in os.listdir(self.tempPath):
            os.unlink(self.tempPath + file)
    
        
    def conectar_sap(f):
        def wrap(*args, **kwargs):
            try:
                SapGuiAuto: win32com.client.CDispatch = win32com.client.GetObject("SAPGUI")# type: ignore
                application: win32com.client.CDispatch = SapGuiAuto.GetScriptingEngine# type: ignore
                connection: win32com.client.CDispatch = application.Children(0)# type: ignore
                session: win32com.client.CDispatch = connection.Children(0)# type: ignore
                kwargs["session"] = session
                retorno = f(*args, **kwargs)
                return retorno
            except TypeError:
                raise TypeError("faltando o **Kargs na função principal")
            except Exception as error:
                raise Exception(f"não foi possivel conectar ao sap --> {type(error).__class__} -> {error}")
        return wrap
    
    def _listar_empreendimentos(self) -> list:
        path:str = f"C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\Janela da Engenharia Controle de Obras - Base de Dados - Geral\\Informações de Obras.xlsx"
        if os.path.exists(path):
            df = pd.read_excel(path)['Código da Obra']
            return df.unique().tolist()
        raise FileExistsError(f"arquivo não encontrado -> {path}")
    
    @conectar_sap  # type: ignore  
    def gerarRelatorio(self, *args, **kwargs) -> None:
        agora = datetime.now()
        
        
        relatorios = [".po"]
        for empre in self._listar_empreendimentos():
            for relatorio in relatorios:
                empreendimento:str = empre + relatorio
                print(f"{datetime.now().strftime('%d/%m/%Y - %H:%M:%S')} {empreendimento} -> Iniciado")
                try:
                    session: win32com.client.CDispatch = kwargs["session"]
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
                        app = xw.Book(file)
                        app.close()
                        os.unlink(self.tempPath + file)
                    sleep(1)
                    session.findById("wnd[0]").sendVKey(43)
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
                    session.findById("wnd[1]/usr/ctxtDY_PATH").text = self.tempPath
                    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = empreendimento.upper() + ".xlsx"
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
                    
                    sleep(3)
                    if os.path.exists(file):
                        app = xw.Book(file)
                        app.close()
                        sleep(1)
                    print(f"{datetime.now().strftime('%d/%m/%Y - %H:%M:%S')}            Finalizado!")
                except Exception as error:
                    print(f"{datetime.now().strftime('%d/%m/%Y - %H:%M:%S')}            error -> {type(error)} -> {error}")
                    continue
    
        tempo = f"tempo de execução: {datetime.now() - agora}"
        print(tempo)
        with open("temp.txt", "w")as _file:
            _file.write(tempo)               
        
        

if __name__ == "__main__":
    speak=True
    
    try:
        bot: CJI3 = CJI3()
        print(bot.gerarRelatorio())
    except Exception:
        error = traceback.format_exc()
        print(error)
        with open("temp.txt", "w")as _file:
            _file.write(error)               
    input()
    #print(bot.tempPath)
    