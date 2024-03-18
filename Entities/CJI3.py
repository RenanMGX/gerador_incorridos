import os
import psutil
import subprocess
import win32com.client
import xlwings as xw

from datetime import datetime
from time import sleep
from credenciais.credenciais import Credential # type: ignore

class CJI3:
    def __init__(self, *, date:datetime) -> None:
        if not isinstance(date, datetime):
            raise TypeError("apenas datetime na instancia 'date'")
        
        self.__date: datetime = date
        self.__dateSTR:str = self.date.strftime("%d.%m.%Y")
        self.__initialDate:str = "01.01.2000"
        
        self.__bases_path:str = os.getcwd() + "\\Bases\\"
        if not os.path.exists(self.bases_path):
            os.makedirs(self.bases_path)
    
    @property
    def date(self):
        return self.__date
    
    @property
    def dateSTR(self):
        return self.__dateSTR

    @property
    def initialDate(self):
        return self.__initialDate
    
    @property
    def bases_path(self):
        return self.__bases_path
    
    @bases_path.setter
    def bases_path(self, value:str):
        if not isinstance(value, str):
            raise TypeError(f"o valor '{value}' atribuido para 'self.bases_path' não é uma string")
        self.__bases_path = value
    
    @property
    def session(self):
        return self.__session
            
    def conectar(self, *, user, password) -> bool:
        try:
            if not self._verificar_sap_aberto():
                subprocess.Popen(r"C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe")
                sleep(5)
            
            SapGuiAuto: win32com.client.CDispatch = win32com.client.GetObject("SAPGUI")# type: ignore
            application: win32com.client.CDispatch = SapGuiAuto.GetScriptingEngine# type: ignore
            connection = application.OpenConnection("S4P", True) # type: ignore
            self.__session: win32com.client.CDispatch = connection.Children(0)# type: ignore
            
            self.session.findById("wnd[0]/usr/txtRSYST-BNAME").text = user # Usuario
            self.session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password # Senha
            self.session.findById("wnd[0]").sendVKey(0)
            
            return True
        except Exception as error:
            raise ConnectionError(f"não foi possivel se conectar ao SAP motivo: {type(error).__class__} -> {error}")
    
    def gerar_relatorios_SAP(self, *, lista:dict, peps:list=[".po"], gerar_quantos:int=987654321) -> None:
        contador_gerados = 1
        if not isinstance(peps, list):
            raise TypeError("apenas listas")
        try:
            lista_executar = lista['executar']
        except KeyError:
            raise KeyError("chave 'executar' não foi encontrada")
        
        agora:datetime = datetime.now()

        try:
            for centro_custo in lista_executar:
                for pep in peps:
                    codigo_empreendimento:str = centro_custo + pep
                    print(f"{datetime.now().strftime('%d/%m/%Y - %H:%M:%S')} {codigo_empreendimento} -> Iniciado")
                    for _ in range(5):
                        try:
                            #executando CJI3
                            self.session.findById("wnd[0]/tbar[0]/okcd").text = ""
                            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/nCJI3"
                            self.session.findById("wnd[0]").sendVKey(0)
                            #import pdb;pdb.set_trace()
                            try:
                                if not "Seleções gestão projetos (Outro perfil BD: ZPS000000001)" in self.session.findById("/app/con[0]/ses[0]/wnd[0]/usr/boxSEL_TEXT").Text:
                                    raise Exception()
                            except:
                                try:
                                    self.session.findById("wnd[1]/usr/sub:SAPLSPO4:0300")
                                    self.session.findById("wnd[0]").sendVKey(4)
                                    self.session.findById("wnd[2]/usr/lbl[6,19]").setFocus()
                                    self.session.findById("wnd[2]").sendVKey(2)
                                    self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                                    self.session.findById("wnd[1]").sendVKey(4)
                                    self.session.findById("wnd[2]/usr/lbl[14,14]").setFocus()
                                    self.session.findById("wnd[2]").sendVKey(2)
                                    self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                                except:
                                    pass

                            #ronda de empresas
                            self.session.findById("wnd[0]/usr/ctxtCN_PSPNR-LOW").text = codigo_empreendimento # empreendimento
                            self.session.findById("wnd[0]/usr/ctxtR_BUDAT-LOW").text = self.initialDate
                            self.session.findById("wnd[0]/usr/ctxtR_BUDAT-HIGH").text = self.dateSTR
                            self.session.findById("wnd[0]/usr/ctxtP_DISVAR").text = "/FABRICIO"
                            self.session.findById("wnd[0]/usr/btnBUT1").press()
                            self.session.findById("wnd[1]/usr/txtKAEP_SETT-MAXSEL").text = "999999999" # valor 999999999
                            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
                            
                            if self.session.findById("wnd[0]/sbar").text == "Não foi selecionado nenhum objeto com os critérios de seleção indicados.":
                                raise FileNotFoundError("Não foi selecionado nenhum objeto com os critérios de seleção indicados.")

                            lista["nomes"][centro_custo]
                            empreendimento_for_save:str = f"{centro_custo} - {lista['nomes'][centro_custo]} - {datetime.now().strftime('%d-%m-%Y')}.xlsx".upper()

                            file:str = self.bases_path + empreendimento_for_save
                            
                            #print(file)
                            if os.path.exists(file):
                                
                                try:
                                    os.unlink(file)
                                except PermissionError:
                                    if self._fechar_excel(file_name=empreendimento_for_save):
                                        os.unlink(file)                              
                                        
                            sleep(1)
                            self.session.findById("wnd[0]").sendVKey(43)
                            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                            self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = self.bases_path
                            self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = empreendimento_for_save
                            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                            
                            sleep(3)
                            
                            self._fechar_excel(file_name=empreendimento_for_save)
                            
                            print(f"{datetime.now().strftime('%d/%m/%Y - %H:%M:%S')}            Finalizado!")  
                            break          
                        except Exception as error:
                            print(f"{datetime.now().strftime('%d/%m/%Y - %H:%M:%S')}            error -> {type(error)} -> {error}")
                            continue
                
                if contador_gerados < gerar_quantos:
                    contador_gerados += 1
                else:
                    break
                #break
                
                
            tempo:str = f"tempo de execução: {datetime.now() - agora}"
            print(tempo)
            # with open("temp.txt", "w")as _file:
            #     _file.write(tempo)   
        
        finally:
            try:
                sleep(1)
                self.session.findById("wnd[0]").close()
                sleep(1)
                self.session.findById('wnd[1]/usr/btnSPOP-OPTION1').press()
            except Exception as error:
                print(f"não foi possivel fechar o SAP {type(error)} | {error}")
    
    
    def _fechar_excel(self, *, file_name:str, timeout=15) -> bool:
        for _ in range(timeout):
            for app in xw.apps:
                for open_file in app.books:
                    if file_name.lower() == open_file.name.lower():
                        open_file.close()
                        if len(xw.apps) <= 0:
                            app.kill()
                        return True
            sleep(1)
        return False
     
        
    def _verificar_sap_aberto(self) -> bool:
        for process in psutil.process_iter(['name']):
            if "saplogon" in process.name().lower():
                return True
        return False    

if __name__ == "__main__":
    date: datetime = datetime.now()
    
    crd = Credential("credencialSAP").load()
    
    bot = CJI3(date=date)
    bot.conectar(user=crd['user'], password=crd['password'])
