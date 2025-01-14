import os
import psutil
import subprocess
import win32com.client
import xlwings as xw

from datetime import datetime
from time import sleep
from typing import Dict, List
from dependencies.sap import SAPManipulation
from dependencies.functions import Functions
from dependencies.credenciais import Credential
from dependencies.config import Config
from dependencies.logs import Logs, traceback

class CJI3(SAPManipulation):
    def __init__(self, *, date:datetime) -> None:
        crd:dict = Credential(Config()['credential']['crd']).load()
        super().__init__(user=crd.get("user"), password=crd.get("password"), ambiente=crd.get("ambiente"))
        
        if not isinstance(date, datetime):
            raise TypeError("apenas datetime na instancia 'date'")
        
        self.__date: datetime = date
        self.__dateSTR:str = self.date.strftime("%d.%m.%Y")
        self.__initialDate:str = "01.01.2000"
        
        self.__bases_path:str = os.getcwd() + "\\Bases\\"
        if not os.path.exists(self.bases_path):
            os.makedirs(self.bases_path)
        for _file in os.listdir(self.bases_path):
            if _file.endswith(".xlsx"):
                try:
                    os.unlink(self.bases_path + _file)
                except PermissionError:
                    Functions.fechar_excel(_file)
                    os.unlink(self.bases_path + _file)
                
    
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
    
            
    @SAPManipulation.start_SAP
    def gerar_relatorios_SAP(self, *, lista:Dict[str, list|dict], peps:list=[".po"], gerar_quantos:int=987654321, numero_relatorios:str="999999999") -> None:
        contador_gerados = 1
        if not isinstance(peps, list):
            raise TypeError("apenas listas")
        if not isinstance(lista['executar'], list):
            raise TypeError("em 'lista['executar']' apenas listas")
        if not isinstance(lista["nomes"], dict):
            raise TypeError("em 'lista['executar']' apenas dicionários")
        
        try:
            lista_executar:list = lista['executar']
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
                            self.session.findById("wnd[1]/usr/txtKAEP_SETT-MAXSEL").text = numero_relatorios
                            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
                            
                            if self.session.findById("wnd[0]/sbar").text == "Não foi selecionado nenhum objeto com os critérios de seleção indicados.":
                                raise FileNotFoundError("Não foi selecionado nenhum objeto com os critérios de seleção indicados.")
                            
                            if (error:=self.session.findById("wnd[0]/sbar").text) == "Memória escassa. Encerrar a transação antes de pausa !":
                                raise Exception(error)

                            lista["nomes"][centro_custo]
                            empreendimento_for_save:str = f"{centro_custo} - {lista['nomes'][centro_custo]} - {datetime.now().strftime('%d-%m-%Y')}.xlsx".upper()

                            file:str = self.bases_path + empreendimento_for_save
                            
                            #print(file)
                            if os.path.exists(file):
                                try:
                                    os.unlink(file)
                                except PermissionError:
                                    if Functions.fechar_excel(empreendimento_for_save):
                                        os.unlink(file)                              
                                        
                            sleep(1)
                            self.session.findById("wnd[0]").sendVKey(43)
                            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                            self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = self.bases_path
                            self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = empreendimento_for_save
                            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                            
                            sleep(3)
                            
                            Functions.fechar_excel(empreendimento_for_save)
                            
                            print(f"{datetime.now().strftime('%d/%m/%Y - %H:%M:%S')}            Finalizado!")  
                            break          
                        except Exception as error:
                            print(f"{datetime.now().strftime('%d/%m/%Y - %H:%M:%S')}            error -> {type(error)} -> {error}")
                            Logs().register(status='Report', description=f"erro ao gerar relatório {codigo_empreendimento} -> {type(error)} -> {error}", exception=traceback.format_exc())
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
    

if __name__ == "__main__":
    pass
    # date: datetime = datetime.now()
    
    # crd = Credential('SAP_PRD').load()
    
    # bot = CJI3(date=date)
    # bot.conectar(user=crd['user'], password=crd['password'])
