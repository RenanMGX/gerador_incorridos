import pandas as pd
import os
from time import sleep
from CJI3 import mountDefaultPath
from getpass import getuser
from typing import List,Dict
from datetime import datetime
from shutil import copy2
import xlwings as xw
from selenium import webdriver
from selenium.webdriver.common.by import By

def medir_tempo(f):
    def wrap(*args, **kwargs):
        agora = datetime.now()
        result = f(*args, **kwargs)
        print(f"\ntempo de execução: {datetime.now() - agora}\n")
        return result
    return wrap

def _find_element(browser:webdriver.Chrome, method:By, target:str, timeout=60):
    for x in range(timeout):
        try:
            result = browser.find_element(method, target)
            return result
        except:
            sleep(2)
        

class Files():
    def __init__(self, path="CJI3") -> None:
        self.tempPath: str = mountDefaultPath(path)
    
    
    def _ler_arquivos(self) -> Dict[str, pd.DataFrame]:
        dictionary: dict = {}
        for file in os.listdir(self.tempPath):
            if file.endswith(".xlsx"):
                fileName = file.replace(".xlsx", "").replace(".PO", "")
                print(f"{fileName} -> Executando")
                try:
                    df = pd.read_excel(self.tempPath + file)
                except:
                    continue
                dictionary[fileName] = df
                break
        return dictionary

    def calcular_pep_por_data(self, date:datetime, df:pd.DataFrame, termo:str) -> float:
        df = df[(df['Data de lançamento'].dt.year == date.year) & (df['Data de lançamento'].dt.month == date.month)]
        df = df[df['Elemento PEP'].str.contains(termo, case=False)]
        
        valores:float = 0
        for valor in df['Valor/moeda objeto'].tolist():
            valores += valor
        
        return round(valores, 2)
        
    @medir_tempo
    def gerar_arquivos(self) -> None:
        self.incc = self.incc_valor()
        for name,df in self._ler_arquivos().items():
            path_new_file = f"C:\\Users\\renan.oliveira\\Downloads\\Incorridos_{datetime.now().strftime('%d-%m-%Y')}"
            path_file = path_new_file + "\\" + (name + ".xlsx")
            
            if not os.path.exists(path_new_file):
                os.mkdir(path_new_file)
            
            #df['Data de lançamento'] = pd.to_datetime(df['Data de lançamento'])
            datas:List[pd.Timestamp] = df['Data de lançamento'].unique().tolist()
            datas.pop(datas.index(pd.NaT))
            
            datas = [data.replace(day=1) for data in datas]
            datas = set(datas)
            datas = list(datas)
            datas = sorted(datas)
            datas.reverse()
            
            if os.path.exists(path_file):
                try:
                    os.unlink(path_file)
                except PermissionError:
                    app = xw.Book(path_file)
                    app.close()
                    os.unlink(path_file)
            
            copy2("modelo planilha\\PEP a PEP - Incorridos - Modelo.xlsx", path_file)
            
            app = xw.App(visible=False)
            with app.books.open(path_file)as wb:
                sheet_principal = wb.sheets['PEP A PEP']
                sheet_temp = wb.sheets['temp']
                
                etapas = len(datas)
                etapa = 1
                for date in datas:
                    print(f"{etapa} / {etapas} --> {date}")
                    
                    sheet_principal.range('N1').api.EntireColumn.Insert()
                    sheet_temp.range('A:A').copy()
                    #sheet_principal.range('N1').select()
                    sheet_principal.range('N1').paste()
                    app.api.CutCopyMode = False
                    
                    sheet_principal.range('E2').value = name #Nome
                    
                    sheet_principal.range('N6').value = date #data
                    
                    sheet_principal.range('N11:N13').value = [[self.calcular_pep_por_data(date, df, "POCI")], [self.calcular_pep_por_data(date, df, "POCD")], [self.calcular_pep_por_data(date, df, "POSP")]]
                    
                    sheet_principal.range('N16:N22').value = [
                                                        [self.calcular_pep_por_data(date, df, "POCRCIPJ")],
                                                        [self.calcular_pep_por_data(date, df, "POCRCISP")],
                                                        [self.calcular_pep_por_data(date, df, "POCRCIIP")],
                                                        [self.calcular_pep_por_data(date, df, "POCRCIIPR")],
                                                        [self.calcular_pep_por_data(date, df, "POCRCIEQ")],
                                                        [self.calcular_pep_por_data(date, df, "POCRCIMO")],
                                                        [self.calcular_pep_por_data(date, df, "POCRCICO")]                                                       
                                                        ]                    
                    
                    sheet_principal.range('N24:N53').value = [
                                                            [self.calcular_pep_por_data(date, df, "POCRCD01")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD02")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD03")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD03")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD05")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD06")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD07")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD08")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD09")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD10")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD11")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD12")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD13")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD14")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD15")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD16")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD17")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD18")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD19")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD20")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD21")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD22")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD23")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD24")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD25")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD26")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD27")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD28")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD29")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD30")]
                                                            ]
                    
                    sheet_principal.range('N55').value = self.calcular_pep_por_data(date, df, "PONI") # ""
                    
                    if etapa == etapas:
                        texto_incc = "Valor Mensal do INCC"
                    else:
                        texto_incc = ""
                    sheet_principal.range('N63').value = texto_incc # ""
                    
                    try:
                        sheet_principal.range('N64').value = self.incc[date.to_pydatetime()]
                    except:
                        sheet_principal.range('N64').value = 0
                    
                    etapa += 1
                    #break
                
                wb.save()
            app.kill()            
            #import pdb; pdb.set_trace()
    
    def incc_valor(self):
        with webdriver.Chrome()as _navegador:
            _navegador.get("https://extra-ibre.fgv.br/autenticacao_produtos_licenciados/?ReturnUrl=%2fautenticacao_produtos_licenciados%2flista-produtos.aspx")
            
            _find_element(_navegador, By.ID, 'ctl00_content_hpkGratuito').click()
            _find_element(_navegador, By.ID, 'dlsCatalogoFixo_imbOpNivelUm_0').click()
            _find_element(_navegador, By.ID, 'dlsCatalogoFixo_imbOpNivelDois_4').click()
            _find_element(_navegador, By.ID, 'dlsMovelCorrente_imbIncluiItem_1').click()
            _find_element(_navegador, By.ID, 'butCatalogoMovelFecha').click()
            
            _find_element(_navegador, By.ID, 'cphConsulta_dlsSerie_lblNome_0')
            _find_element(_navegador, By.ID, 'cphConsulta_rbtSerieHistorica').click()
            _find_element(_navegador, By.ID, 'cphConsulta_butVisualizarResultado').click()
            sleep(1)
            _navegador.get("https://extra-ibre.fgv.br/IBRE/sitefgvdados/VisualizaConsultaFrame.aspx")
            
            tabela = _find_element(_navegador, By.ID, 'xgdvConsulta_DXMainTable').text.split('\n')
        
        tabela.pop(0)
        tabela.pop(0)
        
        tabela = {datetime.strptime(x.split(" ")[0], "%m/%Y"):float(x.split(" ")[1].replace(",",".")) for x in tabela}
            
        return tabela    
        
        
if __name__ == "__main__":
    bot = Files()
    
    #print(f"\n\n{bot.incc_valor()}")
    print(f"\n\n{bot.gerar_arquivos()}")