import pandas as pd
import os
from CJI3 import mountDefaultPath
from getpass import getuser
from typing import List,Dict
from datetime import datetime
from shutil import copy2
import xlwings as xw

def medir_tempo(f):
    def wrap(*args, **kwargs):
        agora = datetime.now()
        result = f(*args, **kwargs)
        print(f"tempo de execução: {datetime.now() - agora}")
        return result
    return wrap

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
        
        return valores
        
    @medir_tempo
    def gerar_arquivos(self) -> None:
        for name,df in self._ler_arquivos().items():
            path_new_file = f"C:\\Users\\renan.oliveira\\Downloads\\Incorridos_{datetime.now().strftime('%d-%m-%Y')}"
            path_file = path_new_file + "\\" + (name + ".xlsx")
            
            if not os.path.exists(path_new_file):
                os.mkdir(path_new_file)
            
            
            
            #df['Data de lançamento'] = pd.to_datetime(df['Data de lançamento'])
            datas:List[datetime] = df['Data de lançamento'].unique().tolist()
            datas.pop(datas.index(pd.NaT))
            
            datas = [data.replace(day=1) for data in datas]
            datas = set(datas)
            datas = list(datas)
            datas = sorted(datas)
            
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
                
                texto_incc = "Valor Mensal do INCC"
                etapas = len(datas)
                etapa = 1
                for date in datas:
                    valor_pep = self.preprara_pep(date, df)
                    
                    print(f"{etapa} / {etapas} --> {date}")
                    etapa += 1
                    sheet_principal.range('N1').api.EntireColumn.Insert()
                    sheet_temp.range('A:A').copy()
                    #sheet_principal.range('N1').select()
                    sheet_principal.range('N1').paste()
                    app.api.CutCopyMode = False
                    
                    sheet_principal.range('N6').value = date #data
                    
                    sheet_principal.range('N11:N13').value = [valor_pep["POCI"], valor_pep["POCD"], valor_pep["POSP"]]
                    
                    sheet_principal.range('N12').value = valor_pep["POCD"] # "POCD"
                    
                    sheet_principal.range('N13').value = valor_pep["POSP"] # "POSP"
                    
                    sheet_principal.range('N16').value = valor_pep["POCRCIPJ"] # "POCRCIPJ"
                    sheet_principal.range('N17').value = valor_pep["POCRCISP"] # "POCRCISP"
                    sheet_principal.range('N18').value = valor_pep["POCRCIIP"] # "POCRCIIP"
                    sheet_principal.range('N19').value = valor_pep["POCRCIIPR"] # "POCRCIIPR"
                    sheet_principal.range('N20').value = valor_pep["POCRCIEQ"] # "POCRCIEQ"
                    sheet_principal.range('N21').value = valor_pep["POCRCIMO"] # "POCRCIMO"
                    sheet_principal.range('N22').value = valor_pep["POCRCICO"] # "POCRCICO"
                    
                    sheet_principal.range('N24').value = valor_pep["POCRCD01"] # "POCRCD01"
                    sheet_principal.range('N25').value = valor_pep["POCRCD02"] # "POCRCD02"
                    sheet_principal.range('N26').value = valor_pep["POCRCD03"] # "POCRCD03"
                    sheet_principal.range('N27').value = valor_pep["POCRCD04"] # "POCRCD04"
                    sheet_principal.range('N28').value = valor_pep["POCRCD05"] # "POCRCD05"
                    sheet_principal.range('N29').value = valor_pep["POCRCD06"] # "POCRCD06"
                    sheet_principal.range('N30').value = valor_pep["POCRCD07"] # "POCRCD07"
                    sheet_principal.range('N31').value = valor_pep["POCRCD08"] # "POCRCD08"
                    sheet_principal.range('N32').value = valor_pep["POCRCD09"] # "POCRCD09"
                    sheet_principal.range('N33').value = valor_pep["POCRCD10"] # "POCRCD10"
                    sheet_principal.range('N34').value = valor_pep["POCRCD11"] # "POCRCD11"
                    sheet_principal.range('N35').value = valor_pep["POCRCD12"] # "POCRCD12"
                    sheet_principal.range('N36').value = valor_pep["POCRCD13"] # "POCRCD13"
                    sheet_principal.range('N37').value = valor_pep["POCRCD14"] # "POCRCD14"
                    sheet_principal.range('N38').value = valor_pep["POCRCD15"] # "POCRCD15"
                    sheet_principal.range('N39').value = valor_pep["POCRCD16"] # "POCRCD16"
                    sheet_principal.range('N40').value = valor_pep["POCRCD17"] # "POCRCD17"
                    sheet_principal.range('N41').value = valor_pep["POCRCD18"] # "POCRCD17"
                    sheet_principal.range('N42').value = valor_pep["POCRCD19"] # "POCRCD17"
                    sheet_principal.range('N43').value = valor_pep["POCRCD20"] # "POCRCD17"
                    sheet_principal.range('N44').value = valor_pep["POCRCD21"] # "POCRCD17"
                    sheet_principal.range('N45').value = valor_pep["POCRCD22"] # "POCRCD17"
                    sheet_principal.range('N46').value = valor_pep["POCRCD23"] # "POCRCD17"
                    sheet_principal.range('N47').value = valor_pep["POCRCD24"] # "POCRCD17"
                    sheet_principal.range('N48').value = valor_pep["POCRCD25"] # "POCRCD17"
                    sheet_principal.range('N49').value = valor_pep["POCRCD26"] # "POCRCD17"
                    sheet_principal.range('N50').value = valor_pep["POCRCD27"] # "POCRCD17"
                    sheet_principal.range('N51').value = valor_pep["POCRCD28"] # "POCRCD17"
                    sheet_principal.range('N52').value = valor_pep["POCRCD29"] # "POCRCD17"
                    sheet_principal.range('N53').value = valor_pep["POCRCD30"] # "POCRCD17"
                    
                    sheet_principal.range('N55').value = valor_pep["PONI"] # ""
                    
                    sheet_principal.range('N63').value = texto_incc # ""
                    texto_incc = ""
                    
                    sheet_principal.range('N64').value = "Valor INCC" 
                    
                    
                    #break
                
                wb.save()
            app.kill()            
            #import pdb; pdb.set_trace()
    
    def preprara_pep(self, date, df) -> dict:
        lista_pep:list = [
                "POCI",
                "POCD",
                "POSP",
                "POCRCIPJ",
                "POCRCISP",
                "POCRCIIP",
                "POCRCIIPR",
                "POCRCIEQ",
                "POCRCIMO",
                "POCRCICO",
                "POCRCD01",
                "POCRCD02",
                "POCRCD03",
                "POCRCD04",
                "POCRCD05",
                "POCRCD06",
                "POCRCD07",
                "POCRCD08",
                "POCRCD09",
                "POCRCD10",
                "POCRCD11",
                "POCRCD12",
                "POCRCD13",
                "POCRCD14",
                "POCRCD15",
                "POCRCD16",
                "POCRCD17",
                "POCRCD18",
                "POCRCD19",
                "POCRCD20",
                "POCRCD21",
                "POCRCD22",
                "POCRCD23",
                "POCRCD24",
                "POCRCD25",
                "POCRCD26",
                "POCRCD27",
                "POCRCD28",
                "POCRCD29",
                "POCRCD30",
                "PONI",
            ]
        
        dicionario_pep = {pep:self.calcular_pep_por_data(date, df, pep) for pep in lista_pep}
        
        return dicionario_pep

            
            
        
if __name__ == "__main__":
    bot = Files()
    
    
    print(f"\n\n{bot.gerar_arquivos()}")