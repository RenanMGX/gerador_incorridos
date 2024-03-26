import os
import xlwings as xw
import json
import mysql.connector
import pandas as pd

from datetime import datetime
from typing import Dict, List
from shutil import copy2
from time import sleep

class Files():
    def __init__(self, date:datetime) -> None:
        self.date:datetime = date
        
        self.__path_bases:str = os.getcwd() + "\\Bases\\"
        if not os.path.exists(self.path_bases):
            os.makedirs(self.path_bases)
        for _file in os.listdir(self.path_bases):
            try:
                os.unlink(self.path_bases + _file)
            except:
                pass
        self.__path_incorridos:str = os.getcwd() + "\\incorridos_gerados\\"
        if not os.path.exists(self.path_incorridos):
            os.makedirs(self.path_incorridos)
        for _file in os.listdir(self.path_incorridos):
            try:
                os.unlink(self.path_incorridos)
            except:
                pass
            
        self.__files_base = self._listar_arquivos()
            
    @property
    def path_incorridos(self):
        return self.__path_incorridos
    
    @path_incorridos.setter
    def path_incorridos(self, value:str):
        if not isinstance(value, str):
            raise TypeError(f"o valor '{value}' atribuido para 'self.path_incorridos' não é uma string")
        self.__path_incorridos = value
        
    @property
    def path_bases(self):
        return self.__path_bases
    
    @path_bases.setter
    def path_bases(self, value:str):
        if not isinstance(value, str):
            raise TypeError(f"o valor '{value}' atribuido para 'self.path_bases' não é uma string")
        self.__path_bases = value
    
    @property
    def files_base(self):
        return self.__files_base
    
    def gerar_incorridos(self, *, infor:dict):
        incc = self._incc_valor()
        for name, file_path in self.__files_base.items():
            df:pd.DataFrame = self._carregar_base(path=file_path, incc_fonte=incc)
            print(f"{name} -> Executando")
            
            nome_empeendimento = infor['nomes'][name]
            path_incorrido:str = self.path_incorridos + f"Incorrido - {name} - {nome_empeendimento}.xlsx"# nome do arquivo
    
            datas:List[pd.Timestamp] = df['Data de lançamento'].unique().tolist()
            datas.pop(datas.index(pd.NaT))# type: ignore
            
            datas = [data.replace(day=1) for data in datas]
            datas = set(datas)# type: ignore
            datas = list(datas)
            datas = sorted(datas)
            datas.reverse()

            if os.path.exists(path_incorrido):
                try:
                    os.unlink(path_incorrido)
                except PermissionError:
                    for open_file in xw.apps:
                        if open_file.books[0].name == path_incorrido:
                            open_file.kill()
                            os.unlink(path_incorrido)
            
            for _ in range(5*60):
                try:
                    copy2("modelo planilha\\PEP a PEP - Incorridos - Modelo.xlsx", path_incorrido)
                    break
                except PermissionError:
                    print(f"feche a planilha '{path_incorrido}'") 
                sleep(1)
            
            #pocrcito = df['Denominação de objeto'][df['Elemento PEP'].str.contains('POCRCITO', case=False)].unique().tolist()[0]
            #import pdb; pdb.set_trace()
            app = xw.App(visible=False)
            with app.books.open(path_incorrido)as wb:
                sheet_principal = wb.sheets['PEP A PEP']
                sheet_temp = wb.sheets['temp']
                
                sheet_principal.range('E2').value = f"{name} - {nome_empeendimento}" #Nome
                
                sheet_principal.range('E3').value = self.date.strftime('%d/%m/%Y') #Data referencia
                
                etapas:int = len(datas)
                etapa:int = 1
                for date in datas:
                    agora = datetime.now()
                    print(f"{etapa} / {etapas} --> {date}")
                    
                    formula_coluna_h = sheet_principal.range('H1:H130').formula
                    sheet_principal.range('N1').api.EntireColumn.Insert()
                    sheet_temp.range('A:A').copy()
                    sheet_principal.range('N1').paste()
                    app.api.CutCopyMode = False
                
                    sheet_principal.range('N6').value = date #data
                    
                    sheet_principal.range('N11:N13').value = [[self._calcular_pep_por_data(date, df, "POCI")], [self._calcular_pep_por_data(date, df, "POCD")], [self._calcular_pep_por_data(date, df, "POSP")]]
                
                    sheet_principal.range('N16:N23').value = [
                                                        [self._calcular_pep_por_data(date, df, "POCRCIPJ")],
                                                        [self._calcular_pep_por_data(date, df, "POCRCISP")],
                                                        [self._calcular_pep_por_data(date, df, "POCRCIIP")],
                                                        [self._calcular_pep_por_data(date, df, "POCRCIPR")],
                                                        [self._calcular_pep_por_data(date, df, "POCRCIEQ")],
                                                        [self._calcular_pep_por_data(date, df, "POCRCIMO")],
                                                        [self._calcular_pep_por_data(date, df, "POCRCICO")],
                                                        [self._calcular_pep_por_data(date, df, "POCRCITO")]                                                      
                                                        ]
                    
                    # try:
                    #     sheet_principal.range('E23').value = df['Denominação de objeto'][df['Elemento PEP'].str.contains('POCRCITO', case=False)].unique().tolist()[0]
                    # except:
                    #     pass

                    sheet_principal.range('N25:N54').value = [
                                                            [self._calcular_pep_por_data(date, df, "POCRCD01")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD02")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD03")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD04")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD05")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD06")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD07")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD08")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD09")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD10")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD11")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD12")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD13")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD14")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD15")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD16")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD17")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD18")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD19")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD20")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD21")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD22")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD23")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD24")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD25")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD26")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD27")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD28")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD29")],
                                                            [self._calcular_pep_por_data(date, df, "POCRCD30")]
                                                            ]

                    sheet_principal.range('N56').value = [[self._calcular_pep_por_data(date, df, "PONI")]]
                    
                    sheet_principal.range('N58:N60').value = [[self._calcular_pep_por_data(date, df, "POPZKT")],
                                                              [self._calcular_pep_por_data(date, df, "POPZOP")],
                                                              [self._calcular_pep_por_data(date, df, "POPZMD")]
                                                               ] # ""
                    
                    sheet_principal.range('N64').value = "Valor Mensal do INCC" if etapa == etapas else "" 

                    try:
                        sheet_principal.range('N65').value = incc[date.to_pydatetime()]
                    except:
                        sheet_principal.range('N65').value = 0
                    
                    sheet_principal.range('H1:H120').formula = formula_coluna_h
                    etapa += 1

                wb.sheets['temp'].delete()
                wb.save()
            app.kill()
            for open_file in xw.apps:
                if open_file.books[0].name == path_incorrido:
                    open_file.kill()
            print("        Finalizado!")

    def salvar_no_destino(self, destino:str):
        if not (destino.endswith("\\")) or (destino.endswith("/")):
            destino += "\\"
        
        bases_path = destino + "Bases\\"
        if not os.path.exists(bases_path):
            os.makedirs(bases_path)
        bases_path += self.date.strftime('%d-%m-%Y\\')
        if not os.path.exists(bases_path):
            os.makedirs(bases_path)
            
        for file in os.listdir(self.path_bases):
            if file.endswith(".xlsx"):
                copy2(self.path_bases+file, bases_path)
        
        incorridos_path = destino + "Incorridos\\"
        if not os.path.exists(incorridos_path):
            os.makedirs(incorridos_path)
        
        for file2 in os.listdir(self.path_incorridos):
            if file2.endswith(".xlsx"):
                copy2(self.path_incorridos+file2, incorridos_path)
                
    def _calcular_pep_por_data(self, date:datetime, df:pd.DataFrame, termo:str) -> float:
        df = df[(df['Data de lançamento'].dt.year == date.year) & (df['Data de lançamento'].dt.month == date.month)]
        df = df[df['Elemento PEP'].str.contains(termo, case=False)]
        
        return round(sum(df['Valor/moeda objeto'].tolist()), 2)
    
    def _carregar_base(self, *, path:str, incc_fonte:dict) -> pd.DataFrame:
        df: pd.DataFrame = pd.read_excel(path)
        
        # adicionando INCC na base e salvando ela
        incc:list = []
        calculo_incc:list = []
        for linha in df.values:
            try:
                valor_incc:float = incc_fonte[datetime.fromisoformat(str(linha[0])).replace(day=1)]
                incc.append(valor_incc)
                calculo_incc.append(linha[3] / valor_incc)
            except:
                incc.append(0.0)
                calculo_incc.append(0.0)
        try:
            del df['incc']
        except KeyError:
            pass
        try:
            del df['Valor_moeda objeto / incc']
        except KeyError:
            pass
        
        for _ in range(5*60):
            try:
                df.to_excel(path, index=False)
                break
            except PermissionError:
                print(f"feche a planilha '{path}'")
            sleep(1)
        
        #Tratando Base antes de retornar na função
        df = df.replace(float('nan'), "")
        df = df[~df['Classe de custo'].astype(str).str.startswith('60')]
        df = df[df['Elemento PEP'] != "POCRCIAI"]
        df = df[df['Denomin.da conta de contrapartida'] != "CUSTO DE TERRENO"]
        df = df[df['Denomin.da conta de contrapartida'] != "TERRENOS"]
        df = df[df['Denomin.da conta de contrapartida'] != "ESTOQUE DE TERRENOS"]
        df = df[df['Denomin.da conta de contrapartida'] != "ESTOQUE DE TERRENO"]
        df = df[df['Denomin.da conta de contrapartida'] != "T. ESTOQUE INICIAL"]
        df = df[df['Denomin.da conta de contrapartida'] != "T.  EST. TERRENOS"]
        df = df[df['Denomin.da conta de contrapartida'] != "T. EST. TERRENOS"]
        
        return df
    
    def _incc_valor(self):
        with open("db_connection.json", 'r')as _file:
            db_config:dict = json.load(_file)
        
        connection = mysql.connector.connect(
            host=db_config['host'],
            user=db_config['user'],
            password=db_config['password'],
            database=db_config['database']
        )
        
        cursor = connection.cursor()
        cursor.execute("SELECT mes, valor FROM incc")
        resultado:list = cursor.fetchall()
        
        indices = {}
        for indice in resultado:
            date = datetime(year=indice[0].year, month=indice[0].month, day=indice[0].day)
            indices[date] = indice[1]
        
        return indices   
    
    def _listar_arquivos(self) -> Dict[str,str]:
            
        lista:dict = {}
        for file in os.listdir(self.path_bases):
            new_file = file.replace(".XLSX",".xlsx")
            os.rename((self.path_bases + file),(self.path_bases + new_file))
            file = new_file
            if file.lower().endswith(".xlsx"):
                for file_open in xw.apps:
                    if file_open.books[0].name.lower() == file.lower():
                        file_open.kill()
            else:
                continue
            file_name:str = file[0:4]
            lista[file_name] = self.path_bases + file
        return lista
    
if __name__ == "__main__":
    pass
    #print(f"\n\n{bot.gerar_arquivos()}")
    #print(bot.copiar_destino(f"C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\Janela da Engenharia Controle de Obras - Incorridos - SAP\\"))
    #print(f"\n\n{bot.incc_valor()}")
    