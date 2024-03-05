from numpy import object_
import pandas as pd
import os
from time import sleep
from CJI3 import mountDefaultPath # type: ignore
from getpass import getuser
from typing import List,Dict
from datetime import datetime
from shutil import copy2
import xlwings as xw # type: ignore
from selenium import webdriver
from selenium.webdriver.common.by import By
import json

def medir_tempo(f):
    def wrap(*args, **kwargs):
        agora: datetime = datetime.now()# type: ignore
        result = f(*args, **kwargs)
        print(f"\ntempo de execução: {datetime.now() - agora}\n")
        return result
    return wrap

def _find_element(browser:webdriver.Chrome, method, target:str, timeout=60):
    """auxiliador para o selenium ele tenta encontrar o objeto por alguns segundos ajustavel
       caso encontre retorna o objeto caso não encontre ele vai gerar o erro mas só depois do tempo acabar

    Args:
        browser (webdriver.Chrome): objeto do chrome seria o navegador
        method (object): qual methodo ira procurar
        target (str): endereço que ira procurar
        timeout (int, optional): tempo limite para tentar achar o objeto. Defaults to 60.

    Returns:
        webdriver.Chrome : se encontrar retorna o object do webdriver
    """
    for x in range(timeout):
        try:
            result = browser.find_element(method, target)
            return result
        except:
            sleep(2)
    raise Exception(f"'{target}' não foi encontrado")
        

class Files():
    def __init__(self, path="CJI3") -> None:
        self.tempPath: str = mountDefaultPath(path)
        
        
        pasta = "incorridos_gerados"
        if not os.path.exists(pasta):
            os.mkdir(pasta)
    
    def _ler_arquivos(self) -> Dict[str, str]:
        """le todos os arquivos na pasta selencionada separa apenas os excel e salva em um dict

        Returns:
            Dict[str, pd.DataFrame]: nome do arquivo, objeto pd.Dataframe
        """
        dictionary: dict = {}
        for file in os.listdir(self.tempPath):
            print(file)
            if file.endswith(".xlsx"):
                fileName:str = file.replace(".xlsx", "").replace(".PO", "")
                try:
                    caminho = self.tempPath + file
                    #df: pd.DataFrame = pd.read_excel(self.tempPath + file)
                except:
                    continue
                dictionary[fileName] = caminho
                #break
        return dictionary
    
    def calcular_pep_por_data(self, date:datetime, df:pd.DataFrame, termo:str) -> float:
        """ira fazer os calculor dos valores

        Args:
            date (datetime): data do filtro
            df (pd.DataFrame): o dataframe para ser filtrado
            termo (str): qual é a coluna que será filtrada

        Returns:
            float: valor filtrado encontrado
        """
        df = df[(df['Data de lançamento'].dt.year == date.year) & (df['Data de lançamento'].dt.month == date.month)]
        df = df[df['Elemento PEP'].str.contains(termo, case=False)]
        
        valores:float = 0
        for valor in df['Valor/moeda objeto'].tolist():
            valores += valor
        
        return round(valores, 2)
        
    @medir_tempo
    def gerar_arquivos(self, path_new_file:str = f"incorridos_gerados\\Incorridos_{datetime.now().strftime('%d-%m-%Y')}") -> None:
        r"""ira gerar as planilhas e alimentando com os dados calculados

        Args:
            path_new_file (str, optional): onde será salvo as planilhas. Defaults to f"C:\Users\{getuser()}\Downloads\Incorridos_{datetime.now().strftime('%d-%m-%Y')}".
        """
        infor_path_base:str = f"C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\"
        infor_path:str = [(infor_path_base + x + '\\Informações de Obras.xlsx') for x in os.listdir(infor_path_base) if 'Base de Dados - Geral' in x][0]        
        infor = pd.read_excel(infor_path)
        
        self.__path_new_file = path_new_file
        self.incc: dict = self.incc_valor()
        for name,caminho_arquivo in self._ler_arquivos().items():
            #df = pd.read_excel(caminho_arquivo)
            df = self.tratar_base(caminho=caminho_arquivo, incc_fonte=self.incc)

            print(f"{name} -> Executando")
            
            temp_name = infor[infor['Código da Obra'] == name[0:4]] # nome pesquisado pelo centro de custo
            temp_name = temp_name['Nome da Obra'].values[0]
            
            path_file:str = self.__path_new_file + "\\" + (f"Incorrido - {name} - {temp_name} - R00.xlsx") # nome do arquivo
            
            if not os.path.exists(self.__path_new_file):
                os.mkdir(self.__path_new_file)
                
            #df['Data de lançamento'] = pd.to_datetime(df['Data de lançamento'])
            datas:List[pd.Timestamp] = df['Data de lançamento'].unique().tolist()
            datas.pop(datas.index(pd.NaT))# type: ignore
            
            datas = [data.replace(day=1) for data in datas]
            datas = set(datas)# type: ignore
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
            
            #import pdb; pdb.set_trace()            
            app = xw.App(visible=False)
            with app.books.open(path_file)as wb:
                sheet_principal = wb.sheets['PEP A PEP']
                sheet_temp = wb.sheets['temp']
                
                sheet_principal.range('E2').value = f"{name} - {temp_name}" #Nome
                
                sheet_principal.range('E3').value = datetime.now().strftime('%d/%m/%Y') #Data referencia
                
                etapas:int = len(datas)
                etapa:int = 1
                for date in datas:
                    agora = datetime.now()
                    print(f"{etapa} / {etapas} --> {date}")
                    
                    formula_coluna_h = sheet_principal.range('H1:H120').formula
                    sheet_principal.range('N1').api.EntireColumn.Insert()
                    sheet_temp.range('A:A').copy()
                    #sheet_principal.range('N1').select()
                    sheet_principal.range('N1').paste()
                    app.api.CutCopyMode = False
                    
                    sheet_principal.range('N6').value = date #data
                    
                    sheet_principal.range('N11:N13').value = [[self.calcular_pep_por_data(date, df, "POCI")], [self.calcular_pep_por_data(date, df, "POCD")], [self.calcular_pep_por_data(date, df, "POSP")]]
                    
                    sheet_principal.range('N16:N22').value = [
                                                        [self.calcular_pep_por_data(date, df, "POCRCIPJ")],
                                                        [self.calcular_pep_por_data(date, df, "POCRCISP")],
                                                        [self.calcular_pep_por_data(date, df, "POCRCIIP")],
                                                        [self.calcular_pep_por_data(date, df, "POCRCIPR")],
                                                        [self.calcular_pep_por_data(date, df, "POCRCIEQ")],
                                                        [self.calcular_pep_por_data(date, df, "POCRCIMO")],
                                                        [self.calcular_pep_por_data(date, df, "POCRCICO")]                                                       
                                                        ]                    
                    
                    sheet_principal.range('N24:N53').value = [
                                                            [self.calcular_pep_por_data(date, df, "POCRCD01")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD02")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD03")],
                                                            [self.calcular_pep_por_data(date, df, "POCRCD04")],
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
                    #print(f"           tempo de execução: {datetime.now() - agora}")
                    sheet_principal.range('N55').value = [[self.calcular_pep_por_data(date, df, "PONI")]]
                    
                    sheet_principal.range('N57:N59').value = [[self.calcular_pep_por_data(date, df, "POPZKT")],
                                                              [self.calcular_pep_por_data(date, df, "POPZOP")],
                                                              [self.calcular_pep_por_data(date, df, "POPZMD")]
                                                               ] # ""
                    
                    sheet_principal.range('N63').value = "Valor Mensal do INCC" if etapa == etapas else "" 
                    
                    try:
                        sheet_principal.range('N64').value = self.incc[date.to_pydatetime()]
                    except:
                        sheet_principal.range('N64').value = 0
                    
                    sheet_principal.range('H1:H120').formula = formula_coluna_h
                    etapa += 1
                    #break
                
                wb.sheets['temp'].delete()
                wb.save()
            app.kill()    
            #import pdb; pdb.set_trace()
    
    def incc_valor(self) -> dict:
        """acessa o site da FGB para extrair o valor do INCC

        Returns:
            dict: data do indice, valor do indice
        """

        with open("db_connection.json", 'r')as _file:
            db_config:dict = json.load(_file)
        
        import mysql.connector
        
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

    def tratar_base(self, caminho:str, incc_fonte:dict) -> pd.DataFrame:
        """le e salva na base os valores do INCC e a divisão do Valor/modeda pelo valor INCC

        Args:
            caminho (str): caminho de onde está o arquivo
            incc_fonte (dict): dicionario com os valores INCC

        Returns:
            pd.DataFrame: DataFrame com a base já tratada
        """
        df:pd.DataFrame = pd.read_excel(caminho)
        
        incc:list = []
        calculo_incc:list = []
        for dados in df.values:
            try:
                valor_incc:float = incc_fonte[datetime.fromisoformat(str(dados[0])).replace(day=1)]
                incc.append(valor_incc)
                
                calculo_incc.append(dados[3] / valor_incc)
                
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
        
        df['incc'] = pd.DataFrame(incc)
        df['Valor_moeda objeto / incc'] = pd.DataFrame(calculo_incc)
        
        df.to_excel(caminho, index=False)
        return  df   

    def copiar_destino(self, destino:str) -> None:
        """copia as planilhas para uma pasta no sharepoint

        Args:
            destino (str): caminho do destino
        """
        #pasta_destino = destino + self.__path_new_file.split("\\")[-1] + "\\"
        pasta_destino:str = destino + "Incorridos\\"
        
        #import pdb; pdb.set_trace()
        
        if not os.path.exists(pasta_destino):
            os.makedirs(pasta_destino)
        
        for file in os.listdir(self.__path_new_file):
            file_path = self.__path_new_file + "\\" + file
            copy2(file_path, pasta_destino)
            os.unlink(file_path)
        if len(os.listdir(self.__path_new_file)) == 0 :
            os.rmdir(self.__path_new_file)
        
        destino_base:str = destino + "Bases\\"
        if not os.path.exists(destino_base):
            os.makedirs(destino_base)
        
        destino_base_por_data:str =  f"{destino_base}\\{datetime.now().strftime('%d-%m-%Y')}"
        if not os.path.exists(destino_base_por_data):
            os.makedirs(destino_base_por_data)
            
        for file_base in os.listdir(self.tempPath):
            copy2(self.tempPath + file_base, destino_base_por_data)
            os.unlink(self.tempPath + file_base)

        
if __name__ == "__main__":
    """como usar
    """
    bot = Files()
    
    print(bot.tratar_base(caminho='C:\\Users\\renan.oliveira\\.bot_ti\\CJI3\\A026.PO.xlsx', incc_fonte=bot.incc_valor()))
    #print(f"\n\n{bot.gerar_arquivos()}")
    #print(bot.copiar_destino(f"C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\Janela da Engenharia Controle de Obras - Incorridos - SAP\\"))
    #print(f"\n\n{bot.incc_valor()}")
    