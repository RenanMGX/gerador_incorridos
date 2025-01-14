import os
import xlwings as xw
import mysql.connector
import pandas as pd

from datetime import datetime
from typing import Dict, List
from shutil import copy2
from time import sleep
from getpass import getuser
from .dependencies.credenciais import Credential
from .dependencies.config import Config
from dateutil.relativedelta import relativedelta
from dependencies.functions import Functions

class Files():
    """
    Classe para manipulação de arquivos e geração de relatórios de incorridos.
    """
    
    def __init__(self, date: datetime, description_sap_tags_path: str = "") -> None:
        """
        Inicializa a classe Files.

        Args:
            date (datetime): Data de referência.
            description_sap_tags_path (str): Caminho para o arquivo Excel com descrições SAP.
        """
        if not isinstance(date, datetime):
            raise TypeError("apenas datas no formato datetime")
        self.date: datetime = date
        
        # Define o caminho para a pasta de bases
        self.__path_bases: str = os.getcwd() + "\\Bases\\"
        if not os.path.exists(self.path_bases):
            os.makedirs(self.path_bases)
            
        # Define o caminho para a pasta de incorridos gerados
        self.__path_incorridos: str = os.getcwd() + "\\incorridos_gerados\\"
        if not os.path.exists(self.path_incorridos):
            os.makedirs(self.path_incorridos)
        for _file in os.listdir(self.path_incorridos):
            try:
                os.unlink(self.path_incorridos + _file)
            except:
                pass
        
        # Carrega as descrições SAP, se o caminho for fornecido
        self.__description_sap_tags: pd.DataFrame
        if not description_sap_tags_path == "":
            self.__description_sap_tags = pd.read_excel(description_sap_tags_path)
        else:
            self.__description_sap_tags = pd.DataFrame()
            
        self.__files_base: dict = self._listar_arquivos()
            
    @property
    def path_incorridos(self):
        return self.__path_incorridos
    
    @path_incorridos.setter
    def path_incorridos(self, value: str):
        if not isinstance(value, str):
            raise TypeError(f"o valor '{value}' atribuido para 'self.path_incorridos' não é uma string")
        self.__path_incorridos = value
        
    @property
    def path_bases(self):
        return self.__path_bases
    
    @path_bases.setter
    def path_bases(self, value: str):
        if not isinstance(value, str):
            raise TypeError(f"o valor '{value}' atribuido para 'self.path_bases' não é uma string")
        self.__path_bases = value
    
    @property
    def files_base(self):
        return self.__files_base
    
    @property
    def description_sap_tags(self):
        return self.__description_sap_tags
    
    def descript(self, *, codigo: str, centro_custo: str) -> str:
        """
        Retorna a descrição de um código SAP para um centro de custo específico.

        Args:
            codigo (str): Código SAP.
            centro_custo (str): Centro de custo.

        Returns:
            str: Descrição do código SAP.
        """
        if self.description_sap_tags.empty:
            return ""
        df: pd.DataFrame = self.description_sap_tags

        df_codigo = df[df['Código SAP'] == codigo]
        if not df_codigo.empty:
            df_centro_custo = df_codigo[df_codigo['Código da Obra'] == centro_custo]
            if not df_centro_custo.empty:
                return str(df_centro_custo['Descrição'].values[0])
            else:
                return ""
        else:
            return ""
    
    @staticmethod
    def __date_verify(datas: list):
        """
        Verifica e ajusta a lista de datas para garantir que todas as datas estejam presentes em sequência mensal.

        Args:
            datas (list): Lista de datas.

        Returns:
            list: Lista de datas ajustada.
        """
        datas = list(set(datas))
        datas = sorted(datas)
        result_datas: list = []
        last_date = ""
        
        for data in datas:
            try:
                if not last_date:
                    result_datas.append(data)
                    last_date = data
                    continue
                
                while not (data - relativedelta(months=1)) == last_date:
                    if last_date > datas[-1]:
                        break
                    last_date = last_date + (relativedelta(months=1))
                    result_datas.append(last_date)
                
                result_datas.append(data)
                last_date = data
            except:
                pass
                    
        return result_datas
       
    def gerar_incorridos(self, *, infor: dict):
        """
        Gera os relatórios de incorridos com base nas informações fornecidas.

        Args:
            infor (dict): Dicionário contendo informações das obras.
        """
        incc = self._incc_valor()
        
        for name, file_path in self.__files_base.items():
            if os.path.isdir(file_path):
                continue
            df: pd.DataFrame = self._carregar_base(path=file_path, incc_fonte=incc)
            print(f"{name} -> Executando")
            
            nome_empeendimento = infor['nomes'][name]
            path_incorrido: str = self.path_incorridos + f"Incorrido - {name} - {nome_empeendimento}.xlsx"  # nome do arquivo
    
            datas: List[pd.Timestamp] = df['Data de lançamento'].unique().tolist()
            datas.pop(datas.index(pd.NaT))  # type: ignore
            
            datas = [data.replace(day=1) for data in datas]
            datas = Files.__date_verify(datas)
            datas.reverse()

            if os.path.exists(path_incorrido):
                try:
                    os.unlink(path_incorrido)
                except PermissionError:
                    self._fechar_excel(file_name=path_incorrido)
                    os.unlink(path_incorrido)
            
            for _ in range(5*60):
                try:
                    copy2("modelo planilha\\PEP a PEP - Incorridos - Modelo.xlsx", path_incorrido)
                    break
                except PermissionError:
                    print(f"feche a planilha '{path_incorrido}'") 
                sleep(1)
            
            app = xw.App(visible=False)
            with app.books.open(path_incorrido) as wb:
                sheet_principal = wb.sheets['PEP A PEP']
                sheet_temp = wb.sheets['temp']
                
                sheet_principal.range('E2').value = f"{name} - {nome_empeendimento}"  # Nome

                sheet_principal.range('E3').value = self.date.strftime('%m/%d/%Y')  # Data referencia
                
                # descrição codigos
                sheet_principal.range('E47:E54').value = [
                    [self.descript(codigo="POCRCD23", centro_custo=name)],
                    [self.descript(codigo="POCRCD24", centro_custo=name)],
                    [self.descript(codigo="POCRCD25", centro_custo=name)],
                    [self.descript(codigo="POCRCD26", centro_custo=name)],
                    [self.descript(codigo="POCRCD27", centro_custo=name)],
                    [self.descript(codigo="POCRCD28", centro_custo=name)],
                    [self.descript(codigo="POCRCD29", centro_custo=name)],
                    [self.descript(codigo="POCRCD30", centro_custo=name)]
                ]
                
                etapas: int = len(datas)
                etapa: int = 1
                for date in datas:
                    agora = datetime.now()
                    print(f"{etapa} / {etapas} --> {date}")
                    
                    formula_coluna_h = sheet_principal.range('H1:H61').formula
                    sheet_principal.range('N1').api.EntireColumn.Insert()
                    sheet_temp.range('A:A').copy()
                    sheet_principal.range('N1').paste()
                    app.api.CutCopyMode = False
                    
                    sheet_principal.range('N6').value = date  # data
                    
                    sheet_principal.range('N11:N13').value = [
                        [self._calcular_pep_por_data(date, df, "POCI")],
                        [self._calcular_pep_por_data(date, df, "POCD")],
                        [self._calcular_pep_por_data(date, df, "POSP")]
                    ]
                
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
                    
                    sheet_principal.range('N58:N60').value = [
                        [self._calcular_pep_por_data(date, df, "POPZKT")],
                        [self._calcular_pep_por_data(date, df, "POPZOP")],
                        [self._calcular_pep_por_data(date, df, "POPZMD")]
                    ]
                    
                    # sheet_principal.range('N64').value = "Valor Mensal do INCC" if etapa == etapas else "" 

                    # try:
                    #     sheet_principal.range('N65').value = incc[date.to_pydatetime()]
                    # except:
                    #     sheet_principal.range('N65').value = 0
                    
                    sheet_principal.range('H1:H61').formula = formula_coluna_h
                    etapa += 1

                wb.sheets['temp'].delete()
                wb.save()
            app.kill()
            self._fechar_excel(file_name=path_incorrido)
            print("        Finalizado!")

    def salvar_no_destino(self, destino: str):
        """
        Salva os arquivos gerados no destino especificado.

        Args:
            destino (str): Caminho do destino onde os arquivos serão salvos.
        """
        if not (destino.endswith("\\") or destino.endswith("/")):
            destino += "\\"
        
        def criar_pasta(caminho):
            if not os.path.exists(caminho):
                os.makedirs(caminho)
        
        bases_path = destino + "Bases\\" + self.date.strftime('%d-%m-%Y\\')
        criar_pasta(bases_path)
            
        for file in os.listdir(self.path_bases):
            if file.endswith(".xlsx"):
                copy2(self.path_bases + file, bases_path)
        
        incorridos_path = destino + "Incorridos\\"
        criar_pasta(incorridos_path)
        
        for file2 in os.listdir(self.path_incorridos):
            if file2.endswith(".xlsx"):
                copy2(self.path_incorridos + file2, incorridos_path)
                
    def _calcular_pep_por_data(self, date: datetime, df: pd.DataFrame, termo: str) -> float:
        """
        Calcula o valor total de um termo específico em uma data específica.

        Args:
            date (datetime): Data de referência.
            df (pd.DataFrame): DataFrame contendo os dados.
            termo (str): Termo a ser calculado.

        Returns:
            float: Valor total calculado.
        """
        df = df[(df['Data de lançamento'].dt.year == date.year) & (df['Data de lançamento'].dt.month == date.month)]
        df = df[df['Elemento PEP'].str.contains(termo, case=False)]
        
        if not df.empty:
            return round(sum(df['Valor/moeda objeto'].tolist()), 2)
        return 0.0
    
    def _carregar_base(self, *, path: str, incc_fonte: dict) -> pd.DataFrame:
        """
        Carrega a base de dados a partir de um arquivo Excel e adiciona o valor do INCC.

        Args:
            path (str): Caminho para o arquivo Excel.
            incc_fonte (dict): Dicionário contendo os valores do INCC.

        Returns:
            pd.DataFrame: DataFrame contendo os dados carregados e ajustados.
        """
        df: pd.DataFrame = pd.read_excel(path, engine="openpyxl")
        # adicionando INCC na base e salvando ela
        incc: list = []
        calculo_incc: list = []
        for linha in df.values:
            try:
                valor_incc: float = incc_fonte[datetime.fromisoformat(str(linha[0])).replace(day=1)]
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
        
        # Tratando Base antes de retornar na função
        df = df.replace(float('nan'), "")
        
        df = df[
            (~df['Classe de custo'].astype(str).str.startswith('60')) &
            (df['Elemento PEP'] != "POCRCIAI") &
            (df['Denomin.da conta de contrapartida'].str.lower().replace(' ', '') != "CUSTO DE TERRENO".lower().replace(' ', '')) &
            (df['Denomin.da conta de contrapartida'].str.lower().replace(' ', '') != "TERRENOS".lower().replace(' ', '')) &
            (df['Denomin.da conta de contrapartida'].str.lower().replace(' ', '') != "ESTOQUE DE TERRENOS".lower().replace(' ', '')) &
            (df['Denomin.da conta de contrapartida'].str.lower().replace(' ', '') != "ESTOQUE DE TERRENO".lower().replace(' ', '')) &
            (df['Denomin.da conta de contrapartida'].str.lower().replace(' ', '') != "T. ESTOQUE INICIAL".lower().replace(' ', '')) &
            (df['Denomin.da conta de contrapartida'].str.lower().replace(' ', '') != "T.  EST. TERRENOS".lower().replace(' ', ''))             
        ]
        
        return df
    
    def _incc_valor(self):
        """
        Carrega os valores do INCC a partir do banco de dados.

        Returns:
            dict: Dicionário contendo os valores do INCC.
        """
        db_config: dict = Credential(Config()['credential']['db']).load()
        
        connection = mysql.connector.connect(
            host=db_config['host'],
            user=db_config['user'],
            password=db_config['password'],
            database=db_config['database']
        )
        
        cursor = connection.cursor()
        cursor.execute("SELECT mes, valor FROM incc")
        resultado: List[List[datetime]] = cursor.fetchall()#type: ignore
        
        indices = {}
        for indice in resultado:
            date = datetime(year=indice[0].year, month=indice[0].month, day=indice[0].day)
            indices[date] = indice[1]
        
        return indices   
    
    def _listar_arquivos(self) -> Dict[str,str]:
        """
        Lista os arquivos na pasta de bases e renomeia os arquivos, se necessário.

        Returns:
            Dict[str, str]: Dicionário contendo os nomes dos arquivos e seus caminhos completos.
        """
        lista:dict = {}
        for file in os.listdir(self.path_bases):
            if "~$" in file:
                continue
            print(file)
            new_file = file.replace(".XLSX",".xlsx")
            
            try:
                os.rename((self.path_bases + file),(self.path_bases + new_file))
            except PermissionError:
                print("aberto")
                self._fechar_excel(file_name=new_file, timeout=2)
                os.rename((self.path_bases + file),(self.path_bases + new_file))
                        
            file = new_file
            file_name:str = file[0:4]
            lista[file_name] = self.path_bases + file
        return lista
    
    def _fechar_excel(self, *, file_name: str, timeout=15) -> bool:
        """
        Fecha o arquivo Excel especificado, se estiver aberto.

        Args:
            file_name (str): Nome do arquivo a ser fechado.
            timeout (int): Tempo máximo de espera para fechar o arquivo (em segundos).

        Returns:
            bool: True se o arquivo foi fechado com sucesso, False caso contrário.
        """
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
    
if __name__ == "__main__":
    from sharePointFolder import SharePointFolder # type: ignore
    
    infor = SharePointFolder.infor_obras(path=f"C:/Users/renan.oliveira/PATRIMAR ENGENHARIA S A/Janela da Engenharia Controle de Obras - _Base de Dados - Geral/Informações de Obras.xlsx")
    date = datetime.now()
    bot = Files(date, description_sap_tags_path=f"C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\Janela da Engenharia Controle de Obras - Incorridos - SAP\\Descrição SAP.xlsx")
    bot.gerar_incorridos(infor=infor)
    #print(f"\n\n{bot.gerar_arquivos()}")
    #print(bot.copiar_destino(f"C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\Janela da Engenharia Controle de Obras - Incorridos - SAP\\"))
    #print(f"\n\n{bot.incc_valor()}")
