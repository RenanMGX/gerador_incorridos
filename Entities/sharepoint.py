import pandas as pd
import os

from typing import Literal
from Entities.dependencies.functions import Functions
from Entities.dependencies.config import Config
from Entities.dependencies.credenciais import Credential
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext


class SharePoint:
    @property
    def df(self) -> pd.DataFrame:
        try:
            return self.__df
        except AttributeError:
            self.consultar()
            return self.__df
    
    @property
    def download_path(self):
        download_path:str = os.path.join(os.getcwd(), "Attachments_Download")
        if not os.path.exists(download_path):
            os.makedirs(download_path)
        return download_path
    
    def __init__(self) -> None:        
        crd:dict = Credential(Config()['credential']['sharepoint']).load()
        url = Config()['credential']['url']
        lista = Config()['credential']['lista']
        
        self.__ctx_auth = AuthenticationContext(url)
        if self.__ctx_auth.acquire_token_for_user(crd['email'], crd['password']):
            self.__ctx = ClientContext(url, self.__ctx_auth)
        else:
            raise PermissionError("não foi possivel acessar a lista")
        
        self.__lista = self.__ctx.web.lists.get_by_title(lista)
        
        self.consultar()
        
    def consultar(self, with_attachment:bool=False):
        items = self.__lista.get_items()
        self.__ctx.load(items)
        self.__ctx.execute_query()
        
        self.limpar_pasta_download() if with_attachment else None
        
        list_valid = []
        for item in items:
            list_valid.append(item.properties)
                    
        self.__df = pd.DataFrame(list_valid)
            
        return self
    
            
    def alterar(self, id, *, valor:str, coluna:str) -> None:
        item_to_update = self.__lista.get_item_by_id(id)
        
        # Atualizando os campos do item
        item_to_update.set_property(coluna, valor)
        
        #item_to_update.set_property("OutroCampo", "Novo Valor")
        item_to_update.update()
            
        # Executando a atualização no servidor
        self.__ctx.execute_query()
        
        self.consultar()
        
    def limpar_pasta_download(self) -> None:
        for file in os.listdir(self.download_path):
            file:str = os.path.join(self.download_path, file)
            
            if os.path.isfile(file):
                try:
                    os.unlink(file)
                except PermissionError as error:
                    print(error)
                    Functions.fechar_excel(file)
                    os.unlink(file)    

if __name__ == "__main__":
    pass
