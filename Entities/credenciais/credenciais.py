import os
import json

class Credenciais:
    def __init__(self, path_file:str="Entities/credenciais/crd.json") -> None:
        """metodo construtor da classe

        Args:
            path_file (str, optional): endereço do arquivo json onde será salvo a senha. Defaults to "Entities/credenciais/crd.json".
        """
        self.path_file:str = path_file
        if not os.path.exists(self.path_file):
            with open(self.path_file, 'w')as _file:
                json.dump({"user" : "user", "pass" : "pass", "key" : 0}, _file)
    
    def read(self) -> dict:
        """le o arquivo json e descriptografa e salva em um dict para ser utilizado

        Raises:
            FileNotFoundError: caso não encontre o arquivo json

        Returns:
            dict: dicionario com os dados para login
        """
        if os.path.exists(self.path_file):
            with open(self.path_file, 'r')as _file:
                dados:dict = json.load(_file)
                dados['pass'] = self.decifrar(dados['pass'], dados['key'])
                return dados
        raise FileNotFoundError(f"arquivo .json não encontrado no caminho '{self.path_file}'")
    
    def cifrar(self, text:str, key:int=1, response_json:bool=False) -> str:
        """criptografa a string informada orientada pela Key

        Args:
            text (str): texto a ser criptografado
            key (int, optional): chave para criptografia. Defaults to 1.
            response_json (bool, optional): retorna a string em formato json. Defaults to False.

        Returns:
            str: valor criptografado
        """
        if not isinstance(key, int):
            key = int(key)
        result:str = ""
        for letra in text:
            codigo:int = ord(letra) + key
            result += chr(codigo)
        
        if response_json:    
            return json.dumps(result)
        return result
    
    def decifrar(self, text:str, key:int) -> str:
        """descriptografa a string

        Args:
            text (str): texto a ser descriptografado
            key (int): chave para descriptografar

        Returns:
            str: texto descriptografado
        """
        return self.cifrar(text, -key)

if __name__ == "__main__":
    bot = Credenciais()
    
    print(bot.read())
    
