import pandas as pd
import os
from CJI3 import mountDefaultPath
from getpass import getuser
from typing import List,Dict

class Files():
    def __init__(self, path="CJI3") -> None:
        self.tempPath: str = mountDefaultPath(path)
    
    def ler_arquivos(self) -> Dict[str, pd.DataFrame]:
        dictionary: dict = {}
        for file in os.listdir(self.tempPath):
            if file.endswith(".xlsx"):
                fileName = file.replace(".xlsx", "")
                df = pd.read_excel(self.tempPath + file)
                dictionary[fileName] = df
        return dictionary
        
        
        

if __name__ == "__main__":
    bot = Files()
    
    
    print(f"\n\n{bot.ler_arquivos()}")