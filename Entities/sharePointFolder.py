import pandas as pd

class SharePointFolder:
    """
    Classe para manipulação de dados de uma pasta do SharePoint.
    """

    @staticmethod
    def infor_obras(*, path: str) -> dict:
        """
        Extrai informações das obras a partir de um arquivo Excel.

        Args:
            path (str): Caminho para o arquivo Excel.

        Returns:
            dict: Dicionário contendo os códigos das obras a serem executadas e os nomes das obras.
        """
        result: dict = {}
        df: pd.DataFrame = pd.read_excel(path)
        
        # empreendimentos para execução
        emp_exec = df[['Código da Obra', 'Geração de Incorridos']]
        emp_exec = emp_exec[emp_exec['Geração de Incorridos'].str.lower() == "Sim".lower()]
        result['executar'] = emp_exec['Código da Obra'].tolist()
        
        # Nomes empreendimentos
        nome_emp = df[['Código da Obra', 'Nome da Obra']]
        nomes: dict = {}
        for dados in nome_emp.to_dict(orient='records'):
            nomes[dados['Código da Obra']] = dados['Nome da Obra']
        result['nomes'] = nomes
        
        return result

if __name__ == "__main__":
    from getpass import getuser

    # Exemplo de uso do método infor_obras
    infor = SharePointFolder.infor_obras(path=f"C:/Users/renan.oliveira/PATRIMAR ENGENHARIA S A/Janela da Engenharia Controle de Obras - _Base de Dados - Geral/Informações de Obras.xlsx")
    
    print(infor)
    # folder.show_folders(_print=True)
    # print(os.path.isdir())
