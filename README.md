# Descrição do Programa de Gerenciamento de Incorridos

O programa analisado é composto por quatro arquivos principais: `gerador_incorridos.py`, `CJI3.py`, `FilesManipulation.py` e `credenciais.py`, trabalhando em conjunto para automatizar a geração de relatórios de incorridos a partir de dados extraídos do SAP e outras fontes.

## gerador_incorridos.py

Este arquivo serve como ponto de entrada do programa. Ele instancia e executa os métodos principais das classes `CJI3` e `Files`, encapsuladas nos arquivos `CJI3.py` e `FilesManipulation.py`, respectivamente. Em caso de exceção, registra o erro em um arquivo de log.

## CJI3.py

Define a classe `CJI3`, responsável pela conexão com o sistema SAP, extração de dados e geração de relatórios relacionados a empreendimentos. Utiliza um decorador para garantir a conexão com o SAP antes da execução de suas funções principais. Os relatórios são salvos em uma pasta específica, definida no momento da instanciação da classe.

## FilesManipulation.py

Contém a classe `Files`, focada na manipulação de arquivos Excel gerados pelo `CJI3.py`, além de realizar cálculos específicos com os dados extraídos e gerar novos relatórios consolidados. Esta classe também acessa a web para extrair valores atualizados do INCC, utilizando Selenium para navegação e extração de dados.

## credenciais.py

Este arquivo define a classe `Credenciais`, usada para gerenciar as credenciais de acesso ao sistema SAP. As credenciais são armazenadas de forma segura em um arquivo JSON e acessadas de maneira criptografada para uso na conexão com o SAP.

O programa executa um fluxo automatizado que abrange a conexão com o SAP, extração de dados de empreendimentos, geração de relatórios detalhados de incorridos, manipulação e análise desses dados em planilhas Excel, e finalmente, a geração de relatórios consolidados com informações atualizadas, incluindo índices financeiros externos como o INCC.