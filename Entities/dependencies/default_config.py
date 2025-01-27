from typing import Dict
from getpass import getuser

default:Dict[str, Dict[str,object]] = {
    'credential': {
        'crd' : 'SAP_PRD',
        'db': 'MYSQL_DB',
        'sharepoint' : 'Microsoft-RPA',
        'url' : 'https://patrimar.sharepoint.com/sites/janeladaengenharia',
        'lista' : 'Lista de Obras',
    },
    'log': {
        'hostname': 'Patrimar-RPA',
        'port': '80',
        'token': 'Central-RPA'
    },
    'paths': {
        'sap': f"C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\Janela da Engenharia Controle de Obras - Incorridos - SAP",
        'sharepoint_incorrido' : f"C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\Janela da Engenharia Controle de Obras - PEP a PEP"

    }
}