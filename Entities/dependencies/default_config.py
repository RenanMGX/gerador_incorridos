from typing import Dict
from getpass import getuser

default:Dict[str, Dict[str,object]] = {
    'credential': {
        'crd' : 'SAP_PRD',
        'db': 'MYSQL_DB'
    },
    'log': {
        'hostname': 'Patrimar-RPA',
        'port': '80',
        'token': 'Central-RPA'
    },
    'paths': {
        'sap': f"C:\\Users\\{getuser()}\\PATRIMAR ENGENHARIA S A\\Janela da Engenharia Controle de Obras - Incorridos - SAP"
    }
}