'''
Autora: Valesca Gomes Soares
'''

import sys
import json
import requests
from requests_kerberos import HTTPKerberosAuth, OPTIONAL
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from SapLogin import LoginProducao
from transacoes import aberturaNota

import cred

def SapNota():
    username=cred.l_geral
    senhaSAP=cred.s_sap_p

    #autenticação kerberos para acessar o piwebapi
    auth = HTTPKerberosAuth(force_preemptive=True, mutual_authentication=OPTIONAL, delegate=True)

    #variáveis para realização das requisições: path(caminho do piwebapi)
    path = "https://ms38013801p.westrock.com/piwebapi"
    database = "F1RDVqZOHaG2zUe9KivanpRAOwsGfawQCD1Eyb03Twtf_ihwUEtHLVRSRTE
		    tUElBUDJcMzgwMQ"

    #url da requisição do batch
    url = "https://ms38013801p.westrock.com/piwebapi/batch"

    #cabeça da requisição
    headers = {
    'X-Requested-With': '',
    'Content-Type': 'application/json'
    }

    #Login no SAP
    LoginProducao(username, senhaSAP)

    id_equip = sys.argv[5]
    
    local_instalacao=sys.argv[1]
    nome_equip=sys.argv[2]
    area=sys.argv[3]
    horas_recomendadas=sys.argv[4]
    horas_reais=sys.argv[6]
    
    #abertura das notas
    descricao = "Realizar ensaio para verificar as condições do equipamento " + nome_equip + ", pois está perto de operar/operou mais que o recomendado sem manutenção. Horas de operação recomendada pelo fabricante entre manutenções: " + horas_recomendadas + " horas, porém operou " + horas_reais + " horas."
    nota1 = aberturaNota(local_instalacao, area, descricao=descricao)
    descricao = "Realizar análise de preditiva para verificar as condições do equipamento " + nome_equip + ", pois está perto de operar/operou mais que o recomendado sem manutenção. Horas de operação recomendada pelo fabricante entre manutenções: " + horas_recomendadas + " horas, porém operou " + horas_reais + " horas."
    nota2 = aberturaNota(local_instalacao, area, descricao=descricao)
    
    nota = nota1+"/"+nota2

    if "erro" in nota1 or "erro" in nota2:
        arquivo = open("erronota.txt", "r")
        conteudo = arquivo.readlines()
        conteudo.append(nota)
        arquivo = open("erronota.txt", "w")
        arquivo.writelines(conteudo)
        arquivo.close()
        return

    #corpo da requisição
    payload = json.dumps({
        "GetAttribute": {
            "Method": "GET",
            "Resource": path + "/attributes/search?databaseWebId=" + database +"&query=Element:{Id:=" + id_equip + "} Name:='Last work order created automatically'"
        },
        "PutValue": {
            "Method": "PUT",
            "ParentIds": ["GetAttribute"],
            "RequestTemplate": {
                "Resource": path + "/attributes/{0}/value",
            },
            "Content": "{\"Value\":\"" + nota + "\"}",
            "Parameters": ["$.GetAttribute.Content.Items[*].WebId"],
        }
    })

    #reposta da requisição
    response = requests.request("POST", url, headers=headers, data=payload, auth=auth,verify=False)
    
    print("Foi criada uma nota do ",sys.argv[2], " que está no local de instalação ", sys.argv[1], " - número da nota: ", nota)

    print("SAPNOTA - acabou")

SapNota()

