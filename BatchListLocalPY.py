'''
Autora: Valesca Gomes Soares
Lista dos locais de instalação de todos os equipamentos do banco de dados 3801 que estão em PM#4
'''

import json
import requests
from requests.auth import HTTPBasicAuth
from requests_kerberos import HTTPKerberosAuth, OPTIONAL
from getpass import getpass
from urllib.parse import unquote

import cred

def listLocal():
  #autenticação kerberos para acessar o piwebapi
  auth = HTTPKerberosAuth(force_preemptive=True, mutual_authentication=OPTIONAL, delegate=True)

  #url da requisição do batch
  url = "https://ms38013801p.westrock.com/piwebapi/batch"

  #variáveis para realização das requisições: path(caminho do piwebapi), database(webid do banco de dados), attributename(atributo desejado) 
  path= "https://ms38013801p.westrock.com/piwebapi"
  database = "F1RDVqZOHaG2zUe9KivanpRAOwsGfawQCD1Eyb03Twtf_ihwUEtHLVRSRTEt
UElBUDJcMzgwMQ"
  attributename = "'SAP Local Number'"

  #corpo da requisição
  payload = json.dumps({
    "GetAllAttributes": {
          "Method": "GET",
          "Resource": path + "/attributes/search?dataBasewebID=" + database +"&query=Element:{Name:='*' Root:='Tres Barras\Paper Mill\Paper Machines\PM%234'} Name:=" + attributename
      },
      "GetAllValue": {
          "Method": "GET",
          "ParentIds": ["GetAllAttributes"],
          "RequestTemplate": {
              "Resource": path + "/streams/{0}/value?selectedFields=Value"
          },
          "Parameters": ["$.GetAllAttributes.Content.Items[*].WebId"],
      }
  })

  #cabeça da requisição
  headers = {
    'X-Requested-With': '',
    'Content-Type': 'application/json'
  }

  #resposta da requisição
  response = requests.request("POST", url, headers=headers, data=payload, auth=auth,verify=False)
  status1 = response.json()['GetAllValue']['Status']

  if(status1 == 207):
    saida = response.json()['GetAllValue']['Content']['Items']
    saida_nome = response.json()['GetAllAttributes']['Content']['Items']
    locais = {}

    for item,item2 in zip(saida,saida_nome):
      valor = item['Content']['Value']
      valor2 = item2['Links']['Element']
      if valor != "0" and valor != 0:
        nome = unquote(valor2)[unquote(valor2).rfind('/')+1:]
        locais[nome] = valor

    return locais

  else:
    print("deu errado listar locais, status: ", status1)
    print(response.json()['GetAllValue']['Content'])
    return False