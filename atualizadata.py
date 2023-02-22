# -*- coding: utf-8 -*-
"""
Autora: Valesca Gomes Soares

Insere as novas datas de manutenção corretiva e preventivas em tags dos equipamentos de acordo com os valores do SAP
"""

import json
import requests
from requests_kerberos import HTTPKerberosAuth, OPTIONAL
import transacoes
from datetime import date
import urllib3
import ferramentas
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

hoje = date.today().strftime("%d.%m.%Y, %H:%M:%S")

def generalMaintDate(lista):
    
    #autenticação kerberos para acessar o piwebapi
    auth = HTTPKerberosAuth(force_preemptive=True, mutual_authentication=OPTIONAL, delegate=True)

    #variáveis para realização das requisições: path(caminho do piwebapi)
    path = "https://ms38013801p.westrock.com/piwebapi"
    
    #url da requisição do batch
    url = "https://ms38013801p.westrock.com/piwebapi/batch"

    #cabeça da requisição
    headers = {
    'X-Requested-With': '',
    'Content-Type': 'application/json'
    }
    
    if lista:
        for item in lista.items():
    
            #últimas datas de manutenção corretiva e preventiva
            datas = transacoes.notas(item[1])
            
            if datas:
            
                codigo_area = item[1].split('.')[0]
                if (codigo_area == "45" or codigo_area == "42" or codigo_area == "44" or codigo_area == "46"):
                    area = "PM4"
                elif (codigo_area == "41" or codigo_area == "43"):
                    area = "PM3"
                elif (codigo_area == "67" or codigo_area == "68"):
                    area = "RB3"
                elif (codigo_area == "77"):
                    area = "PB3"
                else:
                    return False
                
                tag = "\\\pkg-tre1-pihp1\B103-" + area + "-" + item[1].replace('.','') + "LMD-AF"
                
                #corpo da requisição
                payload = json.dumps({
                    "GetWebId": {
                        "Method": "GET",
                        "Resource": path + "/points?path=" + tag +"&selectedFields=WebId"
                    },
                    "GetValue": {
                        "Method": "GET",
                        "ParentIds": ["GetWebId"],
                        "RequestTemplate": {
                            "Resource": path + "/streams/{0}/recorded?
				    SelectedFields=Items.Value;Items.Timestamp
				    &startTime=01/01/2015&endTime=*"
                        },
                        "Parameters": ["$.GetWebId.Content.WebId"],
                    }
                })
                
                #reposta da requisição
                response = requests.request("POST", url=url, data=payload, headers=headers, auth=auth,verify=False)
                
                status = response.status_code
                
                if status == 207:
                    try:
                        valores = response.json()['GetValue']['Content']
				['Items'][0]['Content']['Items']
                    except:
                        print("tag:", tag, "valor: ", response.json()['GetValue']['Content'])
                        continue
                    
                    webid = response.json()["GetWebId"]["Content"]
			          ["WebId"]
                    
                    url_post = path + "/streams/"+ webid + "/value"
                        
                    
                    for data in datas.items():
                        
                        #corpo da requisição
                        payload_post = json.dumps({
                        "Value": data[1]
                        })
                        
                        
                        if not any(d['Value'] == data[1] for d in valores):
                        
                            #reposta da requisição
                            response = requests.request("POST", url=url_post, headers=headers, data=payload_post, auth=auth,verify=False)
                            status = response.status_code
                            
                            if status == 202:
                                print("Nova data de preventiva+corretiva cadastrada ", hoje ," para ", item[1], "novo valor: ", data[1])
                            else:
                                print("POST deu errado na geral para ", item[1] ," status: ", status, " dado novo: ", data[1])
                                
                                
                else:
                    print("deu errado: ", item[1], "status: ", status)