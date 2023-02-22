'''
Autora: Valesca Gomes Soares
'''

from BatchListLocalPY import listLocal
from UpdateMTBF import generalMaintDate
from SapLogin import LoginProducao

import cred

print("########## Coleta de locais de instalação ###########")
lista = listLocal()

print("########## Começo coleta de datas ###########")

username = cred.l_geral
senhaSAP= cred.s_sap_p

LoginProducao(username, senhaSAP)

#atualiza tempo entre preventivas+corretivas
print("########## Manutenção geral ###########")
generalMaintDate(lista)


print("PIAFLOCAIS - acabou")