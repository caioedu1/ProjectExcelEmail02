import pandas as pd
import win32com.client as win32

### Objetivos:
## 1. Criar uma tabela incluindo apenas as minas que tiveram mais que 100 mil horas de trabalho.
## 2. Criar uma tabela incluindo apenas as minas que tiveram mais que 300 mil de produção
## 3. Criar uma tabela que inclui tanto as minas com mais de 300 mil de produção, quantos as que tem mais de 100 mil horas de trabalho.
## 4. Criar um arquivo excel com as novas tabelas
## 5. Mandar os arquivo automaticamente para o email
## Obs: Deve contar o id das minas indicadas


## 1. Criar uma tabela incluindo apenas as minas que tiveram mais que 100 mil horas de trabalho.
arquivo = pd.read_excel('coalpublic2013.xlsx')
pd.set_option("display.max_columns", None)

arquivo_sorted_horas = arquivo.query('Labor_Hours > 100000')
arquivo_sem_colunas = arquivo_sorted_horas.drop(["Year", "Production"], axis=1)
arquivo_final_horas = arquivo_sem_colunas.sort_values(by='Labor_Hours', ascending=True).reset_index(drop=True)
# print(arquivo_final_horas)


## 2. Criar uma tabela incluindo apenas as minas que tiveram mais que 300 mil de produção
arquivo_sorted_producao = arquivo.query('Production > 300000')
arquivo_sem_colunas2 = arquivo_sorted_producao.drop(["Year", "Labor_Hours"], axis=1)
arquivo_final_producao = arquivo_sem_colunas2.sort_values(by='Production', ascending=True).reset_index(drop=True)
# print (arquivo_final_producao)

## 3. Criar uma tabela que inclui tanto as minas com mais de 300 mil de produção, quantos as que tem mais de 100 mil horas de trabalho.
tabela_nova = arquivo.drop("Year", axis=1)
tabela_filtrada1 = tabela_nova.query('Labor_Hours > 100000')
tabela_filtrada2 = tabela_filtrada1.query('Production > 300000')
tabela_final = tabela_filtrada2.sort_values(by='Production',ascending=True).reset_index(drop=True)
#print(tabela_final)


## 4. Criar um arquivo excel com as novas tabelas
# Tabela 1
nome_do_arquivo1 = 'arquivo_horas.xlsx'
arquivo_final_horas.to_excel(nome_do_arquivo1, index=False)

# Tabela 2
nome_do_arquivo2 = 'arquivo_producao.xlsx'
arquivo_final_producao.to_excel(nome_do_arquivo2, index=False)

# Tabela 3
nome_do_arquivo3 = 'tabela_horas_producao.xlsx'
tabela_final.to_excel(nome_do_arquivo3, index=False)


## 5. Mandar os arquivo automaticamente para o email

# Enviar email

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'caioedu1031@gmail.com'
mail.Subject = 'Relatório das Minas de Carvão'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue relatório solicitado acerca das minas de carvão </p>

<p>Tabela das minas que tiveram mais que 100.000 horas de trabalho </p>
<p>{arquivo_final_horas.to_html()}</p>

<p>Tabela das minas que tiveram mais que 300.000 de produção </p>
<p>{arquivo_final_producao.to_html()}</p>
<p>Ambas as tabelas concatenadas</p>
<p>{tabela_final.to_html()} </p>


Qualquer dúvida estou à disposição
Att.,
Caio
'''

mail.Send()

print ("E-mail enviado!")