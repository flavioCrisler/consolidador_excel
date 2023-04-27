import pandas as pd
import os
import datetime

# criando um dataframe vazio com a estrutura final do consolidado
colunas = [
    'Segmento',
    'País',
    'Produto',
    'Qtde de Unidades Vendidas',	
    'Preço Unitário',	
    'Valor Total',	
    'Desconto',	
    'Valor Total c/ Desconto',	
    'Custo Total',	
    'Lucro',	
    'Data',	
    'Mês',	
    'Ano'
]
consolidado = pd.DataFrame(columns=colunas)

# recebe o diretório que quero listar os arquivos
arquivos = os.listdir(r"C:\Users\fcris\OneDrive\Documentos\PYTHON\python_aplicado\projeto01\planilhas")

data = datetime.datetime.now()

# Tratando erros

for arquivo in arquivos:

    # evitar que o código quebre caso utilize um arquivo que não é excel
    # sempre que ao ler o arquivo e não for um .xlsx não vai retornar erro, vai passar para o próximo
    if arquivo.endswith(".xlsx"):
        dados_arquivo = arquivo.split('-')
        segmento = dados_arquivo[0]
        pais = dados_arquivo[1].replace('.xlsx', '') # substitui o .xlsx por nada
        
        try:
            # informado o caminho da pasta planihas para ser lida e realizei um loop com cada arquivo
            df = pd.read_excel(f"planilhas\\{arquivo}")
            df.insert(0, 'Segmento', segmento)
            df.insert(1, 'País', pais)
        except:
            with open("log_erros.txt", "a") as file:
                file.write(f"Erro ao tentar consolidar o arquivo {arquivo}. ")            
    else:
        with open("log_erros.txt", "a") as file:
            file.write(f"O arquivo {arquivo} não é um arquivo Excel válido! ") 

    consolidado = pd.concat([consolidado, df])

# coluna data em formato de exibição br
consolidado["Data"] = consolidado["Data"].dt.strftime("%d/%m/%Y")

# exportando o DataFrame consolidado para um arquivo Excel.
consolidado.to_excel(f"Report_consolidado-{data.strftime('%d-%m-%Y')}.xlsx", 
                     index=False, # não irá colocar uma primeira coluna de índice
                     sheet_name='Report consolidado') # nome da aba.

