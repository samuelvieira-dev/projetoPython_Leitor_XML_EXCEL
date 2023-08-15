import xmltodict
import os
import pandas as pd

def pegar_infos(nome_arquivo, valores):
    # Abre o arquivo XML para leitura binária
    with open(f'nfs/{nome_arquivo}', "rb") as arquivo_xml:
        # Converte o XML em um dicionário usando a biblioteca xmltodict
        dic_arquivo = xmltodict.parse(arquivo_xml)
        
        # Verifica se o XML tem a estrutura "NFe" ou "nfeProc" para encontrar as informações da nota
        if "NFe" in dic_arquivo:
            infos_nf = dic_arquivo["NFe"]["infNFe"]
        else:
            infos_nf = dic_arquivo['nfeProc']["NFe"]["infNFe"]
        
        # Extrai informações específicas do dicionário
        numero_nota = infos_nf["@Id"]
        empresa_emissora = infos_nf['emit']['xNome']
        nome_cliente = infos_nf['dest']['xNome'] 
        endereco = infos_nf['dest']['enderDest']
        
        # Verifica se a informação de peso bruto existe dentro da seção de transporte
        if 'vol' in infos_nf['transp']:
            peso_bruto = infos_nf['transp']['vol']['pesoB']
        else:
            peso_bruto = "Não informado!"
        
        # Adiciona as informações extraídas à lista "valores"
        valores.append([numero_nota, empresa_emissora, nome_cliente, endereco, peso_bruto])

# Lista todos os arquivos no diretório "nfs"
lista_arquivos = os.listdir("nfs")

# Define os nomes das colunas para o DataFrame
colunas = ["numero_nota", "empresa_emissora", "nome_cliente", "endereco", "peso_bruto"]
valores = []

# Itera sobre cada arquivo na lista de arquivos
for arquivo in lista_arquivos:
    # Chama a função "pegar_infos" para extrair informações do arquivo atual
    pegar_infos(arquivo, valores)

# Cria um DataFrame usando as colunas e valores extraídos
tabela = pd.DataFrame(columns=colunas, data=valores)

# Converte o DataFrame em um arquivo Excel chamado "NotasFiscais.xlsx"
tabela.to_excel("NotasFiscais.xlsx", index=False)
