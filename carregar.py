import pandas as pd 
import os
import glob
from datetime import datetime
import openpyxl

def encontrar_arquivo_mais_recente(pasta, extensoes = ['*.xlsx', '*.xls']):
    try:
        if not os.path.exists(pasta):
            raise FileNotFoundError(f"Pasta não encontrada: {pasta}")
        
        arquivos = []
        for extensao in extensoes:
            padrao = os.path.join(pasta,extensao)
            arquivos.extend(glob.glob(padrao))
        if not arquivos:
            print(f"Nenhum arquivo Excel encontrado na pasta: {pasta}")
            return None
        
        arquivo_mais_recente = max(arquivos, key = os.path.getmtime)

        data_modificacao = datetime.fromtimestamp(os.path.getmtime(arquivo_mais_recente))
        tamanho_mb = os.path.getsize(arquivo_mais_recente) / (1024*1024)

        print(f"Arquivo mais recente encontrado: ")
        print(f" • Nome: {os.path.basename(arquivo_mais_recente)}")
        print(f" • Data de modificação: {data_modificacao.strftime('%d/%m/%Y %H:%M:%S')}")
        print(f" • Tamanho: {tamanho_mb:.2f} MB")

        return arquivo_mais_recente
    
    except Exception as e:
        print(f"Erro ao buscar arquivo mais recente: {e}")
        return None

def listar_arquivos_pasta(pasta, extensoes =  ['*.xlsx', '*.xls']):
    try:
         arquivos_info = []

         for extensao in extensoes:
              padrao = os.path.join(pasta,extensao)
              arquivos = glob.glob(padrao)

              for arquivo in arquivos:
                   data_mod = datetime.fromtimestamp(os.path.getmtime(arquivo))
                   tamanho = os.path.getsize(arquivo) / (1024*1024)
                   arquivos_info.append((arquivo, data_mod, tamanho))
         arquivos_info.sort(key = lambda x: x[1], reverse = True)
    
         return arquivos_info
    
    except Exception as e:
         print(f"Erro ao listar arquivos: {e}")
         return []
        
def carregar_excel(caminho_arquivo = None, pasta = None):
    try:
         if pasta:
              caminho_arquivo = encontrar_arquivo_mais_recente(pasta)
              if not caminho_arquivo:
                   return None
         if not caminho_arquivo or not os.path.exists(caminho_arquivo):
            raise FileNotFoundError(f"Arquivo não encontrado: {caminho_arquivo}")
         
         if not caminho_arquivo.lower().endswith(('.xlsx', '.xls')):
              raise ValueError("Arquivo deve ter extensão .xlsx ou .xls")
         
         print(f"Carregando arquivo: {os.path.basename(caminho_arquivo)}")

         df = pd.read_excel(caminho_arquivo)

         if df.empty:
             raise ValueError("Arquivo Excel está vazio!")
         
         print(f"Arquivo carregado com sucesso! Linhas: {len(df)}, Colunas: {len(df.columns)}")
    
         return df

    except Exception as e:
        print(f"Erro ao carregar arquivo Excel: {e}")
        return None 
        
def verificar_estrutura_arquivo(df: pd.DataFrame, colunas_obrigatorias):
     colunas_arquivo = df.columns.tolist()
     colunas_faltando = [col for col in colunas_obrigatorias if col not in colunas_arquivo]

     return len(colunas_faltando) == 0, colunas_faltando