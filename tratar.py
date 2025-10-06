import pandas as pd 
import os 
from datetime import datetime
from carregar import carregar_excel, verificar_estrutura_arquivo

COLUNAS_OBRIGATORIAS = [
    'Chamado', 'Titulo', 'Subcategoria', 'Servico', 'Tipo', 'Canal', 
    'Justificativa N3', 'Atendente Abertura', 'Atendente Fechamento', 
    'Data Abertura', 'Data Fechamento', 'Data Encerramento', 
    'Fila Fechamento', 'Cliente', 'Fechamento (InMin)'
]

def tratar_excel(df):
    try:
        if df is None or df.empty:
            print("DataFrame está vazio ou é None")
            return None 
        
        print(f"Verificando estrutura do arquivo...")

        possui_colunas, colunas_faltando = verificar_estrutura_arquivo(df, COLUNAS_OBRIGATORIAS)

        if not possui_colunas:
            print(f"Colunas obrigatórias faltando: {colunas_faltando}")
            print(f"Colunas disponíveis no arquivo: {df.columns.tolist()}")
            return None 
        
        print(f"✅ Estrutura do arquivo validada")
        
        if 'Fila Fechamento' not in df.columns:
            print("Coluna 'Fila Fechamento' não encontrada")
            return None
         
        df_filtrado = df[df['Fila Fechamento'] == 'Equipe SAD-SEMAMP'].copy()

        if df_filtrado.empty:
            print("Nenhum registro encontrado para 'Equipe SAD-SEMAMP'")
            return None 
        
        df_tratado = df_filtrado[COLUNAS_OBRIGATORIAS].copy()

        novos_nomes_columns = {
            'Chamado'              : 'Número do Chamado',
            'Titulo'               : 'Andar',
            'Subcategoria'         : 'Subcategoria do Serviço',
            'Servico'              : 'Serviço',
            'Tipo'                 : 'Tipo do Chamado',
            'Canal'                : 'Método de Abertura',
            'Justificativa N3'     : 'Contagem de Objetos',
            'Atendente Abertura'   : 'Atendente de Abertura',
            'Atendente Fechamento' : 'Atendente de Fechamento',
            'Data Abertura'        : 'Data de Abertura',
            'Data Fechamento'      : 'Data de Fechamento',
            'Data Encerramento'    : 'Data de Encerramento',
            'Fila Fechamento'      : 'Fila de Fechamento',
            'Cliente'              : 'Cliente do Chamado',
            'Fechamento (InMin)'   : 'Tempo de Fechamento'
        }

        df_tratado.rename(columns=novos_nomes_columns, inplace=True)

        print(f"Tratando dados...")

        df_tratado['Contagem de Objetos'] = df_tratado['Contagem de Objetos'].fillna(0)
        df_tratado['Tempo de Fechamento'] = pd.to_numeric(df_tratado['Tempo de Fechamento'], errors='coerce').fillna(0).astype(int)

        df_tratado['Número do Chamado'] = df_tratado['Número do Chamado'].astype(str)

        colunas_data = ['Data de Abertura', 'Data de Fechamento', 'Data de Encerramento']
        for coluna in colunas_data:
            df_tratado[coluna] = pd.to_datetime(df_tratado[coluna], format='%Y-%m-%d %H:%M:%S', errors='coerce')

        print(f"✅ Dados tratados com sucesso!")
        print(f"📊 Registros finais: {len(df_tratado)}")
        
        if 'Data de Abertura' in df_tratado.columns and not df_tratado['Data de Abertura'].isna().all():
            data_min = df_tratado['Data de Abertura'].min()
            data_max = df_tratado['Data de Abertura'].max()
            print(f"📅 Período: {data_min.strftime('%d/%m/%Y')} a {data_max.strftime('%d/%m/%Y')}")
        
        return df_tratado

    except Exception as e:
        print(f"Erro no tratamento dos dados: {e}")
        return None 

def salvar_relatorio(df: pd.DataFrame, pasta_saida: str = "relatorios", nome_arquivo: str = None):
    try:
        if df is None or df.empty:
            print("DataFrame está vazio, não é possível salvar")
            return False
            
        if nome_arquivo is None:
            data_atual = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_arquivo = f"relatorio_SAD_SEMAMP_{data_atual}.xlsx"

        if not nome_arquivo.endswith('.xlsx'):
            nome_arquivo += '.xlsx'

        caminho_completo = os.path.join(pasta_saida, nome_arquivo)

        os.makedirs(pasta_saida, exist_ok=True)

        df.to_excel(caminho_completo, index=False)

        tamanho_mb = os.path.getsize(caminho_completo) / (1024 * 1024)

        print(f"💾 Arquivo salvo com sucesso!")
        print(f"   • Local: {caminho_completo}")
        print(f"   • Registros: {len(df)}")
        print(f"   • Tamanho: {tamanho_mb:.2f} MB")

        return True 
        
    except Exception as e:
        print(f"Erro ao salvar arquivo: {e}")
        return False