import pandas as pd 
import time
import os
import glob
from datetime import datetime
from carregar import carregar_excel
from tratar import tratar_excel, salvar_relatorio

class ProcessarDados():
    def __init__(self, pasta_entrada = None, pasta_saida = 'relatorios'):
        self.pasta_entrada   = pasta_entrada or self._obter_pasta_padrao()
        self.pasta_saida     = pasta_saida
        self.df              = None
        self.df_tratado      = None
        self.inicio_execucao = None
        self.logs            = []
        self.arquivo_atual   = None 

    def _obter_pasta_padrao(self):
        return os.path.dirname(os.path.abspath(__file__))

    def registrar_log(self, mensagem):
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_mensagem = f"[{timestamp}] {mensagem}"
        print(log_mensagem)
        self.logs.append(log_mensagem)

    def listar_arquivos_pasta(self, pasta, extensoes = ['*.xlsx', '*.xls']):
        try:
            arquivos_info = []
            
            for extensao in extensoes:
                padrao = os.path.join(pasta, extensao)
                arquivos = glob.glob(padrao)
                
                for arquivo in arquivos:
                    data_mod = datetime.fromtimestamp(os.path.getmtime(arquivo))
                    tamanho = os.path.getsize(arquivo) / (1024 * 1024)
                    arquivos_info.append((arquivo, data_mod, tamanho))
            
            arquivos_info.sort(key=lambda x: x[1], reverse=True)
            return arquivos_info
            
        except Exception as e:
            self.registrar_log(f"Erro ao listar arquivos: {e}")
            return []

    def listar_arquivos_disponíveis(self):
        try:
            self.registrar_log(f"Verificando pasta: {self.pasta_entrada}")

            arquivos_info = self.listar_arquivos_pasta(self.pasta_entrada)

            if not arquivos_info:
                self.registrar_log("Nenhum arquivo Excel encontrado na pasta")
                return False 
            
            self.registrar_log(f"Arquivos Excel encontrados ({len(arquivos_info)})")

            for i, (arquivo, data_mod, tamanho) in enumerate(arquivos_info, 1):
                nome_arquivo = os.path.basename(arquivo)
                data_str = data_mod.strftime("%d/%m/%Y %H:%M:%S")
                status = "MAIS RECENTE" if i == 1 else ""

                print(f" {i}. {nome_arquivo}")
                print(f" Modificado: {data_str}")
                print(f" Tamanho: {tamanho:.2f} MB {status}")
                print()

            return True
            
        except Exception as e:
            self.registrar_log(f"Erro ao listar arquivos: {e}")
            return False

    def carregar_dados(self):
        try:
            if not self.listar_arquivos_disponíveis():
                return False
            
            self.registrar_log("Carregando dados do arquivo Excel...")
            self.df = carregar_excel(pasta = self.pasta_entrada)
            
            if self.df is not None:
                self.registrar_log(f"Dados carregados com sucesso! Total: {len(self.df)} registros ")
                return True
            else:
                self.registrar_log("Falha no carregamento dos dados")
                return False 
             
        except Exception as e:
            self.registrar_log(f"Erro ao carregar dados: {e}")
            return False 
        
    def processar_dados(self):
        try:
            self.registrar_log("Processando e tratando dados...")
            self.df_tratado = tratar_excel(self.df)

            if self.df_tratado is not None:
                self.registrar_log(f"Dados processados com sucesso! Registro válidos: {len(self.df_tratado)}")
                return True 
            else:
                self.registrar_log("Falha no processamento dos dados")
                return False 
        except Exception as e:
            self.registrar_log(f"Erro no processamento dos dados: {e}")
            return False
        
    def gerar_relatorios(self):
        try:
            self.registrar_log("Gerando Relatório Excel...")

            data_atual = datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_arquivo = f"relatorio_SAD_SEMAMP_{data_atual}.xlsx"

            if salvar_relatorio(self.df_tratado, self.pasta_saida, nome_arquivo):
                self.registrar_log("Relatório gerado com sucesso!")
                return True 
            else:
                self.registrar_log("Falha na geração do relatório")
                return False 
        
        except Exception as e:
            self.registrar_log(f"Erro na geração do relatório: {e}")
            return False 
        
    def executar_processamento(self):
        self.inicio_execucao = time.time()

        print("=" * 80)
        print("SISTEMA AUTOMÁTICO DE GERAÇÃO DE RELATÓRIOS - INICIADO")
        print("Processamento Mensal - EQUIPE SAD-SEMAMP")
        print("=" * 80)

        etapas = [
            ("Carregamento de Dados", self.carregar_dados),
            ("Processamento de Dados", self.processar_dados),
            ("Geração de Relatório", self.gerar_relatorios)
        ]

        sucessos = 0
        total_etapas = len(etapas)

        for i, (nome, funcao) in enumerate(etapas, 1):
            self.registrar_log(f"Etapa {i}/{total_etapas}: {nome}")

            if funcao():
                sucessos += 1
                self.registrar_log(f"Etapa '{nome}' concluído com sucesso")
            else:
                self.registrar_log(f"Etapa '{nome}' falhou.")
                if i == 1:
                    self.registrar_log("Carregamento Interrompido - falha no carregamento de dados")
                    break
        
        tempo_total = time.time() - self.inicio_execucao

        self._exibir_resumo(tempo_total, sucessos, total_etapas)
        
    def _exibir_resumo(self,tempo_total, sucessos, total_etapas):
        print("\n" + "=" * 80)
        print("Resumo da execução")
        print("\n" + "=" * 80)
        print(f"Tempo total de execução: {tempo_total:.2f} segundos")
        print(f"Processos executados com sucesso: {sucessos}/{total_etapas}")

        if sucessos == total_etapas:
            print("Carregamento Concluído com SUCESSO!")
            print("\n Arquivos gerados:")
            print(" • relatorio.xlsx")
        
            if self.df_tratado is not None:
                print(f"\n Estatísticas:")
                print(f" • Total de registros processados: {len(self.df_tratado)}")

                if 'Data de Abertura' in self.df_tratado.columns:
                    data_min = self.df_tratado['Data de Abertura'].min()
                    data_max = self.df_tratado['Data de Abertura'].max()
                    if pd.notna(data_min) and pd.notna(data_max):
                        print(f" • Período: {data_min.strftime('%d/%m/%Y')} a {data_max.strftime('%d/%m/%Y')}")
        else:
            print("Processamento concluído com FALHAS")
            print("Verifique os logs acima para identificar o problema")

        print("=" * 80)

    def configurar_pastas(self):
        print("\n CONFIGURAÇÃO DE PASTAS")
        print("=" * 40)

        nova_pasta = input(f"Pasta de Entrada Atual:{self.pasta_entrada}\n Digite nova pasta: ").strip()
        if nova_pasta and os.path.exists(nova_pasta):
            self.pasta_entrada = nova_pasta 
            print(f"Pasta de entrada atualizada: {self.pasta_entrada}")
        elif nova_pasta:
            print(f"Pasta não encontrada: {nova_pasta}")

        nova_saida = input(f"Pasta de saída atual: {self.pasta_saida}\n Digite nova pasta: ").strip()
        if nova_saida:
            self.pasta_saida = nova_saida
            print(f"Pasta de saída atualizada: {self.pasta_saida}")

    def salvar_logs(self, nome_arquivo = None):
        try:
            if nome_arquivo is None:
                data_atual = datetime.now().strftime("%Y%m%d_%H%M%S")
                nome_arquivo = f"logs_processamento_{data_atual}.txt"
            caminho_logs = os.path.join(self.pasta_saida, nome_arquivo)

            os.makedirs(self.pasta_saida, exist_ok=True)

            with open(caminho_logs, 'w', encoding = 'utf-8') as f:
                f.write("LOGS DO PROCESSAMENTO AUTOMÁTICO DE DADOS\n")
                f.write("=" * 50 + "\n\n")
                f.write(f"Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
                f.write(f"Pasta de entrada: {self.pasta_entrada}\n")
                f.write(f"Pasta de saída: {self.pasta_saida}\n")

                for log in self.logs:
                    f.write(log + "\n")

                print(f"Logs salvos em: {caminho_logs}")

        except Exception as e:
            print(f"Erro ao salvar logs: {e}")

def main():
    print("SISTEMA AUTOMÁTICO DE GERAÇÃO DE RELATÓRIOS")
    print("Processamento Mensal - Equipe SAD-SEMAMP")
    print("-" * 60) 

    processador = ProcessarDados()

    while True:
        print("\n OPÇÕES:")
        print("1. Executar processamento automático")
        print("2. Configurar pastas")
        print("3. Listar arquivos disponíveis")
        print("4. Sair")

        opcao = input("\n Escolha uma opção(1-4):").strip()

        if opcao == "1":
            processador.executar_processamento()

            salvar_logs = input("\n Deseja salvar os logs?(s/n)").strip().lower()
            if salvar_logs in  ['s', 'sim', 'y', 'yes']:
                processador.salvar_logs()
            break 
        elif opcao == "2":
            processador.configurar_pastas()
        elif opcao == "3":
            processador.listar_arquivos_disponíveis()
        elif opcao == "4":
            print("Saindo do Sistema...")
            break
        else:
            print(f"Opção inválida! Tente novamente.")

if __name__ == "__main__":
    main()