import os
import pandas as pd
import logging
import json
from datetime import datetime

# Configuração do logging
logging.basicConfig(filename=os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'logfile.log'),
                    level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

def log_error(message):
    logging.error(message)

def log_info(message):
    logging.info(message)

def validar_e_limpar_dados(df):
    try:
        df = df.fillna('NULL')
        df = df.drop_duplicates()
        log_info("Validação e limpeza de dados concluídas.")
        return df
    except Exception as e:
        log_error(f"Erro ao validar e limpar dados: {str(e)}")
        print(f"Erro: {str(e)}")
        return None

def selecionar_colunas(df):
    try:
        print("\n--- Seleção de Colunas ---")
        print("Colunas disponíveis:", list(df.columns))
        colunas_selecionadas = input("Informe os nomes das colunas que deseja exportar separados por vírgula ou 'todas' para selecionar todas: ").strip()

        if colunas_selecionadas.lower() == 'todas':
            return df
        else:
            colunas_selecionadas = [col.strip() for col in colunas_selecionadas.split(',')]
            return df[colunas_selecionadas]
    except Exception as e:
        log_error(f"Erro ao selecionar colunas: {str(e)}")
        print(f"Erro: {str(e)}")
        return None

def aplicar_filtro(df):
    try:
        print("\n--- Filtro de Dados ---")
        coluna_filtro = input("Informe o nome da coluna para aplicar o filtro: ").strip()
        valor_filtro = input(f"Informe o valor para filtrar na coluna '{coluna_filtro}': ").strip()
        df_filtrado = df[df[coluna_filtro] == valor_filtro]
        print("\nRegistros filtrados:")
        print(df_filtrado)
        return df_filtrado
    except Exception as e:
        log_error(f"Erro ao aplicar filtro: {str(e)}")
        print(f"Erro: {str(e)}")
        return None

def salvar_arquivo(df, nome_arquivo, formato):
    try:
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        caminho_completo = os.path.join(desktop, f"{nome_arquivo}.{formato}")

        if formato == 'csv':
            df.to_csv(caminho_completo, index=False, sep=';')
        elif formato == 'xlsx':
            df.to_excel(caminho_completo, index=False)
        elif formato == 'txt':
            df.to_csv(caminho_completo, index=False, sep='\t')
        else:
            raise ValueError("Formato de exportação inválido.")
        
        log_info(f"Arquivo salvo em: {caminho_completo}")
        print(f"Arquivo salvo em: {caminho_completo}")
    except Exception as e:
        log_error(f"Erro ao salvar arquivo: {str(e)}")
        print(f"Erro: {str(e)}")

def criar_insert(df, nome_tabela, tipo_banco):
    try:
        print("\n--- Criação de Comandos INSERT ---")
        
        colunas_tabela = []
        for coluna in df.columns:
            nome_coluna_tabela = input(f"Informe o nome da coluna na tabela para a coluna '{coluna}': ").strip()
            colunas_tabela.append(nome_coluna_tabela)

        nome_arquivo_sql = input("Informe o nome do arquivo SQL (sem extensão): ").strip()
        caminho_arquivo_sql = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', f"{nome_arquivo_sql}.sql")
        
        with open(caminho_arquivo_sql, 'w') as file:
            for index, row in df.iterrows():
                valores = []
                for value in row:
                    if pd.isna(value):
                        valores.append("NULL")
                    elif isinstance(value, str):
                        valores.append(f"'{value.replace("'", "''")}'")
                    elif isinstance(value, pd.Timestamp):
                        valores.append(formatar_data(value, tipo_banco))
                    else:
                        valores.append(str(value))
                insert_sql = f"INSERT INTO {nome_tabela} ({', '.join(colunas_tabela)}) VALUES ({', '.join(valores)});\n"
                file.write(insert_sql)
        
        log_info(f"Arquivo SQL salvo em: {caminho_arquivo_sql}")
        print(f"Arquivo SQL salvo em: {caminho_arquivo_sql}")
    except Exception as e:
        log_error(f"Erro ao criar comandos INSERT: {str(e)}")
        print(f"Erro: {str(e)}")

def formatar_data(value, tipo_banco):
    if tipo_banco == 'mysql':
        return f"'{value.strftime('%Y-%m-%d %H:%M:%S')}'"
    elif tipo_banco == 'postgresql':
        return f"'{value.strftime('%Y-%m-%d %H:%M:%S')}'"
    elif tipo_banco == 'sqlite':
        return f"'{value.strftime('%Y-%m-%d %H:%M:%S')}'"
    elif tipo_banco == 'oracle':
        return f"TO_DATE('{value.strftime('%Y-%m-%d %H:%M:%S')}', 'YYYY-MM-DD HH24:MI:SS')"
    else:
        raise ValueError("Tipo de banco de dados não suportado para datas.")

def gerar_relatorio(df):
    try:
        num_registros = len(df)
        colunas = list(df.columns)
        print("\n--- Relatório ---")
        print(f"Número de registros: {num_registros}")
        print(f"Colunas: {', '.join(colunas)}")
        log_info(f"Relatório gerado: {num_registros} registros, Colunas: {', '.join(colunas)}")
    except Exception as e:
        log_error(f"Erro ao gerar relatório: {str(e)}")
        print(f"Erro: {str(e)}")

def exportar_configuracoes(configuracoes, nome_arquivo):
    try:
        caminho_arquivo = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', f"{nome_arquivo}.json")
        with open(caminho_arquivo, 'w') as file:
            json.dump(configuracoes, file, indent=4)
        log_info(f"Configurações exportadas para: {caminho_arquivo}")
        print(f"Configurações exportadas para: {caminho_arquivo}")
    except Exception as e:
        log_error(f"Erro ao exportar configurações: {str(e)}")
        print(f"Erro: {str(e)}")

def importar_configuracoes(nome_arquivo):
    try:
        caminho_arquivo = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', f"{nome_arquivo}.json")
        with open(caminho_arquivo, 'r') as file:
            configuracoes = json.load(file)
        log_info(f"Configurações importadas de: {caminho_arquivo}")
        return configuracoes
    except Exception as e:
        log_error(f"Erro ao importar configurações: {str(e)}")
        print(f"Erro: {str(e)}")
        return {}

def processar_arquivo():
    while True:
        try:
            caminho_arquivo = input("Informe o caminho do arquivo Excel: ").strip()
            df = pd.read_excel(caminho_arquivo)
            df = df.loc[:, ~df.columns.str.contains('^Unnamed')]  # Remove colunas "Unnamed"

            df = validar_e_limpar_dados(df)
            if df is None:
                return  # Retorna ao menu principal se houve erro

            print("\nRegistros lidos:")
            print(df)

            aplicar_filtro_decisao = input("\nDeseja aplicar um filtro nos dados? (sim/não): ").strip().lower()
            if aplicar_filtro_decisao == 'sim':
                df = aplicar_filtro(df)
                if df is None:
                    return  # Retorna ao menu principal se houve erro

            print("\nRegistros após o filtro:")
            print(df)

            selecionar_colunas_decisao = input("\nDeseja selecionar colunas específicas? (sim/não): ").strip().lower()
            if selecionar_colunas_decisao == 'sim':
                df = selecionar_colunas(df)
                if df is None:
                    return  # Retorna ao menu principal se houve erro

            gerar_relatorio(df)

            acao = input("\nVocê deseja criar um insert SQL, exportar as colunas ou exportar/importar configurações? (insert/exportar/configuracoes): ").strip().lower()

            if acao == 'insert':
                tipo_banco = input("Informe o tipo de banco de dados (mysql/postgresql/sqlite/oracle): ").strip().lower()
                if tipo_banco not in ['mysql', 'postgresql', 'sqlite', 'oracle']:
                    print("Tipo de banco de dados não suportado.")
                    continue  # Volta ao menu principal
                nome_tabela = input("Informe o nome da tabela: ").strip()
                criar_insert(df, nome_tabela, tipo_banco)
            elif acao == 'exportar':
                formato = input("Informe o formato de exportação (csv, xlsx, txt): ").strip().lower()
                if formato not in ['csv', 'xlsx', 'txt']:
                    print("Formato de exportação inválido.")
                    continue  # Volta ao menu principal
                nome_arquivo = input("Informe o nome do arquivo a ser exportado (sem extensão): ").strip()
                salvar_arquivo(df, nome_arquivo, formato)
            elif acao == 'configuracoes':
                subacao = input("Você deseja exportar ou importar configurações? (exportar/importar): ").strip().lower()
                if subacao == 'exportar':
                    configuracoes = {
                        'colunas': list(df.columns),
                        'data_lido': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    }
                    nome_arquivo = input("Informe o nome do arquivo para exportar configurações (sem extensão): ").strip()
                    exportar_configuracoes(configuracoes, nome_arquivo)
                elif subacao == 'importar':
                    nome_arquivo = input("Informe o nome do arquivo para importar configurações (sem extensão): ").strip()
                    configuracoes = importar_configuracoes(nome_arquivo)
                    print(f"Configurações importadas: {configuracoes}")
                else:
                    print("Ação inválida.")
            else:
                print("Ação inválida.")
        
        except Exception as e:
            log_error(f"Ocorreu um erro ao processar o arquivo: {str(e)}")
            print(f"Ocorreu um erro: {str(e)}")
            return  # Retorna ao menu principal

def menu_principal():
    while True:
        try:
            print("\n--- Menu Principal ---")
            print("1. Processar novo arquivo")
            print("2. Sair")
            opcao = input("Selecione uma opção: ").strip()

            if opcao == '1':
                processar_arquivo()
            elif opcao == '2':
                break
            else:
                print("Opção inválida.")
        except Exception as e:
            log_error(f"Erro no menu principal: {str(e)}")
            print(f"Erro: {str(e)}")

menu_principal()
