import os
import pandas as pd
import logging
from datetime import datetime

# Configuração do logging
logging.basicConfig(filename=os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', 'logfile.log'),
                    level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')

def log_error(message):
    logging.error(message)

def log_info(message):
    logging.info(message)

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
        
        # Solicita o nome das colunas da tabela para cada coluna do DataFrame
        colunas_tabela = []
        for coluna in df.columns:
            nome_coluna_tabela = input(f"Informe o nome da coluna na tabela para a coluna '{coluna}': ").strip()
            colunas_tabela.append(nome_coluna_tabela)

        # Solicita o nome do arquivo SQL
        nome_arquivo_sql = input("Informe o nome do arquivo SQL (sem extensão): ").strip()
        caminho_arquivo_sql = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop', f"{nome_arquivo_sql}.sql")
        
        with open(caminho_arquivo_sql, 'w') as file:
            # Itera sobre as linhas do DataFrame
            for index, row in df.iterrows():
                valores = []
                for value in row:
                    if pd.isna(value):
                        valores.append("NULL")
                    elif isinstance(value, str):
                        # Escapa as aspas simples dentro das strings
                        valores.append(f"'{value.replace("'", "''")}'")
                    elif isinstance(value, pd.Timestamp):
                        # Formata a data no formato SQL
                        if tipo_banco == 'mysql':
                            valores.append(f"'{value.strftime('%Y-%m-%d %H:%M:%S')}'")
                        elif tipo_banco == 'postgresql':
                            valores.append(f"'{value.strftime('%Y-%m-%d %H:%M:%S')}'")
                        elif tipo_banco == 'sqlite':
                            valores.append(f"'{value.strftime('%Y-%m-%d %H:%M:%S')}'")
                        else:
                            raise ValueError("Tipo de banco de dados não suportado para datas.")
                    else:
                        valores.append(str(value))
                # Cria o comando INSERT
                insert_sql = f"INSERT INTO {nome_tabela} ({', '.join(colunas_tabela)}) VALUES ({', '.join(valores)});\n"
                file.write(insert_sql)
        
        log_info(f"Arquivo SQL salvo em: {caminho_arquivo_sql}")
        print(f"Arquivo SQL salvo em: {caminho_arquivo_sql}")
    except Exception as e:
        log_error(f"Erro ao criar comandos INSERT: {str(e)}")
        print(f"Erro: {str(e)}")

def processar_arquivo():
    try:
        caminho_arquivo = input("Informe o caminho do arquivo Excel: ").strip()
        df = pd.read_excel(caminho_arquivo)
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]  # Remove colunas "Unnamed"

        print("\nRegistros lidos:")
        print(df)

        aplicar_filtro_decisao = input("\nDeseja aplicar um filtro nos dados? (sim/não): ").strip().lower()
        if aplicar_filtro_decisao == 'sim':
            df = aplicar_filtro(df)

        print("\nRegistros após o filtro:")
        print(df)

        selecionar_colunas_decisao = input("\nDeseja selecionar colunas específicas? (sim/não): ").strip().lower()
        if selecionar_colunas_decisao == 'sim':
            df = selecionar_colunas(df)

        acao = input("\nVocê deseja criar um insert SQL ou apenas exportar as colunas? (insert/exportar): ").strip().lower()

        if acao == 'insert':
            tipo_banco = input("Informe o tipo de banco de dados (mysql/postgresql/sqlite): ").strip().lower()
            if tipo_banco not in ['mysql', 'postgresql', 'sqlite']:
                raise ValueError("Tipo de banco de dados não suportado.")
            nome_tabela = input("Informe o nome da tabela: ").strip()
            criar_insert(df, nome_tabela, tipo_banco)
        elif acao == 'exportar':
            formato = input("Informe o formato de exportação (csv, xlsx, txt): ").strip().lower()
            nome_arquivo = input("Informe o nome do arquivo a ser exportado (sem extensão): ").strip()
            salvar_arquivo(df, nome_arquivo, formato)
        else:
            print("Ação inválida.")
            log_error("Ação inválida selecionada.")
    except Exception as e:
        log_error(f"Ocorreu um erro ao processar o arquivo: {str(e)}")
        print(f"Erro: {str(e)}")

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
                log_error("Opção inválida no menu principal.")
        except Exception as e:
            log_error(f"Erro no menu principal: {str(e)}")
            print(f"Erro: {str(e)}")

menu_principal()
