import os
import pandas as pd

def selecionar_colunas(df):
    print("\n--- Seleção de Colunas ---")
    print("Colunas disponíveis:", list(df.columns))
    colunas_selecionadas = input("Informe os nomes das colunas que deseja exportar separados por vírgula ou 'todas' para selecionar todas: ").strip()

    if colunas_selecionadas.lower() == 'todas':
        return df
    else:
        colunas_selecionadas = [col.strip() for col in colunas_selecionadas.split(',')]
        return df[colunas_selecionadas]

def aplicar_filtro(df):
    print("\n--- Filtro de Dados ---")
    coluna_filtro = input("Informe o nome da coluna para aplicar o filtro: ").strip()
    valor_filtro = input(f"Informe o valor para filtrar na coluna '{coluna_filtro}': ").strip()
    df_filtrado = df[df[coluna_filtro] == valor_filtro]
    print("\nRegistros filtrados:")
    print(df_filtrado)
    return df_filtrado

def salvar_arquivo(df, nome_arquivo, formato):
    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
    caminho_completo = os.path.join(desktop, f"{nome_arquivo}.{formato}")

    if formato == 'csv':
        df.to_csv(caminho_completo, index=False, sep=';')
    elif formato == 'xls':
        df.to_excel(caminho_completo, index=False)
    elif formato == 'txt':
        df.to_csv(caminho_completo, index=False, sep='\t')
    else:
        print("Formato de exportação inválido.")

    print(f"Arquivo salvo em: {caminho_completo}")

def criar_insert(df, nome_tabela):
    for index, row in df.iterrows():
        valores = []
        for value in row:
            if pd.isna(value):
                valores.append("NULL")
            elif isinstance(value, str):
                valores.append(f"'{value}'")
            else:
                valores.append(str(value))
        insert_sql = f"INSERT INTO {nome_tabela} ({', '.join(df.columns)}) VALUES ({', '.join(valores)});"
        print(insert_sql)

def processar_arquivo():
    caminho_arquivo = input("Informe o caminho do arquivo Excel: ").strip()
    
    try:
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
            nome_tabela = input("Informe o nome da tabela: ").strip()
            criar_insert(df, nome_tabela)
        elif acao == 'exportar':
            formato = input("Informe o formato de exportação (csv, xls, txt): ").strip().lower()
            nome_arquivo = input("Informe o nome do arquivo a ser exportado (sem extensão): ").strip()
            salvar_arquivo(df, nome_arquivo, formato)
        else:
            print("Ação inválida.")

    except Exception as e:
        print(f"Ocorreu um erro ao ler o arquivo: {str(e)}")

def menu_principal():
    while True:
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

menu_principal()
