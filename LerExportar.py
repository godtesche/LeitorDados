import pandas as pd

def ler_arquivo_excel(caminho_arquivo):

    # Lê o arquivo Excel
    df = pd.read_excel(caminho_arquivo)
    print("Colunas disponíveis no arquivo:")
    print(df.columns.tolist())
    
    # Exibir registros lidos
    while True:
        exibir_registros = input("Deseja exibir os registros lidos? (sim/não): ").strip().lower()
        if exibir_registros == 'sim':
            print("\nRegistros lidos:")
            print(df)
            break
        elif exibir_registros == 'não':
            break
        else:
            print("Resposta inválida. Tente novamente.")
    
    return df

def selecionar_colunas(df):
    colunas_selecionadas = []
    while True:
        print("\nSelecione as colunas que deseja exportar:")
        print(df.columns.tolist())
        colunas = input("Informe os nomes das colunas separados por vírgula ou 'todas' para selecionar todas: ")
        if colunas.lower() == 'todas':
            colunas_selecionadas = df.columns.tolist()
            break
        else:
            colunas_selecionadas = [col.strip() for col in colunas.split(',')]
            if all(col in df.columns for col in colunas_selecionadas):
                break
            else:
                print("Uma ou mais colunas não existem. Tente novamente.")
    return df[colunas_selecionadas]

def criar_insert_sql(df, nome_tabela):
    colunas_arquivo = df.columns.tolist()
    print("\nMapeie as colunas do arquivo para as colunas da tabela:")
    colunas_mapeadas = {}
    for coluna in colunas_arquivo:
        coluna_tabela = input(f"Informe o nome da coluna na tabela para '{coluna}' (ou pressione Enter para manter o mesmo nome): ")
        if coluna_tabela.strip() == "":
            coluna_tabela = coluna
        colunas_mapeadas[coluna] = coluna_tabela

    with open('script_inserts.sql', 'w') as f:
        for _, row in df.iterrows():
            valores = ', '.join([f"'{str(val)}'" if pd.notnull(val) else 'NULL' for val in row])
            colunas_str = ', '.join([colunas_mapeadas[col] for col in colunas_arquivo])
            insert_stmt = f"INSERT INTO {nome_tabela} ({colunas_str}) VALUES ({valores});\n"
            f.write(insert_stmt)
        print("\nScript de inserção SQL gerado com sucesso: script_inserts.sql")

def exportar_colunas(df):
    caminho_exportacao = input("Informe o caminho e nome do arquivo para exportação (ex: exportado.csv): ")
    df.to_csv(caminho_exportacao, index=False)
    print(f"Colunas exportadas com sucesso para {caminho_exportacao}")

def main():
    caminho_arquivo = input("Informe o caminho do arquivo Excel: ")
    df = ler_arquivo_excel(caminho_arquivo)
    
    df_selecionado = selecionar_colunas(df)
    
    acao = input("\nVocê deseja criar um insert SQL ou apenas exportar as colunas? (insert/exportar): ").strip().lower()
    
    if acao == 'insert':
        nome_tabela = input("Informe o nome da tabela no banco de dados: ").strip()
        criar_insert_sql(df_selecionado, nome_tabela)
    elif acao == 'exportar':
        exportar_colunas(df_selecionado)
    else:
        print("Ação inválida. Tente novamente.")

if __name__ == "__main__":
    main()
