import pandas as pd

def ler_planilha(caminho_arquivo):
    try:
        # Lê a planilha (XLSX, CSV, etc.)
        df = pd.read_excel(caminho_arquivo)

        # Remover colunas "Unnamed"
        df = df.loc[:, ~df.columns.str.contains('^Unnamed')]

        # Remover linhas que contêm apenas valores "N/A"
        df.dropna(how='all', inplace=True)

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
    except Exception as e:
        print(f"Erro ao ler o arquivo: {e}")
        return None

def tratar_nulos(df):
    try:
        print("\nTratando colunas nulas...")
        # Preenchendo valores nulos com um valor padrão ou mantendo como estão
        df = df.fillna("N/A")  # ou outro valor de preenchimento, como 0, "", etc.
        print("Tratamento de nulos concluído:")
        print(df)
        return df
    except Exception as e:
        print(f"Erro ao tratar colunas nulas: {e}")
        return df

def separar_colunas(df):
    try:
        # Exibe as colunas disponíveis
        print("\nColunas disponíveis:")
        print(df.columns.tolist())

        # Seleciona as colunas desejadas
        colunas_selecionadas = []
        while True:
            colunas = input("Informe os nomes das colunas que deseja separar, separados por vírgula (ou 'todas' para selecionar todas): ")
            if colunas.lower() == 'todas':
                colunas_selecionadas = df.columns.tolist()
                break
            else:
                colunas_selecionadas = [col.strip() for col in colunas.split(',')]
                if all(col in df.columns for col in colunas_selecionadas):
                    break
                else:
                    print("Uma ou mais colunas não existem. Tente novamente.")
        df_separado = df[colunas_selecionadas]
        print("\nRegistros separados:")
        print(df_separado)
        return df_separado
    except Exception as e:
        print(f"Erro ao separar colunas: {e}")
        return df

def aplicar_filtro(df):
    try:
        while True:
            coluna_filtro = input("Digite o nome da coluna na qual deseja aplicar um filtro (ou 'nenhum' para pular): ").strip()
            if coluna_filtro.lower() == 'nenhum':
                break
            elif coluna_filtro in df.columns:
                valor_filtro = input(f"Digite o valor do filtro para a coluna '{coluna_filtro}': ").strip()
                df_filtrado = df[df[coluna_filtro] == valor_filtro]
                print("\nRegistros filtrados:")
                print(df_filtrado)
                return df_filtrado
            else:
                print("Coluna inválida. Tente novamente.")
        return df
    except Exception as e:
        print(f"Erro ao aplicar filtro: {e}")
        return df

def exibir_linhas(df):
    try:
        while True:
            qtd_linhas = input("Quantas linhas você deseja exibir? (Digite um número ou 'todas' para exibir todas): ").strip().lower()
            if qtd_linhas == 'todas':
                print("\nTodas as linhas selecionadas:")
                print(df)
                break
            else:
                try:
                    qtd_linhas = int(qtd_linhas)
                    if 0 < qtd_linhas <= len(df):
                        print(f"\nExibindo as primeiras {qtd_linhas} linhas:")
                        print(df.head(qtd_linhas))
                        break
                    else:
                        print(f"Por favor, insira um número entre 1 e {len(df)}.")
                except ValueError:
                    print("Entrada inválida. Tente novamente.")
    except Exception as e:
        print(f"Erro ao exibir linhas: {e}")

def criar_insert_sql(df, nome_tabela):
    try:
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
    except Exception as e:
        print(f"Erro ao criar script de inserção SQL: {e}")

def exportar_dados(df):
    try:
        formato = input("Em qual formato deseja exportar? (txt/csv/xlsx): ").strip().lower()
        caminho_exportacao = input("Informe o caminho e nome do arquivo para exportação (ex: exportado.csv): ")

        if formato == 'txt':
            df.to_csv(caminho_exportacao, index=False, sep='\t')
            print(f"Dados exportados com sucesso para {caminho_exportacao}")
        elif formato == 'csv':
            df.to_csv(caminho_exportacao, index=False)
            print(f"Dados exportados com sucesso para {caminho_exportacao}")
        elif formato == 'xlsx':
            df.to_excel(caminho_exportacao, index=False)
            print(f"Dados exportados com sucesso para {caminho_exportacao}")
        else:
            print("Formato inválido. Tente novamente.")
    except Exception as e:
        print(f"Erro ao exportar dados: {e}")

def gerar_relatorio(df, caminho_arquivo):
    try:
        # Gera um relatório simples em formato CSV ou Excel
        caminho_relatorio = caminho_arquivo.replace('.xlsx', '_relatorio.xlsx')
        with pd.ExcelWriter(caminho_relatorio) as writer:
            df.to_excel(writer, sheet_name='Dados Tratados', index=False)
            summary = df.describe(include='all').transpose()
            summary.to_excel(writer, sheet_name='Resumo')
        print(f"Relatório gerado com sucesso: {caminho_relatorio}")
    except Exception as e:
        print(f"Erro ao gerar relatório: {e}")

def main():
    try:
        caminho_arquivo = input("Informe o caminho do arquivo Excel: ")
        df = ler_planilha(caminho_arquivo)
        
        if df is None:
            return

        df = tratar_nulos(df)
        
        df_selecionado = separar_colunas(df)
        
        df_filtrado = aplicar_filtro(df_selecionado)
        
        exibir_linhas(df_filtrado)

        acao = input("\nVocê deseja criar um insert SQL ou apenas exportar os dados? (insert/exportar): ").strip().lower()
        
        if acao == 'insert':
            nome_tabela = input("Informe o nome da tabela no banco de dados: ").strip()
            criar_insert_sql(df_filtrado, nome_tabela)
        elif acao == 'exportar':
            exportar_dados(df_filtrado)
        else:
            print("Ação inválida. Tente novamente.")
        
        gerar_relatorio(df_filtrado, caminho_arquivo)

    except Exception as e:
        print(f"Ocorreu um erro: {e}")

if __name__ == "__main__":
    main()
