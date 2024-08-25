# LeitorDados
Claro! Aqui está um exemplo de README para o código:

---

# Excel Processor & SQL Inserter

## Descrição

Este script Python permite a leitura e processamento de planilhas Excel (.xls e .xlsx), com a opção de filtrar registros, selecionar colunas, exportar dados em diferentes formatos (CSV, XLS, TXT), ou gerar scripts de inserção SQL com base nos dados processados. O script inclui um menu interativo e suporta o processamento contínuo de múltiplos arquivos.

## Funcionalidades

- **Leitura de Planilhas Excel**: Importa arquivos Excel e processa os dados contidos nas planilhas.
- **Filtragem de Registros**: Permite aplicar filtros para restringir os registros com base em critérios específicos.
- **Seleção de Colunas**: Oferece a opção de escolher quais colunas dos registros filtrados serão utilizadas para exportação ou inserção SQL.
- **Exportação de Registros**: Exporta os registros selecionados em diferentes formatos (CSV, XLS, TXT).
- **Geração de Scripts SQL**: Gera scripts de inserção SQL com base nos dados processados.
- **Menu Interativo**: Interface de menu que guia o usuário pelas diferentes etapas do processo.
- **Processamento Contínuo**: Possibilidade de processar múltiplos arquivos de forma contínua até que o usuário decida encerrar.

## Pré-requisitos

- Python 3.7 ou superior
- Bibliotecas Python:
  - `pandas`
  - `openpyxl`
  - `xlrd`

## Instalação

1. Clone este repositório ou baixe os arquivos.
2. Instale as dependências necessárias:
   ```bash
   pip install pandas openpyxl xlrd
   ```

## Como Usar

1. Execute o script Python.
2. O menu principal será exibido com as seguintes opções:
   - **Processar novo arquivo**: Importa um arquivo Excel para processamento.
   - **Sair**: Encerra o script.
3. Após selecionar "Processar novo arquivo":
   - Informe o caminho do arquivo Excel.
   - O script lerá os dados e oferecerá a opção de filtrar os registros.
   - Visualize os registros filtrados e escolha se deseja selecionar colunas específicas ou utilizar todas.
   - Escolha entre gerar um script de inserção SQL ou exportar os dados.
   - Se optar por exportar, selecione o formato (CSV, XLS, TXT) e o script perguntará o nome do arquivo que será salvo na área de trabalho.
4. O script volta ao menu principal para permitir o processamento de outros arquivos.

## Exemplo de Uso

1. Selecione a opção de processar um novo arquivo.
2. Informe o caminho do arquivo Excel.
3. Aplique filtros conforme necessário.
4. Escolha as colunas desejadas para exportação ou geração de script SQL.
5. Escolha o formato de exportação ou opte por gerar um script SQL.
6. O arquivo será salvo na área de trabalho com o nome especificado.

## Considerações

- Certifique-se de que o arquivo Excel esteja formatado corretamente e que as colunas que você deseja processar estejam claramente identificadas.
- O script é robusto para lidar com arquivos que contenham colunas "Unnamed", removendo-as automaticamente.
- Ao gerar scripts SQL, o script identifica automaticamente se os valores são strings ou números, aplicando as devidas formatações.

## Contribuição

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou pull requests.

## Licença

Este projeto é distribuído sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes.

---

Esse README fornece uma visão geral do projeto, instruções de instalação, exemplos de uso e detalhes técnicos importantes. Se precisar de alguma alteração ou informação adicional, é só me avisar!
