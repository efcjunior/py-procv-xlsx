# py-procv-xlsx

# Como a biblioteca funciona

1. Carrega a planilha do arquivo XLSX origem;
2. Carrega a planilha do arquivo XLSX destino;
2. Procura um valor da coluna A do arquivo origem na coluna I do arquivo destino;
3. Caso haja uma igualdade, então atualiza a coluna AH na linha correspondente com valor da coluna B do arquivo origem;

# Pré-requisitos

* Python >= 3.10
* Pacote openpyxl

# Como utilizar e executar a biblioteca

## Executando diretamente script

1. Abra o código fonte `excel_file_updater_script`
2. Configure as variáveis que indicam os nomes dos arquivos, suas respectivas planilhas e o índices das colunas de pesquisa de atualização:

```
_source_file = 'files/source_file.xlsx'
_source_sheet = 'Sheet1'
_source_value_column_index = 0  # column 'A'
_source_content_column_index = 1  # column 'B'

_target_file = 'files/target_file.xlsx'
_target_sheet = 'BillingOffer'
_target_value_column_index = 8  # column 'I'
_target_content_column_index = 33  # column 'AH'
```

## Importando a biblioteca como módulo no seu script

1. Abra o código fonte `main.py` (ou seu programa)
2. Importe o módulo `from excel_file_updater import ExcelFileUpdater`
3. Configure as variáveis que indicam os nomes dos arquivos, suas respectivas planilhas e o índices das colunas de pesquisa de atualização:

```
_source_file = 'files/source_file.xlsx'
_source_sheet = 'Sheet1'
_source_value_column_index = 0  # column 'A'
_source_content_column_index = 1  # column 'B'

_target_file = 'files/target_file.xlsx'
_target_sheet = 'BillingOffer'
_target_value_column_index = 8  # column 'I'
_target_content_column_index = 33  # column 'AH'
```

4. Instancie o objeto `ExcelFileUpdater`

```
updater = ExcelFileUpdater(
    _source_file,
    _source_sheet,
    _source_value_column_index,
    _source_content_column_index,
    _target_file,
    _target_sheet,
    _target_value_column_index,
    _target_content_column_index)
```
5. Chame o método/função

```
updater.update_target_file()
```