# Gerador de Relatórios Excel

## Descrição

O **Gerador de Relatórios Excel** é um script Python que lê duas planilhas Excel, uma contendo dados de clientes (`CMCL904-CLIENTE-CC.xlsx`) e outra com dados de fornecedores (`CMCL904-FORNECEDOR.xlsx`), e gera um novo arquivo Excel com relatórios consolidados. O script processa os dados, limpa valores monetários e de data, e cria várias abas com informações detalhadas e gráficos, incluindo:

- **EMISSOES**: Detalhes das emissões de passagens aéreas.
- **EMISSÃO E REEMISSAO**: Resumo de emissões e reemissões com totais e percentuais.
- **TOTAL POR EMPRESAS**: Total gasto por empresa com percentuais.
- **TOTAL POR CENTRO DE CUSTO**: Total por centro de custo com quantidade de bilhetes e percentuais.
- **TOTAL POR CIA AEREA**: Total por companhia aérea com gráficos de barras e pizza.
- **TOTAL POR CIA E TRECHO**: Total por companhia aérea e trecho.
- **TOTAL POR SOLICITANTE**: Total por solicitante com percentuais.
- **TOTAL CREDITOS DISPONIVEIS**: Informações sobre créditos disponíveis (atualmente apenas com cabeçalhos).

O script inclui uma interface gráfica simples criada com `tkinter` para facilitar a seleção dos arquivos de entrada e saída.

## Arquivos de Entrada

- **`CMCL904-CLIENTE-CC.xlsx`**: Contém dados de clientes, como Razão Social, CNPJ, Centro de Custo, Fornecedor, Tarifas, Taxas, Passageiro, Solicitante, Documento, Trecho, Emissão, Ida e Volta.
- **`CMCL904-FORNECEDOR.xlsx`**: Contém dados de fornecedores, como Fornecedor, Tarifas, Taxas e Total.

Dois arquivos de teste com dados anonimizados estão disponíveis na pasta `test_data`:
- `CMCL904-CLIENTE-CC_anonimizado.xlsx`
- `CMCL904-FORNECEDOR_anonimizado.xlsx`

## Pré-requisitos

O script foi desenvolvido e testado exclusivamente em sistemas operacionais **Windows**. Para executá-lo, é necessário instalar as dependências listadas abaixo.

### Dependências

As bibliotecas Python necessárias são:

- `pandas`
- `openpyxl`
- `tkinter` (geralmente incluído na instalação padrão do Python)

Para instalar as dependências, execute o seguinte comando no terminal (certifique-se de ter o Python e o `pip` instalados):

```bash
pip install pandas openpyxl
```

### Python

O script foi testado com Python 3.8 ou superior. Certifique-se de ter uma versão compatível instalada. Você pode baixar o Python em [python.org](https://www.python.org/downloads/).

## Como Usar

### Executando o Script Python

1. Clone ou baixe o repositório contendo o script `AppGeraRel.py` e os arquivos de teste.
2. Certifique-se de que as dependências estão instaladas (veja acima).
3. Execute o script com o comando:

```bash
python AppGeraRel.py
```

4. A interface gráfica será aberta. Siga as instruções abaixo:
   - Clique em "Selecionar" para escolher o arquivo de dados do cliente (`CMCL904-CLIENTE-CC.xlsx` ou o arquivo anonimizado).
   - Clique em "Selecionar" para escolher o arquivo de dados do fornecedor (`CMCL904-FORNECEDOR.xlsx` ou o arquivo anonimizado).
   - Clique em "Selecionar" para definir o local e o nome do arquivo de saída (extensão `.xlsx`).
   - Clique em "Gerar Relatório" para processar os dados e criar o arquivo Excel.
   - Clique em "Sair" para fechar a aplicação.

### Usando o Executável (Windows)

Um executável foi gerado para facilitar o uso em sistemas Windows, eliminando a necessidade de instalar o Python ou as dependências manualmente.

#### Gerando o Executável

Para gerar o executável a partir do script, siga estas etapas:

1. Instale o `PyInstaller`:

```bash
pip install pyinstaller
```

2. Navegue até o diretório onde está o arquivo `AppGeraRel.py` e execute:

```bash
pyinstaller --onefile AppGeraRel.py
```
ou
```bash
pyinstaller --onefile AppGeraRel.py
```

3. O executável será gerado na pasta `dist` com o nome `AppGeraRel.exe`.

#### Executando o Executável

1. Copie o executável `AppGeraRel.exe` para o mesmo diretório que contém os arquivos de entrada (ou especifique os caminhos completos na interface).
2. Execute o arquivo `AppGeraRel.exe` clicando duas vezes.
3. Siga os mesmos passos da interface gráfica descritos acima.

**Nota**: O executável foi testado apenas em sistemas Windows.

## Estrutura do Arquivo de Saída

O arquivo Excel gerado contém as seguintes abas:

- **EMISSOES**: Lista detalhada de emissões com informações como Razão Social, CNPJ, Fornecedor, Tarifas, Taxas, etc.
- **EMISSÃO E REEMISSAO**: Resumo de emissões e reemissões, incluindo totais e ticket médio.
- **TOTAL POR EMPRESAS**: Total gasto por empresa com percentuais.
- **TOTAL POR CENTRO DE CUSTO**: Quantidade de bilhetes e valores totais por centro de custo.
- **TOTAL POR CIA AEREA**: Total por companhia aérea, com gráfico de barras (tarifas por mês) e gráfico de pizza (percentual por companhia).
- **TOTAL POR CIA E TRECHO**: Total por companhia aérea e trecho, com quantidade de bilhetes e ticket médio.
- **TOTAL POR SOLICITANTE**: Total por solicitante com percentuais.
- **TOTAL CREDITOS DISPONIVEIS**: Estrutura para créditos disponíveis (atualmente apenas com cabeçalhos).

## Observações

- **Formato dos Arquivos de Entrada**: Os arquivos de entrada devem seguir o formato esperado, com as colunas especificadas no script. Use os arquivos anonimizados como referência.
- **Limpeza de Dados**: O script realiza a limpeza de valores monetários (ex.: remove "R$" e converte para float) e datas (ex.: padroniza para o ano atual se necessário).
- **Gráficos**: Os gráficos de barras e pizza na aba "TOTAL POR CIA AEREA" são gerados automaticamente com base nos dados processados.
- **Testes**: Os arquivos de teste anonimizados estão na pasta `test_data` e podem ser usados para verificar o funcionamento do script.

## Desenvolvimento

Este projeto foi criado de forma colaborativa com o auxílio de **Grok**, uma inteligência artificial desenvolvida pela xAI, que ajudou na construção e otimização do código.

## Limitações

- O script foi testado apenas em sistemas Windows.
- A aba "TOTAL CREDITOS DISPONIVEIS" está incompleta, contendo apenas os cabeçalhos.
- O script espera que as colunas dos arquivos de entrada sigam exatamente os nomes especificados. Qualquer discrepância pode causar erros.

## Contato

Para dúvidas ou sugestões, entre em contato com o desenvolvedor ou abra uma issue no repositório do projeto.