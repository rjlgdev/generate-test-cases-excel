# Gerador de Planilha de Casos de Teste a partir de Arquivos Gherkin

Este script Python (`generate_test_case_excel.py`) automatiza a criação de uma planilha Excel formatada de Casos de Teste, extraindo as informações diretamente de um arquivo de *feature* no formato Gherkin.

## Funcionalidades

*   **Parsing Gherkin:** Analisa cenários (`Scenario`) e passos (`Given`, `When`, `Then`, `And`, `But`).
*   **Extração de Metadados:** Extrai links de evidência e status de resultado a partir de comentários específicos no arquivo `.feature` (ex: `#Evidência:` e `#Resultado:`).
*   **Formatação Profissional:** Gera uma planilha Excel (`.xlsx`) com cabeçalhos, formatação de cores (para o status) e ajuste automático de colunas.
*   **Localização (Português):** Todos os cabeçalhos e as palavras-chave Gherkin são traduzidos para o português (Dado, Quando, Então, E, Mas).
*   **Separação Lógica:** Separa os passos em "Pré-Condição", "Passos de Teste" e "Resultado Esperado" de forma inteligente, baseada na ordem das palavras-chave Gherkin.

## Pré-requisitos

Para executar este script, você precisa ter o **Python 3** instalado e a biblioteca `openpyxl`.

### Instalação da Biblioteca

Abra seu terminal ou prompt de comando e execute:

```bash
pip install openpyxl
```

## Como Usar

### 1. Preparar o Arquivo `.feature`

Certifique-se de que seu arquivo Gherkin (`.feature`) contenha os metadados de Evidência e Resultado como comentários, logo acima do `Scenario` correspondente.

**Exemplo de Formato:**

```gherkin
Feature: Nome da Feature

  #Evidência: https://link.para.a.evidencia/
  #Resultado: SUCESSO✅
  Scenario: 01) Título do Caso de Teste
    Given que o usuário está logado
    When ele clica no botão "X"
    And preenche o formulário
    Then deve ser exibida a mensagem de sucesso
```

### 2. Executar o Script

Execute o script `generate_test_case_excel.py` no seu terminal, passando o caminho completo ou relativo para o seu arquivo `.feature` como argumento.

```bash
python generate_test_case_excel.py /caminho/para/seu/arquivo.feature
```

**Alternativas para Windows:**

Se o comando `python` não funcionar, tente usar `python3` ou `py`:

```bash
python3 generate_test_case_excel.py /caminho/para/seu/arquivo.feature
# OU
py generate_test_case_excel.py /caminho/para/seu/arquivo.feature
```

### 3. Resultado

O script irá gerar um arquivo Excel no mesmo diretório de execução, com o nome baseado na sua Feature (ex: `Nome_da_Feature_Casos_de_Teste_v3.xlsx`).

A planilha gerada terá as seguintes colunas em português:

| ID do Caso de Teste | Caso de Teste | Pré-Condição | Passos de Teste | Resultado Esperado | Evidência | Status |
| :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| TC_01 | Título do Caso... | Dado que... | 1. Quando... | 1. Então... | Link | SUCESSO✅ |

---

*Criado por Rafael La Guardia*
