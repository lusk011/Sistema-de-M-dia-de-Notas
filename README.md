# Sistema de Média de Notas

Aplicação web feita em Python com Streamlit para cadastrar alunos, registrar notas, calcular médias e gerar uma planilha Excel automaticamente com os resultados da turma.

## Funcionalidades

- Cadastro da turma
- Cadastro de alunos
- Registro de três notas por aluno
- Cálculo automático da média
- Classificação automática da situação do aluno:
  - Aprovado
  - Recuperação
  - Reprovado
- Correção dos dados antes de gerar a planilha
- Geração de planilha Excel com os dados da turma
- Resumo final com total de alunos, aprovados, recuperação, reprovados e média geral

## Tecnologias usadas

- Python
- Streamlit
- Pandas
- OpenPyXL

## Como executar o projeto

Clone o repositório:

```bash
git clone https://github.com/lusk011/Sistema-de-M-dia-de-Notas.git
```

Entre na pasta do projeto:

```bash
cd Sistema-de-M-dia-de-Notas
```

Instale as dependências:

```bash
pip install streamlit pandas openpyxl
```

Execute o Streamlit:

```bash
streamlit run streamlit_app.py
```

Se o arquivo principal tiver outro nome, como `verificadormédia.py`, execute:

```bash
streamlit run verificadormédia.py
```

## Geração da planilha

Após cadastrar todos os alunos, o sistema mostra uma tabela de resultados. Antes de baixar a planilha, é possível corrigir nomes e notas diretamente na interface.

Depois das correções, basta clicar em **Gerar e baixar planilha Excel**.

A planilha gerada contém:

- Nome dos alunos
- Notas
- Média
- Situação
- Observação
- Resumo da turma

## Estrutura básica do projeto

```text
Sistema-de-M-dia-de-Notas/
├── verificadormédia.py
└── README.md
```

## Objetivo

Este projeto foi criado para facilitar o registro de notas escolares e automatizar a criação de uma planilha final com o desempenho da turma.