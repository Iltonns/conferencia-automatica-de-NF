﻿# Conferência Automática de NF

Automação para conferência de faturamento com leitura de notas fiscais em PDF e extração de códigos de produto, gerando planilha Excel consolidada por dia.

## Objetivo

Processar as pastas de faturamento de um dia específico, identificar códigos de produto nas NFs e entregar um arquivo `.xlsx` pronto para conferência operacional.

## Escopo atual

- Calcula a data alvo automaticamente:
  - segunda-feira: usa sexta anterior
  - demais dias: usa ontem
- Navega na estrutura de faturamento por mês/dia
- Percorre pastas de clientes
- Lê o primeiro PDF encontrado em cada pasta de cliente
- Extrai códigos de produto com filtros de ruído fiscal
- Gera planilha Excel consolidada

## Arquitetura atual

Projeto minimalista com um único entrypoint:

- `script.py`

Funções principais:

- `calcular_data_alvo()`: define dia de processamento
- `extrair_dados_pdf(caminho_pdf)`: parse de tabelas PDF e filtragem de códigos
- `main()`: orquestra leitura de diretórios, extração e exportação

## Requisitos

- Python 3.10+
- Dependências Python:
  - `pandas`
  - `pdfplumber`
  - `openpyxl` (engine de escrita `.xlsx` do pandas)

## Instalação

```bash
python -m venv .venv
.venv\Scripts\activate
pip install pandas pdfplumber openpyxl
```

## Configuração

No estado atual, o caminho raiz do faturamento está hardcoded em `script.py`:

```python
PASTA_BASE_FATURAMENTO = r"G:\Drives compartilhados\TW_Faturamento\2026\FATURAMENTO - M30 SC"
```

Ajuste para o ambiente da operação antes de executar.

### Estrutura esperada de diretórios

O script espera o padrão abaixo:

```text
<PASTA_BASE_FATURAMENTO>/
  <MM - NOME_DO_MES>/
    <DD>/
      <CLIENTE_A>/
        nota1.pdf
      <CLIENTE_B>/
        nota1.pdf
```

## Execução

```bash
python script.py
```

Saída esperada:

- Arquivo `Conferencia_DD_MM_AAAA.xlsx` no diretório atual
- Em caso de arquivo aberto, fallback para `Conferencia_DD_MM_AAAA_REPETIDO.xlsx`

## Regras de extração de código

A extração aplica validações para reduzir falsos positivos:

- usa primeira coluna de cada linha extraída das tabelas PDF
- ignora cabeçalhos e textos fiscais por lista de termos bloqueados
- aceita apenas valores com 4 a 18 caracteres
- descarta valores apenas numéricos

## Limitações conhecidas

- Processa somente o primeiro PDF de cada cliente
- Lista de termos ignorados é heurística e depende do layout das NFs
- Não há parametrização por variável de ambiente/CLI
- Não há suíte de testes automatizados

## Segurança e boas práticas

- `credentials.json`, `.env`, `.venv`, `__pycache__` e `*.pyc` já estão no `.gitignore`
- Evite salvar dados sensíveis em logs
- Garanta permissão mínima de leitura no drive de faturamento

## Próximos passos recomendados

- Externalizar `PASTA_BASE_FATURAMENTO` para variável de ambiente
- Processar múltiplos PDFs por cliente
- Adicionar logging estruturado
- Criar testes de regressão para parse de PDFs
- Criar `requirements.txt` e pipeline de execução