# Conferência Automática de NF

Automação operacional para conferência de faturamento com leitura de Notas Fiscais em PDF, extração de códigos de produto e geração de planilha Excel consolidada por data de faturamento.

O projeto foi preparado para execução manual e execução automática pelo **Agendador de Tarefas do Windows** usando o arquivo `Executar_Conferencia.bat`.

## Objetivo

Processar a pasta de faturamento da data operacional correta, identificar itens de produto nas NFs dos clientes e gerar uma planilha `.xlsx` pronta para conferência.

## Escopo atual

- Calcula automaticamente a data alvo:
  - segunda-feira: processa a sexta-feira anterior;
  - terça a sexta-feira: processa o dia anterior.
- Navega na estrutura de faturamento por mês e dia.
- Percorre as pastas de clientes.
- Seleciona o PDF mais provável de ser a Nota Fiscal.
- Ignora boletos/cobranças quando existir PDF de NF mais confiável.
- Extrai códigos e descrições de produto das tabelas do PDF.
- Remove ruídos fiscais, cabeçalhos, CNPJs mascarados e duplicidades.
- Gera planilha Excel consolidada.
- Gera logs por execução.

## Arquitetura

Projeto minimalista com dois entrypoints operacionais:

- `script.py`: rotina principal de conferência.
- `Executar_Conferencia.bat`: wrapper para execução pelo Windows.

Funções principais em `script.py`:

- `calcular_data_alvo()`: define a data operacional da conferência.
- `montar_caminho_dia(data_alvo)`: monta a pasta diária do faturamento.
- `selecionar_pdf_nota_fiscal(arquivos_pdf)`: escolhe o PDF mais provável de NF.
- `extrair_numero_nota_nome_arquivo(nome_arquivo)`: identifica o número da nota pelo nome do arquivo.
- `extrair_dados_pdf(caminho_pdf)`: extrai itens válidos do PDF.
- `salvar_excel(dados_finais, data_alvo)`: exporta a planilha consolidada.
- `processar_conferencia()`: orquestra o fluxo completo.

## Requisitos

- Windows com acesso ao drive de faturamento.
- Python 3.10+.
- Dependências Python:
  - `pandas`
  - `pdfplumber`
  - `openpyxl`

## Instalação

```bash
python -m venv .venv
.venv\Scripts\activate
pip install pandas pdfplumber openpyxl
```

> Se o Agendador usar um Python global, instale as dependências nesse mesmo interpretador ou altere o `.bat` para apontar para o Python da `.venv`.

## Configuração

O caminho raiz do faturamento está definido em `script.py`:

```python
PASTA_BASE_FATURAMENTO = Path(
    r"G:\Drives compartilhados\TW_Faturamento\2026\FATURAMENTO - M30 SC"
)
```

Ajuste esse valor conforme o ambiente da operação.

### Estrutura esperada de diretórios

```text
<PASTA_BASE_FATURAMENTO>/
  <MM - NOME_DO_MES>/
    <DD>/
      <CLIENTE_A>/
        nota_fiscal.pdf
      <CLIENTE_B>/
        nota_fiscal.pdf
```

Exemplo:

```text
G:\Drives compartilhados\TW_Faturamento\2026\FATURAMENTO - M30 SC\
  06 - JUNHO\
    10\
      CLIENTE EXEMPLO\
        35260600000000000000550010000012341000012345.pdf
```

## Execução manual

Execute diretamente pelo Python:

```bash
python script.py
```

Ou execute pelo BAT:

```bash
Executar_Conferencia.bat
```

Saídas esperadas:

- `Conferencia_DD_MM_AAAA.xlsx` no diretório do projeto.
- `Conferencia_DD_MM_AAAA_HHMMSS.xlsx` quando a planilha padrão estiver aberta.
- `logs/conferencia_AAAA_MM_DD.log` com log estruturado da execução.
- `conferencia.log` com stdout/stderr redirecionado pelo `.bat`.

## Automação pelo Agendador de Tarefas do Windows

### 1. Validar o BAT

Antes de agendar, execute manualmente:

```bash
Executar_Conferencia.bat
```

Confirme se foram gerados:

- arquivo `.xlsx`;
- log diário em `logs/`;
- registro em `conferencia.log`.

### 2. Criar tarefa

No Windows:

1. Abra **Agendador de Tarefas**.
2. Clique em **Criar Tarefa**.
3. Aba **Geral**:
   - Nome: `Conferência Automática de NF`
   - Marque: **Executar somente quando o usuário estiver conectado** se o drive `G:` depender da sessão do usuário.
   - Marque: **Executar com privilégios mais altos**.
4. Aba **Disparadores**:
   - Novo disparador: diário ou semanal.
   - Frequência recomendada: segunda a sexta-feira.
   - Horário recomendado: após a conclusão do faturamento e sincronização do drive.
5. Aba **Ações**:
   - Ação: **Iniciar um programa**.
   - Programa/script:

```text
C:\Users\Dell\OneDrive\Documentos\GitHub\conferencia-automatica-de-NF\Executar_Conferencia.bat
```

   - Iniciar em:

```text
C:\Users\Dell\OneDrive\Documentos\GitHub\conferencia-automatica-de-NF
```

6. Aba **Condições**:
   - Desmarque restrições de energia se o processo precisar rodar em notebook fora da tomada.
   - Marque opção para ativar o computador se necessário.
7. Aba **Configurações**:
   - Permitir execução sob demanda.
   - Interromper se executar por tempo excessivo, conforme política operacional.
   - Reexecutar em caso de falha se houver instabilidade no drive.

### 3. Atenção ao drive mapeado

O caminho atual usa `G:\...`. Drives mapeados podem não existir quando a tarefa roda sem sessão interativa.

Para produção, use uma destas abordagens:

- manter a tarefa como **Executar somente quando o usuário estiver conectado**;
- substituir `G:\...` por caminho UNC/rede equivalente, se disponível;
- garantir que o agente de sincronização do drive esteja ativo antes do horário da tarefa.

## Regras de seleção do PDF

A seleção da Nota Fiscal segue a prioridade abaixo:

1. PDF cujo nome contenha chave NFe de 44 dígitos.
2. PDF que não contenha marcadores de boleto/cobrança.
3. Primeiro PDF em ordem alfabética como fallback.

Marcadores ignorados:

- `_bol_`
- `boleto`
- `_cobranca_`
- `_cobrança_`

## Regras de extração de código

A extração aplica filtros para reduzir falsos positivos:

- usa as duas primeiras colunas limpas das tabelas do PDF como código e descrição;
- ignora cabeçalhos e textos fiscais por lista de termos bloqueados;
- aceita códigos alfanuméricos com formato compatível com produto;
- descarta valores apenas numéricos;
- descarta CNPJs mascarados;
- exige descrição preenchida;
- remove duplicidades dentro da mesma NF.

## Logs e auditoria

O script registra eventos relevantes:

- início e fim da execução;
- data de faturamento processada;
- pasta acessada;
- quantidade de clientes encontrados;
- NF lida por cliente;
- itens extraídos;
- erros de leitura de PDF;
- falha por pasta inexistente;
- caminho da planilha gerada.

Arquivos de log:

```text
logs/conferencia_AAAA_MM_DD.log
conferencia.log
```

## Manutenção

O código foi documentado com docstrings nas funções principais e comentários nos pontos de decisão operacional:

- cálculo de data alvo;
- seleção de PDF de NF;
- extração e limpeza de linhas do PDF;
- deduplicação de itens;
- fallback de exportação quando o Excel está aberto.

## Limitações conhecidas

- Processa um PDF de NF por cliente.
- A lista de termos ignorados é heurística e depende do layout das NFs.
- O caminho do faturamento ainda está fixo no código.
- Não há parametrização por CLI ou variável de ambiente.
- Não há suíte de testes automatizados.

## Segurança e boas práticas

- `credentials.json`, `.env`, `.venv`, `__pycache__` e `*.pyc` estão no `.gitignore`.
- Não registrar dados fiscais sensíveis além do necessário para auditoria operacional.
- Usar permissão mínima de leitura no drive de faturamento.
- Validar periodicamente se o Agendador está executando com o mesmo usuário que possui acesso ao drive.

## Próximos passos recomendados

- Externalizar `PASTA_BASE_FATURAMENTO` para variável de ambiente.
- Permitir configuração de data alvo por argumento CLI.
- Processar múltiplos PDFs por cliente quando necessário.
- Criar `requirements.txt`.
- Criar testes de regressão com PDFs anonimizados.
- Centralizar rotação/retenção de logs.
