"""
Script automático de conferência de faturamento.

Roda de segunda a sexta pelo Agendador de Tarefas do Windows.

Regras:
- Segunda-feira: confere a sexta-feira anterior.
- Terça a sexta: confere o dia anterior.
- Não usa input(), pois precisa rodar sozinho.
- Gera Excel em uma pasta de saída.
- Gera log da execução.
"""

import os
import re
import logging
import datetime
from pathlib import Path

import pandas as pd
import pdfplumber


# =========================
# CONFIGURAÇÕES
# =========================

PASTA_BASE_FATURAMENTO = Path(
    r"G:\Drives compartilhados\TW_Faturamento\2026\FATURAMENTO - M30 SC"
)

PASTA_SAIDA = Path(__file__).parent
PASTA_LOGS = PASTA_SAIDA / "logs"

PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
PASTA_LOGS.mkdir(parents=True, exist_ok=True)


# =========================
# LOG
# =========================

data_log = datetime.datetime.now().strftime("%Y_%m_%d")
arquivo_log = PASTA_LOGS / f"conferencia_{data_log}.log"

logging.basicConfig(
    filename=arquivo_log,
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding="utf-8",
)

console = logging.StreamHandler()
console.setLevel(logging.INFO)
console.setFormatter(logging.Formatter("%(message)s"))
logging.getLogger("").addHandler(console)


# =========================
# PADRÕES
# =========================

REGEX_CHAVE_NFE = re.compile(r"(?<!\d)\d{44}(?!\d)")
MARCADORES_BOLETO = ("_bol_", "boleto", "_cobranca_", "_cobrança_")

PADRAO_CODIGO_PRODUTO = re.compile(
    r"^[A-Z0-9][A-Z0-9./_-]{2,29}$",
    re.IGNORECASE,
)

PADRAO_CNPJ_MASCARADO = re.compile(
    r"^[A-Z]?\d{2,3}\.\d{3}/\d{4}-\d{2}$",
    re.IGNORECASE,
)


TERMOS_PARA_IGNORAR = [
    "CÓDIGO",
    "DADOS",
    "IDENTIFICAÇÃO",
    "TW SISTEMAS",
    "RUA 14",
    "FLORIANOPOLIS",
    "CHAVE",
    "CONSULTA",
    "WWW.NFE",
    "NATUREZA",
    "VENDA DE",
    "INSCRIÇÃO",
    "BASE DE",
    "VALOR",
    "NOME",
    "AZUL LINHAS",
    "AVENIDA",
    "QUANTIDADE",
    "RIÇÃO",
    "BRUTO",
    "ESO LÍQUIDO",
    "CPF",
    "CNPJ",
    "MUNICÍPIO",
    "ENDEREÇO",
    "BAIRRO",
    "TRANSPORTADOR",
    "FRETE",
    "PLACA",
    "UF",
    "ESTADUAL",
    "IÇÃO",
    "ÃO ESTADUAL",
    "CRIÇÃO",
]


# =========================
# FUNÇÕES
# =========================

def calcular_data_alvo() -> datetime.datetime:
    """Calcula a data operacional que deve ser conferida na execução atual.

    A automação roda em dias úteis. Na segunda-feira, a rotina precisa voltar
    para sexta-feira; nos demais dias úteis, processa o faturamento do dia
    anterior.
    """
    hoje = datetime.datetime.now()

    if hoje.weekday() == 0:
        return hoje - datetime.timedelta(days=3)

    return hoje - datetime.timedelta(days=1)


def obter_nome_mes_ptbr(data: datetime.datetime) -> str:
    """Retorna o nome do mês em português usado na estrutura de pastas."""
    meses = {
        1: "JANEIRO",
        2: "FEVEREIRO",
        3: "MARÇO",
        4: "ABRIL",
        5: "MAIO",
        6: "JUNHO",
        7: "JULHO",
        8: "AGOSTO",
        9: "SETEMBRO",
        10: "OUTUBRO",
        11: "NOVEMBRO",
        12: "DEZEMBRO",
    }

    return meses[data.month]


def eh_boleto(nome_arquivo: str) -> bool:
    """Identifica se o nome do arquivo PDF representa boleto/cobrança."""
    nome = Path(nome_arquivo).stem.lower()
    return any(marcador in nome for marcador in MARCADORES_BOLETO)


def selecionar_pdf_nota_fiscal(arquivos_pdf: list[str]) -> str | None:
    """Seleciona o PDF mais provável de ser a Nota Fiscal dentro da pasta.

    Prioridade:
    1. Arquivo cujo nome contenha chave NFe de 44 dígitos.
    2. Primeiro PDF que não pareça boleto/cobrança.
    3. Primeiro PDF em ordem alfabética como fallback operacional.
    """
    if not arquivos_pdf:
        return None

    pdfs_ordenados = sorted(arquivos_pdf, key=str.lower)

    # Chave NFe no nome é o sinal mais confiável para diferenciar NF de anexos.
    candidatos_chave_nfe = [
        arq
        for arq in pdfs_ordenados
        if REGEX_CHAVE_NFE.search(Path(arq).stem)
    ]

    if candidatos_chave_nfe:
        return candidatos_chave_nfe[0]

    candidatos_nao_boleto = [
        arq for arq in pdfs_ordenados if not eh_boleto(arq)
    ]

    if candidatos_nao_boleto:
        return candidatos_nao_boleto[0]

    return pdfs_ordenados[0]


def extrair_chave_nfe(nome_arquivo: str) -> str | None:
    """Extrai a chave NFe de 44 dígitos presente no nome do arquivo."""
    base = Path(nome_arquivo).stem
    match = REGEX_CHAVE_NFE.search(base)
    return match.group(0) if match else None


def extrair_numero_nota_nome_arquivo(nome_arquivo: str) -> str | None:
    """Obtém o número da NF a partir do padrão do nome do arquivo.

    A função suporta nomes com `doc_<numero>` e nomes baseados na chave NFe.
    Quando a chave existe, o número da nota é extraído das posições oficiais
    da chave de acesso.
    """
    base = Path(nome_arquivo).stem

    match_doc = re.search(r"doc_(\d+)", base, re.IGNORECASE)
    if match_doc:
        return str(int(match_doc.group(1)))

    chave = extrair_chave_nfe(nome_arquivo)

    if chave:
        numero_chave = chave[25:34]

        try:
            return str(int(numero_chave))
        except ValueError:
            return numero_chave

    return None


def codigo_ou_descricao_deve_ignorar(codigo: str, descricao: str) -> bool:
    """Valida se código ou descrição pertencem a blocos fiscais irrelevantes."""
    codigo_upper = codigo.upper()
    descricao_upper = descricao.upper()

    if any(termo in codigo_upper for termo in TERMOS_PARA_IGNORAR):
        return True

    if any(termo in descricao_upper for termo in TERMOS_PARA_IGNORAR):
        return True

    return False


def codigo_produto_valido(codigo: str, descricao: str) -> bool:
    """Confirma se o par código/descrição possui formato de item de produto."""
    if not PADRAO_CODIGO_PRODUTO.match(codigo):
        return False

    if codigo.isdigit():
        return False

    if PADRAO_CNPJ_MASCARADO.match(codigo):
        return False

    if not descricao or len(descricao) < 3:
        return False

    return True


def extrair_dados_pdf(caminho_pdf: Path) -> list[dict]:
    """Extrai itens de produto das tabelas do PDF da Nota Fiscal.

    A extração trabalha com tabelas detectadas pelo `pdfplumber` e considera
    as duas primeiras colunas limpas como código e descrição. Os filtros
    removem cabeçalhos, textos fiscais, CNPJs mascarados, códigos numéricos
    puros e duplicidades dentro do mesmo PDF.
    """
    itens = []
    itens_unicos = set()

    try:
        with pdfplumber.open(str(caminho_pdf)) as pdf:
            for page in pdf.pages:
                tabelas = page.extract_tables()

                for tabela in tabelas:
                    for linha in tabela:
                        if not linha:
                            continue

                        linha_limpa = [
                            str(item).strip()
                            for item in linha
                            if item is not None and str(item).strip() != ""
                        ]

                        if len(linha_limpa) < 2:
                            continue

                        # Layouts de NF costumam posicionar código e descrição nas primeiras colunas.
                        codigo = linha_limpa[0]
                        descricao = linha_limpa[1]

                        if codigo_ou_descricao_deve_ignorar(codigo, descricao):
                            continue

                        if not codigo_produto_valido(codigo, descricao):
                            continue

                        chave_item = (codigo, descricao)

                        # Evita repetir o mesmo item quando o parser encontra a tabela mais de uma vez.
                        if chave_item in itens_unicos:
                            continue

                        itens_unicos.add(chave_item)

                        logging.info(f"      Item encontrado: {codigo} | {descricao}")

                        itens.append(
                            {
                                "codigo_produto": codigo,
                                "descricao_produto": descricao,
                            }
                        )

    except Exception as erro:
        logging.error(
            f"Erro ao ler o PDF {caminho_pdf.name}: {erro}"
        )

    return itens


def montar_caminho_dia(data_alvo: datetime.datetime) -> Path:
    """Monta o caminho da pasta diária de faturamento conforme padrão M30."""
    mes_str = f"{data_alvo.strftime('%m')} - {obter_nome_mes_ptbr(data_alvo)}"
    dia_str = data_alvo.strftime("%d")

    return PASTA_BASE_FATURAMENTO / mes_str / dia_str


def salvar_excel(dados_finais: list[dict], data_alvo: datetime.datetime) -> Path | None:
    """Gera a planilha final de conferência e retorna o caminho criado.

    Caso o arquivo padrão esteja aberto no Excel, cria uma versão alternativa
    com timestamp para preservar a execução automática sem intervenção manual.
    """
    if not dados_finais:
        logging.warning("Processo concluído, mas nenhum código válido foi extraído.")
        return None

    df = pd.DataFrame(dados_finais)

    dia = data_alvo.strftime("%d")
    mes = data_alvo.strftime("%m")
    ano = data_alvo.strftime("%Y")

    nome_excel = f"Conferencia_{dia}_{mes}_{ano}.xlsx"
    caminho_excel = PASTA_SAIDA / nome_excel

    try:
        df.to_excel(caminho_excel, index=False)
        logging.info(f"Planilha gerada com sucesso: {caminho_excel}")
        return caminho_excel

    except PermissionError:
        timestamp = datetime.datetime.now().strftime("%H%M%S")
        nome_alternativo = f"Conferencia_{dia}_{mes}_{ano}_{timestamp}.xlsx"
        caminho_alternativo = PASTA_SAIDA / nome_alternativo

        df.to_excel(caminho_alternativo, index=False)

        logging.warning(
            f"A planilha original estava aberta. Arquivo salvo como: {caminho_alternativo}"
        )

        return caminho_alternativo


def processar_conferencia() -> None:
    """Orquestra o processo completo de conferência automática.

    Fluxo principal:
    - calcula a data operacional;
    - localiza a pasta diária no drive de faturamento;
    - percorre as pastas de clientes;
    - seleciona o PDF de NF;
    - extrai itens válidos;
    - exporta a planilha consolidada.
    """
    logging.info("=" * 80)
    logging.info("Iniciando conferência automática de faturamento")

    data_alvo = calcular_data_alvo()

    ano_str = data_alvo.strftime("%Y")
    mes_str = data_alvo.strftime("%m")
    dia_str = data_alvo.strftime("%d")

    data_faturamento = f"{dia_str}/{mes_str}/{ano_str}"
    caminho_completo_dia = montar_caminho_dia(data_alvo)

    logging.info(f"Data de conferência: {data_faturamento}")
    logging.info(f"Pasta acessada: {caminho_completo_dia}")

    if not caminho_completo_dia.exists():
        logging.error(f"Pasta não encontrada: {caminho_completo_dia}")
        return

    try:
        pastas_clientes = [
            item
            for item in caminho_completo_dia.iterdir()
            if item.is_dir()
        ]

    except Exception as erro:
        logging.error(f"Erro ao listar pastas de clientes: {erro}")
        return

    logging.info(f"Pastas de clientes encontradas: {len(pastas_clientes)}")

    dados_finais = []

    for caminho_cliente in pastas_clientes:
        cliente = caminho_cliente.name

        logging.info("-" * 60)
        logging.info(f"Cliente: {cliente}")

        try:
            arquivos_cliente = list(caminho_cliente.iterdir())

            pdfs = [
                arquivo.name
                for arquivo in arquivos_cliente
                if arquivo.is_file() and arquivo.suffix.lower() == ".pdf"
            ]

            if not pdfs:
                logging.warning("Nenhuma Nota Fiscal PDF encontrada nesta pasta.")
                continue

            nome_pdf = selecionar_pdf_nota_fiscal(pdfs)

            if not nome_pdf:
                logging.warning("Nenhum PDF válido para Nota Fiscal foi encontrado.")
                continue

            caminho_pdf = caminho_cliente / nome_pdf

            logging.info(f"Lendo Nota Fiscal: {nome_pdf}")

            numero_nota = extrair_numero_nota_nome_arquivo(nome_pdf)
            itens_nota = extrair_dados_pdf(caminho_pdf)

            if not itens_nota:
                logging.warning("Nenhum item de produto válido encontrado nesta NF.")
                continue

            for item in itens_nota:
                dados_finais.append(
                    {
                        "Data Faturamento": data_faturamento,
                        "Cliente": cliente,
                        "Número Nota": numero_nota,
                        "Arquivo NF": nome_pdf,
                        "Código Produto": item["codigo_produto"],
                        "Descrição Produto": item["descricao_produto"],
                    }
                )

        except Exception as erro:
            logging.error(f"Erro ao processar cliente {cliente}: {erro}")

    caminho_excel = salvar_excel(dados_finais, data_alvo)

    if caminho_excel:
        logging.info("Processo finalizado com sucesso.")
    else:
        logging.warning("Processo finalizado sem geração de Excel.")

    logging.info("=" * 80)


# =========================
# EXECUÇÃO
# =========================

if __name__ == "__main__":
    processar_conferencia()
