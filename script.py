"""Script de conferência de faturamento com extração de códigos de produtos de PDFs.

Resumo para iniciantes:
1. Descobre qual dia deve ser conferido (ontem, ou sexta se hoje for segunda).
2. Entra na pasta desse dia e percorre cada cliente.
3. Lê a primeira nota fiscal PDF de cada cliente.
4. Extrai códigos de produtos válidos e monta uma tabela.
5. Exporta o resultado em Excel.
"""

import os
import datetime
import pandas as pd
import pdfplumber


def extrair_dados_pdf(caminho_pdf):
    """Extrai códigos de produto a partir de tabelas de um PDF de nota fiscal.

    Args:
        caminho_pdf (str): Caminho completo do arquivo PDF da nota fiscal.

    Returns:
        list[str]: Lista com os códigos de produto válidos encontrados.

    Como funciona:
    - O script percorre todas as páginas do PDF.
    - Em cada página, tenta extrair tabelas.
    - Para cada linha da tabela, pega o primeiro valor como possível código.
    - Aplica filtros para remover cabeçalhos e textos fiscais.
    """
    codigos = []

    # Palavras comuns em cabeçalhos/campos fiscais que NÃO são código de produto.
    # Se o texto da coluna 1 tiver qualquer um desses termos, a linha é ignorada.
    termos_para_ignorar = [
        "CÓDIGO", "DADOS", "IDENTIFICAÇÃO", "TW SISTEMAS", "RUA 14", "FLORIANOPOLIS",
        "CHAVE", "CONSULTA", "WWW.NFE", "NATUREZA", "VENDA DE", "INSCRIÇÃO", "BASE DE",
        "VALOR", "NOME", "AZUL LINHAS", "AVENIDA", "QUANTIDADE", "RIÇÃO", "BRUTO", "ESO LÍQUIDO",
        "CPF", "CNPJ", "MUNICÍPIO", "ENDEREÇO", "BAIRRO", "TRANSPORTADOR", "FRETE", "PLACA", "UF",
        "ESTADUAL", "IÇÃO", "ÃO ESTADUAL", "CRIÇÃO"
    ]

    try:
        # Abre o PDF com gerenciador de contexto.
        # Isso garante que o arquivo será fechado automaticamente ao final.
        with pdfplumber.open(caminho_pdf) as pdf:
            for page in pdf.pages:
                tabelas = page.extract_tables()

                for tabela in tabelas:
                    for linha in tabela:
                        if not linha:
                            continue

                        # Remove valores nulos e espaços sobrando em cada coluna da linha.
                        linha_limpa = [
                            str(item).strip()
                            for item in linha
                            if item is not None and str(item).strip() != ""
                        ]

                        if not linha_limpa:
                            continue

                        # Regra principal: a primeira coluna da linha é tratada como código.
                        codigo = linha_limpa[0]

                        # Filtro 1: remove cabeçalhos/textos fiscais conhecidos.
                        devera_ignorar = any(
                            termo in codigo.upper() for termo in termos_para_ignorar
                        )
                        if devera_ignorar:
                            continue

                        # Filtro 2: valida formato do código.
                        # - entre 4 e 18 caracteres
                        # - não pode ser só números
                        # Isso elimina textos longos e campos fiscais.
                        if 4 <= len(codigo) <= 18 and not codigo.isdigit():
                            print(f"      🔹 Código de produto encontrado: {codigo}")
                            codigos.append(codigo)
    except Exception as e:
        # Se houver erro de leitura do PDF, apenas avisa e segue o fluxo.
        print(f"    ❌ Erro ao ler o PDF {os.path.basename(caminho_pdf)}: {e}")

    return codigos


def calcular_data_alvo():
    """Calcula a data de faturamento a ser processada.

    Returns:
        datetime.datetime: Data alvo para buscar as pastas.

    Regra de negócio:
    - Segunda-feira: usa a sexta anterior (3 dias atrás).
    - Outros dias: usa o dia anterior.
    """
    hoje = datetime.datetime.now()

    # weekday() retorna: segunda=0, terça=1, ..., domingo=6
    if hoje.weekday() == 0:
        data_alvo = hoje - datetime.timedelta(days=3)
    else:
        data_alvo = hoje - datetime.timedelta(days=1)

    return data_alvo


def main():
    """Função principal que orquestra todo o processo de conferência.

    Fluxo:
    1. Calcula a data alvo.
    2. Monta o caminho da pasta do dia.
    3. Lista clientes.
    4. Extrai códigos dos PDFs.
    5. Salva tudo em Excel.
    """
    PASTA_BASE_FATURAMENTO = r"G:\Drives compartilhados\TW_Faturamento\2026\FATURAMENTO - M30 SC"

    data_alvo = calcular_data_alvo()
    ano_str = data_alvo.strftime("%Y")

    # Gera nome do mês no padrão usado pelas pastas, por exemplo: "05 - MAIO".
    nome_mes_extenso = pd.to_datetime(data_alvo).month_name(locale="pt_BR").upper()
    mes_str = f"{data_alvo.strftime('%m')} - {nome_mes_extenso}"
    dia_str = data_alvo.strftime("%d")

    caminho_completo_dia = os.path.join(PASTA_BASE_FATURAMENTO, mes_str, dia_str)

    print(f"\n📆 Data de conferência: {dia_str}/{data_alvo.strftime('%m')}/{ano_str}")
    print(f"📂 Acessando a pasta: {caminho_completo_dia}")

    # Caso não encontre a pasta automática, oferece entrada manual do dia.
    if not os.path.exists(caminho_completo_dia):
        print(f"\n⚠ Não encontrei a pasta automática no caminho: {caminho_completo_dia}")
        resposta = input("Deseja digitar o dia manualmente? (Ex: 22) ou aperte Enter para sair: ").strip()
        if resposta:
            dia_str = resposta.zfill(2)
            caminho_completo_dia = os.path.join(PASTA_BASE_FATURAMENTO, mes_str, dia_str)
            if not os.path.exists(caminho_completo_dia):
                print(f"❌ Erro: O caminho manual '{caminho_completo_dia}' também não existe.")
                return
        else:
            return

    try:
        itens = os.listdir(caminho_completo_dia)
        pastas_clientes = [
            i for i in itens if os.path.isdir(os.path.join(caminho_completo_dia, i))
        ]
    except Exception as e:
        print(f"❌ Erro ao listar pastas: {e}")
        return

    print(f"📁 Encontradas {len(pastas_clientes)} pastas de clientes.\n")

    # Estrutura em memória onde cada item será uma linha da planilha final.
    dados_finais = []

    for cliente in pastas_clientes:
        caminho_cliente = os.path.join(caminho_completo_dia, cliente)
        print(f"💼 Cliente: {cliente}")

        try:
            arquivos_cliente = os.listdir(caminho_cliente)
            pdfs = [arq for arq in arquivos_cliente if arq.lower().endswith(".pdf")]

            if pdfs:
                # Usa o primeiro PDF encontrado na pasta do cliente.
                nome_pdf = pdfs[0]
                caminho_completo_pdf = os.path.join(caminho_cliente, nome_pdf)
                print(f"    📄 Lendo Nota Fiscal: {nome_pdf}")

                codigos = extrair_dados_pdf(caminho_completo_pdf)

                # Cada código vira uma linha da planilha.
                for cod in codigos:
                    dados_finais.append(
                        {
                            "Data Faturamento": f"{dia_str}/{data_alvo.strftime('%m')}/{ano_str}",
                            "Cliente": cliente,
                            "Arquivo NF": nome_pdf,
                            "Código Produto": cod,
                        }
                    )
            else:
                print("    ⚠ Nenhuma Nota Fiscal (.pdf) encontrada nesta pasta.")
        except Exception as e:
            print(f"    ❌ Erro ao acessar a pasta deste cliente: {e}")

        print("-" * 60)

    # Salva resultado no Excel.
    if dados_finais:
        df = pd.DataFrame(dados_finais)
        nome_excel = f"Conferencia_{dia_str}_{data_alvo.strftime('%m')}_{ano_str}.xlsx"

        try:
            df.to_excel(nome_excel, index=False)
            print(f"\n🚀 SUCESSO COMPLETO! Planilha '{nome_excel}' gerada perfeitamente e sem sujeiras!")
        except PermissionError:
            # Se o arquivo estiver aberto no Excel, salva com nome alternativo.
            print(f"\n⚠ A planilha '{nome_excel}' está aberta no Excel!")
            nome_alternativo = f"Conferencia_{dia_str}_{data_alvo.strftime('%m')}_{ano_str}_REPETIDO.xlsx"
            df.to_excel(nome_alternativo, index=False)
            print(f"📝 Para não perder o trabalho, salvei como: '{nome_alternativo}'")
    else:
        print("\n⚠ Processo concluído, mas nenhum código válido foi extraído.")


if __name__ == "__main__":
    main()
