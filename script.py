import os
import datetime
import pandas as pd
import pdfplumber

def extrair_dados_pdf(caminho_pdf):
    """Lê o PDF da Nota Fiscal local e retorna APENAS os códigos dos produtos reais."""
    codigos = []
    
    # Palavras-chave completas ou parciais de campos fiscais para bloquear de vez
    termos_para_ignorar = [
        "CÓDIGO", "DADOS", "IDENTIFICAÇÃO", "TW SISTEMAS", "RUA 14", "FLORIANOPOLIS",
        "CHAVE", "CONSULTA", "WWW.NFE", "NATUREZA", "VENDA DE", "INSCRIÇÃO", "BASE DE",
        "VALOR", "NOME", "AZUL LINHAS", "AVENIDA", "QUANTIDADE", "RIÇÃO", "BRUTO", "ESO LÍQUIDO",
        "CPF", "CNPJ", "MUNICÍPIO", "ENDEREÇO", "BAIRRO", "TRANSPORTADOR", "FRETE", "PLACA", "UF",
        "ESTADUAL", "IÇÃO", "ÃO ESTADUAL", "CRIÇÃO"
    ]
    
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            for page in pdf.pages:
                tabelas = page.extract_tables()
                for tabela in tabelas:
                    for linha in tabela:
                        if linha:
                            # Limpa os espaços em branco de cada coluna
                            linha_limpa = [str(item).strip() for item in linha if item is not None and str(item).strip() != '']
                            
                            if linha_limpa:
                                codigo = linha_limpa[0]
                                
                                # 1. Ignora se o código contiver qualquer um dos termos proibidos
                                devera_ignorar = any(termo in codigo.upper() for termo in termos_para_ignorar)
                                if devera_ignorar:
                                    continue
                                
                                # 2. Trava de tamanho máximo e mínimo para códigos de autopeças
                                if 4 <= len(codigo) <= 18 and not codigo.isdigit():
                                    print(f"      🔹 Código de produto encontrado: {codigo}")
                                    codigos.append(codigo)
    except Exception as e:
        print(f"    ❌ Erro ao ler o PDF {os.path.basename(caminho_pdf)}: {e}")
    return codigos

def calcular_data_alvo():
    """Calcula a data correta. Se hoje for Segunda, busca a Sexta anterior (3 dias atrás)."""
    hoje = datetime.datetime.now()
    if hoje.weekday() == 0:  # 0 = Segunda-feira
        data_alvo = hoje - datetime.timedelta(days=3)
    else:
        data_alvo = hoje - datetime.timedelta(days=1)
    return data_alvo

def main():
    PASTA_BASE_FATURAMENTO = r"G:\Drives compartilhados\TW_Faturamento\2026\FATURAMENTO - M30 SC"
    
    data_alvo = calcular_data_alvo()
    ano_str = data_alvo.strftime('%Y')
    
    nome_mes_extenso = pd.to_datetime(data_alvo).month_name(locale='pt_BR').upper()
    mes_str = f"{data_alvo.strftime('%m')} - {nome_mes_extenso}" 
    dia_str = data_alvo.strftime('%d')
    
    caminho_completo_dia = os.path.join(PASTA_BASE_FATURAMENTO, mes_str, dia_str)
    
    print(f"\n📅 Data de conferência: {dia_str}/{data_alvo.strftime('%m')}/{ano_str}")
    print(f"📂 Acessando a pasta: {caminho_completo_dia}")
    
    # IMPORTANTE: Para automações de background (n8n/SSH), se a pasta automática não existir,
    # o script encerra silenciosamente em vez de travar esperando digitação no terminal vazio.
    if not os.path.exists(caminho_completo_dia):
        print(f"\n❌ Erro: Pasta automática não encontrada no caminho: {caminho_completo_dia}")
        return

    try:
        itens = os.listdir(caminho_completo_dia)
        pastas_clientes = [i for i in itens if os.path.isdir(os.path.join(caminho_completo_dia, i))]
    except Exception as e:
        print(f"❌ Erro ao listar pastas: {e}")
        return

    print(f"📁 Encontradas {len(pastas_clientes)} pastas de clientes.\n")
    
    dados_finais = []
    
    for cliente in pastas_clientes:
        caminho_cliente = os.path.join(caminho_completo_dia, cliente)
        print(f"💼 Cliente: {cliente}")
        
        try:
            arquivos_cliente = os.listdir(caminho_cliente)
            pdfs = [arq for arq in arquivos_cliente if arq.lower().endswith('.pdf')]
            
            if pdfs:
                nome_pdf = pdfs[0]
                caminho_completo_pdf = os.path.join(caminho_cliente, nome_pdf)
                print(f"    📄 Lendo Nota Fiscal: {nome_pdf}")
                
                codigos = extrair_dados_pdf(caminho_completo_pdf)
                
                for cod in codigos:
                    dados_finais.append({
                        "Data Faturamento": f"{dia_str}/{data_alvo.strftime('%m')}/{ano_str}",
                        "Cliente": cliente,
                        "Arquivo NF": nome_pdf,
                        "Código Produto": cod
                    })
            else:
                print(f"    ⚠ Nenhuma Nota Fiscal (.pdf) encontrada nesta pasta.")
        except Exception as e:
            print(f"    ❌ Erro ao acessar a pasta deste cliente: {e}")
            
        print("-" * 60)

    # 4. SALVAR OS ARQUIVOS EXCEL (Histórico + Cópia Fixo para o n8n)
    if dados_finais:
        df = pd.DataFrame(dados_finais)
        
        # Nomes dos arquivos
        nome_excel_historico = f"Conferencia_{dia_str}_{data_alvo.strftime('%m')}_{ano_str}.xlsx"
        nome_excel_fixo = "conferencia_atual.xlsx"
        
        # Salva o arquivo de histórico
        try:
            df.to_excel(nome_excel_historico, index=False)
            print(f"📝 Histórico salvo em: '{nome_excel_historico}'")
        except PermissionError:
            print(f"⚠ O arquivo '{nome_excel_historico}' estava aberto. Histórico não sobrescrito.")
            
        # Salva a cópia fixa que o n8n vai ler
        try:
            df.to_excel(nome_excel_fixo, index=False)
            print(f"🚀 Cópia para o n8n atualizada em: '{nome_excel_fixo}'")
        except PermissionError:
            print(f"❌ Erro: O arquivo '{nome_excel_fixo}' está aberto! Feche-o para o bot funcionar.")
            
        print("\n🚀 SUCESSO COMPLETO!")
    else:
        print("\n⚠ Processo concluído, mas nenhum código válido foi extraído.")

if __name__ == '__main__':
    main()