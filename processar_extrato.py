import pandas as pd
import re
import PyPDF2
from datetime import datetime
import os
import tempfile

def extract_text_from_pdf(pdf_file):
    """Extrai texto do arquivo PDF"""
    text = ""
    reader = PyPDF2.PdfReader(pdf_file)
    for page in reader.pages:
        text += page.extract_text()
    return text

def parse_transactions(text):
    """Analisa o texto e extrai informações das transações"""
    transactions = []
    
    # Padrões para identificar diferentes tipos de transações
    pix_pattern = r"Comprovante de Transferência.*?dados do pagador.*?nome do pagador: (.*?)(?:\n|$).*?CPF / CNPJ do pagador: (.*?)(?:\n|$).*?agência/conta: (.*?)(?:\n|$).*?dados do recebedor.*?nome do recebedor: (.*?)(?:\n|$).*?CPF / CNPJ do recebedor: (.*?)(?:\n|$).*?instituição: (.*?)(?:\n|$).*?valor: R\$ ([\d\.,]+).*?data da transferência: (\d{2}/\d{2}/\d{4}).*?tipo de pagamento: (.*?)(?:\n|$).*?transação efetuada em (\d{2}/\d{2}/\d{4}) às (\d{2}:\d{2}:\d{2})"
    
    titulo_pattern = r"Comprovante de Operação\s+- Títulos.*?Nome: (.*?)(?:\n|$).*?Nome do favorecido: (.*?)(?:\n|$).*?CPF/CNPJ do pagador: (.*?)(?:\n|$).*?Valor pago: R\$ ([\d\.,]+).*?Data de vencimento: (\d{2}/\d{2}/\d{4}).*?Pagamento efetuado em (\d{2}\.\d{2}\.\d{4}) às (\d{2}:\d{2}:\d{2})"
    
    qrcode_pattern = r"Comprovante de pagamento QR Code.*?nome do pagador: (.*?)(?:\n|$).*?CPF / CNPJ do pagador: (.*?)(?:\n|$).*?agência/conta: (.*?)(?:\n|$).*?nome do recebedor: (.*?)(?:\n|$).*?CPF / CNPJ do recebedor: (.*?)(?:\n|$).*?valor da transação: ([\d\.,]+).*?Pagamento efetuado em (\d{2}/\d{2}/\d{4}) às (\d{2}:\d{2}:\d{2})"
    
    # Encontrar transações PIX
    pix_matches = re.finditer(pix_pattern, text, re.DOTALL)
    for match in pix_matches:
        transactions.append({
            'Tipo': 'PIX',
            'Data': match.group(10),
            'Hora': match.group(11),
            'Pagador': match.group(1).strip(),
            'CPF/CNPJ Pagador': match.group(2).strip(),
            'Conta Pagador': match.group(3).strip(),
            'Beneficiário': match.group(4).strip(),
            'CPF/CNPJ Beneficiário': match.group(5).strip(),
            'Banco Beneficiário': match.group(6).strip(),
            'Valor': float(match.group(7).replace('.', '').replace(',', '.')),
            'Descrição': match.group(9).strip()
        })
    
    # Encontrar pagamentos de títulos
    titulo_matches = re.finditer(titulo_pattern, text, re.DOTALL)
    for match in titulo_matches:
        transactions.append({
            'Tipo': 'Título Bancário',
            'Data': match.group(6).replace('.', '/'),
            'Hora': match.group(7),
            'Pagador': match.group(1).strip(),
            'CPF/CNPJ Pagador': match.group(3).strip(),
            'Conta Pagador': '',
            'Beneficiário': match.group(2).strip(),
            'CPF/CNPJ Beneficiário': '',
            'Banco Beneficiário': '',
            'Valor': float(match.group(4).replace('.', '').replace(',', '.')),
            'Descrição': f'Vencimento: {match.group(5)}'
        })
    
    # Encontrar pagamentos QR Code
    qrcode_matches = re.finditer(qrcode_pattern, text, re.DOTALL)
    for match in qrcode_matches:
        transactions.append({
            'Tipo': 'QR Code',
            'Data': match.group(7),
            'Hora': match.group(8),
            'Pagador': match.group(1).strip(),
            'CPF/CNPJ Pagador': match.group(2).strip(),
            'Conta Pagador': match.group(3).strip(),
            'Beneficiário': match.group(4).strip(),
            'CPF/CNPJ Beneficiário': match.group(5).strip(),
            'Banco Beneficiário': '',
            'Valor': float(match.group(6).replace('.', '').replace(',', '.')),
            'Descrição': 'Pagamento QR Code'
        })
    
    return transactions

def create_excel(transactions):
    """Cria arquivo Excel com as transações e retorna como bytes"""
    df = pd.DataFrame(transactions)
    
    # Ordenar por data e hora
    try:
        df['Data_Hora'] = pd.to_datetime(df['Data'] + ' ' + df['Hora'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
        df = df.sort_values('Data_Hora')
        df = df.drop('Data_Hora', axis=1)
    except Exception as e:
        print(f"Erro ao ordenar por data e hora: {e}")
    
    # Formatar valor como moeda
    df['Valor'] = df['Valor'].apply(lambda x: f"R$ {x:.2f}")
    
    # Salvar em um buffer de bytes para download
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
        df.to_excel(tmp.name, index=False)
        tmp.flush()
        with open(tmp.name, 'rb') as f:
            return f.read(), len(transactions)

def process_pdf_to_excel(pdf_file):
    """Função principal que processa o PDF e retorna o Excel"""
    try:
        # Extrair texto do PDF
        text = extract_text_from_pdf(pdf_file)
        
        # Analisar transações
        transactions = parse_transactions(text)
        
        # Se não encontrou transações
        if not transactions:
            return None, 0, "Nenhuma transação encontrada no extrato. Verifique se o formato é compatível."
        
        # Criar Excel
        excel_data, num_transactions = create_excel(transactions)
        
        return excel_data, num_transactions, None
        
    except Exception as e:
        return None, 0, f"Erro ao processar o arquivo: {str(e)}"