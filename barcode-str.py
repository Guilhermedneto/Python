import os
import re
import numpy as np
from datetime import datetime, timedelta
from pyzbar.pyzbar import decode
from pdf2image import convert_from_path
from openpyxl import Workbook
import streamlit as st

# Função para calcular o dígito verificador modulo 10
def modulo10(num):
    soma = 0
    peso = 2
    for c in reversed(num):
        parcial = int(c) * peso
        if parcial > 9:
            s = str(parcial)
            parcial = int(s[0]) + int(s[1])
        soma += parcial
        peso = 1 if peso == 2 else 2

    resto10 = soma % 10
    modulo10 = 0 if resto10 == 0 else 10 - resto10
    return modulo10

# Função para montar campo com dígito verificador
def monta_campo(campo):
    campo_dv = "%s%s" % (campo, modulo10(campo))
    return "%s.%s" % (campo_dv[0:5], campo_dv[5:])

# Função para gerar linha digitável
def linha_digitavel(linha):
    return ' '.join([monta_campo(linha[0:4] + linha[19:24]),
                     monta_campo(linha[24:34]),
                     monta_campo(linha[34:44]),
                     linha[4],
                     linha[5:19]])

# Função para extrair informações do código de barras do boleto
def extrair_informacoes(codigo_barras):
    valor = int(codigo_barras[9:19]) / 100.0
    fator_vencimento = int(codigo_barras[5:9])
    data_base = datetime(1997, 10, 7)
    data_vencimento = data_base + timedelta(days=fator_vencimento)
    linha = linha_digitavel(codigo_barras)
    return valor, data_vencimento.strftime('%Y-%m-%d'), linha

# Função para ler códigos de barras a partir de um PDF
def BarcodeReader(pdf_path):
    pages = convert_from_path(pdf_path, 500)
    detected_barcodes = []
    for page in pages:
        raw_img = page.convert('RGB')
        img = np.array(raw_img)
        
        barcodes = decode(img)
        detected_barcodes.extend([barcode.data.decode("utf-8") for barcode in barcodes if barcode.data != "" and barcode.type == 'I25'])
    return detected_barcodes

# Função para processar os boletos e extrair informações
def processar_boletos(uploaded_files):
    wb = Workbook()
    ws = wb.active
    ws.append(["Nome do Arquivo", "Código de Barras", "Valor", "Data de Vencimento", "Linha Digitável", "Data do Pagamento"])

    for uploaded_file in uploaded_files:
        pdf_path = f"temp_{uploaded_file.name}"
        with open(pdf_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        codigos_barras = BarcodeReader(pdf_path)
        
        if codigos_barras:
            for codigo in codigos_barras:
                valor, data_vencimento, linha = extrair_informacoes(codigo)
                data_pagamento = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                ws.append([uploaded_file.name, codigo, valor, data_vencimento, linha, data_pagamento])
            
            st.success(f"Boleto '{uploaded_file.name}' processado com sucesso.")
            st.write("Códigos de Barras Encontrados:")
            for codigo in codigos_barras:
                st.code(codigo)
                valor, data_vencimento, linha = extrair_informacoes(codigo)
                st.write(f"Valor: R$ {valor:.2f}")
                st.write(f"Data de Vencimento: {data_vencimento}")
                st.write(f"Linha Digitável: {linha}")
        else:
            st.warning(f"Nenhum código de barras encontrado no arquivo: {uploaded_file.name}")
        
        os.remove(pdf_path)

    # Salvar a planilha Excel
    excel_file = 'boletos_pagos.xlsx'
    wb.save(excel_file)

    return excel_file

# Inicializando o aplicativo Streamlit
st.title("Leitor de Código de Barras de Boletos")

# Upload de arquivos PDF
uploaded_files = st.file_uploader("Escolha arquivos PDF", type="pdf", accept_multiple_files=True)

if uploaded_files:
    excel_file = processar_boletos(uploaded_files)
    
    with open(excel_file, "rb") as f:
        st.download_button(
            label="Baixar Planilha de Boletos Pagos",
            data=f,
            file_name=excel_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
