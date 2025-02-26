import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
import csv
from io import StringIO
from collections import defaultdict

def process_xml(file):
    """Processa um arquivo XML e retorna os dados extraídos."""
    try:
        tree = ET.parse(file)
        root = tree.getroot()
        
        data = []
        item_summary = defaultdict(int)
        headers = ["Grupo", "Entrega", "LojaCompradora", "Item", "CodigoFab", "DescricaoResumida", "Qtde", "CNPJLojaCompradora"]
        
        for pedido in root.findall(".//Pedidos"):
            grupo = pedido.find("Grupo").text if pedido.find("Grupo") is not None else ""
            entrega = pedido.find("Entrega").text if pedido.find("Entrega") is not None else ""
            loja_compradora = pedido.find("LojaCompradora").text if pedido.find("LojaCompradora") is not None else ""
            item = pedido.find("Item").text if pedido.find("Item") is not None else ""
            codigo_fab = pedido.find("CodigoFab").text if pedido.find("CodigoFab") is not None else ""
            descricao_resumida = pedido.find("DescricaoResumida").text if pedido.find("DescricaoResumida") is not None else ""
            qtde = pedido.find("Qtde").text if pedido.find("Qtde") is not None else ""
            qtde = qtde.replace(".", ",") if qtde else ""
            cnpj = pedido.find("CNPJLojaCompradora").text if pedido.find("CNPJLojaCompradora") is not None else ""
            cnpj = f"'{cnpj}" if cnpj else ""
            
            data.append([grupo, entrega, loja_compradora, item, codigo_fab, descricao_resumida, qtde, cnpj])
            
            # Atualiza o resumo de quantidade total por item
            item_summary[(codigo_fab, descricao_resumida)] += int(qtde.replace(",", "")) if qtde else 0
        
        return headers, data, item_summary
    except ET.ParseError as e:
        st.error(f"Erro ao processar o arquivo XML: {e}")
        return None, None, None

# Interface do Streamlit
st.title("Extração de Dados de XML")

uploaded_files = st.file_uploader("Envie os arquivos XML", accept_multiple_files=True, type=["xml"])

if uploaded_files:
    all_data = []
    headers = None
    total_summary = defaultdict(int)
    
    for file in uploaded_files:
        file_headers, file_data, file_summary = process_xml(file)
        if file_data:
            headers = file_headers
            all_data.extend(file_data)
            for key, value in file_summary.items():
                total_summary[key] += value
    
    if all_data:
        df = pd.DataFrame(all_data, columns=headers)
        st.write("### Dados Extraídos")
        st.dataframe(df)
        
        # Exibir resumo de quantidade total por item
        summary_data = [[codigo_fab, descricao_resumida, qtde_total] for (codigo_fab, descricao_resumida), qtde_total in total_summary.items()]
        summary_df = pd.DataFrame(summary_data, columns=["CodigoFab", "DescricaoResumida", "Qtde Total"])
        st.write("### Quantidade Total por Item")
        st.dataframe(summary_df)
        
        # Converter DataFrame para CSV
        csv_buffer = StringIO()
        df.to_csv(csv_buffer, sep=';', index=False, encoding='utf-8', quoting=csv.QUOTE_MINIMAL)
        csv_buffer.seek(0)
        
        # Botão para baixar o CSV
        st.download_button(label="Baixar CSV", data=csv_buffer, file_name="pedidos_extracao.csv", mime="text/csv")
    else:
        st.warning("Nenhum dado foi extraído dos arquivos XML.")
