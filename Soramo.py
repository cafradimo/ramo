import os
import re
import tempfile
import shutil
from io import BytesIO
from collections import defaultdict
import pandas as pd
import streamlit as st
import pdfplumber
from fpdf import FPDF
from datetime import datetime
from PIL import Image
from tempfile import NamedTemporaryFile
from PyPDF2 import PdfReader

# =================== CONFIGURAÇÃO ===================
st.set_page_config(page_title="CREA-RJ", layout="wide", page_icon="")

# =================== FUNÇÕES AUXILIARES ===================
def criar_temp_dir():
    """Cria diretório temporário"""
    return tempfile.mkdtemp()

def limpar_temp_dir(temp_dir):
    """Remove diretório temporário"""
    shutil.rmtree(temp_dir, ignore_errors=True)

def extrair_data_relatorio(texto):
    """Extrai a data do relatório do texto do PDF"""
    # Procura pelo padrão "Data Relatório : DD/MM/YYYY"
    padrao_data = re.search(r'Data\s+Relatório\s*:\s*(\d{2}/\d{2}/\d{4})', texto, re.IGNORECASE)
    if padrao_data:
        try:
            data = datetime.strptime(padrao_data.group(1), '%d/%m/%Y')
            return data
        except ValueError:
            pass
    
    # Se não encontrar no formato específico, tenta outros padrões
    padrao_data_alternativo = re.search(r'Data\s+Relatório\s*:\s*(\d{2}-\d{2}-\d{4})', texto, re.IGNORECASE)
    if padrao_data_alternativo:
        try:
            data = datetime.strptime(padrao_data_alternativo.group(1), '%d-%m-%Y')
            return data
        except ValueError:
            pass
    
    return None

# =================== MÓDULO DE EXTRAÇÃO ===================
def extrair_dados_ramo_atividade(texto, filename):
    """Extrai dados para Ramo de Atividade"""
    dados = {
        'Arquivo': filename,
        'Ramo': '',
        'Qtd. Ramo': '',
        'Data': None
    }
    
    # Extrai data do relatório
    dados['Data'] = extrair_data_relatorio(texto)
    
    secao = re.search(r'04\s*-\s*Identificação.*?(?=05\s*-|$)', texto, re.DOTALL|re.IGNORECASE)
    if secao:
        ramos = re.findall(r'Ramo\s*Atividade\s*:\s*(.*?)(?=\n|$)', secao.group(), re.IGNORECASE)
        if ramos:
            contagem = defaultdict(int)
            for ramo in [r.strip() for r in ramos if r.strip()]:
                contagem[ramo] += 1
            
            dados['Ramo'] = ", ".join(contagem.keys())
            dados['Qtd. Ramo'] = ", ".join(map(str, contagem.values()))
    
    return dados

def extrair_fiscal(texto):
    """Extrai o nome do fiscal do texto do PDF"""
    fiscal = re.search(r'Agente\s+de\s+Fiscalização\s*:\s*([^\n]+)', texto)
    return fiscal.group(1).strip() if fiscal else "Não identificado"

# =================== GERADORES DE RELATÓRIO PDF ===================
def gerar_relatorio_ramo_atividade(df, fiscal, primeira_data, ultima_data):
    """Gera PDF para Ramo de Atividade"""
    pdf = FPDF()
    pdf.add_page()
    
    # Adiciona logo
    try:
        logo_path = "10.png"
        if os.path.exists(logo_path):
            pdf.image(logo_path, x=60, y=10, w=90)  # Centralizado
    except:
        pass
    
    pdf.set_y(40)  # Espaço para o logo
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "RELATÓRIO DE RAMOS DE ATIVIDADE", 0, 1, 'C')
    
    # Informação do Agente de Fiscalização
    pdf.set_font("Arial", '', 12)
    pdf.cell(0, 10, f"Agente de Fiscalização: {fiscal}", 0, 1, 'L')
    
    # Informação do Período
    if primeira_data and ultima_data:
        pdf.cell(0, 10, f"Período: {primeira_data.strftime('%d/%m/%Y')} a {ultima_data.strftime('%d/%m/%Y')}", 0, 1, 'L')
    
    pdf.ln(10)
    
    # Processa dados
    contagem = defaultdict(int)
    for _, row in df[df['Arquivo'] != 'TOTAL GERAL'].iterrows():
        if row['Ramo'] and row['Qtd. Ramo']:
            for ramo, qtd in zip(row['Ramo'].split(','), row['Qtd. Ramo'].split(',')):
                if ramo.strip() and qtd.strip().isdigit():
                    contagem[ramo.strip()] += int(qtd.strip())
    
    # Ordena por quantidade
    ramos_ordenados = sorted(contagem.items(), key=lambda x: x[1], reverse=True)
    total = sum(contagem.values())
    
    # Cabeçalho
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(120, 8, "RAMO DE ATIVIDADE", 1, 0, 'C')
    pdf.cell(30, 8, "QUANTIDADE", 1, 0, 'C')
    pdf.cell(30, 8, "PORCENTAGEM", 1, 1, 'C')
    
    # Dados
    pdf.set_font("Arial", size=9)
    for ramo, qtd in ramos_ordenados:
        pdf.cell(120, 8, ramo[:60] + ('...' if len(ramo) > 60 else ''), 1)
        pdf.cell(30, 8, str(qtd), 1, 0, 'C')
        pdf.cell(30, 8, f"{(qtd/total)*100:.1f}%" if total > 0 else "0%", 1, 1, 'C')
    
    # Rodapé
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(120, 8, "TOTAL GERAL", 1)
    pdf.cell(30, 8, str(total), 1, 0, 'C')
    pdf.cell(30, 8, "100%", 1, 1, 'C')
    
    pdf.ln(10)
    pdf.set_font("Arial", 'I', 8)
    pdf.cell(0, 10, f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}", 0, 0, 'C')
    
    return pdf.output(dest='S').encode('latin1')

# =================== MÓDULO PRINCIPAL ===================
def extrator_pdf_consolidado():
    st.title(" Extrator PDF - Ramos de Atividade")
    st.markdown("""
    **Extrai automaticamente dados de:**
    - Ramos de Atividade  
    """)

    uploaded_files = st.file_uploader("Selecione os PDFs", type="pdf", accept_multiple_files=True)
    
    if uploaded_files:
        temp_dir = criar_temp_dir()
        try:
            with st.spinner("Processando arquivos..."):
                dados_ra = []
                fiscais = set()
                datas = []  # Lista para armazenar todas as datas encontradas
                
                for file in uploaded_files:
                    temp_path = os.path.join(temp_dir, file.name)
                    with open(temp_path, "wb") as f:
                        f.write(file.getbuffer())
                    
                    with pdfplumber.open(temp_path) as pdf:
                        texto = "\n".join(p.extract_text() or "" for p in pdf.pages)
                    
                    # Extrai dados de Ramo de Atividade
                    dados_arquivo = extrair_dados_ramo_atividade(texto, file.name)
                    dados_ra.append(dados_arquivo)
                    
                    # Armazena a data se encontrada
                    if dados_arquivo['Data']:
                        datas.append(dados_arquivo['Data'])
                    
                    # Extrai nome do fiscal
                    fiscal = extrair_fiscal(texto)
                    fiscais.add(fiscal)
                    
                    os.unlink(temp_path)
                
                # Determina a primeira e última data
                primeira_data = min(datas) if datas else None
                ultima_data = max(datas) if datas else None
                
                # Cria DataFrame
                df_ra = pd.DataFrame(dados_ra)
                
                # Adiciona totais
                total_ra = sum(int(q) for r in dados_ra for q in r['Qtd. Ramo'].split(',') if r['Qtd. Ramo'] and q.strip().isdigit())
                df_ra = pd.concat([df_ra, pd.DataFrame({
                    'Arquivo': ['TOTAL GERAL'],
                    'Ramo': [''],
                    'Qtd. Ramo': [str(total_ra)]
                })], ignore_index=True)
                
                # Exibição
                st.dataframe(df_ra)
                
                # Obtém o fiscal principal (primeiro da lista)
                fiscal_principal = list(fiscais)[0] if fiscais else "Não identificado"
                
                # Geração de relatórios
                pdf_ra = gerar_relatorio_ramo_atividade(df_ra, fiscal_principal, primeira_data, ultima_data)
                
                # Excel consolidado
                excel_buffer = BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                    df_ra.to_excel(writer, sheet_name='Ramo Atividade', index=False)
                
                # Download
                st.success("Processamento concluído!")
                
                # Exibe informações do fiscal e período
                st.markdown(f"**Agente de Fiscalização:** {fiscal_principal}")
                if primeira_data and ultima_data:
                    st.markdown(f"**Período:** {primeira_data.strftime('%d/%m/%Y')} a {ultima_data.strftime('%d/%m/%Y')}")
                
                st.download_button(
                    "⬇️ Baixar Excel Completo",
                    excel_buffer.getvalue(),
                    "ramos_atividade.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                with st.expander("⬇️ Baixar Relatório (PDF)"):
                    st.download_button(
                        "Ramo Atividade",
                        pdf_ra,
                        "relatorio_ramos.pdf"
                    )
        
        finally:
            limpar_temp_dir(temp_dir)

# =================== INTERFACE PRINCIPAL ===================
def main():
    # Configuração visual
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        try:
            logo = Image.open("10.png")
            st.image(logo, width=400)  # Logo centralizado
        except:
            st.write("CREA-RJ - Conselho Regional de Engenharia e Agronomia do Rio de Janeiro")
    
    st.markdown("---")
    
    # Exibe apenas o módulo principal
    extrator_pdf_consolidado()

if __name__ == "__main__":
    main()