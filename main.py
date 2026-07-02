import streamlit as st
import pandas as pd
import re
import zipfile
import io

# Configuração inicial da página web com visual otimizado
st.set_page_config(
    page_title="Consolidador de Anexos Pro", 
    page_icon="📦", 
    layout="centered"
)

# ESTILIZAÇÃO ULTRA-MODERNA (APENAS ESTÉTICA, SEM MEXER NA ESTRUTURA NATIVA)
st.markdown("""
    <style>
        /* Importação da Fonte Premium Inter */
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
        
        * {
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif !important;
        }
        
        /* Fundo do App Moderno em Degradê Sutil */
        .stApp {
            background: linear-gradient(180deg, #f8fafc 0%, #f1f5f9 100%) !important;
        }
        
        /* Ajuste do bloco central */
        .block-container {
            padding-top: 3rem !important;
            padding-bottom: 4rem !important;
            max-width: 680px !important;
        }
        
        /* Cabeçalho Corporativo Elegante */
        .header-title {
            background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            font-size: 2.3rem;
            font-weight: 800;
            letter-spacing: -0.04em;
            margin-bottom: 6px;
        }
        .header-subtitle {
            color: #64748b;
            font-size: 1rem;
            font-weight: 500;
            margin-bottom: 2rem;
        }
        
        /* Estilização das Labels Tradicionais (Idênticas ao seu original, mas com fonte linda) */
        div[data-testid="stWidgetLabel"] p {
            color: #1e293b !important;
            font-size: 1.05rem !important;
            font-weight: 600 !important;
            letter-spacing: -0.01em !important;
        }
        
        /* Modernização das Caixas de Upload (Layout Padrão Mantido) */
        div[data-testid="stFileUploader"] section {
            background-color: #ffffff !important;
            border: 2px dashed #cbd5e1 !important;
            border-radius: 12px !important;
            padding: 1rem !important;
            box-shadow: 0 1px 3px rgba(0,0,0,0.02) !important;
            transition: all 0.2s ease !important;
        }
        div[data-testid="stFileUploader"] section:hover {
            border-color: #3b82f6 !important;
            background-color: #f0f7ff !important;
        }
        
        /* Modernização do Campo de Input de Texto */
        div[data-testid="stTextInput"] input {
            background-color: #ffffff !important;
            border: 1px solid #cbd5e1 !important;
            border-radius: 10px !important;
            padding: 10px 14px !important;
            font-size: 0.95rem !important;
            color: #334155 !important;
        }
        div[data-testid="stTextInput"] input:focus {
            border-color: #3b82f6 !important;
            box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.15) !important;
        }
        
        /* BOTÃO PRINCIPAL ULTRA-PRO MODERNO */
        div.stButton > button:first-child {
            background: linear-gradient(135deg, #1d4ed8 0%, #2563eb 100%) !important;
            color: white !important;
            border: none !important;
            padding: 12px 24px !important;
            font-weight: 700 !important;
            font-size: 1.05rem !important;
            border-radius: 10px !important;
            width: 100% !important;
            box-shadow: 0 4px 12px rgba(37, 99, 235, 0.25) !important;
            transition: all 0.2s ease !important;
            margin-top: 1rem;
        }
        div.stButton > button:first-child:hover {
            background: linear-gradient(135deg, #1e40af 0%, #1d4ed8 100%) !important;
            box-shadow: 0 6px 16px rgba(37, 99, 235, 0.35) !important;
            transform: translateY(-1px) !important;
        }
        
        /* Customização sutil dos blocos de métricas */
        div[data-testid="stMetric"] {
            background: #ffffff !important;
            padding: 14px 18px !important;
            border-radius: 12px !important;
            border: 1px solid #e2e8f0 !important;
        }
    </style>
""", unsafe_allow_html=True)

# Topo do Site (Header Premium)
st.markdown('<div class="header-title">Consolidador de Anexos</div>', unsafe_allow_html=True)
st.markdown('<div class="header-subtitle">Tratamento e agrupamento inteligente de relatórios para o Power Automate Desktop.</div>', unsafe_allow_html=True)

def extrair_codigo_orgao(nome_arquivo):
    padrao = re.search(r'\b\d{4}\b', nome_arquivo)
    if padrao:
        return padrao.group(0)
    padrao_livre = re.search(r'\d{4}', nome_arquivo)
    if padrao_livre:
        return padrao_livre.group(0)
    return None

# --- PASSO 1: Base de Destinatários ---
# Mantido exatamente igual ao seu código original estruturalmente
arquivo_emails = st.file_uploader("📂 Passo 1: Selecione a Base de Destinatários (.xlsx)", type=["xlsx"])

st.markdown("<br>", unsafe_allow_html=True)

# --- PASSO 2: Relatórios e Documentos ---
# Mantido exatamente igual ao seu código original estruturalmente
arquivos_soltos = st.file_uploader("📄 Passo 2: Arraste os Relatórios e Anexos (Em Lote)", type=["xlsx", "xls", "csv", "txt"], accept_multiple_files=True)

# --- MÉTRICAS MODERNAS ---
if arquivos_soltos:
    total_arquivos = len(arquivos_soltos)
    tamanho_total_bytes = sum(arq.size for arq in arquivos_soltos)
    tamanho_total_mb = tamanho_total_bytes / (1024 * 1024)
    
    col1, col2 = st.columns(2)
    with col1:
        st.metric(label="Arquivos Carregados", value=f"{total_arquivos} un")
    with col2:
        st.metric(label="Volume do Lote", value=f"{tamanho_total_mb:.2f} MB")
    st.markdown("<br>", unsafe_allow_html=True)
else:
    st.caption("ℹ️ Aguardando inserção de arquivos do Passo 2 para análise de metadados...")

st.markdown("<br>", unsafe_allow_html=True)

# --- PASSO 3: Diretório do Robô Local ---
# Mantido exatamente igual ao seu código original estruturalmente
caminho_local_pc = st.text_input("⚙️ Passo 3: Insira o Diretório do Robô Local", placeholder="Ex: C:\\RoboAutomate\\Arquivos").strip().strip('"')

st.markdown("<br>", unsafe_allow_html=True)

# --- PROCESSAMENTO ---
if st.button("🚀 Iniciar Processamento e Gerar Carga"):
    if not arquivo_emails:
        st.error("Por favor, faça o upload do arquivo de e-mails.")
    elif not arquivos_soltos:
        st.error("Por favor, faça o upload de pelo menos um arquivo para agrupar.")
    elif not caminho_local_pc:
        st.error("Por favor, digite o caminho da pasta local para a planilha final.")
    else:
        try:
            with st.spinner("Compilando pacotes binários e gerando índices estruturados..."):
                df_emails_raw = pd.read_excel(arquivo_emails)
                col_orgao = df_emails_raw.columns[0]
                col_email = df_emails_raw.columns[1]
                
                df_emails_raw[col_orgao] = df_emails_raw[col_orgao].astype(str).str.strip().str.zfill(4)
                df_emails_raw[col_email] = df_emails_raw[col_email].astype(str).str.strip()
                
                dict_emails = df_emails_raw.groupby(col_orgao)[col_email].apply(lambda x: ";".join(set(x))).to_dict()

                arquivos_por_orgao = {}
                for arquivo in arquivos_soltos:
                    codigo = extrair_codigo_orgao(arquivo.name)
                    if codigo:
                        if codigo not in arquivos_por_orgao:
                            arquivos_por_orgao[codigo] = []
                        arquivos_por_orgao[codigo].append((arquivo.getvalue(), arquivo.name))

                zip_mestre_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_mestre_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_mestre:
                    dados_finais = []
                    todos_orgaos = sorted(list(arquivos_por_orgao.keys()))
                    
                    for orgao in todos_orgaos:
                        nome_zip_orgao = f"{orgao}.zip"
                        
                        zip_orgao_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_orgao_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf_orgao:
                            for conteudo, nome_arq in arquivos_por_orgao[orgao]:
                                zipf_orgao.writestr(nome_arq, conteudo)
                        
                        zip_mestre.writestr(nome_zip_orgao, zip_orgao_buffer.getvalue())
                        
                        caminho_zip_final_pc = f"{caminho_local_pc}\\{nome_zip_orgao}"
                        
                        emails_agrupados = dict_emails.get(orgao, "")
                        dados_finais.append({
                            "Emails": emails_agrupados,
                            "Anexo": caminho_zip_final_pc
                        })
                    
                    df_final = pd.DataFrame(dados_finais, columns=["Emails", "Anexo"])
                    excel_buffer = io.BytesIO()
                    df_final.to_excel(excel_buffer, index=False)
                    
                    zip_mestre.writestr("Resultado_Consolidado.xlsx", excel_buffer.getvalue())
                
                st.success("✨ Lote processado e compactado com sucesso!")
                
                st.download_button(
                    label="📥 BAIXAR PACOTE CONSOLIDADO (.ZIP)",
                    data=zip_mestre_buffer.getvalue(),
                    file_name="pacote_automacao.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                st.warning("⚠️ **Instrução Técnica:** Extraia os arquivos baixados diretamente dentro da pasta local configurada no Passo 3 para manter a integridade do robô.")
                
        except Exception as e:
            st.error(f"Erro crítico no processamento de dados: {e}")
