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

# ESTILIZAÇÃO AVANÇADA UI/UX (CSS PREMIUM INTERNOCIONAL)
st.markdown("""
    <style>
        /* Importação de Fonte Moderna e Configuração Global */
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
        
        * {
            font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif !important;
        }
        
        /* Fundo do App em Degradê Sutil (Design Moderno) */
        .stApp {
            background: linear-gradient(180deg, #f8fafc 0%, #f1f5f9 100%) !important;
        }
        
        /* Ajuste do bloco central */
        .block-container {
            padding-top: 4rem !important;
            padding-bottom: 4rem !important;
            max-width: 680px !important;
        }
        
        /* Cabeçalho de Alta Conversão / Corporativo */
        .header-title {
            background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            font-size: 2.5rem;
            font-weight: 800;
            letter-spacing: -0.05em;
            margin-bottom: 4px;
        }
        .header-subtitle {
            color: #64748b;
            font-size: 1.05rem;
            font-weight: 500;
            margin-bottom: 2.5rem;
        }
        
        /* Cards de Passos com Indicador de Borda Neon */
        .step-card {
            background: #ffffff;
            padding: 24px;
            border-radius: 16px;
            border: 1px solid #e2e8f0;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.02), 0 2px 4px -1px rgba(0, 0, 0, 0.02);
            margin-bottom: 1.5rem;
            position: relative;
            transition: all 0.3s ease;
        }
        .step-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 12px 20px -8px rgba(0, 0, 0, 0.08);
            border-color: #cbd5e1;
        }
        .step-tag {
            background: #eff6ff;
            color: #2563eb;
            font-size: 0.75rem;
            font-weight: 700;
            padding: 4px 10px;
            border-radius: 20px;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            display: inline-block;
            margin-bottom: 12px;
        }
        .step-title {
            color: #0f172a;
            font-size: 1.2rem;
            font-weight: 700;
            margin-bottom: 6px;
        }
        .step-desc {
            color: #64748b;
            font-size: 0.9rem;
            line-height: 1.5;
            margin-bottom: 0px;
        }
        
        /* Modernização das Caixas de Upload nativas do Streamlit */
        div[data-testid="stFileUploader"] {
            padding: 6px 0 0 0;
        }
        div[data-testid="stFileUploader"] section {
            background-color: #f8fafc !important;
            border: 2px dashed #cbd5e1 !important;
            border-radius: 12px !important;
            padding: 1.5rem !important;
            transition: all 0.2s ease;
        }
        div[data-testid="stFileUploader"] section:hover {
            border-color: #3b82f6 !important;
            background-color: #f0f7ff !important;
        }
        
        /* Remove qualquer margem ou label residual invisível que cause fantasmas visuais */
        div[data-testid="stFileUploader"] label {
            display: none !important;
        }
        
        /* Modernização do Campo de Input de Texto */
        div[data-testid="stTextInput"] input {
            background-color: #f8fafc !important;
            border: 1px solid #e2e8f0 !important;
            border-radius: 10px !important;
            padding: 12px 16px !important;
            font-size: 0.95rem !important;
            color: #334155 !important;
        }
        div[data-testid="stTextInput"] input:focus {
            border-color: #3b82f6 !important;
            box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.15) !important;
        }
        
        /* BOTÃO PRINCIPAL ULTRA-PRO (Estilo Dashboard Moderno) */
        div.stButton > button:first-child {
            background: linear-gradient(135deg, #1d4ed8 0%, #2563eb 100%) !important;
            color: white !important;
            border: none !important;
            padding: 14px 28px !important;
            font-weight: 700 !important;
            font-size: 1.05rem !important;
            border-radius: 12px !important;
            width: 100% !important;
            box-shadow: 0 4px 12px rgba(37, 99, 235, 0.25) !important;
            transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1) !important;
            letter-spacing: -0.01em;
            margin-top: 1rem;
        }
        div.stButton > button:first-child:hover {
            background: linear-gradient(135deg, #1e40af 0%, #1d4ed8 100%) !important;
            box-shadow: 0 8px 20px rgba(37, 99, 235, 0.4) !important;
            transform: translateY(-2px) !important;
        }
        div.stButton > button:first-child:active {
            transform: translateY(0px) !important;
        }
        
        /* Customização dos Widgets de Métricas Modernas */
        div[data-testid="stMetric"] {
            background: #ffffff !important;
            padding: 16px 20px !important;
            border-radius: 14px !important;
            border: 1px solid #e2e8f0 !important;
            box-shadow: 0 1px 3px rgba(0,0,0,0.01) !important;
        }
        div[data-testid="stMetricLabel"] p {
            color: #64748b !important;
            font-weight: 600 !important;
            font-size: 0.85rem !important;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }
        div[data-testid="stMetricValue"] div {
            color: #0f172a !important;
            font-weight: 800 !important;
            font-size: 1.6rem !important;
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

# --- PASSO 1: Lista de E-mails ---
st.markdown("""
    <div class="step-card">
        <span class="step-tag">Etapa 01</span>
        <div class="step-title">Base de Destinatários</div>
        <div class="step-desc">Selecione o arquivo mestre do Excel (.xlsx) contendo a relação de Órgãos e E-mails correspondentes.</div>
    </div>
""", unsafe_allow_html=True)
arquivo_emails = st.file_uploader("", type=["xlsx"], label_visibility="collapsed")

# --- PASSO 2: Upload dos arquivos/anexos ---
st.markdown("""
    <div class="step-card">
        <span class="step-tag">Etapa 02</span>
        <div class="step-title">Relatórios e Documentos</div>
        <div class="step-desc">Arraste e solte todos os arquivos soltos em lote que o sistema deve processar (.xlsx, .csv, .txt).</div>
    </div>
""", unsafe_allow_html=True)
arquivos_soltos = st.file_uploader("", type=["xlsx", "xls", "csv", "txt"], accept_multiple_files=True, label_visibility="collapsed")

# --- MÉTRICAS MODERNAS (UX PREMIUM) ---
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

# --- PASSO 3: Input do caminho local ---
st.markdown("""
    <div class="step-card">
        <span class="step-tag">Etapa 03</span>
        <div class="step-title">Diretório do Robô Local</div>
        <div class="step-desc">Insira a pasta exata do Windows onde o Power Automate executará a extração e leitura física dos itens.</div>
    </div>
""", unsafe_allow_html=True)
caminho_local_pc = st.text_input(
    "",
    placeholder="Ex: C:\\RoboAutomate\\Arquivos",
    label_visibility="collapsed"
).strip().strip('"')

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
