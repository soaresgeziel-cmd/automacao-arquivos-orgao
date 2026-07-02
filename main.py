import streamlit as st
import pandas as pd
import re
import zipfile
import io

# Configuração inicial da página web com visual otimizado
st.set_page_config(
    page_title="Consolidador de Anexos", 
    page_icon="📦", 
    layout="centered"
)

# Estilização CSS Customizada para UX/UI Premium
st.markdown("""
    <style>
        /* Fundo e Container Geral */
        .main { background-color: #fcfcfd; }
        .block-container { padding-top: 3rem; max-width: 720px; }
        
        /* Títulos e Tipografia */
        h1 { color: #0f172a; font-weight: 800; font-size: 2.2rem; letter-spacing: -0.05em; }
        p.subtitle { color: #475569; font-size: 1.1rem; margin-top: -10px; margin-bottom: 2rem; }
        
        /* Cards de Passos (UX) */
        .step-card {
            background-color: #ffffff;
            padding: 24px;
            border-radius: 12px;
            box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.05), 0 1px 2px -1px rgba(0, 0, 0, 0.05);
            border: 1px solid #e2e8f0;
            margin-bottom: 1.5rem;
        }
        .step-title {
            color: #1e293b;
            font-size: 1.15rem;
            font-weight: 700;
            margin-bottom: 4px;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .step-desc { color: #64748b; font-size: 0.875rem; margin-bottom: 16px; }
        
        /* Ajuste do botão principal do Streamlit via CSS */
        div.stButton > button:first-child {
            background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
            color: white;
            border: none;
            padding: 12px 24px;
            font-weight: 600;
            font-size: 1rem;
            border-radius: 8px;
            width: 100%;
            box-shadow: 0 4px 6px -1px rgba(37, 99, 235, 0.2);
            transition: all 0.2s;
        }
        div.stButton > button:first-child:hover {
            background: linear-gradient(135deg, #1d4ed8 0%, #1e40af 100%);
            box-shadow: 0 10px 15px -3px rgba(37, 99, 235, 0.3);
            transform: translateY(-1px);
        }
    </style>
""", unsafe_allow_html=True)

# Topo do Site (Header)
st.markdown("<h1>📦 Consolidador de Anexos</h1>", unsafe_allow_html=True)
st.markdown("<p class='subtitle'>Tratamento e agrupamento inteligente de arquivos para automação com o Power Automate Desktop.</p>", unsafe_allow_html=True)

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
        <div class="step-title">📂 1. Base de Destinatários</div>
        <div class="step-desc">Insira a planilha Excel (.xlsx) contendo o mapeamento de Órgãos e seus respectivos E-mails.</div>
    </div>
""", unsafe_allow_html=True)
arquivo_emails = st.file_uploader("Upload e-mails", type=["xlsx"], label_visibility="collapsed")

# --- PASSO 2: Upload dos arquivos/anexos ---
st.markdown("""
    <div class="step-card">
        <div class="step-title">📄 2. Relatórios e Anexos</div>
        <div class="step-desc">Arraste de uma só vez todos os relatórios soltos que o robô deverá agrupar (.xlsx, .csv, .txt).</div>
    </div>
""", unsafe_allow_html=True)
arquivos_soltos = st.file_uploader("Upload arquivos", type=["xlsx", "xls", "csv", "txt"], accept_multiple_files=True, label_visibility="collapsed")

# --- MÉTRICAS VISUAIS (MUDANÇA RADICAL DE UX) ---
if arquivos_soltos:
    total_arquivos = len(arquivos_soltos)
    tamanho_total_bytes = sum(arq.size for arq in arquivos_soltos)
    tamanho_total_mb = tamanho_total_bytes / (1024 * 1024)
    
    # Exibe os dados em colunas limpas e modernas, em vez de uma caixa de alerta pesada
    col1, col2 = st.columns(2)
    with col1:
        st.metric(label="Arquivos Carregados", value=f"{total_arquivos} itens")
    with col2:
        st.metric(label="Volume de Dados", value=f"{tamanho_total_mb:.2f} MB")
    st.markdown("<br>", unsafe_allow_html=True)
else:
    st.caption("ℹ️ Aguardando a inclusão de arquivos para análise...")

# --- PASSO 3: Input do caminho local ---
st.markdown("""
    <div class="step-card">
        <div class="step-title">⚙️ 3. Diretório de Destino Local</div>
        <div class="step-desc">Informe o caminho da pasta física do seu computador onde o Power Automate fará a leitura dos arquivos.</div>
    </div>
""", unsafe_allow_html=True)
caminho_local_pc = st.text_input(
    "Caminho local:",
    placeholder="Ex: C:\\RoboAutomate\\Arquivos",
    label_visibility="collapsed"
).strip().strip('"')

st.markdown("<br>", unsafe_allow_html=True)

# --- PROCESSAMENTO ---
if st.button("🚀 Processar e Gerar Pacote"):
    if not arquivo_emails:
        st.error("Por favor, faça o upload do arquivo de e-mails.")
    elif not arquivos_soltos:
        st.error("Por favor, faça o upload de pelo menos um arquivo para agrupar.")
    elif not caminho_local_pc:
        st.error("Por favor, digite o caminho da pasta local para a planilha final.")
    else:
        try:
            with st.spinner("Engrenagens rodando... Formatando dados e compactando arquivos."):
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
                
                st.success("✨ Processamento concluído com sucesso!")
                
                st.download_button(
                    label="📥 BAIXAR PACOTE COMPLETO (.ZIP)",
                    data=zip_mestre_buffer.getvalue(),
                    file_name="pacote_automacao.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                st.warning("⚠️ **Próxima etapa:** Extraia o conteúdo deste arquivo .zip diretamente dentro da pasta local configurada no Passo 3.")
                
        except Exception as e:
            st.error(f"Ocorreu um erro ao processar: {e}")
