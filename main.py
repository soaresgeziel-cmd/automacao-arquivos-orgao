import streamlit as st
import pandas as pd
import re
import zipfile
import io

# Configuração da página
st.set_page_config(
    page_title="Consolidador de Anexos", 
    page_icon="📦", 
    layout="centered"
)

# Estilização CSS corrigida
estilo_css = """
<style>
    .main { background-color: #f8f9fa; }
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    h1 { color: #1e3a8a; font-weight: 700; }
    .step-box { 
        background-color: #ffffff; 
        padding: 20px; 
        border-radius: 10px; 
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05); 
        margin-bottom: 20px; 
        border-left: 5px solid #1e3a8a; 
    }
    .stButton>button { 
        background-color: #1e3a8a; 
        color: white; 
        border-radius: 8px; 
        padding: 10px 24px; 
        font-weight: bold; 
        width: 100%; 
        border: none; 
    }
</style>
"""
st.markdown(estilo_css, unsafe_allow_html=True)

# Topo do Site / Header
st.title("📦 Gerenciador de Anexos")
st.markdown("<p style='color: #64748b; font-size: 1.1rem; margin-top: -10px;'>Prepare seus arquivos para disparo automatizado no Power Automate Desktop com poucos cliques.</p>", unsafe_allow_html=True)
st.markdown("---")

def extrair_codigo_orgao(nome_arquivo):
    padrao = re.search(r'\b\d{4}\b', nome_arquivo)
    if padrao:
        return padrao.group(0)
    padrao_livre = re.search(r'\d{4}', nome_arquivo)
    if padrao_livre:
        return padrao_livre.group(0)
    return None

# --- PASSO 1 ---
st.markdown('<div class="step-box">### 📂 1. Base de Destinatários<br><small style="color: #64748b;">Faça o upload do arquivo Excel contendo a relação de Órgãos e E-mails.</small></div>', unsafe_allow_html=True)
arquivo_emails = st.file_uploader("Selecione a planilha de Órgãos x E-mails", type=["xlsx"], label_visibility="collapsed")

# --- PASSO 2 ---
st.markdown('<div class="step-box">### 📄 2. Arquivos para Agrupamento<br><small style="color: #64748b;">Arraste e solte todos os relatórios soltos de uma só vez (.xlsx, .csv, .txt).</small></div>', unsafe_allow_html=True)
arquivos_soltos = st.file_uploader("Arraste os arquivos aqui", type=["xlsx", "xls", "csv", "txt"], accept_multiple_files=True, label_visibility="collapsed")

# --- PASSO 3 ---
st.markdown('<div class="step-box">### ⚙️ 3. Configuração do Diretório Local<br><small style="color: #64748b;">Informe a pasta do seu computador onde o Power Automate fará a leitura física dos arquivos.</small></div>', unsafe_allow_html=True)
caminho_local_pc = st.text_input(
    "Caminho local da pasta:",
    placeholder="Ex: C:\\RoboAutomate\\Arquivos",
    label_visibility="collapsed"
).strip().strip('"')

st.markdown("<br>", unsafe_allow_html=True)

# --- PROCESSAMENTO ---
if st.button("🚀 PROCESSAR E GERAR PACOTE DE DOWNLOAD"):
    if not arquivo_emails:
        st.error("❌ Por favor, faça o upload do arquivo de e-mails no Passo 1.")
    elif not arquivos_soltos:
        st.error("❌ Por favor, selecione os arquivos para agrupar no Passo 2.")
    elif not caminho_local_pc:
        st.error("❌ Por favor, preencha o caminho da pasta local no Passo 3.")
    else:
        try:
            with st.spinner("Engrenagens rodando... Compactando arquivos e gerando o índice."):
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
                
                st.balloons()
                st.success("✨ Sucesso! Tudo processado e organizado perfeitamente.")
                
                st.download_button(
                    label="📥 BAIXAR PACOTE COMPLETO (.ZIP)",
                    data=zip_mestre_buffer.getvalue(),
                    file_name="pacote_automacao.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                st.warning("⚠️ **Próximo passo:** Extraia os arquivos baixados diretamente dentro da pasta local informada no Passo 3 para que o Power Automate funcione.")
                
        except Exception as e:
            st.error(f"Ocorreu um erro ao processar os arquivos: {e}")
