import streamlit as st
import pandas as pd
import re
import zipfile
import io

# Configuração inicial da página web
st.set_page_config(page_title="Consolidador de Arquivos", page_icon="📦", layout="centered")

st.title("📦 Consolidador de Anexos por Órgão")
st.write("Faça o upload dos arquivos e gere o pacote pronto para o Power Automate Desktop.")

def extrair_codigo_orgao(nome_arquivo):
    """Busca um padrão de 4 dígitos (código do órgão) no nome do arquivo."""
    padrao = re.search(r'\b\d{4}\b', nome_arquivo)
    if padrao:
        return padrao.group(0)
    padrao_livre = re.search(r'\d{4}', nome_arquivo)
    if padrao_livre:
        return padrao_livre.group(0)
    return None

# --- PASSO 1: Upload da lista de e-mails ---
st.subheader("1. Lista de E-mails")
arquivo_emails = st.file_uploader("Selecione o arquivo Excel (.xlsx) de Órgãos x E-mails", type=["xlsx"])

# --- PASSO 2: Upload dos arquivos/anexos ---
st.subheader("2. Arquivos para Agrupar")
arquivos_soltos = st.file_uploader(
    "Arraste e solte TODOS os arquivos soltos aqui (.xlsx, .csv, .txt)", 
    type=["xlsx", "xls", "csv", "txt"], 
    accept_multiple_files=True
)

# --- PASSO 3: Input do caminho local ---
st.subheader("3. Diretório de Destino Local")
caminho_local_pc = st.text_input(
    "Digite o caminho da pasta do seu computador onde você vai extrair os arquivos:",
    placeholder="Ex: C:\\RoboAutomate\\Arquivos"
).strip().strip('"')

# --- PROCESSAMENTO ---
if st.button("Processar e Gerar Pacote", type="primary"):
    if not arquivo_emails:
        st.error("Por favor, faça o upload do arquivo de e-mails.")
    elif not arquivos_soltos:
        st.error("Por favor, faça o upload de pelo menos um arquivo para agrupar.")
    elif not caminho_local_pc:
        st.error("Por favor, digite o caminho da pasta local para a planilha final.")
    else:
        try:
            with st.spinner("Processando dados e compactando arquivos..."):
                # 1. Ler e estruturar lista de e-mails
                df_emails_raw = pd.read_excel(arquivo_emails)
                col_orgao = df_emails_raw.columns[0]
                col_email = df_emails_raw.columns[1]
                
                df_emails_raw[col_orgao] = df_emails_raw[col_orgao].astype(str).str.strip().str.zfill(4)
                df_emails_raw[col_email] = df_emails_raw[col_email].astype(str).str.strip()
                
                dict_emails = df_emails_raw.groupby(col_orgao)[col_email].apply(lambda x: ";".join(set(x))).to_dict()

                # 2. Agrupar arquivos enviados por órgão
                arquivos_por_orgao = {}
                for arquivo in arquivos_soltos:
                    codigo = extrair_codigo_orgao(arquivo.name)
                    if codigo:
                        if codigo not in arquivos_por_orgao:
                            arquivos_por_orgao[codigo] = []
                        # Guarda o conteúdo em memória do arquivo e o nome dele
                        arquivos_por_orgao[codigo].append((arquivo.getvalue(), arquivo.name))

                # 3. Criar os ZIPs e a planilha na memória (para o usuário baixar tudo junto)
                zip_mestre_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_mestre_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_mestre:
                    dados_finais = []
                    todos_orgaos = sorted(list(arquivos_por_orgao.keys()))
                    
                    for orgao in todos_orgaos:
                        nome_zip_orgao = f"{orgao}.zip"
                        
                        # Criar o arquivo .zip do órgão em memória
                        zip_orgao_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_orgao_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf_orgao:
                            for conteudo, nome_arq in arquivos_por_orgao[orgao]:
                                zipf_orgao.writestr(nome_arq, conteudo)
                        
                        # Salva APENAS o .zip do órgão dentro do ZIP Mestre
                        zip_mestre.writestr(nome_zip_orgao, zip_orgao_buffer.getvalue())
                        
                        # Monta o caminho local dinâmico que o usuário digitou na tela
                        caminho_zip_final_pc = f"{caminho_local_pc}\\{nome_zip_orgao}" if caminho_local_pc.endswith('\\') else f"{caminho_local_pc}\\{nome_zip_orgao}"
                        
                        emails_agrupados = dict_emails.get(orgao, "")
                        dados_finais.append({
                            "Emails": emails_agrupados,
                            "Anexo": caminho_zip_final_pc
                        })
                    
                    # Criar a planilha Excel final em memória
                    df_final = pd.DataFrame(dados_finais, columns=["Emails", "Anexo"])
                    excel_buffer = io.BytesIO()
                    df_final.to_excel(excel_buffer, index=False)
                    
                    # Salva a planilha final dentro do ZIP Mestre
                    zip_mestre.writestr("Resultado_Consolidado.xlsx", excel_buffer.getvalue())
                
                st.success("✨ Processamento concluído!")
                
                # Botão de download do pacote completo (.zip mestre)
                st.download_button(
                    label="📥 Baixar Pacote Completo (.ZIP)",
                    data=zip_mestre_buffer.getvalue(),
                    file_name="pacote_automacao.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                st.info("💡 Após baixar, extraia o conteúdo deste arquivo .zip diretamente dentro da pasta local que você configurou no Passo 3.")
                
        except Exception as e:
            st.error(f"Ocorreu um erro ao processar: {e}")
