import streamlit as st
import pandas as pd
import win32com.client as win32
import pythoncom
import os
from streamlit_quill import st_quill

st.set_page_config(page_title="Alfredo do email", layout="wide")

# --- MINI MANUAL ATUALIZADO ---
with st.expander("📖 MANUAL: Como selecionar a pasta de anexos"):
    st.markdown("""
    1. **Abra a pasta** onde estão seus arquivos no Windows Explorer.
    2. Clique na **barra de endereços** no topo da pasta (onde aparece o caminho).
    3. Copie o texto (ex: `C:\\MeusDocumentos\\Certificados`).
    4. **Cole** esse texto no campo 'Caminho da pasta' na barra lateral do Alfredo.
    5. O programa removerá as aspas automaticamente se houver.
    """)

st.title("✉️ Alfredo do email")

# --- SIDEBAR ---
with st.sidebar:
    st.header("1. Configurações")
    uploaded_file = st.file_uploader("Suba sua Planilha", type=["xlsx", "csv"])

    # Campo de texto para o caminho da pasta (Substitui o buscador de pastas)
    caminho_input = st.text_input("Caminho da pasta de anexos (Copie e cole aqui)")
    caminho_anexos = caminho_input.replace('"', '').strip()  # Limpeza de aspas

if uploaded_file:
    if uploaded_file.name.endswith('xlsx'):
        excel_file = pd.ExcelFile(uploaded_file)
        aba = st.selectbox("Selecione a Aba", excel_file.sheet_names)
        df = pd.read_excel(uploaded_file, sheet_name=aba)
    else:
        df = pd.read_csv(uploaded_file)

    st.write("### 📋 Dados Identificados", df.head(3))
    colunas = df.columns.tolist()

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        col_email = st.selectbox("Coluna 'Para'", colunas)
    with c2:
        col_cc = st.selectbox("Coluna 'Cópia (CC)'", [None] + colunas)
    with c3:
        col_bcc = st.selectbox("Coluna 'Cópia Oculta (BCC)'", [None] + colunas)
    with c4:
        col_arq = st.selectbox("Coluna do Anexo", [None] + colunas)

    st.divider()

    st.subheader("📝 Redigir Mensagem")
    subject = st.text_input("Assunto do E-mail")

    content = st_quill(
        placeholder="Escreva seu e-mail aqui... use {Tags} para personalizar.",
        html=True,
        key="quill_editor"
    )

    # --- BOTÕES ---
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        btn_teste = st.button("🧪 Gerar 1 TESTE", use_container_width=True)
    with col_t2:
        btn_massa = st.button("🚀 Gerar TODOS", type="primary", use_container_width=True)

    if btn_teste or btn_massa:
        df_proc = df.head(1) if btn_teste else df

        try:
            pythoncom.CoInitialize()
            try:
                outlook = win32.GetActiveObject("Outlook.Application")
            except:
                outlook = win32.Dispatch('outlook.application')

            sucesso = 0
            erros_anexo = []
            progress = st.progress(0)

            for index, row in df_proc.iterrows():
                email_dest = str(row[col_email]).strip()
                if "@" not in email_dest: continue

                mail = outlook.CreateItem(0)
                mail.Display()  # Abre para carregar a assinatura

                # Preenchimento de destinatários (Garante que apareçam no Outlook)
                mail.To = email_dest
                if col_cc and pd.notna(row[col_cc]):
                    mail.CC = str(row[col_cc]).strip()
                if col_bcc and pd.notna(row[col_bcc]):
                    mail.BCC = str(row[col_bcc]).strip()

                # Substituição de Tags
                assunto_f = subject
                corpo_f = content
                for col in colunas:
                    tag = "{" + col + "}"
                    val = str(row[col]) if pd.notna(row[col]) else ""
                    assunto_f = assunto_f.replace(tag, val)
                    corpo_f = corpo_f.replace(tag, val)

                mail.Subject = assunto_f
                mail.HTMLBody = f"<div style='font-family: Calibri; font-size: 11pt;'>{corpo_f}</div><br>" + mail.HTMLBody

                # Lógica de Anexo corrigida
                if caminho_anexos and col_arq and pd.notna(row[col_arq]):
                    nome_arquivo = str(row[col_arq]).strip()
                    arq_path = os.path.join(caminho_anexos, nome_arquivo)

                    if os.path.exists(arq_path):
                        mail.Attachments.Add(os.path.abspath(arq_path))
                    else:
                        erros_anexo.append(f"Arquivo não encontrado: {nome_arquivo}")

                mail.Save()
                mail.Close(0)
                sucesso += 1
                progress.progress((index + 1) / len(df_proc))

            st.success(f"✅ Finalizado! {sucesso} rascunhos criados no Outlook.")
            if erros_anexo:
                with st.expander("⚠️ Ver arquivos não localizados"):
                    for err in erros_anexo: st.warning(err)

        except Exception as e:
            st.error(f"Erro no processamento: {e}")
else:
    st.info("Aguardando planilha...")