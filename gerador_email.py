import streamlit as st
import pandas as pd
import win32com.client as win32
import pythoncom
import os
from streamlit_quill import st_quill

st.set_page_config(page_title="Alfredo do Email", layout="wide")

# --- MINI MANUAL ---
with st.expander("📖 MANUAL: Como selecionar a pasta de anexos"):
    st.markdown("""
    1. **Abra a pasta** onde estão seus arquivos no Windows Explorer.
    2. Clique na **barra de endereços** no topo da pasta (onde aparece o caminho).
    3. Copie o texto (ex: `C:\\MeusDocumentos\\Certificados`).
    4. **Cole** esse texto no campo "Caminho da pasta" na barra lateral do Alfredo.
    5. O programa removerá as aspas automaticamente se houver.
    """)

st.title("✉️ Alfredo do Email")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("1. Planilha")
    uploaded_file = st.file_uploader("Envie sua planilha", type=["xlsx", "csv"])

    st.header("2. Anexos")
    caminho_input = st.text_input(
        "Caminho da pasta de anexos",
        placeholder="Ex: C:\\Documentos\\Certificados",
        help="Copie o caminho da pasta no Windows Explorer e cole aqui.",
    )
    caminho_anexos = caminho_input.replace('"', '').strip()

if uploaded_file:
    if uploaded_file.name.endswith('xlsx'):
        excel_file = pd.ExcelFile(uploaded_file)
        aba = st.selectbox("Selecione a aba", excel_file.sheet_names)
        df = pd.read_excel(uploaded_file, sheet_name=aba)
    else:
        df = pd.read_csv(uploaded_file)

    st.write("### 📋 Pré-visualização dos dados", df.head(3))
    colunas = df.columns.tolist()

    st.subheader("🗂️ Mapeamento de Colunas")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        col_email = st.selectbox("Destinatário (Para)", colunas)
    with c2:
        col_cc = st.selectbox("Cópia (CC)", [None] + colunas)
    with c3:
        col_bcc = st.selectbox("Cópia Oculta (CCO)", [None] + colunas)
    with c4:
        col_arq = st.selectbox("Coluna do Anexo", [None] + colunas)

    st.divider()

    st.subheader("📝 Redigir Mensagem")
    subject = st.text_input("Assunto do e-mail", placeholder="Ex: Seu certificado, {NomeDaColuna}!")

    content = st_quill(
        placeholder="Escreva seu e-mail aqui... use {NomeDaColuna} para personalizar.",
        html=True,
        key="quill_editor"
    )

    # --- BOTÕES ---
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        btn_teste = st.button("🧪 Gerar 1 Teste", use_container_width=True)
    with col_t2:
        btn_massa = st.button("🚀 Gerar Todos", type="primary", use_container_width=True)

    if btn_teste or btn_massa:
        if not subject.strip():
            st.warning("⚠️ O assunto do e-mail está vazio. Preencha antes de continuar.")
        elif not (content and content.strip()):
            st.warning("⚠️ O corpo do e-mail está vazio. Escreva a mensagem antes de continuar.")
        else:
            df_proc = df.head(1) if btn_teste else df
            total = len(df_proc)

            try:
                pythoncom.CoInitialize()
                try:
                    outlook = win32.GetActiveObject("Outlook.Application")
                except Exception:
                    outlook = win32.Dispatch('outlook.application')

                sucesso = 0
                erros_anexo = []
                progress = st.progress(0, text="Gerando rascunhos...")

                for contador, (_, row) in enumerate(df_proc.iterrows(), start=1):
                    email_dest = str(row[col_email]).strip()
                    if "@" not in email_dest:
                        continue

                    mail = outlook.CreateItem(0)
                    mail.Display()  # Abre para carregar a assinatura padrão do Outlook

                    mail.To = email_dest
                    if col_cc and pd.notna(row[col_cc]):
                        mail.CC = str(row[col_cc]).strip()
                    if col_bcc and pd.notna(row[col_bcc]):
                        mail.BCC = str(row[col_bcc]).strip()

                    # Substituição de tags pelo valor da coluna correspondente
                    assunto_f = subject
                    corpo_f = content
                    for col in colunas:
                        tag = "{" + col + "}"
                        val = str(row[col]) if pd.notna(row[col]) else ""
                        assunto_f = assunto_f.replace(tag, val)
                        corpo_f = corpo_f.replace(tag, val)

                    mail.Subject = assunto_f
                    mail.HTMLBody = (
                        f"<div style='font-family: Calibri; font-size: 11pt;'>{corpo_f}</div><br>"
                        + mail.HTMLBody
                    )

                    # Anexo
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
                    progress.progress(contador / total, text=f"Processando {contador} de {total}...")

                progress.empty()
                st.success(f"✅ Concluído! {sucesso} rascunho(s) criado(s) no Outlook.")
                if erros_anexo:
                    with st.expander(f"⚠️ {len(erros_anexo)} arquivo(s) não localizado(s) — clique para ver"):
                        for err in erros_anexo:
                            st.warning(err)

            except Exception as e:
                st.error(f"❌ Erro ao processar: {e}")
            finally:
                pythoncom.CoUninitialize()
else:
    st.info("👈 Envie uma planilha na barra lateral para começar.")