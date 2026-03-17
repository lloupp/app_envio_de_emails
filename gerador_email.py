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

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("1. Configurações")
    arquivo_carregado = st.file_uploader("Suba sua Planilha", type=["xlsx", "csv"])

    # Campo de texto para o caminho da pasta (Substitui o buscador de pastas)
    caminho_input = st.text_input("Caminho da pasta de anexos (Copie e cole aqui)")
    caminho_anexos = caminho_input.replace('"', '').strip()  # Limpeza de aspas

if arquivo_carregado:
    if arquivo_carregado.name.endswith('xlsx'):
        arquivo_excel = pd.ExcelFile(arquivo_carregado)
        aba = st.selectbox("Selecione a Aba", arquivo_excel.sheet_names)
        df = pd.read_excel(arquivo_carregado, sheet_name=aba)
    else:
        df = pd.read_csv(arquivo_carregado)

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
    assunto = st.text_input("Assunto do E-mail")

    conteudo = st_quill(
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
        df_processado = df.head(1) if btn_teste else df

        try:
            pythoncom.CoInitialize()
            try:
                outlook = win32.GetActiveObject("Outlook.Application")
            except:
                outlook = win32.Dispatch('outlook.application')

            sucesso = 0
            erros_anexo = []
            barra_progresso = st.progress(0)

            for indice, linha in df_processado.iterrows():
                destinatario = str(linha[col_email]).strip()
                if "@" not in destinatario: continue

                rascunho = outlook.CreateItem(0)
                rascunho.Display()  # Abre para carregar a assinatura

                # Preenchimento de destinatários (Garante que apareçam no Outlook)
                rascunho.To = destinatario
                if col_cc and pd.notna(linha[col_cc]):
                    rascunho.CC = str(linha[col_cc]).strip()
                if col_bcc and pd.notna(linha[col_bcc]):
                    rascunho.BCC = str(linha[col_bcc]).strip()

                # Substituição de Tags
                assunto_formatado = assunto
                corpo_formatado = conteudo
                for col in colunas:
                    marcador = "{" + col + "}"
                    valor = str(linha[col]) if pd.notna(linha[col]) else ""
                    assunto_formatado = assunto_formatado.replace(marcador, valor)
                    corpo_formatado = corpo_formatado.replace(marcador, valor)

                rascunho.Subject = assunto_formatado
                rascunho.HTMLBody = f"<div style='font-family: Calibri; font-size: 11pt;'>{corpo_formatado}</div><br>" + rascunho.HTMLBody

                # Lógica de Anexo corrigida
                if caminho_anexos and col_arq and pd.notna(linha[col_arq]):
                    nome_arquivo = str(linha[col_arq]).strip()
                    caminho_arquivo = os.path.join(caminho_anexos, nome_arquivo)

                    if os.path.exists(caminho_arquivo):
                        rascunho.Attachments.Add(os.path.abspath(caminho_arquivo))
                    else:
                        erros_anexo.append(f"Arquivo não encontrado: {nome_arquivo}")

                rascunho.Save()
                rascunho.Close(0)
                sucesso += 1
                barra_progresso.progress((indice + 1) / len(df_processado))

            st.success(f"✅ Finalizado! {sucesso} rascunhos criados no Outlook.")
            if erros_anexo:
                with st.expander("⚠️ Ver arquivos não localizados"):
                    for mensagem_erro in erros_anexo: st.warning(mensagem_erro)

        except Exception as erro:
            st.error(f"Erro no processamento: {erro}")
else:
    st.info("Aguardando planilha...")