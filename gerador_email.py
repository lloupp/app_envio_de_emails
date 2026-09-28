import os
import smtplib
import ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
import pythoncom
import streamlit as st
import win32com.client as win32
from streamlit_quill import st_quill


st.set_page_config(page_title="Alfredo do email", layout="wide")

# Toolbar do Quill restrita às opções cujo HTML resultante o Outlook (motor Word)
# realmente renderiza. "formula", "code" e "code-block" usam CSS/KaTeX que o
# Outlook não recebe, então o botão aparecia mas o efeito sumia no rascunho final.
TOOLBAR_QUILL = [
    [{"header": [1, 2, 3, False]}, {"size": ["small", False, "large", "huge"]}],
    ["bold", "italic", "underline", "strike", {"script": "sub"}, {"script": "super"}],
    [{"color": []}, {"background": []}, {"font": []}],
    [{"list": "ordered"}, {"list": "bullet"}, {"indent": "-1"}, {"indent": "+1"}, {"align": []}],
    ["blockquote", "link", "image", "clean"],
]

# O Quill aplica header/size/font/color/background/align/indent via classes CSS
# (ql-size-large, ql-font-serif, ...) definidas em quill.core.css, que só existe
# dentro do editor. Sem isso embutido no HTMLBody, essas formatações desaparecem
# no rascunho do Outlook. Replicamos aqui as mesmas regras, escopadas ao corpo do e-mail.
CSS_QUILL_EMAIL = """
<style>
.alfredo-corpo .ql-font-serif { font-family: Georgia, 'Times New Roman', serif; }
.alfredo-corpo .ql-font-monospace { font-family: Monaco, 'Courier New', monospace; }
.alfredo-corpo .ql-size-small { font-size: 0.75em; }
.alfredo-corpo .ql-size-large { font-size: 1.5em; }
.alfredo-corpo .ql-size-huge { font-size: 2.5em; }
.alfredo-corpo .ql-align-center { text-align: center; }
.alfredo-corpo .ql-align-right { text-align: right; }
.alfredo-corpo .ql-align-justify { text-align: justify; }
.alfredo-corpo .ql-indent-1 { padding-left: 3em; }
.alfredo-corpo .ql-indent-2 { padding-left: 6em; }
.alfredo-corpo .ql-indent-3 { padding-left: 9em; }
.alfredo-corpo .ql-color-white { color: #fff; }
.alfredo-corpo .ql-color-red { color: #e60000; }
.alfredo-corpo .ql-color-orange { color: #f90; }
.alfredo-corpo .ql-color-yellow { color: #ff0; }
.alfredo-corpo .ql-color-green { color: #008a00; }
.alfredo-corpo .ql-color-blue { color: #06c; }
.alfredo-corpo .ql-color-purple { color: #93f; }
.alfredo-corpo .ql-bg-black { background-color: #000; }
.alfredo-corpo .ql-bg-red { background-color: #e60000; }
.alfredo-corpo .ql-bg-orange { background-color: #f90; }
.alfredo-corpo .ql-bg-yellow { background-color: #ff0; }
.alfredo-corpo .ql-bg-green { background-color: #008a00; }
.alfredo-corpo .ql-bg-blue { background-color: #06c; }
.alfredo-corpo .ql-bg-purple { background-color: #93f; }
</style>
"""


def valor_celula(row, coluna):
    if coluna and pd.notna(row[coluna]):
        return str(row[coluna]).strip()
    return ""


def email_valido(email):
    return bool(email and "@" in email and "." in email.split("@")[-1])


def personalizar_texto(texto, row, colunas):
    texto_final = texto or ""
    for coluna in colunas:
        tag = "{" + coluna + "}"
        valor = str(row[coluna]) if pd.notna(row[coluna]) else ""
        texto_final = texto_final.replace(tag, valor)
    return texto_final


def montar_corpo_html(corpo):
    return (
        f"{CSS_QUILL_EMAIL}"
        f"<div class='alfredo-corpo' style='font-family: Calibri; font-size: 11pt;'>{corpo or ''}</div>"
    )


def obter_caminho_anexo(caminho_anexos, nome_arquivo):
    nome_limpo = str(nome_arquivo).strip()
    caminho = os.path.join(caminho_anexos, nome_limpo)
    return nome_limpo, caminho


def anexar_arquivo_outlook(mail, caminho_anexos, nome_arquivo):
    nome_limpo, caminho = obter_caminho_anexo(caminho_anexos, nome_arquivo)
    if os.path.exists(caminho):
        mail.Attachments.Add(os.path.abspath(caminho))
        return None
    return f"Arquivo não encontrado: {nome_limpo}"


def criar_rascunho_outlook(outlook, dados_email):
    mail = outlook.CreateItem(0)
    mail.Display()

    mail.To = dados_email["para"]
    mail.CC = dados_email["cc"]
    mail.BCC = dados_email["bcc"]
    mail.Subject = dados_email["assunto"]
    mail.HTMLBody = f"{dados_email['corpo_html']}<br>{mail.HTMLBody}"

    for anexo in dados_email["anexos"]:
        erro_anexo = anexar_arquivo_outlook(mail, anexo["pasta"], anexo["nome"])
        if erro_anexo:
            dados_email["erros_anexo"].append(erro_anexo)

    mail.Save()
    mail.Close(0)


def validar_configuracao_smtp(servidor, porta, usuario, senha, remetente):
    erros = []

    if not servidor.strip():
        erros.append("Informe o servidor SMTP.")
    if porta <= 0:
        erros.append("Informe uma porta SMTP válida.")
    if not usuario.strip():
        erros.append("Informe o usuário SMTP.")
    if not senha:
        erros.append("Informe a senha SMTP.")
    if not email_valido(remetente):
        erros.append("Informe um remetente válido.")

    return erros


def testar_conexao_smtp(servidor, porta, usuario, senha, usar_starttls):
    with smtplib.SMTP(servidor, porta, timeout=20) as smtp:
        smtp.ehlo()
        if usar_starttls:
            contexto = ssl.create_default_context()
            smtp.starttls(context=contexto)
            smtp.ehlo()
        if usuario or senha:
            smtp.login(usuario, senha)


def montar_email_mime_smtp(dados_email, remetente):
    msg = MIMEMultipart("mixed")
    msg["From"] = remetente
    msg["To"] = dados_email["para"]

    if dados_email["cc"]:
        msg["Cc"] = dados_email["cc"]
    if dados_email["bcc"]:
        msg["Bcc"] = dados_email["bcc"]

    msg["Subject"] = dados_email["assunto"]

    corpo = MIMEMultipart("alternative")
    corpo.attach(MIMEText(dados_email["corpo_html"], "html", "utf-8"))
    msg.attach(corpo)

    for anexo in dados_email["anexos"]:
        nome_limpo, caminho = obter_caminho_anexo(anexo["pasta"], anexo["nome"])
        if not os.path.exists(caminho):
            dados_email["erros_anexo"].append(f"Arquivo não encontrado: {nome_limpo}")
            continue

        with open(caminho, "rb") as arquivo:
            parte = MIMEBase("application", "octet-stream")
            parte.set_payload(arquivo.read())

        encoders.encode_base64(parte)
        parte.add_header("Content-Disposition", f'attachment; filename="{nome_limpo}"')
        msg.attach(parte)

    return msg


def montar_dados_email(row, colunas, col_email, col_cc, col_bcc, col_arq, subject, content, caminho_anexos):
    para = valor_celula(row, col_email)
    cc = valor_celula(row, col_cc)
    bcc = valor_celula(row, col_bcc)

    assunto = personalizar_texto(subject, row, colunas)
    corpo = personalizar_texto(content, row, colunas)

    anexos = []
    if caminho_anexos and col_arq and pd.notna(row[col_arq]):
        anexos.append({"pasta": caminho_anexos, "nome": row[col_arq]})

    return {
        "para": para,
        "cc": cc,
        "bcc": bcc,
        "assunto": assunto,
        "corpo_html": montar_corpo_html(corpo),
        "anexos": anexos,
        "erros_anexo": [],
    }


def mostrar_avisos(titulo, avisos):
    if avisos:
        with st.expander(titulo):
            for aviso in avisos:
                st.warning(aviso)


# --- MINI MANUAL ATUALIZADO ---
with st.expander("📖 MANUAL: Como selecionar a pasta de anexos"):
    st.markdown(
        """
    1. **Abra a pasta** onde estão seus arquivos no Windows Explorer.
    2. Clique na **barra de endereços** no topo da pasta (onde aparece o caminho).
    3. Copie o texto (ex: `C:\\MeusDocumentos\\Certificados`).
    4. **Cole** esse texto no campo 'Caminho da pasta' na barra lateral do Alfredo.
    5. O programa removerá as aspas automaticamente se houver.
    """
    )

st.title("✉️ Alfredo do email")

# --- SIDEBAR ---
with st.sidebar:
    st.header("1. Configurações")
    uploaded_file = st.file_uploader("Suba sua Planilha", type=["xlsx", "csv"])

    modo_envio = st.radio(
        "Modo de saída",
        ["Outlook Desktop", "SMTP"],
        help="Outlook cria rascunhos no aplicativo desktop. SMTP prepara mensagens MIME e valida a conexão, sem enviar e-mails nesta versão.",
    )

    caminho_input = st.text_input("Caminho da pasta de anexos (Copie e cole aqui)")
    caminho_anexos = caminho_input.replace('"', "").strip()

    smtp_config = {}
    if modo_envio == "SMTP":
        st.divider()
        st.subheader("SMTP")
        smtp_config["servidor"] = st.text_input("Servidor SMTP", placeholder="smtp.seudominio.com")
        smtp_config["porta"] = st.number_input("Porta", min_value=1, max_value=65535, value=587, step=1)
        smtp_config["usuario"] = st.text_input("Usuário")
        smtp_config["senha"] = st.text_input("Senha", type="password")
        smtp_config["remetente"] = st.text_input("Remetente", placeholder="voce@seudominio.com")
        smtp_config["usar_starttls"] = st.checkbox("Usar STARTTLS", value=True)
        smtp_config["testar_conexao"] = st.checkbox("Testar conexão/autenticação SMTP ao preparar", value=False)

        st.info("O modo SMTP prepara e valida as mensagens, mas não envia e-mails automaticamente.")

if uploaded_file:
    if uploaded_file.name.endswith("xlsx"):
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
        toolbar=TOOLBAR_QUILL,
        key="quill_editor",
    )

    if modo_envio == "Outlook Desktop":
        texto_teste = "🧪 Gerar 1 TESTE"
        texto_massa = "🚀 Gerar TODOS"
    else:
        texto_teste = "🧪 Preparar 1 TESTE"
        texto_massa = "🚀 Preparar TODOS"

    col_t1, col_t2 = st.columns(2)
    with col_t1:
        btn_teste = st.button(texto_teste, use_container_width=True)
    with col_t2:
        btn_massa = st.button(texto_massa, type="primary", use_container_width=True)

    if btn_teste or btn_massa:
        df_proc = df.head(1) if btn_teste else df
        total_linhas = len(df_proc)

        if modo_envio == "SMTP":
            erros_config = validar_configuracao_smtp(
                smtp_config["servidor"],
                smtp_config["porta"],
                smtp_config["usuario"],
                smtp_config["senha"],
                smtp_config["remetente"],
            )

            if erros_config:
                st.error("Revise a configuração SMTP antes de preparar as mensagens.")
                mostrar_avisos("⚠️ Ver problemas de configuração SMTP", erros_config)
                st.stop()

            if smtp_config["testar_conexao"]:
                try:
                    testar_conexao_smtp(
                        smtp_config["servidor"],
                        smtp_config["porta"],
                        smtp_config["usuario"],
                        smtp_config["senha"],
                        smtp_config["usar_starttls"],
                    )
                    st.success("Conexão SMTP validada com sucesso.")
                except Exception as e:
                    st.error(f"Erro ao validar conexão SMTP: {e}")
                    st.stop()

        try:
            outlook = None
            if modo_envio == "Outlook Desktop":
                pythoncom.CoInitialize()
                try:
                    outlook = win32.GetActiveObject("Outlook.Application")
                except Exception:
                    outlook = win32.Dispatch("outlook.application")

            sucesso = 0
            mensagens_smtp = []
            erros_anexo = []
            avisos_destinatario = []
            progress = st.progress(0)

            for posicao, (index, row) in enumerate(df_proc.iterrows(), start=1):
                dados_email = montar_dados_email(
                    row,
                    colunas,
                    col_email,
                    col_cc,
                    col_bcc,
                    col_arq,
                    subject,
                    content,
                    caminho_anexos,
                )

                if not email_valido(dados_email["para"]):
                    avisos_destinatario.append(f"Linha {index + 1}: destinatário inválido ou vazio.")
                    progress.progress(posicao / total_linhas)
                    continue

                if modo_envio == "Outlook Desktop":
                    criar_rascunho_outlook(outlook, dados_email)
                else:
                    mensagem = montar_email_mime_smtp(dados_email, smtp_config["remetente"])
                    mensagens_smtp.append(mensagem)

                erros_anexo.extend(dados_email["erros_anexo"])
                sucesso += 1
                progress.progress(posicao / total_linhas)

            if modo_envio == "Outlook Desktop":
                st.success(f"✅ Finalizado! {sucesso} rascunhos criados no Outlook.")
            else:
                st.success(f"✅ Finalizado! {sucesso} mensagens SMTP preparadas. Nenhum e-mail foi enviado.")
                if mensagens_smtp:
                    primeira = mensagens_smtp[0]
                    with st.expander("👀 Prévia técnica da primeira mensagem SMTP"):
                        st.write("**De:**", primeira.get("From", ""))
                        st.write("**Para:**", primeira.get("To", ""))
                        st.write("**CC:**", primeira.get("Cc", ""))
                        st.write("**BCC:**", primeira.get("Bcc", ""))
                        st.write("**Assunto:**", primeira.get("Subject", ""))
                        st.code(primeira.as_string()[:5000], language="eml")

            mostrar_avisos("⚠️ Ver destinatários ignorados", avisos_destinatario)
            mostrar_avisos("⚠️ Ver arquivos não localizados", erros_anexo)

        except Exception as e:
            st.error(f"Erro no processamento: {e}")
else:
    st.info("Aguardando planilha...")
