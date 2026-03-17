import streamlit as st
import pandas as pd
import win32com.client as win32
import pythoncom
import os
from streamlit_quill import st_quill

st.set_page_config(
    page_title="Alfredo do Email",
    page_icon="✉️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- ESTILOS CUSTOMIZADOS ---
st.markdown("""
<style>
    /* Cabeçalho principal */
    .bloco-titulo {
        background: linear-gradient(135deg, #1a73e8 0%, #0d47a1 100%);
        padding: 2rem 2.5rem;
        border-radius: 12px;
        color: white;
        margin-bottom: 1.5rem;
    }
    .bloco-titulo h1 { color: white; margin: 0; font-size: 2rem; }
    .bloco-titulo p  { color: rgba(255,255,255,0.85); margin: 0.4rem 0 0; font-size: 1rem; }

    /* Badges de etapa */
    .etapa-badge {
        display: inline-flex;
        align-items: center;
        gap: 0.5rem;
        background: #e8f0fe;
        color: #1a73e8;
        font-weight: 700;
        font-size: 0.85rem;
        padding: 0.3rem 0.8rem;
        border-radius: 20px;
        margin-bottom: 0.6rem;
    }
    .etapa-badge.concluida { background: #e6f4ea; color: #137333; }

    /* Cards de seção */
    .card-secao {
        background: #ffffff;
        border: 1px solid #e0e0e0;
        border-radius: 10px;
        padding: 1.5rem;
        margin-bottom: 1.2rem;
        box-shadow: 0 1px 4px rgba(0,0,0,0.07);
    }

    /* Dica de tags */
    .dica-tags {
        background: #fff8e1;
        border-left: 4px solid #f9a825;
        border-radius: 4px;
        padding: 0.6rem 1rem;
        margin-top: 0.8rem;
        font-size: 0.88rem;
    }

    /* Botão primário customizado */
    div[data-testid="stButton"] button[kind="primary"] {
        background: linear-gradient(135deg, #1a73e8, #0d47a1);
        border: none;
        font-size: 1rem;
        font-weight: 600;
        padding: 0.6rem 1.2rem;
        border-radius: 8px;
    }

    /* Barra lateral */
    section[data-testid="stSidebar"] {
        background: #f8f9fa;
    }
    section[data-testid="stSidebar"] h2 {
        color: #1a73e8;
    }
    .separador-lateral { border-top: 1px solid #dee2e6; margin: 1rem 0; }
</style>
""", unsafe_allow_html=True)


# ── CABEÇALHO ──────────────────────────────────────────────────────────────────
st.markdown("""
<div class="bloco-titulo">
    <h1>✉️ Alfredo do Email</h1>
    <p>Automatize o envio de e-mails personalizados pelo Outlook a partir de uma planilha Excel ou CSV.</p>
</div>
""", unsafe_allow_html=True)


# ── BARRA LATERAL ──────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Configuração")
    st.caption("Preencha os campos abaixo para começar.")
    st.markdown('<div class="separador-lateral"></div>', unsafe_allow_html=True)

    # Etapa 1 – Planilha
    st.markdown("**📂 Etapa 1 – Planilha de destinatários**")
    arquivo_carregado = st.file_uploader(
        "Arraste ou clique para enviar",
        type=["xlsx", "csv"],
        help="Formatos aceitos: .xlsx e .csv",
    )

    st.markdown('<div class="separador-lateral"></div>', unsafe_allow_html=True)

    # Etapa 2 – Pasta de anexos
    st.markdown("**📎 Etapa 2 – Pasta de anexos** *(opcional)*")
    caminho_input = st.text_input(
        "Cole o caminho da pasta aqui",
        placeholder=r"Ex.: C:\Documentos\Certificados",
        help="Abra a pasta no Windows Explorer, clique na barra de endereços e copie o caminho completo.",
    )
    caminho_anexos = caminho_input.replace('"', '').strip()

    if caminho_input and not os.path.isdir(caminho_anexos):
        st.warning("⚠️ Pasta não encontrada. Verifique o caminho informado.")

    st.markdown('<div class="separador-lateral"></div>', unsafe_allow_html=True)

    # Mini-manual recolhível
    with st.expander("📖 Como copiar o caminho da pasta?"):
        st.markdown("""
1. **Abra** a pasta no Windows Explorer.
2. Clique na **barra de endereços** (onde aparece o caminho).
3. **Copie** o texto — ex.: `C:\\MeusDocumentos\\Certificados`.
4. **Cole** no campo acima.
> O programa remove aspas extras automaticamente.
        """)

    st.markdown('<div class="separador-lateral"></div>', unsafe_allow_html=True)
    st.caption("Alfredo do E-mail · v2.0")


# ── CONTEÚDO PRINCIPAL ─────────────────────────────────────────────────────────
if not arquivo_carregado:
    # Estado vazio – orientação visual clara
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.markdown("""
<div class="card-secao" style="text-align:center;">
    <div style="font-size:2.5rem;">📂</div>
    <strong>Etapa 1</strong><br>
    Faça upload da sua planilha na barra lateral.
</div>""", unsafe_allow_html=True)
    with col_b:
        st.markdown("""
<div class="card-secao" style="text-align:center;">
    <div style="font-size:2.5rem;">✏️</div>
    <strong>Etapa 2</strong><br>
    Mapeie as colunas e redija seu e-mail com <code>{Tags}</code>.
</div>""", unsafe_allow_html=True)
    with col_c:
        st.markdown("""
<div class="card-secao" style="text-align:center;">
    <div style="font-size:2.5rem;">🚀</div>
    <strong>Etapa 3</strong><br>
    Clique em <em>Gerar TODOS</em> e os rascunhos aparecerão no Outlook.
</div>""", unsafe_allow_html=True)

    st.info("👆 Comece enviando sua planilha na **barra lateral** à esquerda.")
    st.stop()

# ── LEITURA DA PLANILHA ────────────────────────────────────────────────────────
if arquivo_carregado.name.endswith("xlsx"):
    arquivo_excel = pd.ExcelFile(arquivo_carregado)
    if len(arquivo_excel.sheet_names) > 1:
        aba = st.selectbox(
            "📑 Selecione a aba da planilha",
            arquivo_excel.sheet_names,
            help="Escolha qual aba contém os dados de destinatários.",
        )
    else:
        aba = arquivo_excel.sheet_names[0]
    df = pd.read_excel(arquivo_carregado, sheet_name=aba)
else:
    df = pd.read_csv(arquivo_carregado)

colunas = df.columns.tolist()
total_registros = len(df)

# ── ETAPA 1 – VISUALIZAR DADOS ─────────────────────────────────────────────────
st.markdown('<div class="etapa-badge concluida">✅ Planilha carregada</div>', unsafe_allow_html=True)
with st.expander(f"📋 Ver dados da planilha  —  {total_registros} linha(s) encontrada(s)", expanded=False):
    st.dataframe(df.head(5), use_container_width=True)

st.divider()

# ── ETAPA 2 – MAPEAMENTO DE COLUNAS ───────────────────────────────────────────
st.markdown('<div class="etapa-badge">📌 Etapa 2 — Mapeie as colunas</div>', unsafe_allow_html=True)
st.markdown("""
<div class="card-secao">
""", unsafe_allow_html=True)

c1, c2, c3, c4 = st.columns(4)
with c1:
    col_email = st.selectbox(
        "📧 Para (obrigatório)",
        colunas,
        help="Coluna que contém os endereços de e-mail dos destinatários.",
    )
with c2:
    col_cc = st.selectbox(
        "👁️ Cópia (CC)",
        [None] + colunas,
        help="Coluna com os e-mails em cópia. Deixe em branco se não houver.",
    )
with c3:
    col_bcc = st.selectbox(
        "🔒 Cópia Oculta (BCC)",
        [None] + colunas,
        help="Coluna com e-mails em cópia oculta. Os destinatários não se verão.",
    )
with c4:
    col_arq = st.selectbox(
        "📎 Anexo",
        [None] + colunas,
        help="Coluna com o nome do arquivo a anexar (ex.: certificado_joao.pdf). Configure também a pasta na barra lateral.",
    )

st.markdown("</div>", unsafe_allow_html=True)

# Validação rápida da coluna de e-mail
emails_invalidos = df[col_email].apply(lambda x: "@" not in str(x)).sum()
if emails_invalidos:
    st.warning(f"⚠️ {emails_invalidos} linha(s) sem '@' na coluna **{col_email}** serão ignoradas.")

st.divider()

# ── ETAPA 3 – REDIGIR MENSAGEM ─────────────────────────────────────────────────
st.markdown('<div class="etapa-badge">✏️ Etapa 3 — Redija a mensagem</div>', unsafe_allow_html=True)

assunto = st.text_input(
    "📌 Assunto do e-mail",
    placeholder="Ex.: Olá {Nome}, seu certificado está disponível!",
    help="Use {NomeDaColuna} para personalizar o assunto de cada destinatário.",
)

st.markdown("""
<div class="dica-tags">
    💡 <strong>Dica de Tags:</strong> use <code>{NomeDaColuna}</code> no assunto e no corpo para inserir dados da planilha em cada e-mail.
    Colunas disponíveis: <strong>{colunas}</strong>
</div>
""".format(colunas="</strong>, <strong>".join(["{" + c + "}" for c in colunas])), unsafe_allow_html=True)

st.markdown("")
conteudo = st_quill(
    placeholder="Escreva seu e-mail aqui… Ex.: Olá {Nome}, segue em anexo seu documento.",
    html=True,
    key="quill_editor",
)

# Pré-visualização da primeira linha
if assunto or conteudo:
    with st.expander("👁️ Pré-visualização com a 1ª linha da planilha", expanded=False):
        primeira_linha = df.iloc[0]
        assunto_prev = assunto
        corpo_prev = conteudo or ""
        for col in colunas:
            marcador = "{" + col + "}"
            valor = str(primeira_linha[col]) if pd.notna(primeira_linha[col]) else ""
            assunto_prev = assunto_prev.replace(marcador, valor)
            corpo_prev = corpo_prev.replace(marcador, valor)
        st.markdown(f"**Para:** `{str(primeira_linha[col_email]).replace('`', '')}`")
        st.markdown(f"**Assunto:** {assunto_prev}")
        st.markdown("**Corpo:**")
        st.markdown(corpo_prev, unsafe_allow_html=True)

st.divider()

# ── ETAPA 4 – GERAR RASCUNHOS ─────────────────────────────────────────────────
st.markdown('<div class="etapa-badge">🚀 Etapa 4 — Gerar rascunhos no Outlook</div>', unsafe_allow_html=True)

validos = total_registros - emails_invalidos
col_t1, col_t2, col_espacador = st.columns([1, 1, 2])
with col_t1:
    btn_teste = st.button(
        "🧪 Gerar 1 Teste",
        use_container_width=True,
        help="Cria apenas o rascunho da 1ª linha para você revisar antes de gerar todos.",
    )
with col_t2:
    btn_massa = st.button(
        f"🚀 Gerar TODOS  ({validos})",
        type="primary",
        use_container_width=True,
        help=f"Cria {validos} rascunho(s) no Outlook de uma vez.",
    )

if btn_teste or btn_massa:
    if not assunto:
        st.error("❌ Preencha o **assunto** do e-mail antes de continuar.")
        st.stop()
    if not conteudo:
        st.error("❌ O **corpo** do e-mail está vazio. Escreva sua mensagem antes de continuar.")
        st.stop()

    df_processado = df.head(1) if btn_teste else df
    quantidade = len(df_processado)

    if quantidade == 0:
        st.warning("⚠️ Nenhuma linha válida para processar.")
        st.stop()

    try:
        pythoncom.CoInitialize()
        try:
            outlook = win32.GetActiveObject("Outlook.Application")
        except Exception:
            outlook = win32.Dispatch("outlook.application")

        sucesso = 0
        erros_anexo = []

        with st.status(f"⏳ Gerando rascunhos… 0 / {quantidade}", expanded=True) as status_geracao:
            barra_progresso = st.progress(0)

            for indice, linha in df_processado.iterrows():
                destinatario = str(linha[col_email]).strip()
                if "@" not in destinatario:
                    continue

                rascunho = outlook.CreateItem(0)
                rascunho.Display()  # Abre para carregar a assinatura padrão

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
                rascunho.HTMLBody = (
                    f"<div style='font-family: Calibri; font-size: 11pt;'>{corpo_formatado}</div><br>"
                    + rascunho.HTMLBody
                )

                # Lógica de Anexo
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

                progresso = sucesso / quantidade
                barra_progresso.progress(progresso)
                status_geracao.update(label=f"⏳ Gerando rascunhos… {sucesso} / {quantidade}")

            status_geracao.update(label=f"✅ Concluído! {sucesso} rascunho(s) criado(s).", state="complete")

        st.success(f"✅ **Pronto!** {sucesso} rascunho(s) criado(s) na pasta **Rascunhos** do Outlook.")

        if erros_anexo:
            with st.expander(f"⚠️ {len(erros_anexo)} arquivo(s) não localizado(s) — clique para ver"):
                for mensagem_erro in erros_anexo:
                    st.warning(mensagem_erro)

    except Exception as erro:
        st.error(f"❌ Erro ao conectar com o Outlook: {erro}")
        st.info("💡 Certifique-se de que o Microsoft Outlook está instalado e configurado neste computador.")
    finally:
        pythoncom.CoUninitialize()