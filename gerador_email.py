"""
Alfredo do Email - Sistema de Envio de Emails em Massa via Outlook
Autor: Sistema automatizado de envio de emails personalizados
Versão: 2.0
"""

import streamlit as st
import pandas as pd
import win32com.client as win32
import pythoncom
import os
import re
import logging
from typing import Optional, Tuple, List, Dict, Any
from pathlib import Path
from streamlit_quill import st_quill

# Configuração de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('email_generator.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Configurações da aplicação
st.set_page_config(page_title="Alfredo do email", layout="wide")


def validate_email(email: str) -> bool:
    """
    Valida se uma string é um endereço de e-mail válido.

    Args:
        email: String contendo o endereço de e-mail a ser validado

    Returns:
        bool: True se o e-mail é válido, False caso contrário
    """
    if not email or not isinstance(email, str):
        return False

    email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(email_pattern, email.strip()))


def sanitize_path(path: str) -> Optional[str]:
    """
    Sanitiza e valida um caminho de arquivo/pasta.

    Args:
        path: Caminho a ser sanitizado

    Returns:
        str: Caminho sanitizado ou None se inválido
    """
    if not path:
        return None

    # Remove aspas e espaços extras
    clean_path = path.replace('"', '').replace("'", '').strip()

    # Verifica se o caminho existe
    if not os.path.exists(clean_path):
        logger.warning(f"Caminho não existe: {clean_path}")
        return None

    return clean_path


def validate_attachment(file_path: str, allowed_extensions: Optional[List[str]] = None) -> bool:
    """
    Valida se um arquivo de anexo é seguro e existe.

    Args:
        file_path: Caminho completo do arquivo
        allowed_extensions: Lista de extensões permitidas (None = todas)

    Returns:
        bool: True se o arquivo é válido, False caso contrário
    """
    if not os.path.exists(file_path):
        return False

    if not os.path.isfile(file_path):
        return False

    if allowed_extensions:
        file_ext = Path(file_path).suffix.lower()
        if file_ext not in allowed_extensions:
            logger.warning(f"Extensão não permitida: {file_ext}")
            return False

    return True


def load_dataframe(uploaded_file) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    """
    Carrega um DataFrame de um arquivo Excel ou CSV.

    Args:
        uploaded_file: Arquivo enviado pelo Streamlit

    Returns:
        Tuple: (DataFrame ou None, mensagem de erro ou None)
    """
    try:
        if uploaded_file.name.endswith('.xlsx'):
            excel_file = pd.ExcelFile(uploaded_file)

            # Se houver apenas uma aba, carrega automaticamente
            if len(excel_file.sheet_names) == 1:
                df = pd.read_excel(uploaded_file, sheet_name=0)
                logger.info(f"Arquivo Excel carregado: {uploaded_file.name}")
                return df, None
            else:
                # Retorna None para que o usuário selecione a aba
                return None, None
        else:
            df = pd.read_csv(uploaded_file, encoding='utf-8')
            logger.info(f"Arquivo CSV carregado: {uploaded_file.name}")
            return df, None

    except Exception as e:
        error_msg = f"Erro ao carregar arquivo: {str(e)}"
        logger.error(error_msg)
        return None, error_msg


def get_outlook_instance() -> Optional[win32.CDispatch]:
    """
    Obtém uma instância do Outlook.

    Returns:
        Instância do Outlook ou None se houver erro
    """
    try:
        pythoncom.CoInitialize()
        try:
            outlook = win32.GetActiveObject("Outlook.Application")
            logger.info("Conectado ao Outlook ativo")
        except:
            outlook = win32.Dispatch('outlook.application')
            logger.info("Nova instância do Outlook criada")
        return outlook
    except Exception as e:
        logger.error(f"Erro ao conectar ao Outlook: {e}")
        return None


def replace_tags(text: str, row: pd.Series, columns: List[str]) -> str:
    """
    Substitui tags {NomeColuna} pelos valores da linha.

    Args:
        text: Texto com tags a serem substituídas
        row: Linha do DataFrame com os valores
        columns: Lista de colunas disponíveis

    Returns:
        str: Texto com tags substituídas
    """
    result = text
    for col in columns:
        tag = "{" + col + "}"
        value = str(row[col]) if pd.notna(row[col]) else ""
        result = result.replace(tag, value)
    return result


def create_email_draft(
    outlook: win32.CDispatch,
    email_to: str,
    subject: str,
    body_html: str,
    email_cc: Optional[str] = None,
    email_bcc: Optional[str] = None,
    attachment_path: Optional[str] = None
) -> Tuple[bool, Optional[str]]:
    """
    Cria um rascunho de e-mail no Outlook.

    Args:
        outlook: Instância do Outlook
        email_to: E-mail do destinatário
        subject: Assunto do e-mail
        body_html: Corpo do e-mail em HTML
        email_cc: E-mail para cópia (opcional)
        email_bcc: E-mail para cópia oculta (opcional)
        attachment_path: Caminho do anexo (opcional)

    Returns:
        Tuple: (sucesso, mensagem de erro ou None)
    """
    try:
        # Valida e-mail principal
        if not validate_email(email_to):
            error_msg = f"E-mail inválido: {email_to}"
            logger.warning(error_msg)
            return False, error_msg

        # Cria o item de e-mail
        mail = outlook.CreateItem(0)
        mail.Display()  # Abre para carregar a assinatura padrão

        # Define destinatários
        mail.To = email_to

        if email_cc and validate_email(email_cc):
            mail.CC = email_cc

        if email_bcc and validate_email(email_bcc):
            mail.BCC = email_bcc

        # Define assunto e corpo
        mail.Subject = subject
        mail.HTMLBody = f"<div style='font-family: Calibri; font-size: 11pt;'>{body_html}</div><br>" + mail.HTMLBody

        # Adiciona anexo se existir
        if attachment_path and os.path.exists(attachment_path):
            if validate_attachment(attachment_path):
                mail.Attachments.Add(os.path.abspath(attachment_path))
                logger.info(f"Anexo adicionado: {attachment_path}")
            else:
                error_msg = f"Anexo inválido: {attachment_path}"
                logger.warning(error_msg)
                return False, error_msg

        # Salva o rascunho
        mail.Save()
        mail.Close(0)

        logger.info(f"Rascunho criado com sucesso para: {email_to}")
        return True, None

    except Exception as e:
        error_msg = f"Erro ao criar rascunho: {str(e)}"
        logger.error(error_msg)
        return False, error_msg


def process_emails(
    df: pd.DataFrame,
    outlook: win32.CDispatch,
    col_email: str,
    subject_template: str,
    body_template: str,
    columns: List[str],
    col_cc: Optional[str] = None,
    col_bcc: Optional[str] = None,
    col_attachment: Optional[str] = None,
    attachment_folder: Optional[str] = None,
    progress_bar = None
) -> Dict[str, Any]:
    """
    Processa múltiplos e-mails a partir de um DataFrame.

    Args:
        df: DataFrame com os dados dos e-mails
        outlook: Instância do Outlook
        col_email: Nome da coluna com e-mails dos destinatários
        subject_template: Template do assunto com tags
        body_template: Template do corpo com tags
        columns: Lista de colunas disponíveis
        col_cc: Nome da coluna com e-mails CC (opcional)
        col_bcc: Nome da coluna com e-mails BCC (opcional)
        col_attachment: Nome da coluna com nomes de anexos (opcional)
        attachment_folder: Pasta com os arquivos de anexo (opcional)
        progress_bar: Barra de progresso do Streamlit (opcional)

    Returns:
        Dict com estatísticas do processamento
    """
    stats = {
        'success_count': 0,
        'error_count': 0,
        'errors': [],
        'attachment_errors': []
    }

    total_rows = len(df)

    for index, row in df.iterrows():
        # Obtém e valida e-mail principal
        email_to = str(row[col_email]).strip()

        if not validate_email(email_to):
            error_msg = f"Linha {index + 1}: E-mail inválido - {email_to}"
            stats['errors'].append(error_msg)
            stats['error_count'] += 1
            logger.warning(error_msg)
            continue

        # Substitui tags no assunto e corpo
        subject = replace_tags(subject_template, row, columns)
        body = replace_tags(body_template, row, columns)

        # Obtém CC e BCC se disponíveis
        email_cc = None
        if col_cc and pd.notna(row[col_cc]):
            email_cc = str(row[col_cc]).strip()

        email_bcc = None
        if col_bcc and pd.notna(row[col_bcc]):
            email_bcc = str(row[col_bcc]).strip()

        # Prepara anexo se disponível
        attachment_path = None
        if attachment_folder and col_attachment and pd.notna(row[col_attachment]):
            filename = str(row[col_attachment]).strip()
            attachment_path = os.path.join(attachment_folder, filename)

            if not os.path.exists(attachment_path):
                error_msg = f"Arquivo não encontrado: {filename}"
                stats['attachment_errors'].append(error_msg)
                logger.warning(error_msg)
                attachment_path = None

        # Cria o rascunho
        success, error = create_email_draft(
            outlook=outlook,
            email_to=email_to,
            subject=subject,
            body_html=body,
            email_cc=email_cc,
            email_bcc=email_bcc,
            attachment_path=attachment_path
        )

        if success:
            stats['success_count'] += 1
        else:
            stats['error_count'] += 1
            stats['errors'].append(f"Linha {index + 1}: {error}")

        # Atualiza barra de progresso
        if progress_bar:
            progress_bar.progress((index + 1) / total_rows)

    return stats


def render_manual():
    """Renderiza o manual de uso na interface."""
    with st.expander("📖 MANUAL: Como usar o Alfredo do Email"):
        st.markdown("""
        ### Como selecionar a pasta de anexos
        1. **Abra a pasta** onde estão seus arquivos no Windows Explorer.
        2. Clique na **barra de endereços** no topo da pasta (onde aparece o caminho).
        3. Copie o texto (ex: `C:\\MeusDocumentos\\Certificados`).
        4. **Cole** esse texto no campo 'Caminho da pasta' na barra lateral do Alfredo.
        5. O programa removerá as aspas automaticamente se houver.

        ### Como usar tags personalizadas
        - Use `{NomeDaColuna}` no assunto ou corpo do e-mail
        - As tags serão substituídas pelos valores da planilha
        - Exemplo: "Olá {Nome}, segue seu documento {TipoDocumento}"

        ### Modo de teste
        - Use o botão "🧪 Gerar 1 TESTE" para criar apenas o primeiro e-mail
        - Verifique o rascunho no Outlook antes de gerar todos
        - Isso evita erros em envios em massa
        """)


def render_sidebar() -> Tuple[Optional[pd.DataFrame], Dict[str, Any]]:
    """
    Renderiza a barra lateral e retorna configurações.

    Returns:
        Tuple: (DataFrame ou None, dicionário de configurações)
    """
    config = {
        'uploaded_file': None,
        'attachment_folder': None,
        'df': None,
        'columns': [],
        'col_email': None,
        'col_cc': None,
        'col_bcc': None,
        'col_attachment': None,
        'aba': None
    }

    with st.sidebar:
        st.header("1. Configurações")

        # Upload de arquivo
        config['uploaded_file'] = st.file_uploader(
            "Suba sua Planilha",
            type=["xlsx", "csv"],
            help="Envie um arquivo Excel (.xlsx) ou CSV"
        )

        # Caminho de anexos
        caminho_input = st.text_input(
            "Caminho da pasta de anexos (Copie e cole aqui)",
            help="Cole o caminho completo da pasta com os arquivos"
        )
        config['attachment_folder'] = sanitize_path(caminho_input)

        if config['attachment_folder']:
            st.success(f"✅ Pasta válida: {config['attachment_folder']}")
        elif caminho_input:
            st.error("❌ Caminho inválido ou pasta não encontrada")

    return config


def main():
    """Função principal da aplicação."""

    # Renderiza manual
    render_manual()

    # Título
    st.title("✉️ Alfredo do Email")
    st.markdown("*Sistema inteligente de envio de e-mails personalizados via Outlook*")

    # Renderiza barra lateral e obtém configurações
    config = render_sidebar()

    if not config['uploaded_file']:
        st.info("👆 Aguardando upload da planilha...")
        st.markdown("""
        ### 🚀 Como começar:
        1. Faça upload de uma planilha Excel ou CSV
        2. Configure as colunas de destinatários
        3. Digite o assunto e o corpo do e-mail
        4. Clique em "Gerar TODOS" para criar os rascunhos no Outlook
        """)
        return

    # Carrega DataFrame
    df = None
    if config['uploaded_file'].name.endswith('.xlsx'):
        excel_file = pd.ExcelFile(config['uploaded_file'])
        if len(excel_file.sheet_names) > 1:
            aba = st.selectbox("Selecione a Aba", excel_file.sheet_names)
            df = pd.read_excel(config['uploaded_file'], sheet_name=aba)
        else:
            df = pd.read_excel(config['uploaded_file'], sheet_name=0)
    else:
        df, error = load_dataframe(config['uploaded_file'])
        if error:
            st.error(error)
            return

    if df is None or df.empty:
        st.error("❌ Erro ao carregar planilha ou planilha vazia")
        return

    # Mostra preview dos dados
    st.write("### 📋 Preview dos Dados")
    st.dataframe(df.head(3), use_container_width=True)
    st.info(f"📊 Total de registros: {len(df)}")

    columns = df.columns.tolist()

    # Configuração de colunas
    st.write("### ⚙️ Mapeamento de Colunas")
    c1, c2, c3, c4 = st.columns(4)

    with c1:
        col_email = st.selectbox("📧 Coluna 'Para' *", columns)
    with c2:
        col_cc = st.selectbox("📋 Coluna 'Cópia (CC)'", [None] + columns)
    with c3:
        col_bcc = st.selectbox("🔒 Coluna 'Cópia Oculta (BCC)'", [None] + columns)
    with c4:
        col_arq = st.selectbox("📎 Coluna do Anexo", [None] + columns)

    st.divider()

    # Editor de mensagem
    st.subheader("📝 Redigir Mensagem")

    col_subject, col_tags = st.columns([3, 1])
    with col_subject:
        subject = st.text_input(
            "Assunto do E-mail",
            placeholder="Ex: Documento {TipoDocumento} - {Nome}",
            help="Use {NomeColuna} para inserir valores da planilha"
        )
    with col_tags:
        st.markdown("**Tags disponíveis:**")
        with st.expander("Ver colunas"):
            for col in columns:
                st.code(f"{{{col}}}")

    content = st_quill(
        placeholder="Escreva seu e-mail aqui... use {Tags} para personalizar.",
        html=True,
        key="quill_editor"
    )

    # Validação antes de gerar
    if not subject or not content:
        st.warning("⚠️ Preencha o assunto e o corpo do e-mail antes de gerar")
        return

    # Botões de ação
    st.divider()
    col_btn1, col_btn2 = st.columns(2)

    with col_btn1:
        btn_teste = st.button(
            "🧪 Gerar 1 TESTE",
            use_container_width=True,
            help="Gera apenas o primeiro e-mail para testar"
        )
    with col_btn2:
        btn_massa = st.button(
            "🚀 Gerar TODOS",
            type="primary",
            use_container_width=True,
            help="Gera todos os e-mails da planilha"
        )

    # Processamento
    if btn_teste or btn_massa:
        df_proc = df.head(1) if btn_teste else df

        with st.spinner('🔄 Conectando ao Outlook...'):
            outlook = get_outlook_instance()

        if not outlook:
            st.error("❌ Erro ao conectar com o Outlook. Verifique se está instalado e configurado.")
            return

        st.info(f"📨 Processando {len(df_proc)} e-mail(s)...")
        progress = st.progress(0)

        # Processa e-mails
        stats = process_emails(
            df=df_proc,
            outlook=outlook,
            col_email=col_email,
            subject_template=subject,
            body_template=content,
            columns=columns,
            col_cc=col_cc,
            col_bcc=col_bcc,
            col_attachment=col_arq,
            attachment_folder=config['attachment_folder'],
            progress_bar=progress
        )

        # Mostra resultados
        if stats['success_count'] > 0:
            st.success(f"✅ Finalizado! {stats['success_count']} rascunho(s) criado(s) no Outlook.")

        if stats['error_count'] > 0:
            st.warning(f"⚠️ {stats['error_count']} e-mail(s) com erro")
            with st.expander("Ver erros"):
                for error in stats['errors']:
                    st.error(error)

        if stats['attachment_errors']:
            with st.expander("⚠️ Arquivos não localizados"):
                for error in stats['attachment_errors']:
                    st.warning(error)

        logger.info(f"Processamento concluído: {stats['success_count']} sucessos, {stats['error_count']} erros")


if __name__ == "__main__":
    main()
