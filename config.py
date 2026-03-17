"""
Configurações da aplicação Alfredo do Email
"""

# Configurações de Email
EMAIL_CONFIG = {
    'allowed_attachment_extensions': [
        '.pdf', '.docx', '.xlsx', '.pptx', '.txt',
        '.jpg', '.jpeg', '.png', '.zip', '.rar'
    ],
    'max_attachment_size_mb': 25,  # Limite do Outlook
}

# Configurações de Logging
LOGGING_CONFIG = {
    'log_file': 'email_generator.log',
    'log_level': 'INFO',
    'log_format': '%(asctime)s - %(levelname)s - %(message)s'
}

# Configurações da Interface
UI_CONFIG = {
    'page_title': 'Alfredo do Email',
    'layout': 'wide',
    'default_font_family': 'Calibri',
    'default_font_size': '11pt'
}

# Mensagens de erro personalizadas
ERROR_MESSAGES = {
    'invalid_email': 'Endereço de e-mail inválido',
    'file_not_found': 'Arquivo não encontrado',
    'outlook_connection': 'Não foi possível conectar ao Outlook. Verifique se está instalado e configurado.',
    'empty_spreadsheet': 'A planilha está vazia ou não pôde ser carregada',
    'invalid_path': 'Caminho inválido ou pasta não encontrada',
}

# Mensagens de sucesso
SUCCESS_MESSAGES = {
    'draft_created': 'Rascunho criado com sucesso',
    'all_processed': 'Todos os e-mails foram processados',
}
