Alfredo do Email
O Alfredo do Email é uma ferramenta de automação desenvolvida com Python e Streamlit para facilitar o envio de e-mails em massa através do Microsoft Outlook. Ele permite a personalização dinâmica do corpo do e-mail e do assunto utilizando colunas de uma planilha Excel ou CSV, além de anexar arquivos automaticamente.

✨ Funcionalidades
Upload de Planilha: Suporte para arquivos .xlsx e .csv.

Mapeamento Dinâmico: Escolha quais colunas representam o destinatário, cópia (CC), cópia oculta (BCC) e o nome do anexo.

Editor de Texto Rico: Utilize o componente streamlit-quill para formatar seu e-mail com negrito, listas e links.

Tags Personalizadas: Use {NomeDaColuna} no assunto ou no corpo para personalizar cada mensagem.

Integração com Outlook: Os e-mails são gerados diretamente como rascunhos no Outlook, preservando sua assinatura padrão.

Modo Teste: Gere apenas o primeiro registro para validar a formatação antes de processar toda a lista.

🛠️ Pré-requisitos
Para rodar este projeto, você precisará de:

Sistema Operacional: Windows (necessário para a integração com win32com).

Microsoft Outlook: Instalado e configurado com uma conta ativa.

Python 3.8+ instalado.

🚀 Como instalar e rodar
Clone o repositório:

Bash
git clone https://github.com/SEU_USUARIO/alfredo-do-email.git
cd alfredo-do-email
Crie um ambiente virtual (recomendado):

Bash
python -m venv venv
venv\Scripts\activate
Instale as dependências:

Bash
pip install streamlit pandas pywin32 streamlit-quill
Execute a aplicação:

Bash
streamlit run seu_arquivo.py
📖 Como usar
Prepare sua planilha: Certifique-se de que ela tenha colunas para o e-mail e, se necessário, uma coluna com o nome exato dos arquivos de anexo (ex: documento_01.pdf).

Configure os Anexos: Na barra lateral, cole o caminho da pasta onde os arquivos estão salvos no seu computador.

Redija o E-mail: Use o formato {Coluna} para inserir dados da planilha. Exemplo: Olá {Nome}, segue seu boleto.

Gere os Rascunhos: Clique em "Gerar Todos". O Alfredo abrirá o Outlook em segundo plano e salvará os e-mails na sua pasta de Rascunhos.

📝 Observações Técnicas
O script utiliza pythoncom.CoInitialize() para garantir a estabilidade da conexão com a API do Windows.

O uso de mail.Display() seguido de mail.Save() é proposital para garantir que a assinatura padrão do seu Outlook seja carregada no corpo do e-mail.
