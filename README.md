# ✉️ Alfredo do Email

O **Alfredo do Email** é uma ferramenta de automação desenvolvida com Python e Streamlit para facilitar o envio de e-mails em massa através do Microsoft Outlook. Ele permite a personalização dinâmica do corpo e do assunto do e-mail utilizando colunas de uma planilha Excel ou CSV, além de anexar arquivos automaticamente.

---

## ✨ Funcionalidades

- **Upload de Planilha:** Suporte para arquivos `.xlsx` e `.csv`.
- **Mapeamento Dinâmico:** Escolha quais colunas representam o destinatário, cópia (CC), cópia oculta (CCO) e o nome do anexo.
- **Editor de Texto Rico:** Utilize o componente `streamlit-quill` para formatar seu e-mail com negrito, listas e links.
- **Tags Personalizadas:** Use `{NomeDaColuna}` no assunto ou no corpo para personalizar cada mensagem.
- **Integração com Outlook:** Os e-mails são gerados diretamente como rascunhos no Outlook, preservando sua assinatura padrão.
- **Modo Teste:** Gere apenas o primeiro rascunho para validar a formatação antes de processar toda a lista.

---

## 🛠️ Pré-requisitos

- **Sistema Operacional:** Windows (necessário para a integração com `win32com`).
- **Microsoft Outlook** instalado e configurado com uma conta ativa.
- **Python 3.8+** instalado.

---

## 🚀 Como instalar e rodar

**1. Clone o repositório:**

```bash
git clone https://github.com/lloupp/app_envio_de_emails.git
cd app_envio_de_emails
```

**2. Crie um ambiente virtual (recomendado):**

```bash
python -m venv venv
venv\Scripts\activate
```

**3. Instale as dependências:**

```bash
pip install -r requirements.txt
```

**4. Execute a aplicação:**

```bash
streamlit run gerador_email.py
```

---

## 📖 Como usar

1. **Prepare sua planilha:** Certifique-se de que ela tenha uma coluna com os e-mails dos destinatários e, se necessário, uma coluna com o nome exato dos arquivos de anexo (ex: `documento_01.pdf`).
2. **Configure os Anexos:** Na barra lateral, cole o caminho completo da pasta onde os arquivos estão salvos (ex: `C:\Documentos\Certificados`).
3. **Redija o E-mail:** Use o formato `{NomeDaColuna}` para inserir dados da planilha. Exemplo: `Olá {Nome}, segue seu boleto.`
4. **Gere os Rascunhos:** Clique em **"Gerar Todos"**. O Alfredo abrirá o Outlook em segundo plano e salvará os e-mails na sua pasta de Rascunhos.

---

## 📝 Observações Técnicas

- O script utiliza `pythoncom.CoInitialize()` para garantir a estabilidade da conexão com a API do Windows, e `pythoncom.CoUninitialize()` ao finalizar.
- O uso de `mail.Display()` seguido de `mail.Save()` é proposital: garante que a assinatura padrão do Outlook seja carregada no corpo do e-mail.
- Caso algum arquivo de anexo não seja encontrado, o rascunho ainda é criado e os arquivos faltantes são listados ao final do processamento.

