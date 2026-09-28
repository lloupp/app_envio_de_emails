import sys
import pandas as pd
import win32com.client as win32
import pythoncom
import os

if len(sys.argv) < 2:
    print("Uso: python gerar_certificados.py \"<pasta com planilha_envio_certificados.xlsx e os PDFs>\"")
    sys.exit(1)

PASTA_ANEXOS = sys.argv[1]
PLANILHA = os.path.join(PASTA_ANEXOS, 'planilha_envio_certificados.xlsx')

SUBJECT = "Seu certificado do {Nome do Curso} já está disponível! 🎓"

BODY = """
<p>Olá, {Nome do Destinatário},</p>
<p>Esperamos que este e-mail o(a) encontre bem.</p>
<p>Em nome do Centro de Treinamento Médico, gostaríamos de agradecer imensamente pela sua participação
e dedicação no {Nome do Curso}, realizado em {Data do Curso}.</p>
<p>Sua presença e colaboração como {Tipo Participação} foram fundamentais para o sucesso e a excelência
do nosso evento.</p>
<p>Segue em anexo o seu certificado.</p>
<p>Atenciosamente,<br>Centro de Treinamento Médico</p>
"""

df = pd.read_excel(PLANILHA, sheet_name="Envio")
colunas = df.columns.tolist()

pythoncom.CoInitialize()
try:
    outlook = win32.GetActiveObject("Outlook.Application")
except Exception:
    outlook = win32.Dispatch('outlook.application')

sucesso = 0
erros_anexo = []

for _, row in df.iterrows():
    email_dest = str(row['Email']).strip()
    if "@" not in email_dest:
        continue

    mail = outlook.CreateItem(0)
    mail.Display()
    mail.To = email_dest

    assunto_f = SUBJECT
    corpo_f = BODY
    for col in colunas:
        tag = "{" + col + "}"
        val = str(row[col]) if pd.notna(row[col]) else ""
        assunto_f = assunto_f.replace(tag, val)
        corpo_f = corpo_f.replace(tag, val)

    mail.Subject = assunto_f
    mail.HTMLBody = (
        f"<div style='font-family: Calibri; font-size: 11pt;'>{corpo_f}</div><br>" + mail.HTMLBody
    )

    nome_arquivo = str(row['Anexo']).strip()
    arq_path = os.path.join(PASTA_ANEXOS, nome_arquivo)
    if os.path.exists(arq_path):
        mail.Attachments.Add(os.path.abspath(arq_path))
    else:
        erros_anexo.append(nome_arquivo)

    mail.Save()
    mail.Close(0)
    sucesso += 1

print(f"Rascunhos criados: {sucesso}")
if erros_anexo:
    print("Anexos nao encontrados:", erros_anexo)
