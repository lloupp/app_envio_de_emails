"""
gerar_certificados_lote.py
Gera rascunhos de e-mail no Outlook para um lote de certificados, a partir de
uma planilha com a estrutura REAL do lote:

    Colunas: NOME | E-mail | Modalidade | TEXTO | ESP | DE | ATE | CH | anexo
    Aba: Sheet1

A coluna 'anexo' costuma vir VAZIA neste lote, entao o nome do arquivo e
derivado por convencao:  Certificado_{NOME}_{ESP}
(espacos -> "_", acentos removidos, igual ao gerador dos .docx). Como a pasta
pode conter o .docx original ou a versao ja assinada (salva como .docx.pdf
pelo processo de assinatura), o script tenta as extensoes .docx.pdf, .pdf
e .docx nessa ordem, usando a primeira que existir.

Os e-mails sao salvos como RASCUNHOS no Outlook (nao sao enviados).

Uso:
    python gerar_certificados_lote.py "<caminho da planilha .xlsx>" "<pasta dos certificados>"

    Opcional: --dry-run  -> so mostra o que seria feito, sem abrir o Outlook
              (automatico fora do Windows, onde nao ha win32com).

Exemplo (lote assinado):
    python gerar_certificados_lote.py ^
      "M:\\2025 COMERCIAL\\CERTIFICADOS PARA ASSINATURA\\LOTE 13-07-2026\\CERTIFICADOS_NOVOS_2026-07-13.xlsx" ^
      "M:\\2025 COMERCIAL\\CERTIFICADOS PARA ASSINATURA\\LOTE 13-07-2026\\ASSINADOS"
"""

import os
import sys
import unicodedata

import unicodedata as _unic

# ---------------------------------------------------------------- config (edicavel)
SHEET_NAME = "Sheet1"
COL_NOME = "NOME"
COL_EMAIL = "E-mail"
COL_ANEXO = "anexo"          # usada se preenchida; caso contrario deriva de NOME+ESP
COL_ESP = "ESP"              # especialidade, usada para derivar o nome do arquivo
SUBJECT = "Certificado - {NOME} - {ESP}"

BODY = """\
<p>Ol&aacute;, {NOME},</p>
<p>Segue em anexo o seu certificado referente &agrave; {Modalidade}: {TEXTO} {ESP}, \
no per&iacute;odo de {DE} a {ATE}, com carga hor&aacute;ria de {CH} hora(s).</p>
<p>Atenciosamente,<br>Centro de Treinamento M&eacute;dico</p>
"""
# ---------------------------------------------------------------------------------


def transform(s):
    """Espacos -> '_', remove acentos (igual ao nomeador dos .docx)."""
    s = _unic.normalize("NFD", str(s))
    s = "".join(c for c in s if _unic.combining(c) == 0)
    return s.replace(" ", "_")


EXTENSOES_CANDIDATAS = [".docx.pdf", ".pdf", ".docx"]


def derivar_nomes(nome, esp):
    """Lista de nomes de arquivo candidatos, na ordem de preferencia.

    O lote pode conter tanto o .docx original quanto a versao assinada
    (salva como .docx.pdf), entao tentamos as extensoes mais comuns.
    """
    base = f"Certificado_{transform(nome)}_{transform(esp)}"
    return [base + ext for ext in EXTENSOES_CANDIDATAS]


def carregar_linhas(planilha):
    # openpyxl eh mais robusto que pandas aqui e ja e dependencia do repo
    try:
        import openpyxl
    except ImportError:
        import pandas as pd
        df = pd.read_excel(planilha, sheet_name=SHEET_NAME).fillna("")
        return df.to_dict("records")
    wb = openpyxl.load_workbook(planilha, data_only=True)
    ws = wb[SHEET_NAME]
    data = list(ws.iter_rows(values_only=True))
    cols = list(data[0])
    return [dict(zip(cols, r)) for r in data[1:]]


def main():
    dry_run = "--dry-run" in sys.argv
    args = [a for a in sys.argv[1:] if not a.startswith("--")]

    if len(args) < 2:
        print("Uso: python gerar_certificados_lote.py \"<planilha.xlsx>\" \"<pasta certificados>\"")
        sys.exit(1)

    planilha, pasta = args[0], args[1]

    if not os.path.exists(planilha):
        print(f"Planilha nao encontrada: {planilha}")
        sys.exit(1)
    if not os.path.isdir(pasta):
        print(f"Pasta de certificados nao encontrada: {pasta}")
        sys.exit(1)

    linhas = carregar_linhas(planilha)
    print(f"Linhas na planilha: {len(linhas)}")

    # Tenta importar o COM; se nao existir (Linux/Mac), forca dry-run
    outlook = None
    if not dry_run:
        try:
            import pythoncom
            import win32com.client as win32
            pythoncom.CoInitialize()
            try:
                outlook = win32.GetActiveObject("Outlook.Application")
            except Exception:
                outlook = win32.Dispatch("outlook.application")
        except Exception as e:
            print(f"[aviso] nao foi possivel iniciar o Outlook ({e}). Usando dry-run.")
            dry_run = True

    sucesso = 0
    pulados = []          # sem e-mail valido
    sem_anexo = []        # arquivo nao encontrado

    for row in linhas:
        nome = row.get(COL_NOME, "")
        email = str(row.get(COL_EMAIL, "")).strip()
        if not email or "@" not in email:
            pulados.append(nome)
            continue

        _anexo_val = row.get(COL_ANEXO, "")
        anexo_cell = "" if _anexo_val is None else str(_anexo_val).strip()
        candidatos = [anexo_cell] if anexo_cell else derivar_nomes(row.get(COL_NOME, ""), row.get(COL_ESP, ""))

        nome_arq = None
        arq_path = None
        for candidato in candidatos:
            caminho = os.path.join(pasta, candidato)
            if os.path.exists(caminho):
                nome_arq = candidato
                arq_path = caminho
                break

        if arq_path is None:
            sem_anexo.append((nome, candidatos[0]))
            continue

        # Substituicao de tags no assunto e no corpo
        assunto_f = SUBJECT
        corpo_f = BODY
        for col, val in row.items():
            tag = "{" + str(col) + "}"
            v = "" if val is None else str(val)
            assunto_f = assunto_f.replace(tag, v)
            corpo_f = corpo_f.replace(tag, v)

        if dry_run:
            print(f"[dry-run] {email}  ->  assunto='{assunto_f}'  anexo='{nome_arq}'")
            sucesso += 1
            continue

        mail = outlook.CreateItem(0)
        mail.Display()                       # carrega a assinatura padrao
        mail.To = email
        mail.Subject = assunto_f
        mail.HTMLBody = (
            f"<div style='font-family: Calibri; font-size: 11pt;'>{corpo_f}</div><br>"
            + mail.HTMLBody
        )
        mail.Attachments.Add(os.path.abspath(arq_path))
        mail.Save()
        mail.Close(0)
        sucesso += 1

    print(f"\nRascunhos criados: {sucesso}")
    if pulados:
        print(f"\nLinhas PULADAS (sem e-mail valido): {len(pulados)}")
        for p in pulados:
            print(f"  - {p}")
    if sem_anexo:
        print(f"\nAnexos NAO encontrados: {len(sem_anexo)}")
        for n, f in sem_anexo:
            print(f"  - {n}  (esperado: {f})")


if __name__ == "__main__":
    main()
