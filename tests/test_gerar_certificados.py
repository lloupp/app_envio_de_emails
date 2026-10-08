import csv
from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import Workbook

from gerar_certificados import (
    CAMPOS_OBRIGATORIOS, carregar_linhas, criar_rascunho,
    localizar_pdf, main, preparar_lote, substituir_tags, texto_celula,
)


class MensagemFalsa:
    def __init__(self):
        self.HTMLBody = "<p>Assinatura institucional</p>"
        self.Attachments = self
        self.anexos = []
        self.saved = False
        self.closed = None

    def Display(self):
        pass

    def Add(self, caminho):
        self.anexos.append(caminho)

    def Save(self):
        self.saved = True

    def Close(self, opcao):
        self.closed = opcao


class OutlookFalso:
    def __init__(self):
        self.mensagens = []

    def CreateItem(self, tipo):
        assert tipo == 0
        mensagem = MensagemFalsa()
        self.mensagens.append(mensagem)
        return mensagem


class CertificadosTestes(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.pasta = Path(self.temp.name)
        self.pdf = self.pasta / "certificado.pdf"
        self.pdf.write_bytes(b"%PDF-1.4 teste")
        self.dados = {
            "Nome do Destinatário": "Ana & Bruno",
            "Email": "ana@example.com",
            "Anexo": "certificado.pdf",
            "Nome do Curso": "Cirurgia <Robótica>",
            "Data do Curso": datetime(2026, 10, 8),
            "Tipo Participação": "Participante",
        }

    def criar_planilha(self, dados=None, colunas=None):
        livro = Workbook()
        aba = livro.active
        aba.title = "Envio"
        colunas = list(colunas or CAMPOS_OBRIGATORIOS)
        aba.append(colunas)
        if dados is not None:
            for registro in dados:
                aba.append([registro.get(coluna) for coluna in colunas])
        livro.save(self.pasta / "planilha_envio_certificados.xlsx")

    def test_valores_excel_e_html_nao_injetam_marcadores(self):
        self.assertEqual(texto_celula(datetime(2026, 10, 8)), "08/10/2026")
        dados = dict(self.dados, **{"Nome do Curso": "<img src=x> {Nome do Destinatário}"})
        corpo = substituir_tags("<p>{Nome do Curso}</p>", dados, html=True)
        self.assertEqual(corpo, "<p>&lt;img src=x&gt; {Nome do Destinatário}</p>")

    def test_planilha_e_rascunho_sem_envio(self):
        self.criar_planilha([self.dados])
        aprovados, relatorio = preparar_lote(self.pasta, carregar_linhas(self.pasta / "planilha_envio_certificados.xlsx"))
        self.assertEqual(len(aprovados), 1)
        self.assertEqual(relatorio[0]["status"], "Pronto")
        self.assertIn("Cirurgia <Robótica>", aprovados[0]["assunto"])
        self.assertIn("Cirurgia &lt;Robótica&gt;", aprovados[0]["corpo"])
        outlook = OutlookFalso()
        criar_rascunho(outlook, aprovados[0])
        self.assertEqual(outlook.mensagens[0].anexos, [str(self.pdf)])
        self.assertTrue(outlook.mensagens[0].saved)
        self.assertEqual(outlook.mensagens[0].closed, 0)
        self.assertIn("Assinatura institucional", outlook.mensagens[0].HTMLBody)

    def test_nao_cria_rascunho_sem_pdf_e_ignora_email_invalido(self):
        sem_pdf = dict(self.dados, **{"Anexo": "inexistente.pdf"})
        sem_email = dict(self.dados, **{"Email": "sem-arroba"})
        prontos, relatorio = preparar_lote(self.pasta, [(2, sem_pdf), (3, sem_email)])
        self.assertEqual(prontos, [])
        self.assertTrue(all(reg["status"] == "Ignorado" for reg in relatorio))

    def test_protege_caminhos_e_formato(self):
        for nome in ("../segredo.pdf", r"..\segredo.pdf", "/tmp/segredo.pdf", r"C:\segredo.pdf", "certificado.docx"):
            with self.subTest(nome=nome), self.assertRaises(ValueError):
                localizar_pdf(self.pasta, nome)
        self.assertEqual(localizar_pdf(self.pasta, "certificado.pdf"), self.pdf)

    def test_impede_registros_duplicados(self):
        prontos, relatorio = preparar_lote(self.pasta, [(2, self.dados), (3, self.dados)])
        self.assertEqual(len(prontos), 1)
        self.assertIn("duplicado", relatorio[1]["detalhe"])

    def test_planilha_sem_coluna_essencial(self):
        self.criar_planilha([self.dados], colunas=["Email", "Anexo"])
        with self.assertRaisesRegex(ValueError, "Colunas ausentes"):
            carregar_linhas(self.pasta / "planilha_envio_certificados.xlsx")

    def test_simulacao_sem_outlook_e_csv(self):
        self.criar_planilha([self.dados])
        destino = self.pasta / "resultado.csv"
        with patch("gerar_certificados.conectar_outlook", side_effect=AssertionError("Outlook foi aberto")):
            retorno = main([str(self.pasta), "--dry-run", "--relatorio-csv", str(destino)])
        self.assertEqual(retorno, 0)
        with destino.open(encoding="utf-8-sig", newline="") as arquivo:
            linhas = list(csv.DictReader(arquivo, delimiter=";"))
        self.assertEqual(linhas[0]["status"], "Pronto")

    def test_simulacao_com_falhas_retorna_codigo_um(self):
        self.criar_planilha([dict(self.dados, **{"Anexo": "ausente.pdf"})])
        retorno = main([str(self.pasta), "--dry-run"])
        self.assertEqual(retorno, 1)


if __name__ == "__main__":
    unittest.main()
