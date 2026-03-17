"""
Testes unitários para o sistema Alfredo do Email
"""

import unittest
import pandas as pd
from unittest.mock import Mock, patch, MagicMock
import os
import sys

# Adiciona o diretório raiz ao path para importar o módulo
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Mock das dependências do Streamlit e win32com antes de importar
sys.modules['streamlit'] = MagicMock()
sys.modules['streamlit_quill'] = MagicMock()
sys.modules['win32com'] = MagicMock()
sys.modules['win32com.client'] = MagicMock()

from gerador_email import (
    validate_email,
    sanitize_path,
    validate_attachment,
    replace_tags,
)


class TestEmailValidation(unittest.TestCase):
    """Testes para validação de e-mails"""

    def test_valid_email(self):
        """Testa e-mails válidos"""
        valid_emails = [
            'usuario@exemplo.com',
            'nome.sobrenome@empresa.com.br',
            'teste123@dominio.org',
            'user+tag@example.co.uk'
        ]
        for email in valid_emails:
            with self.subTest(email=email):
                self.assertTrue(validate_email(email))

    def test_invalid_email(self):
        """Testa e-mails inválidos"""
        invalid_emails = [
            'email_sem_arroba.com',
            '@sem_usuario.com',
            'sem_dominio@',
            'espaço @email.com',
            'email@',
            '@dominio.com',
            '',
            None,
            'email sem arroba',
            'email@dominio',
        ]
        for email in invalid_emails:
            with self.subTest(email=email):
                self.assertFalse(validate_email(email))

    def test_email_with_whitespace(self):
        """Testa e-mail com espaços (deve validar após strip)"""
        self.assertTrue(validate_email('  usuario@exemplo.com  '))


class TestPathSanitization(unittest.TestCase):
    """Testes para sanitização de caminhos"""

    def test_empty_path(self):
        """Testa caminho vazio"""
        self.assertIsNone(sanitize_path(''))
        self.assertIsNone(sanitize_path(None))

    def test_remove_quotes(self):
        """Testa remoção de aspas do caminho"""
        # Cria um diretório temporário para teste
        test_dir = os.path.dirname(os.path.abspath(__file__))
        path_with_quotes = f'"{test_dir}"'
        result = sanitize_path(path_with_quotes)
        self.assertEqual(result, test_dir)

    def test_nonexistent_path(self):
        """Testa caminho que não existe"""
        fake_path = '/caminho/que/nao/existe/xyz123'
        result = sanitize_path(fake_path)
        self.assertIsNone(result)

    def test_existing_path(self):
        """Testa caminho que existe"""
        current_dir = os.path.dirname(os.path.abspath(__file__))
        result = sanitize_path(current_dir)
        self.assertEqual(result, current_dir)


class TestAttachmentValidation(unittest.TestCase):
    """Testes para validação de anexos"""

    def test_nonexistent_file(self):
        """Testa arquivo que não existe"""
        self.assertFalse(validate_attachment('/caminho/arquivo_inexistente.pdf'))

    def test_file_with_allowed_extension(self):
        """Testa arquivo com extensão permitida"""
        # Cria um arquivo temporário
        test_file = 'test_file.pdf'
        with open(test_file, 'w') as f:
            f.write('test')

        try:
            result = validate_attachment(test_file, ['.pdf', '.docx'])
            self.assertTrue(result)
        finally:
            if os.path.exists(test_file):
                os.remove(test_file)

    def test_file_with_disallowed_extension(self):
        """Testa arquivo com extensão não permitida"""
        # Cria um arquivo temporário
        test_file = 'test_file.exe'
        with open(test_file, 'w') as f:
            f.write('test')

        try:
            result = validate_attachment(test_file, ['.pdf', '.docx'])
            self.assertFalse(result)
        finally:
            if os.path.exists(test_file):
                os.remove(test_file)


class TestTagReplacement(unittest.TestCase):
    """Testes para substituição de tags"""

    def test_single_tag_replacement(self):
        """Testa substituição de uma única tag"""
        text = "Olá {Nome}, bem-vindo!"
        row = pd.Series({'Nome': 'João', 'Email': 'joao@email.com'})
        columns = ['Nome', 'Email']
        result = replace_tags(text, row, columns)
        self.assertEqual(result, "Olá João, bem-vindo!")

    def test_multiple_tags_replacement(self):
        """Testa substituição de múltiplas tags"""
        text = "Olá {Nome}, seu e-mail é {Email}"
        row = pd.Series({'Nome': 'Maria', 'Email': 'maria@email.com'})
        columns = ['Nome', 'Email']
        result = replace_tags(text, row, columns)
        self.assertEqual(result, "Olá Maria, seu e-mail é maria@email.com")

    def test_missing_value_replacement(self):
        """Testa substituição com valor ausente (NaN)"""
        text = "Olá {Nome}, seu telefone é {Telefone}"
        row = pd.Series({'Nome': 'Pedro', 'Telefone': None})
        columns = ['Nome', 'Telefone']
        result = replace_tags(text, row, columns)
        self.assertEqual(result, "Olá Pedro, seu telefone é ")

    def test_no_tags(self):
        """Testa texto sem tags"""
        text = "Olá, bem-vindo!"
        row = pd.Series({'Nome': 'Ana'})
        columns = ['Nome']
        result = replace_tags(text, row, columns)
        self.assertEqual(result, "Olá, bem-vindo!")

    def test_repeated_tags(self):
        """Testa tags repetidas no texto"""
        text = "Olá {Nome}, {Nome}!"
        row = pd.Series({'Nome': 'Carlos'})
        columns = ['Nome']
        result = replace_tags(text, row, columns)
        self.assertEqual(result, "Olá Carlos, Carlos!")


class TestDataFrameLoading(unittest.TestCase):
    """Testes para carregamento de DataFrames"""

    def test_empty_dataframe(self):
        """Testa detecção de DataFrame vazio"""
        df = pd.DataFrame()
        self.assertTrue(df.empty)

    def test_valid_dataframe(self):
        """Testa DataFrame válido"""
        df = pd.DataFrame({
            'Nome': ['João', 'Maria'],
            'Email': ['joao@email.com', 'maria@email.com']
        })
        self.assertFalse(df.empty)
        self.assertEqual(len(df), 2)
        self.assertListEqual(df.columns.tolist(), ['Nome', 'Email'])


def run_tests():
    """Executa todos os testes"""
    loader = unittest.TestLoader()
    suite = loader.loadTestsFromModule(sys.modules[__name__])
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)
    return result.wasSuccessful()


if __name__ == '__main__':
    unittest.main()
