from django.core.exceptions import ValidationError
from django.test import SimpleTestCase, TestCase

from viagens.models import Viajante
from viagens.utils.normalize import format_rg, normalize_rg


class NormalizeRgUtilsTests(SimpleTestCase):
    def test_normalize_and_format_12345678x(self):
        canon = normalize_rg("12345678x")
        self.assertEqual(canon, "12345678X")
        self.assertEqual(format_rg(canon), "1.234.567-X")

    def test_normalize_and_format_1212345678(self):
        canon = normalize_rg("1212345678")
        self.assertEqual(canon, "1212345678")
        self.assertEqual(format_rg(canon), "12.123.456-8")

    def test_paste_masked_keeps_format(self):
        canon = normalize_rg("12.123.456-8")
        self.assertEqual(canon, "121234568")
        self.assertEqual(format_rg(canon), "12.123.456-8")

    def test_paste_masked_with_x_uppercase(self):
        canon = normalize_rg("1.234.567-x")
        self.assertEqual(canon, "1234567X")
        self.assertEqual(format_rg(canon), "1.234.567-X")


class NormalizeRgModelValidationTests(TestCase):
    def test_model_rejects_rg_with_less_than_9_chars(self):
        viajante = Viajante(
            nome="SERVIDOR TESTE",
            rg="1.234.567-X",
            cpf="12345678901",
            cargo="Delegado",
            telefone="41999999999",
        )
        with self.assertRaises(ValidationError):
            viajante.full_clean()
