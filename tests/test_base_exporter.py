"""Tests for base_exporter — shared export helpers."""
import unittest
import pandas as pd
import math

from base_exporter import BaseExporter


class TestIsValid(unittest.TestCase):
    def test_nan(self):
        self.assertFalse(BaseExporter.is_valid(float('nan')))

    def test_none(self):
        self.assertFalse(BaseExporter.is_valid(None))

    def test_empty_string(self):
        self.assertFalse(BaseExporter.is_valid(""))

    def test_nan_string(self):
        self.assertFalse(BaseExporter.is_valid("nan"))

    def test_none_string(self):
        self.assertFalse(BaseExporter.is_valid("None"))

    def test_na_string(self):
        self.assertFalse(BaseExporter.is_valid("N/A"))

    def test_valid_number(self):
        self.assertTrue(BaseExporter.is_valid(42))

    def test_valid_string(self):
        self.assertTrue(BaseExporter.is_valid("Hello"))

    def test_valid_zero(self):
        self.assertTrue(BaseExporter.is_valid(0))

    def test_whitespace_only(self):
        self.assertFalse(BaseExporter.is_valid("   "))


class TestToNumber(unittest.TestCase):
    def test_integer(self):
        self.assertEqual(BaseExporter.to_number("42"), 42)

    def test_float(self):
        self.assertAlmostEqual(BaseExporter.to_number("3.14"), 3.14)

    def test_none(self):
        self.assertIsNone(BaseExporter.to_number(None))

    def test_nan_string(self):
        self.assertIsNone(BaseExporter.to_number("nan"))

    def test_not_done(self):
        self.assertIsNone(BaseExporter.to_number("not done"))

    def test_non_numeric(self):
        self.assertIsNone(BaseExporter.to_number("abc"))

    def test_negative(self):
        self.assertEqual(BaseExporter.to_number("-5"), -5)

    def test_no_try_int(self):
        self.assertAlmostEqual(BaseExporter.to_number("42", try_int=False), 42.0)
        self.assertIsInstance(BaseExporter.to_number("42", try_int=False), float)


class TestToFloat(unittest.TestCase):
    def test_integer_string(self):
        self.assertAlmostEqual(BaseExporter.to_float("42"), 42.0)

    def test_float_string(self):
        self.assertAlmostEqual(BaseExporter.to_float("3.14"), 3.14)

    def test_none(self):
        self.assertIsNone(BaseExporter.to_float(None))

    def test_invalid(self):
        self.assertIsNone(BaseExporter.to_float("abc"))

    def test_nan_value(self):
        self.assertIsNone(BaseExporter.to_float(float('nan')))


class TestFormatDate(unittest.TestCase):
    def test_iso_date(self):
        self.assertEqual(BaseExporter.format_date("2025-03-15"), "15-Mar-2025")

    def test_datetime_with_time(self):
        self.assertEqual(BaseExporter.format_date("2025-03-15T10:30:00"), "15-Mar-2025")

    def test_pipe_delimited(self):
        self.assertEqual(BaseExporter.format_date("2025-03-15|2025-04-01"), "15-Mar-2025")

    def test_empty(self):
        self.assertEqual(BaseExporter.format_date(""), "")

    def test_none(self):
        self.assertEqual(BaseExporter.format_date(None), "")

    def test_nan(self):
        self.assertEqual(BaseExporter.format_date(float('nan')), "")


class TestSafeStr(unittest.TestCase):
    def setUp(self):
        self.row = pd.Series({'A': 'hello', 'B': float('nan'), 'C': ''})

    def test_valid_column(self):
        self.assertEqual(BaseExporter.safe_str(self.row, 'A'), 'hello')

    def test_nan_column(self):
        self.assertIsNone(BaseExporter.safe_str(self.row, 'B'))

    def test_empty_column(self):
        self.assertIsNone(BaseExporter.safe_str(self.row, 'C'))

    def test_missing_column(self):
        self.assertIsNone(BaseExporter.safe_str(self.row, 'Z'))


class TestGetPatientRow(unittest.TestCase):
    def setUp(self):
        self.df = pd.DataFrame({
            'Screening #': ['101-01', '102-02'],
            'Name': ['Alice', 'Bob'],
        })
        self.exporter = BaseExporter(self.df)

    def test_found(self):
        row = self.exporter.get_patient_row('101-01')
        self.assertIsNotNone(row)
        self.assertEqual(row['Name'], 'Alice')

    def test_not_found(self):
        row = self.exporter.get_patient_row('999-99')
        self.assertIsNone(row)


class TestGenerateExport(unittest.TestCase):
    """Test the multi-patient export orchestration."""

    def setUp(self):
        self.df = pd.DataFrame({'Screening #': ['101-01', '102-02']})
        self.exporter = _StubExporter(self.df)

    def test_single_patient(self):
        data, ext, pid = self.exporter.generate_export(['101-01'])
        self.assertEqual(ext, 'xlsx')
        self.assertEqual(pid, '101-01')
        self.assertEqual(data, b'excel_101-01')

    def test_multiple_patients(self):
        data, ext, pid = self.exporter.generate_export(['101-01', '102-02'])
        self.assertEqual(ext, 'zip')
        self.assertIsNone(pid)
        # zip data should be non-empty bytes
        self.assertIsInstance(data, bytes)
        self.assertGreater(len(data), 0)

    def test_custom_filename_fmt(self):
        data, ext, pid = self.exporter.generate_export(
            ['101-01', '102-02'],
            filename_fmt=lambda p: f"{p}_labs.xlsx")
        self.assertEqual(ext, 'zip')


class _StubExporter(BaseExporter):
    """Stub exporter for testing generate_export."""
    def process_patient(self, patient_id, **kwargs):
        return f"excel_{patient_id}".encode()


if __name__ == '__main__':
    unittest.main()
