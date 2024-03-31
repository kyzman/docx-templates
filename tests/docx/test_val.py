import unittest
from docx_proc import keys_validation
from docxtpl import DocxTemplate


class TestValues(unittest.TestCase):

    def test_equal_keys_validation(self):
        test_dataset = {'value1': 'значение 1', 'value2': 'значение 2', 'value3': 'значение 4'}
        test_tpl = DocxTemplate('test1.docx')
        result1, result2 = keys_validation(test_dataset.keys(), test_tpl.get_undeclared_template_variables())
        self.assertSetEqual(result1, set())
        self.assertSetEqual(result1, set())
        self.assertSetEqual(result1, result2)

    def test_addition_in_data_keys_validation(self):
        test_dataset = {'value1': 'значение 1', 'value2': 'значение 2', 'value3': 'значение 4', 'value4': 'addition'}
        test_tpl = DocxTemplate('test1.docx')
        result1, result2 = keys_validation(test_dataset.keys(), test_tpl.get_undeclared_template_variables())
        self.assertSetEqual(result1, {'value4'})
        self.assertSetEqual(result2, set())
        self.assertNotEqual(result1, result2)

    def test_addition_in_template_keys_validation(self):
        test_dataset = {'value1': 'значение 1', 'value2': 'значение 2'}
        test_tpl = DocxTemplate('test1.docx')
        result1, result2 = keys_validation(test_dataset.keys(), test_tpl.get_undeclared_template_variables())
        self.assertSetEqual(result1, set())
        self.assertSetEqual(result2, {'value3'})
        self.assertNotEqual(result1, result2)

