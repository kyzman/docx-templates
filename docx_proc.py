import io
import operator
from collections import OrderedDict

from docxtpl import DocxTemplate
from docx import Document


class ProcessTemplate:
    def __init__(self, templ: DocxTemplate, context: dict | list[dict]):
        self.template = templ
        context_keys = context.keys()
        template_keys = templ.get_undeclared_template_variables()
        result1 = template_keys - context_keys
        result2 = context_keys - template_keys
        if result1:
            print(f'[Ошибка!] В шаблоне {templ.template_file} присутствует лишнее поле {result1}!\n'
                  f'Продолжение работы невозможно!')
            raise SystemExit(1)
        if result2:
            print(f"[Внимание!] В шаблоне {templ.template_file} отсутствует: {result2}")
        self.context = context
        self.template.render(self.context)

    def get(self) -> DocxTemplate:
        return self.template


class MemSavedTemplate:
    def __init__(self, rendered_template: DocxTemplate):
        self.template = rendered_template
        self.buffer = io.BytesIO()
        self.template.save(self.buffer)

    def get(self) -> Document:
        self.buffer.seek(0)
        return self.buffer


class DocxPreAssembly:
    def __init__(self, order: dict, buffer: dict):
        ordered = OrderedDict(sorted(order.items(), key=operator.itemgetter(1)))
        self.order = OrderedDict()
        for ordered in ordered.keys():
            self.order[ordered] = buffer[ordered]

    def get(self) -> dict:
        return dict(self.order)

    def get_parts_list(self) -> list:
        result = []
        for key in self.order.keys():
            result.append(str(key).replace(".", "_"))
        return result


class DocxCreateCommonInMem:
    def __init__(self, assembly: DocxPreAssembly):
        self.root = io.BytesIO()
        source = Document()
        for paragraph in assembly.get_parts_list():
            source.add_paragraph("{{{{{p " + paragraph + " }}}}}")
        source.save(self.root)

    def get(self) -> Document:
        self.root.seek(0)
        return self.root


class DocxFinalAssembly:
    def __init__(self, filename, root, assembly: DocxPreAssembly):
        self.assembly = assembly
        self.common = DocxTemplate(root)
        self.filename = filename

    def form_dict(self):
        result = {}
        for key in self.assembly.get().keys():
            result[str(key).replace('.', '_')] = self.common.new_subdoc(self.assembly.order[key])
        return result

    def save(self):
        self.common.render(self.form_dict())
        self.common.save(self.filename)


def create_templates_buffer(files: set, dataset: dict, path: str = '') -> dict[MemSavedTemplate]:
    tmp_buffer = {}
    for file in list(files):
        current_template = DocxTemplate(f"{path}{file}")
        tmp_buffer[file] = MemSavedTemplate(ProcessTemplate(current_template, dataset[file]).get()).get()
    return tmp_buffer


def main():
    files = {'heading_1.docx', 'heading_2.docx', 'body_1.docx', 'body_2.docx', 'bottom_1.docx',  'bottom_2.docx'}
    dataset = {'heading_1.docx': {'company': 'ООО "Первая Компания"', 'Project': 'Проект №1', 'data': 'Что-то особенное!'},
               'heading_2.docx': {'number': '001'},
               'body_1.docx': {'company_name': 'ГК "Госактивы"',
                               'company_address': 'Россия, Москва, пл. Победы, д.1',
                               'court_name': 'Суд Москвы',
                               'date_from': '13.02.2024',
                               'case_number': '54',
                               'financial_organization': 'Банк Москвы',
                               'orgn': '35482100551',
                               'inn': '950643635656',
                               'reg_address': 'Россия, Москва, пр. Банковый, д.1',
                               },
               'body_2.docx': {'bill_number': '67582',
                               'company_name': 'ГК "Госактивы"',
                               'table': [{'pp_number': '1', 'lot_number': '1', 'pp_date': '14.02.2024', 'winner_name': 'Петров А.В.', 'payment_purpose': 'выплата', 'payment_sum': '1503000'},
                                         {'pp_number': '2', 'lot_number': '2', 'pp_date': '14.02.2024', 'winner_name': 'Иванов Д.Я.', 'payment_purpose': 'выплата', 'payment_sum': '570000'},
                                         {'pp_number': '3', 'lot_number': '3', 'pp_date': '15.02.2024', 'winner_name': 'Кошкин С.М.', 'payment_purpose': 'выплата', 'payment_sum': '45000'},
                                         ],
                               'total': '2118000'
                               },
               'bottom_1.docx': {'sign': 'Кочетков', 'day': '26', 'month': 'марта', 'year': '24'},
               'bottom_2.docx': {'full_name': 'Кочетков Пётр Яковлевич', 'credentials': 'Тюняев А.П.'}
               }
    order = {'heading_1.docx': 0, 'heading_2.docx': 1, 'body_1.docx': 2, 'body_2.docx': 3, 'bottom_1.docx': 4,  'bottom_2.docx': 5}

    buffer = create_templates_buffer(files, dataset, 'docx/')
    assembly = DocxPreAssembly(order, buffer)
    asm_buffer = DocxFinalAssembly('result.docx', DocxCreateCommonInMem(assembly).get(), assembly)
    asm_buffer.save()


if __name__ == '__main__':
    main()
