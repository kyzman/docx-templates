import io
import operator
from collections import OrderedDict

from docxtpl import DocxTemplate

import datasets  # Данные для работы сборщика документа
from docx import Document


def keys_validation(dataset_values, template_values) -> tuple[set, set]:
    dataset_vs_template = dataset_values - template_values
    template_vs_dataset = template_values - dataset_values
    return dataset_vs_template, template_vs_dataset


class ProcessTemplate:
    def __init__(self, templ: DocxTemplate, context: dict | list[dict]):
        self.template = templ
        result1, result2 = keys_validation(context.keys(), templ.get_undeclared_template_variables())
        if result2:
            print(f'[Ошибка!] В шаблоне {templ.template_file} присутствуют поля, которых нет в данных:', end=' ')
            print(*list(result2), sep=", ")
            print(f'[Выход] Продолжение работы невозможно!')
            raise SystemExit(1)
        if result1:
            print(f'[Внимание!] В шаблоне {templ.template_file} отсутствуют поля для вставки данных:', end=' ')
            print(*list(result1), sep=", ")
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
    """ Подготовка, создание и сохранение в памяти специального шаблона, по которому будет идти сборка документа """
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
    # создание буфера в памяти из предоставленных шаблонов
    buffer = create_templates_buffer(datasets.files, datasets.dataset, 'docx/')
    # предварительная сборка шаблонов по заданному порядку
    assembly = DocxPreAssembly(datasets.order, buffer)
    # сборка и сохранение финального документа из подготовленной последовательности шаблонов
    asm_buffer = DocxFinalAssembly('result.docx', DocxCreateCommonInMem(assembly).get(), assembly)
    asm_buffer.save()


if __name__ == '__main__':
    main()
