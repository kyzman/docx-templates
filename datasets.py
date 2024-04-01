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