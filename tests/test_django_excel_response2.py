# -*- coding: utf-8 -*-

from django_excel_response import ExcelResponse


class TestDjangoExcelResponse2Commands(object):

    def setup_class(self):
        self.data1 = [{
            'Column 1': 1,
            'Column 2': 2,
        }, {
            'Column 1': 3,
            'Column 2': 4,
        }]
        self.data2 = [
            ['Column 1', 'Column 2'],
            [1, 2],
            [3, 4]
        ]
        self.data3 = [
            ['Column 1', 'Column 2'],
            [1, [2, 3]],
            [3, 4]
        ]

        self.data11 = {
            'Sheet 1': [{
                'Column 1': 1,
                'Column 2': 2,
            }, {
                'Column 1': 3,
                'Column 2': 4,
            }]
        }
        self.data22 = {
            'Sheet 1': [
                ['Column 1', 'Column 2'],
                [1, 2],
                [3, 4]
            ]
        }
        self.data33 = {
            'Sheet 1': [
                ['Column 1', 'Column 2'],
                [1, [2, 3]],
                [3, 4]
            ]
        }

        self.content_type_csv = 'text/csv'
        self.content_type_xls = 'application/vnd.ms-excel'

        self.sheet_data1 = {'Sheet 1': [['Column 1', 'Column 2'], [1, 2], [3, 4]]}
        self.sheet_data2 = {'Sheet 1': [['Column 1', 'Column 2'], [1, [2, 3]], [3, 4]]}

    def test_as_csv(self):
        csv1 = ExcelResponse(self.data1, 'my_data', force_csv=True, font='name SimSum')
        assert csv1.headers['Content-Type'] == self.content_type_csv
        assert csv1.data == self.sheet_data1

        csv2 = ExcelResponse(self.data2, 'my_data', force_csv=True, font='name SimSum')
        assert csv2.headers['Content-Type'] == self.content_type_csv
        assert csv2.data == self.sheet_data1

        csv3 = ExcelResponse(self.data3, 'my_data', force_csv=True, font='name SimSum')
        assert csv3.headers['Content-Type'] == self.content_type_csv
        assert csv3.data == self.sheet_data2

    def test_as_xls(self):
        xls1 = ExcelResponse(self.data1, 'my_data', font='name SimSum')
        assert xls1.headers['Content-Type'] == self.content_type_xls
        assert xls1.data == self.sheet_data1

        xls2 = ExcelResponse(self.data2, 'my_data', font='name SimSum')
        assert xls2.headers['Content-Type'] == self.content_type_xls
        assert xls2.data == self.sheet_data1

        # xls3 = ExcelResponse(self.data3, 'my_data', font='name SimSum')
        # assert xls3.headers['Content-Type'] == self.content_type_xls
        # assert xls3.data == self.sheet_data2

        xls11 = ExcelResponse(self.data11, 'my_data', font='name SimSum')
        assert xls11.headers['Content-Type'] == self.content_type_xls
        assert xls11.data == self.sheet_data1

        xls22 = ExcelResponse(self.data22, 'my_data', font='name SimSum')
        assert xls22.headers['Content-Type'] == self.content_type_xls
        assert xls22.data == self.sheet_data1

    def test_as_row_merge_xls(self):
        xls1 = ExcelResponse(self.data1, 'my_data', font='name SimSum', row_merge=True)
        assert xls1.headers['Content-Type'] == self.content_type_xls
        assert xls1.data == self.sheet_data1

        xls2 = ExcelResponse(self.data2, 'my_data', font='name SimSum', row_merge=True)
        assert xls2.headers['Content-Type'] == self.content_type_xls
        assert xls2.data == self.sheet_data1

        xls3 = ExcelResponse(self.data3, 'my_data', font='name SimSum', row_merge=True)
        assert xls3.headers['Content-Type'] == self.content_type_xls
        assert xls3.data == self.sheet_data2

        xls11 = ExcelResponse(self.data1, 'my_data', font='name SimSum', row_merge=True)
        assert xls11.headers['Content-Type'] == self.content_type_xls
        assert xls11.data == self.sheet_data1

        xls22 = ExcelResponse(self.data2, 'my_data', font='name SimSum', row_merge=True)
        assert xls22.headers['Content-Type'] == self.content_type_xls
        assert xls22.data == self.sheet_data1

        xls33 = ExcelResponse(self.data3, 'my_data', font='name SimSum', row_merge=True)
        assert xls33.headers['Content-Type'] == self.content_type_xls
        assert xls33.data == self.sheet_data2
