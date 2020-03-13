# django-excel-response2
A function extends of Tarken's django-excel-response

## Inherit

    # Since Version 2.0.2
    if 'FileResponse' in names:
        ExcelResponse = type('ExcelResponse', (http.FileResponse, ), dict(__init__=__init__))
    elif 'StreamingHttpResponse' in names:
        ExcelResponse = type('StreamingHttpResponse', (http.StreamingHttpResponse, ), dict(__init__=__init__))
    else:
        ExcelResponse = type('HttpResponse', (http.HttpResponse, ), dict(__init__=__init__))


## Installation

    pip install django-excel-response2


## Usage

    from django_excel_response import ExcelResponse

    def excelview(request):
        objs = SomeModel.objects.all()
        return ExcelResponse(objs)


or

    from django_excel_response import ExcelResponse

    def excelview(request):
        data = [
            ['Column 1', 'Column 2'],
            [1, 2],
            [3, 4]
        ]
        return ExcelResponse(data, 'my_data', font='name SimSum')


or

    from django_excel_response import ExcelResponse

    def excelview(request):
        data = [
            ['Column 1', 'Column 2'],
            [1, [2, 3]],
            [3, 4]
        ]
        return ExcelResponse(data, 'my_data', font='name SimSum', row_merge=True)


## Params

  * font='name SimSum'
    * Set Font as SimSum(宋体)
  * force_csv=True
    * CSV Format? True for Yes, False for No, Default is False


## CSV

  ```python
  datas = [
      [u'中文', ]
  ]
  ```

|                 | Win Excel 2013 | Mac Excel 2011 | Mac Excel 2016 | Mac Numbers |
| --------------- | :------------: | :------------: | :------------: | :---------: |
| UTF8            | Messy          | Messy          | Messy          | Normal      |
| GB18030         | Normal         | Normal         | Normal         | Messy       |
| UTF8 + BOM_UTF8 | Normal         | Messy          | Normal         | Normal      |
| UTF16LE + BOM   |                |                |                |             |
