=====================
django-excel-response
=====================

django-excel-response
=====================

A subclass of HttpResponse which will transform a QuerySet,
or sequence of sequences, into either an Excel spreadsheet or
CSV file formatted for Excel, depending on the amount of data.
All of this is done in-memory and on-the-fly, with no disk writes,
thanks to the StringIO library.

* DjangoSnippets - http://djangosnippets.org/snippets/1151/
* PyPI - https://pypi.python.org/pypi/django-excel-response/1.0

django-excel-response2
======================

When using Tarken’s django-excel-response.
We find that Chinese is messed code when we open .xls in Mac OS.
As discussed in http://segmentfault.com/q/1010000000095546.
We realize django-excel-response2 Based on Tarken’s django-excel-response
to solve this problem By adding a Param named font to set font.

At The Same Time:

* Fix Bug
    * can't subtract offset-naive and offset-aware datetimes

Installation
============

::

    pip install django-excel-response2


Usage
=====

::

    from excel_response2 import ExcelResponse

    def excelview(request):
        objs = SomeModel.objects.all()
        return ExcelResponse(objs)


or::

    from excel_response2 import ExcelResponse

    def excelview(request):
        data = [
            ['Column 1', 'Column 2'],
            [1, 2],
            [3, 4]
        ]
        return ExcelResponse(data, 'my_data', font='name SimSum')


Params
======

* font='name SimSum'
    * Set Font as SimSum(宋体)
* force_csv=True
    * CSV Format? True for Yes, False for No, Default is False
