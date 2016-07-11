# -*- coding:utf-8 -*-

"""
A function extends of Tarken's django-excel-response

django-excel-response
1、djangosnippets - http://djangosnippets.org/snippets/1151/
2、pypi - https://pypi.python.org/pypi/django-excel-response/1.0

When using Tarken's django-excel-response. We find that
Chinese is messed code when we open .xls in Mac OS.
As discussed in http://segmentfault.com/q/1010000000095546.
We realize django-excel-response2
Based on Tarken's django-excel-response to solve this problem
By adding a Param named font to set font.
"""

import datetime

import pytz

from django import http
from django.conf import settings
from django.db.models.query import QuerySet
from django.utils import timezone


# ValuesQuerySet and ValuesListQuerySet have been removed.
# https://docs.djangoproject.com/en/1.9/releases/1.9/#miscellaneous
try:
    from django.db.models.query import ValuesQuerySet
    SUPPORT_ValuesQuerySet = True
except ImportError:
    SUPPORT_ValuesQuerySet = False

try:
    from cStringIO import StringIO
except ImportError:
    try:
        from StringIO import StringIO
    except ImportError:
        from io import StringIO


def __init__(self, data, output_name='excel_data', headers=None, force_csv=False, encoding='utf8', font=''):

    # Make sure we've got the right type of data to work with
    # ``list index out of range`` if data is ``[]``
    valid_data = False
    if SUPPORT_ValuesQuerySet and isinstance(data, ValuesQuerySet):
        data = list(data)
    elif isinstance(data, QuerySet):
        data = list(data.values())
    if hasattr(data, '__getitem__'):
        if isinstance(data[0], dict):
            if headers is None:
                headers = data[0].keys()
            data = [[row[col] for col in headers] for row in data]
            data.insert(0, headers)
        if hasattr(data[0], '__getitem__'):
            valid_data = True
    assert valid_data is True, 'ExcelResponse requires a sequence of sequences'

    output = StringIO()
    # Excel has a limit on number of rows; if we have more than that, make a csv
    use_xls = False
    if len(data) <= 65536 and not force_csv:
        try:
            import xlwt
            use_xls = True
        except ImportError:
            # xlwt doesn't exist; fall back to csv
            pass
    if use_xls:
        book = xlwt.Workbook(encoding=encoding)
        sheet = book.add_sheet('Sheet 1')
        styles = {
            'datetime': xlwt.easyxf(num_format_str='yyyy-mm-dd hh:mm:ss'),
            'date': xlwt.easyxf(num_format_str='yyyy-mm-dd'),
            'time': xlwt.easyxf(num_format_str='hh:mm:ss'),
            'font': xlwt.easyxf('%s %s' % ('font:', font)),
            'default': xlwt.Style.default_style,
        }

        for rowx, row in enumerate(data):
            for colx, value in enumerate(row):
                if isinstance(value, datetime.datetime):
                    if timezone.is_aware(value):
                        value = timezone.make_naive(value, pytz.timezone(settings.TIME_ZONE))
                    cell_style = styles['datetime']
                elif isinstance(value, datetime.date):
                    cell_style = styles['date']
                elif isinstance(value, datetime.time):
                    cell_style = styles['time']
                elif font:
                    cell_style = styles['font']
                else:
                    cell_style = styles['default']
                sheet.write(rowx, colx, value, style=cell_style)
        book.save(output)
        content_type = 'application/vnd.ms-excel'
        file_ext = 'xls'
    else:
        for row in data:
            out_row = []
            for value in row:
                if not isinstance(value, basestring):
                    value = unicode(value)
                value = value.encode(encoding)
                out_row.append(value.replace('"', '""'))
            output.write('"%s"\n' % '","'.join(out_row))
        content_type = 'text/csv'
        file_ext = 'csv'
    output.seek(0)
    super(ExcelResponse, self).__init__(output, content_type=content_type)
    self['Content-Disposition'] = 'attachment;filename="%s.%s"' % (output_name.replace('"', '\"'), file_ext)


names = dir(http)


if 'FileResponse' in names:
    ExcelResponse = type('ExcelResponse', (http.FileResponse, ), dict(__init__=__init__))
elif 'StreamingHttpResponse' in names:
    ExcelResponse = type('StreamingHttpResponse', (http.StreamingHttpResponse, ), dict(__init__=__init__))
else:
    ExcelResponse = type('HttpResponse', (http.HttpResponse, ), dict(__init__=__init__))
