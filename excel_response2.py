# -*- coding:utf-8 -*-

import datetime

import pytz
import xlwt

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


# Min (Max. Rows) for Widely Used Excel
# http://superuser.com/questions/366468/what-is-the-maximum-allowed-rows-in-a-microsoft-excel-xls-or-xlsx
EXCEL_MAXIMUM_ALLOWED_ROWS = 65536
# Column Width Limit For ``xlwt``
# https://github.com/python-excel/xlwt/blob/master/xlwt/Column.py#L22
EXCEL_MAXIMUM_ALLOWED_COLUMN_WIDTH = 65535

# https://github.com/urwid/urwid/blob/master/urwid/old_str_util.py
WIDTHS = [
    (126, 1),
    (159, 0),
    (687, 1),
    (710, 0),
    (711, 1),
    (727, 0),
    (733, 1),
    (879, 0),
    (1154, 1),
    (1161, 0),
    (4347, 1),
    (4447, 2),
    (7467, 1),
    (7521, 0),
    (8369, 1),
    (8426, 0),
    (9000, 1),
    (9002, 2),
    (11021, 1),
    (12350, 2),
    (12351, 1),
    (12438, 2),
    (12442, 0),
    (19893, 2),
    (19967, 1),
    (55203, 2),
    (63743, 1),
    (64106, 2),
    (65039, 1),
    (65059, 0),
    (65131, 2),
    (65279, 1),
    (65376, 2),
    (65500, 1),
    (65510, 2),
    (120831, 1),
    (262141, 2),
    (1114109, 1),
]


def get_width(o):
    """ Return the screen column width for unicode ordinal o. """
    if o == 0xe or o == 0xf:
        return 0
    for num, wid in WIDTHS:
        if o <= num:
            return wid
    return 1


def calc_width(s):
    """ Should Improved """
    return sum([get_width(ord(i)) for i in s])


@property
def as_xls(self):
    book = xlwt.Workbook(encoding=self.encoding)
    sheet = book.add_sheet(self.sheet_name)

    styles = {
        'datetime': xlwt.easyxf(num_format_str='yyyy-mm-dd hh:mm:ss'),
        'date': xlwt.easyxf(num_format_str='yyyy-mm-dd'),
        'time': xlwt.easyxf(num_format_str='hh:mm:ss'),
        'font': xlwt.easyxf('%s %s' % ('font:', self.font)),
        'default': xlwt.Style.default_style,
    }

    widths = {}
    for rowx, row in enumerate(self.data):
        for colx, value in enumerate(row):
            if value is None and self.blanks_for_none:
                value = ''

            if isinstance(value, datetime.datetime):
                if timezone.is_aware(value):
                    value = timezone.make_naive(value, pytz.timezone(settings.TIME_ZONE))
                cell_style = styles['datetime']
            elif isinstance(value, datetime.date):
                cell_style = styles['date']
            elif isinstance(value, datetime.time):
                cell_style = styles['time']
            elif self.font:
                cell_style = styles['font']
            else:
                cell_style = styles['default']

            sheet.write(rowx, colx, value, style=cell_style)

            if self.auto_adjust_width:
                width = calc_width(value) * 256 if isinstance(value, basestring) else calc_width(str(value)) * 256
                if width > widths.get(colx, 0):
                    width = min(width, self.EXCEL_MAXIMUM_ALLOWED_COLUMN_WIDTH)
                    widths[colx] = width
                    sheet.col(colx).width = width

    book.save(self.output)


@property
def as_csv(self):
    for row in self.data:
        out_row = []
        for value in row:
            if value is None and self.blanks_for_none:
                value = ''
            if not isinstance(value, basestring):
                value = unicode(value)
            value = value.encode(self.encoding)
            out_row.append(value.replace('"', '""'))
        self.output.write('"%s"\n' % '","'.join(out_row))


def __init__(self, data, output_name='excel_data', headers=None, force_csv=False, encoding='utf8', font='', sheet_name='Sheet 1', blanks_for_none=True, auto_adjust_width=True):

    self.data = data
    self.output_name = output_name
    self.headers = headers
    self.force_csv = force_csv
    self.encoding = encoding
    self.font = font
    self.sheet_name = sheet_name
    self.blanks_for_none = blanks_for_none
    self.auto_adjust_width = auto_adjust_width

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

    self.output = StringIO()
    # Excel has a limit on number of rows; if we have more than that, make a csv
    use_xls = True if len(self.data) <= self.EXCEL_MAXIMUM_ALLOWED_ROWS and not self.force_csv else False
    _, content_type, file_ext = (self.as_xls, 'application/vnd.ms-excel', 'xls') if use_xls else (self.as_csv, 'text/csv', 'csv')
    self.output.seek(0)
    super(ExcelResponse, self).__init__(self.output, content_type=content_type)
    self['Content-Disposition'] = 'attachment;filename="%s.%s"' % (self.output_name.replace('"', '\"'), file_ext)


names = dir(http)


clsdict = {
    'EXCEL_MAXIMUM_ALLOWED_ROWS': EXCEL_MAXIMUM_ALLOWED_ROWS,
    'EXCEL_MAXIMUM_ALLOWED_COLUMN_WIDTH': EXCEL_MAXIMUM_ALLOWED_COLUMN_WIDTH,
    '__init__': __init__,
    'as_xls': as_xls,
    'as_csv': as_csv,
}


if 'FileResponse' in names:
    ExcelResponse = type('ExcelResponse', (http.FileResponse, ), clsdict)
elif 'StreamingHttpResponse' in names:
    ExcelResponse = type('StreamingHttpResponse', (http.StreamingHttpResponse, ), clsdict)
else:
    ExcelResponse = type('HttpResponse', (http.HttpResponse, ), clsdict)
