# -*- coding: utf-8 -*-
from __future__ import with_statement
from setuptools import setup


version = '1.0.1'


setup(
    name='django-excel-response2',
    version=version,
    keywords='django-excel-response django-excel-response2',
    description="A function extends of Tarken's django-excel-response",
    long_description='',

    url='https://github.com/Brightcells/django-excel-response2',

    author='Hackathon',
    author_email='kimi.huang@brightcells.com',

    py_modules=['excel_response2'],
    install_requires = ['xlwt', 'pytz'],

    classifiers=[
        'Development Status :: 5 - Production/Stable',
        'Environment :: Web Environment',
        'Framework :: Django',
        'Intended Audience :: Developers',
        'Programming Language :: Python',
        'Topic :: Software Development :: Libraries :: Python Modules',
        'Topic :: Office/Business :: Financial :: Spreadsheet',
    ],
)
