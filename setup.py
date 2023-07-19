# -*- coding: utf-8 -*-

from __future__ import with_statement

from setuptools import setup


version = '3.0.6'


setup(
    name='django-excel-response2',
    version=version,
    keywords='django-excel-response django-excel-response2',
    description="A function extends of Tarken's django-excel-response",
    long_description=open('README.md').read(),
    long_description_content_type='text/markdown',

    url='https://github.com/django-xxx/django-excel-response2',

    author='Hackathon',
    author_email='kimi.huang@brightcells.com',

    packages=['django_excel_response'],
    py_modules=[],
    install_requires=['django-six>=1.0.4', 'excel-base>=1.0.3'],

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
