# -*- coding: utf-8 -*-

from __future__ import with_statement
from setuptools import setup


version = '2.0.1'


setup(
    name='CodeConvert',
    version=version,
    keywords='CodeConvert unicode utf8 utf-8 gbk latin1 raw_unicode_escape',
    description="Code Convert for Humans",
    long_description='CodeConvert is a simple code convert script(library) for Python, built for human beings.',

    url='https://github.com/Brightcells/CodeConvert',

    author='Hackathon',
    author_email='kimi.huang@brightcells.com',

    py_modules=['CodeConvert'],
    install_requires=[],

    classifiers=[
        'Development Status :: 5 - Production/Stable',
        'Intended Audience :: Developers',
        'Programming Language :: Python',
        'Topic :: Software Development :: Libraries :: Python Modules',
    ],
)
