# -*- coding:utf-8 -*-
from setuptools import find_packages, setup

setup(
    name='excel-python',
    version='1.0.0',
    url='https://github.com/caoruiy/excel-python/',
    author='nolly',
    author_email='nollyup@gmail.com',
    description=('A simple API of read and write excel \
        file based on xlrd, xlwt and xlutils'),
    license='MIT',
    packages=find_packages(),
    extras_require={
        'xlrd':['xlrd'],
        'xlwt':['xlwt'],
        'xlutils':['xlutils'],
        'traceback':['traceback']
    }
)

