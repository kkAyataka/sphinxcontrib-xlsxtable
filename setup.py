# -*- coding: utf-8 -*-

from setuptools import setup, find_packages

long_desc = ''
with open("README.rst", "r") as fh:
    long_desc = fh.read()

setup(
    name='sphinxcontrib-xlsxtable',
    version='0.1.7',
    url='https://github.com/kkAyataka/sphinxcontrib-xlsxtable',
    download_url='http://pypi.python.org/pypi/sphinxcontrib-xlsxtable',
    license='MIT',
    author='kkAyataka',
    author_email='kk.ayataka@gmail.com',
    description='A sphinx extension for making table from Excel file',
    long_description=long_desc,
    long_description_content_type='text/x-rst',
    zip_safe=False,
    classifiers=[
        'Development Status :: 4 - Beta',
        'Environment :: Console',
        'Environment :: Web Environment',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
        'Programming Language :: Python',
        'Framework :: Sphinx :: Extension',
        'Topic :: Documentation',
        'Topic :: Utilities',
    ],
    platforms='any',
    packages=find_packages(),
    include_package_data=True,
    install_requires=[
        'Sphinx',
        'openpyxl'
    ],
    namespace_packages=['sphinxcontrib'],
)
