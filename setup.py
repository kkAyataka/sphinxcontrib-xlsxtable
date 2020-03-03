# -*- coding: utf-8 -*-

from setuptools import setup, find_packages

long_desc = '''
This package contains the xlsxtable Sphinx extension.

.. add description here ..
'''

requires = ['Sphinx>=0.6']

setup(
    name='sphinxcontrib-xlsxtable',
    version='0.1',
    url='https://github.com/kkAyataka/sphinxcontrib-xlsxtable',
    download_url='http://pypi.python.org/pypi/sphinxcontrib-xlsxtable',
    license='MIT',
    author='kkAyataka',
    author_email='kk.ayataka@gmail.com',
    description='Sphinx "xlsxtable" extension',
    long_description=long_desc,
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
        #'Framework :: Sphinx :: Theme',
        'Topic :: Documentation',
        'Topic :: Utilities',
    ],
    platforms='any',
    packages=find_packages(),
    include_package_data=True,
    install_requires=requires,
    namespace_packages=['sphinxcontrib'],
)
