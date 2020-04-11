=======================
sphinxcontrib-xlsxtable
=======================

A sphinx extension for making table from Excel file.

- Depends on `OpenPyXL <https://openpyxl.readthedocs.io/en/stable/>`__

  - Supports xlsx file

- Supports merged cell
- Supports Japanese

This extension **generates a grid table string internally** from Excel file.


Install and Set up
==================

Install from PyPI.

.. code-block::

   $ pip install sphinxcontrib-xlsxtable

Configure conf.py

.. code-block:: python

   # conf.py
   extensions = [
       'sphinxcontrib.xlsxtable',
   ]


Usage
=====

reStructuredText directive:

.. code-block:: rst

   .. xlsx-table:: Table Caption
      :file: path/to/xlsx/file.xlsx
      :header-rows: 1

Excel file:

.. image:: https://raw.githubusercontent.com/kkAyataka/sphinxcontrib-xlsxtable/master/sample-excel.png

Rendered HTML:

.. image:: https://raw.githubusercontent.com/kkAyataka/sphinxcontrib-xlsxtable/master/sample-rendering.png


Options
=======

.. contents::
   :local:


Caption (optional)
------------------

Specifies table caption string.

.. code-block:: rst

   .. xlsx-table:: Table Caption
      :file: path/to/xlsx/file.xlsx


\:file: (required)
------------------

Specifies path to Excel file. You can use relative path.

.. code-block:: rst

   .. xlsx-table::
      :file: path/to/xlsx/file.xlsx


\:header-rows: (optional)
-------------------------

Specified the number of lines are used as header.

.. code-block:: rst

   .. xlsx-table::
      :file: path/to/xlsx/file.xlsx
      :header-rows: 1


\:sheet: (optional)
-------------------

Generates a table from a sheet with the specified sheet name.

If this option is not specified, current active sheet is used.

.. code-block:: rst

   .. xlsx-table::
      :file: path/to/xlsx/file.xlsx
      :sheet: Sheet1


\:start-row: (optional)
-----------------------

Specifies start row.

.. code-block:: rst

   .. xlsx-table::
      :file: path/to/xlsx/file.xlsx
      :start-row: 2


CLI
===

You can use from CLI.

.. code-block::

   $ python -m sphinxcontrib.xlsxtable --sheet=Sheet1 --header-rows=1 test/_res/sample.xlsx
   +----+-------+-------+--------+
   | A1 | B1    | C1    | D1     |
   +----+-------+-------+--------+
   | A2 | B2:B3 | C2    | D2     |
   +----+       +-------+--------+
   | A3 |       | C3:D3          |
   +----+-------+-------+--------+
   | A4 | B4    | C4    | - D4-1 |
   |    |       |       | - D4-2 |
   +----+-------+-------+--------+


LICENSE
=======

- MIT
