=======================
sphinxcontrib-xlsxtable
=======================

A sphinx extension for making table from Excel file.

- Depends on `OpenPyXL <https://openpyxl.readthedocs.io/en/stable/>`__

  - Supports xlsx file

- Supports merged cell
- Supports images
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

Specifies start row number.

.. code-block:: rst

   .. xlsx-table::
      :file: path/to/xlsx/file.xlsx
      :start-row: 2


\:start-column: (optional)
--------------------------

Specifies start column number.

.. code-block:: rst

   .. xlsx-table::
      :file: path/to/xlsx/file.xlsx
      :start-column: 2


\:include-rows: / :exclude-rows: (optional)
-------------------------------------------

Specifies include or exclude rows.
Exclude setting has priority.

.. code-block:: rst

   .. xlsx-table::
      :file: path/to/xlsx/file.xlsx
      :include-rows: 1-2 4 8
      :exclude-rows: 3 5-7


\:include-columns: / :exclude-columns: (optional)
-------------------------------------------------

Specifies include or exclude columns.
Exclude setting has priority.

.. code-block:: rst

   .. xlsx-table::
      :file: path/to/xlsx/file.xlsx
      :include-columns: A-B 4
      :exclude-columns: C 5-6


CLI
===

You can use from CLI.

.. code-block::

   $ python -m sphinxcontrib.xlsxtable --sheet=Sheet1 --header-rows=1 test/_res/sample.xlsx
   +----+-------+-------+--------+
   | A1 | B1    | C1    | D1     |
   +====+=======+=======+========+
   | A2 | B2:B3 | C2    | D2     |
   +----+       +-------+--------+
   | A3 |       | C3:D3          |
   +----+-------+-------+--------+
   | A4 | B4    | C4    | - D4-1 |
   |    |       |       | - D4-2 |
   +----+-------+-------+--------+


Links
=====

- `sphinxcontrib-xlsxtableの解説 <https://kkayataka.hatenablog.com/entry/2020/03/14/140305>`__
- `sphinxcontrib-xlsxtableのモジュール実行 <https://kkayataka.hatenablog.com/entry/2020/04/11/173717>`__
- `sphinxcontrib-xlsxtableに行・列指定オプションを追加 <https://kkayataka.hatenablog.com/entry/2020/07/25/131440>`__
- `sphinxcontrib-xlsxtableの画像対応 <https://kkayataka.hatenablog.com/entry/2023/07/16/231550>`__

LICENSE
=======

- MIT
