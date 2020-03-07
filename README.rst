=======================
sphinxcontrib-xlsxtable
=======================

A sphinx extension for making table from Excel file.

- Depends on `OpenPyXL <https://openpyxl.readthedocs.io/en/stable/>`__

  - Supports xlsx file

- Supports merged cell
- Supports Japanese

This extension **generates a grid table string internally** from Excel file.

reStructuredText:

.. code-block:: rst

   .. xlsx-table:: Table Caption
      :file: path/to/xlsx/file.xlsx
      :header-rows: 1

Excel file:

.. image:: https://raw.githubusercontent.com/kkAyataka/sphinxcontrib-xlsxtable/master/sample-excel.png

Rendered HTML:

.. image:: https://raw.githubusercontent.com/kkAyataka/sphinxcontrib-xlsxtable/master/sample-rendering.png


LICENSE
=======

- MIT
