Test / Sample for the sphinxcontrib-xlsxtable
=============================================

Sample 1
--------

- Includes vertical merged cell
- Includes horizontal merged cell
- Includes Japanese
- Includes symbol (east asian width is "A")

.. xlsx-table:: Sample 1
   :file: _res/sample.xlsx
   :header-rows: 1
   :sheet: Sheet1


Sample 2
--------

- Specifies sheet name
- Has a merged cell at last row

.. xlsx-table:: Sample 2
   :file: _res/sample.xlsx
   :sheet: Sheet2


Sample 3
--------

- Specifies start-row
- Specifies start-column
- Specifies header-rows

.. xlsx-table:: Sample 3.1
   :file: _res/sample.xlsx
   :sheet: Sheet3
   :start-row: 2
   :start-column: 2

.. xlsx-table:: Sample 3.2
   :file: _res/sample.xlsx
   :sheet: Sheet3
   :start-row: 4
   :header-rows: 2
