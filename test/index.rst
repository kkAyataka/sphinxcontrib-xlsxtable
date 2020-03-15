Test / Sample for the sphinxcontrib-xlsxtable
=============================================

Sample 1
--------

- Includes vertical merged cell
- Includes horizontal merged cell
- Includes Japanese
- Includes symbol (east asian width is "A")

.. xlsx-table:: Sample Table
   :file: _res/sample.xlsx
   :header-rows: 1
   :sheet: Sheet1


Sample 2
--------

- Specifies sheet name
- Has a merged cell at last row

.. xlsx-table:: Sample Table
   :file: _res/sample.xlsx
   :sheet: Sheet2
