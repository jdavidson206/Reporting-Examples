# Reporting-Examples
## Example of Excel Macros and Pivot Tables

This site contains examples of Excel workbooks that use VBA Macros and Pivot Tables to generate reports.  
The main workbook, SampleReports.xlsm can be downloaded here:


The Macro code can be separately viewed here:


The following comments apply to the SampleReports workbook and its functionality.  The "Introduction" worksheet in the workbook contains detailed documentation on the worksheets and their functionality.

**Introduction**

The purpose of this workbook is to provide sales and support revenue results, including pivot tables of the results.  The worksheets contain Macros that perform the calculations and comparisons used in this document.  

It is required for you to have your Excel options set to allow Macros to run.  
The workbook contains instructions for allowing Macros to run in Excel.

The data in this workbook was downloaded from the web:
'http://vader.lab.asu.edu/education/papers/2015/Sample%20-%20Superstore%20Sales%20(Excel).xls

The original workbook came with these worksheets:  Orders, Returns, Users.  All other worksheets have been designed and developed by the author of this workbook.

**Worksheets and Content**

This workbook contains the following worksheets:

* Introduction -- instructions for using the workbook; this worksheet

* Sales -- contains the initial raw data set from Orders worksheet; added a "Managers" column to identify the manager for the sale

* Returns -- contains a list of returned order id's and the status "Returned"

* Users -- contains a list of Managers, their regions, and provinces; changed names, added link between regions and provinces.

* Col List -- a list of the columns of the worksheets used in a Macros; translates column letter to number.

* Macro -- contains a select group of data from the Sales worksheet; validate that the base margin for each product contains the same value; processes returned orders.

* BaseCheck -- contains order id, product name, and base margin from Macro worksheet; used in base margin validation macro.

* SuppRev -- contains product and quantity and calculates the support revenue based on the support values in the worksheets table.

* Pivot Tables -- the workbook contains the following pivot tables

   * PT Sales by Customer & Order -- with discounts and costs
   * PT SuppRev -- support revenue by product subcategory
   * PT Sales by Prod -- sales by product with discounts and costs
   * PT Sales by Manager -- sales by manager with discounts and costs
   * PT Sales by Month -- sales by month, quarter, year with discounts and costs

* Data -- the raw data on the Macro workbook; can be copied to the "Macro" worksheet to rerun the macros.

**View Macro Code**

To view the Macro code, click on "Developer" on the menu, click on "Visual Basic" on the Developer ribbon.
If "Developer" does not appear on the menu, see instructions to display.

**Workbook Macros**

The workbook contains macros to perform the following functions:

* Validate the base margin rate is the same for each product
* Accept the validated base margins
* Process returned orders, and reduce sales amounts
* Calculate the Support Revenue for each order

Additionally, the macros also

* Update the pivot tables
* Save the workbook

**How to Run the Macros**

The Macro worksheet has a "Validate Margins" button to run the Validate Margins macros.

The Macro worksheet has a "Process Returns" button to run the Process Returns macro.

The BaseCheck worksheet has a "Accept Changes" button to run the Accept Changes macro; this must be run after the Validate Margins macro.

The SuppRev worksheet has a "Process Support Revenue" button to calculate the product support revenue; should be run after the Process Returns macro for best results.
