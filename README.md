# Welcome to similarityFinder!

The objective of this python program is to make copying data between two MS Excel workbooks easier. It matches cell values in one/two columns between the workbooks, and outputs the correspoding required column from the input workbook to the output workbook.

Similar to XlookUp but the main feature is : Descriptions do not have to be an equal match, we use the python library FuzzyWuzzy to estimate if two descriptions are similar enough to be considered the same despite having typos and so on.

Note:
* A "CHECK" mark in an output sheet's cell indicates that this cell did not match any input sheet's cells and may require manual entry. 
