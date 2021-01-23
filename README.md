# vba-challenge
### Contains
- Sub-stockCheckAllSheets().vb
- images folder: 
    - 2014subRun.png
    - 2015subRun.png
    - 2016subRun.png

### Description
The purpose of this analysis was to create a vba file that was capable of looping through similarly formatted excel spreadsheets and returning:

- the yearly change, percent change, and total volume of all stocks, with conditional formatting on the yearly change column to show whether an increase or decrease occurred
- the stock with the greatest % increase and greatest % decrease, and the value of that increase/decrease

The Sub-stockCheckAllSheets().vb contains the code created to achieve this purpose, and the images folder contains screenshots of the spreadsheets with the outputs resulting from running the code. 
### Challenges

One of the challenges in this project was properly defining the vol_total column. Large volumes of stocks moved throughout the year, requiring setting the vol_total column as 'LongLong'. This issue illustrates the importance of setting proper data types in VBA: if not set properly, they will stop the code from running. 

A third challenge was dealing with zeros in the percent change formula: if the start value of a stock was zero, this would cause the percent change calculation to throw an error, since dividing by zero equals infinity. In order to give an approximation of the percent change, rather than a non-specific 'ERROR', the denominator in the percent formula was set to 1 in the cases where the start value was zero. 