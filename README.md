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

One of the challenges in this project was properly defining the vol_total column. Large volumes of stocks moved throughout the year, requiring setting the vol_total column as 'LongLong'. This issue illustrates the importance of setting proper data types in VBA: if not set properly, they will stop the code from running. VBA falls on the structured end of a continuum of flexibility in programming languages: data types have to be specifically defined. In this case, there is not much meaning in the distinction between a column being 'LongLong', as opposed to 'Long': understanding about the data is not gained by learning that the data within the column is a larger integer. In other words, a more flexible programming language, where data types do not need to be explicitly defined, would have been just as appropriate to use for analysis as VBA. However, the advantage to using VBA here is its compatability with Excel; one would have to integrate another programming language into Excel if wanting to sidestep the structured nature of VBA.  

Another challenge was using clauses within a for loop in order to mark the change between one stock's data and another. This was accomplished by taking non-intuitive step of setting the for loop to run over the second row of the data to the last row of the data (rw = 2 To lastRow). This takes advantage of the fact that the data in the first row can be captured by referring to 'the row before' (i.e.: the first row) using 'rw-1', and that there are empty rows following the last row with data, allowing for comparison between the last data-bearing row and the row following. Thus there are three main clauses within the for loop: one that applies when it identifies the first row for each stock, another that applies between the first and last row for each stock, and a third that identifies the last row of the whole dataset. 

A third challenge was dealing with zeros in the percent change formula: if the start value of a stock was zero, this would cause the percent change calculation to throw an error, since dividing by zero equals infinity. In order to give an approximation of the percent change, rather than a non-specific 'ERROR', the denominator in the percent formula was set to 1 in the cases where the start value was zero. 