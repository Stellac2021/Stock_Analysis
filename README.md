# Stock_Analysis

## Overview of Project
An original VBA script was created to analyze stock information and the stock trading volumes to determine if a stock is worth investing. To improve work efficiency of the stock analysis, I have refactored the coding based on the original VBA script. 

### Purpose
The purpose of this project is to examine whether refactored VBA script runs faster than the origianl VBA script.

The stock data used in this project are 12 stock information from 2017 and 2018 including the stock tickers, ticke rvalue, stock issued date, starting price, closing price and stock trading volumes. 

## Results

Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

### Analysis
 - **Original script** running time 

 ![](Resources/Original_scirpt_2017.png)
 ![](Resources/Original_scirpt_2018.png)
 
 - **Refactored script** running time 
 
 ![](Resources/VBA_Challenge_2017.png)
 ![](Resources/VBA_Challenge_2018.png)

:point_right: Above is the comparison of program run time comparison between original script and refactored script. As shown in the images, the run time is more than 10 times faster after code refactoring. 

The main refactored area was the loop section. 
  - **Original script**
    
      >'4)Loop through the tickers
        For i = 0 To 11
         ticker = tickers(i)
         totalVolume = 0 
        
         '5) loop through rows in the data
          Worksheets("2018").Activate
          For j = 2 To RowCount
            '5a) Get total volume for current ticker
             If Cells(j, 1).Value = ticker Then

                totalVolume = totalVolume + Cells(j, 8).Value
           End If
            '5b) get starting price for current ticker
             If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
               startingPrice = Cells(j, 6).Value
             End If
            '5c) get ending price for current ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
               endingPrice = Cells(j, 6).Value
           End If
      

## Summary

What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?
### Advantages
### Disadvantages
