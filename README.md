# stock-analysis.
Module 2. VBA
## Overview of Project

The purpose of this analysis is to help Steve analize a data set of green energy stocks from the years 2017 and 2018, as his parents are interested in investing their money in a company DQ New Energy Corp (DQ). Steve promised to look into DQ stock along with several other companies for his parents. In order to assist Steve, a VBA script was developed in Excel to analyse the data set and provide the total daily volume and yearly return for each stock. 

Steve loved the first macro that was created, and in the future he is planning to perform the analysis on larger data sets, therefore, It was decided to refactor the script to see if it is possible to reduce the processing time.

## Results
### Analysis of stocks
- Data set information:
  - Two data set (years 2017 and 2018)
  - Data from 12 green energy companies (tickers)
  - Each data set contaided 3012 rows
  - Attribute information:
    - Ticker, Date, Open, High, Low,	Close,	Adj Close and	Volume
   
The focus of this analysis is to visualize campaign which stocks are the most profitable. In order to do this, the following steps needed to be taken prior to conducting the analysis:

1. Create a draft (pseudocode) 
<img src="/Resources/img1.png" width="50%" height="50%">

2. *Define time variables* and ask for the year to perform the analysis (*inputbox*)
<img src="/Resources/img2.png" width="50%" height="50%">

3. Measure code performance.
  - 3.1 Underneath the *yearValuevariable* set the *startTime* variable equal to the **Timer function**, which will allow us to start the clock
  - 3.2 After the last *Next i* and before the *End Sub* command, set the endTime variable equal to the **Timer function**.
  - 3.3 Create a *messagebox* that displays the elapsed time
  <img src="/Resources/img3.png" width="50%" height="50%">

4. Format the output sheet
  - 4.1 Activate the output worksheet 
  - 4.2 Add headers
  <img src="/Resources/img4.png" width="50%" height="50%">
  
5. Assign each of the tickers to an element in an array
<img src="/Resources/img5.png" width="50%" height="50%">

6. Depending on the year selected in the inputbox, the worksheet is activated and get the number of rows in the worksheet to loop over
<img src="/Resources/img6.png" width="50%" height="50%">

7. Create an index variable and set it equal to zero before iterating over all the rows. Also create the output arrays
<img src="/Resources/img7.png" width="50%" height="50%">

8. Create nested loops to run analyses on all of the stocks
  - Step 2a: Create a for loop to initialize the *tickerVolumes* to zero.
  - Step 2b: Create a for loop that will loop over all the rows in the spreadsheet.
  - Step 3a: Write a script that increases the current *tickerVolumes* variable and adds the ticker volume for the current stock ticker.
  - Step 3b: Write an if-then statement to check if the current row is the first row with the selected *tickerIndex*. If it is, then assign the current starting price to the tickerStartingPrices variable.
  - Step 3c: Write an if-then statement to check if the current row is the last row with the selected *tickerIndex*. If it is, then assign the current closing price to the tickerEndingPrices variable.
  - Step 3d: Increase the *tickerIndex* if the next row’s ticker doesn’t match the previous row’s ticker.
  <img src="/Resources/img8.png" width="50%" height="50%">

9. Use a for loop to loop through your arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the “Ticker,” “Total Daily Volume,” and “Return” columns in your spreadsheet.
<img src="/Resources/img9.png" width="50%" height="50%">

10. Use visual and numeric formatting in outputs for the selected year stock analysis
    - the green color indicates that the result is positive; if the result is negative, it is indicated with red
<img src="/Resources/img10.png" width="50%" height="50%">

11. Assign the macro/vba script to a control button 
<img src="/Resources/img11.png" width="50%" height="50%">


----------------------------------------------------------------------------------------------------------------------------------------------------------------
### Analysis outcome
After running the code for the stock analysis in both years, the output result looks like this:
- 2017
<img src="/Resources/img2017.png" width="30%" height="30%">

- 2018
<img src="/Resources/img2018.png" width="30%" height="30%">

*Note: Green stocks indicate a positive return; if the result is negative, it is indicated with red*

Almost all stocks in 2017 offered a positive return and it is observed that DQ was the company that showed the highest growth at 199.4%, however, as can be seen, the majority of stocks fell in 2018, with DQ being the company that had the largest drop in its shares by 62.2%

It is recommended not to invest in DQ, and from this analysis, it can be seen that these stocks are not a safe first place for an investment, with the exception of RUN, which gained 81.9% and could be a good option to invest.


### VBA Performance (Refactor VBA code)
The data was processed using a VBA script in Excel. Using the original script, the queries were performed in approximately 0.16 seconds for both 2017 and 2018 datasets.

## Summary
Code refactoring is the process of restructuring the original script without changing its external behavior. Refactoring is intended to improve the design, structure, and implementation of the code while preserving its functionality.

### Advantages and disadvantages of refactoring code in general
- Advantages
  -  Improved code readability 
  -  Improved source-code maintenance and scalability
  -  It makes code easier to understand
    
- Disadvantages  
  - Invest time in developing it
  - Chance of mistakes 
  - It's risky when developers do not understand what's all about
    
