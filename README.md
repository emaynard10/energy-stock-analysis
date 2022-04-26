# Energy Stock Analysis
Columbia Data Bootcamp analysis of renewable energy stocks with VBA

## Overview of Project
This project analyzes green energy stocks and compares twelve stocks over two different years, 2017 and 2018 in order to determine the best return on investment. By comparing the stocks over two years there should be a better chance of choosing a stock that may have a positive return on investment. Many of the stocks did much better in 2017 which may indicate some volutility in the green energy stock market. If more data is gathered from subsequent years and added to the excel sheet, the project can be updated to run a similar analysis on those years. By refactoring the original code, run time for the All Stocks Analysis is decreased and made more efficient. It would be worthwhile to get a better idea of the stocks if they were tracked over the course of more that two years; a reasonable way the beter understand which stocks could be a good long term investment would be to see the five and ten year returns for each stock. 

## Results
The results show that over the two years there were two stocks with did well in both years: ENPH and RUN. Based on this the recommendation is not to invest in the DQ stock, which had a great return on investment in 2017, but did not perform well in 2018. It is often advised to diversify investments so investing in both ENPH and RUN.

### Stock Analysis in 2017
![Screen Shot 2022-04-26 at 9 17 09 AM](https://user-images.githubusercontent.com/99676466/165364178-9809ba59-d313-4acb-9a7d-794f8e21e12b.png)

### Stock Analysis in 2018
![Screen Shot 2022-04-26 at 9 17 46 AM](https://user-images.githubusercontent.com/99676466/165363517-2e977a16-e45f-43a6-bdef-b9576791b907.png)

The limitations on this analysis are the short period of study; 2018 saw a significant decrease in return for all the stocks, so it would be worth invstigating why green energy stocks did not perform well this year before making a decision. 

The refactored code decreased run time for analyzing the stocks by changing the code from looping through all the rows of data over and over to going through the data once and looking for specific index criteria. while nested for loops can be a powerful way to look through data over and over in this case the code does three things while looking through once. As the code runs through each row, it adds the volume, checks if the row switches to a different stock in order to set a start price, and checks for the end of that stocks data in order to set an ending price. Once it is finished with a stock, it increases the ticker to move on the next stock. Assigning a tickerIndex and using this index in the code was key to decreasing the run time. Rather than nesting a loop within another loop, the refactored code ends the first for loop before looping through the rows 

## Summary

### 1. What are the advantages or disadvantages of refactoring code?
  Some advatages of refactoring the code is to decrease the run time for the analysis. The code becomes a bit more versatile in the sense that if it were applied to larger data sets in would not have to loop through the data over long periods of time. Another advantage is that the code is cleaner and easier to read. The code couldbe applied to other sheets with data from subsequent years; for example if a sheet for 2019 was added to the data set with the same 12 stocks, it would be easy to run the code with the input box asking which year.
  Disadvantages of the refactored code is that in its current state it would still need to be modified if more stocks were added because the code is only written for 12 stocks. Analyzing more than twleve stocks may give a better picture of what is happening in the green energy market. If there was data for the year 2019 with a different set of stocks, the code would have to be reworked to include arrays for the new stocks. Refactoring code can take time, though running a slow code over and over the time could also add up.
### 2. How do these pros and cons apply to refactoring the original VBA script?
  Refactoring the original VBA stript is an easier way to edit the code since there is already a starting point, but may limit the code and what it can do for future datasets. This dataset is relatively small so the refactored code decreased the run time incrementally; if the dataset was much larger the refactoring would be even more important to be able to efficiently run through all the stock data.
