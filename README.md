# Stock-Analysis

## Overview of Project
### Purpose
The purpose of this project is to empoyer my client, Steve, to automatically calculate total number of shares traded and annual returns for selected stock data in 2017 and 2018. With this data, Steve can help his parents analyze the results to determine how certains stocks did each year and if there are any changes they want to make in their stock portfolio. 

Additionally, the project involves reviewing and refactoring the code used for automation, in hopes for efficiency gains. For example, if the refactored code runs faster, it will be more appealing to use it to process large volumes of stock data. This would be a huge benefit for Steve, allowing him to expand his analysis to the entire stock market for any year.
### Results
#### Stock Performance Between 2017 - 2018
In order to compare the stock performance between 2017 and 2018:
- I first had to roll-up the raw stock data for each year. I created a subroutine to loop through the 2018 data for each of the 12 stock "tickers" and output each ticker's total volume and annual return to the **All Stocks Analysis** worksheet of the [VBA Challenge spreadsheet](/VBA_Challenge.xlsm).
  - The subroutine also included special formatting to display positive annual stock returns in green and negative annual stock returns in red. This allows the user to very quickly determine which stocks had better annual outcomes.
- I then created a new subroutine to allow the user to select between 2017 and 2018 stock data by allowing them to input the year. I created a button in the **All Stocks Analysis** worksheet to allow a user to easily execute this new subroutine. 
- Now that stock data can easily be generated for both 2017 and 2018, we can compare the results.
  - Results from 2017 ![2017 Stock Analysis](/VBA_Challenge_2017.png)  
  - Observations of 2017 data include:
    - 11 of the 12 stocks had postive annual returns
    - 2 of the 11 postive stocks saw single digit returns
    - 5 of the 11 postive stocks saw double digit returns
    - 4 of the 11 postive stocks saw triple digit returns
  - Results from 2018 ![2018 Stock Analysis](/VBA_Challenge_2018.png)  
  - Observations of 2018 data include:
    - Only 2 of the 12 stocks had postive annual returns, 10 had negative annual returns
    - 2 of the 11 postive stocks saw single digit returns
    - 5 of the 11 postive stocks saw double digit returns
    - 4 of the 11 postive stocks saw triple digit returns
  - Comparisons of 2017 and 2018 data
    - Stock tickers ENPH and RUN were the only 2 stocks that saw positive returns in both years
    - ENPH had a triple digit return in 2017 and a double digit return in 2018
    - RUN had a single digit return in 2017 and a double digit return in 2018
  - Recommendation
    - Based on analysis, I'd recommend to my client Steve that his parents consider investing in ENPH and RUN stocks
    
#### Comparing Execution Times Between Original & Refactored Code

### Summary
- What are the advantages or disadvantages of refactoring code?
  - Advantages include:
    - The code is easier for the author and others to read and understand
    - The code is easier to perform troubleshooting to catch issues
    - The code is easier to maintain over time
  - Disadvantages include:
    - The actiion of refactoring can become very time-consuming 
    - Time spent might not be worth the positive gain from refactoring
    - There's a risk of the person refactoring causing bugs in existing code
  
- How do these pros and cons apply to refactoring the original VBA script?
  - 


