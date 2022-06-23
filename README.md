# Challenge 2 / Stock Analysis 

## Overview of Project
This Stock analysis is to determine which of the stocks being monitored by Steve and his family are more likely to be a good investment.  With the information available a daily volume and percentage return per year per ticker could be determined by creating a Macros (VBA Script) that could analyze the data. A macros  is a set of rules included in the background of an excel sheet that preforms tasks (in this case analyzing stock data) included loops, arrays and commands that  process the information and provided an outcome. 

## Data 
Determining which stocks were the better investment would be based on ticker data recorded for the year 2017 and 2018. A stock ticker is a series of letters assigned to a specific stock as an identifier. The stock ticker reports transactions and price data for the specific stock throughout the day (Hayes) the data reported includes the date, the opening price, the high and low for said date, the closing price, the adjusted closing price (adjusts for dividends stocks splits etc.) and the Volume.

## Analysis and Challenges
To analyze the data a Macros was created following the steps provided. The Macros included adding an index of tickers, creating loops, and formatting the sheet where the data was displayed. The display sheet was automated with the addition of input boxes. The original Macros was refracted to make the code more efficient. Loops were added, the loops collect the data for the specific ticker for both years and calculates the total daily value and the return percentage. An input Box added to the sheet helped display the data by year and a timer added measured the codes efficiency.  My greatest challenge has been the order in which items must be placed ie. Keeping things within the same block. Closing the “For” function at the right place has also been a challenge. I also learned that the data must be static! In order for the loops to work the data must remain in order (i.e. tickers in ascending order) or the macros will fail.

## Results
The results were displayed per individual ticker in the "All Stocks Sheet" as tickers, total daily volume, and return. The use of input boxes provide a text field where the year can be entered and the data for the year is displayed. Added formatting provides color coding of stock percentages based on their value. This helps to clearly identify which stocks fared best. Once Refactoring was completed a timer was added to measure the speed in which the code ran. Adding the timer shows that refactoring increases the speed of computation for both data sets by .29. 

# Summary 
In summary the analysis conducted provided easy to use information with visual aids that clearly show a stocks percentage of return. The information will assist Steve and his family in making an informed decision prior to investing. The code can be repurposed and used for any upcoming years with minor modification. 

**Advantages of refactoring code**    
Refactoring a code makes a code faster. It helps reduce the chances of creating mistakes by adding loops and eliminating duplicate code. It reduces the size making it overall more efficient and easier to run.   
**Disadvantages of refactoring code**   
The only disadvantage that I can see in refactoring is the time commitment it requires. I am sure that as we continue to learn it will become easier but it was time consuming.
-The advantages and disadvantages come into play when refactoring the original VBA script



[Refference]

Hayes, A. (2021) *Investopedia* https://www.investopedia.com/ask/answers/12/what-is-a-stock-ticker.asp

