# Challenge 2 / Stock Analysis 

## Overview of Project
This Stock analysis is to determine which of the stocks being monitored by Steve and his family are more likely to be a good investment. A Macros (VBA scipts) was created to analyze the data. A macros  is a set of rules included in the background of an excel sheet that preforms tasks (in this case analyzing stock data) included loops, arrays and commands that  process the information and provided an outcome. The original Macros completed an analysis that provided an overall daily value and return. Refactoring of the original Macros was completed to individualize the report and provided a daily volume and percentage return per year per ticker .  

## Data 
Determining which stocks were the better investment would be based on ticker data recorded for the year 2017 and 2018. A stock ticker is a series of letters assigned to a specific stock as an identifier. The stock ticker reports transactions and price data for the specific stock throughout the day (Hayes) the data reported includes the date, the opening price, the high and low for said date, the closing price, the adjusted closing price (adjusts for dividends stocks splits etc.) and the Volume.

## Analysis and Challenges
To analyze the data a Macros was created following the steps provided. The Macros refactoring included adding an index of tickers and creating loops. Formating to include color coding was also added to the data sheet. The display sheet was automated with the addition of input boxes. The addition of loops in refactoring is done to collect the data for the specific ticker for both years and calculates the total daily value and the return percentage. Input Box added to the sheet help display the data by the year entyered by the end user. A timer added by inserting a "startTime=Timer" to the beggininghg of teh sheet and an "endTime" to the end measures the codes efficiency.  My greatest challenge has been the order in which items must be placed  an example of this is ending the "For" function and keeping things within the same block for. I also learned that the data must be static! failure to keep the values set will result in the macros failing.

## Results
The results were displayed per individual ticker in the "All Stocks Sheet" as tickers, total daily volume, and return. The use of input boxes provide a text field where the year can be entered and the data for the year is displayed. Added formatting provides color coding of stock percentages based on their value. The color formating helps to clearly identify which stocks fared best. Once Refactoring was completed the timer added to measured the speed in which the code runs. The timer  shows that refactoring increases the speed of computation for both data sets by .29. 

![image](https://user-images.githubusercontent.com/104601282/175348843-8595f0bd-1584-4169-9f31-f390c217014e.png)



# Summary 
In summary the analysis conducted provided easy to use information with visual aids that clearly show a stocks percentage of return. The information will assist Steve and his family in making an informed decision prior to investing. The code can be repurposed and used for any upcoming years with minor modification. 

**Advantages of refactoring code**    
Refactoring a code makes a code faster. It helps reduce the chances of creating mistakes by adding loops and eliminating duplicate code. It reduces the size making it overall more efficient and easier to run.   
**Disadvantages of refactoring code**   
The only disadvantage that I can see in refactoring is the time commitment it requires. I am sure that as we continue to learn it will become easier but it was time consuming.




[Refference]

Hayes, A. (2021) *Investopedia* https://www.investopedia.com/ask/answers/12/what-is-a-stock-ticker.asp

