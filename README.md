# VBA-challenge

Resouces:
    https://stackoverflow.com/questions/20648149/what-are-numberformat-options-in-excel-vba


Create a script that loops through all the stocks for each quarter and outputs the following information:
    Loops through all the worksheet with 'for' function and variable 'i'
         1. The ticker symbol
        -Set up the summary table to print out the values for questions 1, 2, 3, and 4
        - Print the current ticker until its value changes.
    
         2. Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
        - Determine openingPrice by determine if the tickers are same or the previousDate is different currentdate (same quarter period).
        - ClosingPrice will be the last value of the current quarter.
        - Quarterly change = closingPrice - openingPrice
        - Format the quarterChange with "Interior.ColorIndex" - if it's positive, prints out with green, and if it's negative, prints out with red.

        3. The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
        - percentChange = quarterChange/ openingPrice (if the openingPrice is different than 0)
        - Otherwise, percentChange = 0
        - Format the output with ("0.00%")

         4. The total stock volume of the stock.
         - Accumulating the total volume for same ticker until it changes to different ticker.
