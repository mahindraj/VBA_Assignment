# VBA_Assignment
This challenge is to practice VBA scripting to analyze stock market date over three years. 
The goals to achieve were: 
1. the ticker symbol
2. The Yearly change and Percentage change from closing price at end of that year to the opening price at the beginning of that year
3. Total Sotck Volume of the Stock
4. Generate functionality to the script to return the stock with both the Greatest % increase, Greatest % decrease along with the greteast total volume

To conduct this assignment and generate code my script, I relied on various sources including Google, Stack Overflow, as well as refering to the different exerices from Module 2. I was able to create the script by first define the "What" and then break it down step by step in order to write my script:

1. Define WS to apply one code to all three worksheets
2. Declare the variables
3. Create Column Headers
5. Set the Ticker Counter to first row since tickers are located below Column A; set start row to 2
6. Loop through all rows; while looping it has to check if tick name changed and then ticker will be generated along summary table
7. Calculate both Yearly and Percentage Change
8. Conditional Formatting(ColorIndex=3 for red for negative values and ColorIndex=4 for green for positive values
9. Calculate and write total volume
10. Preperare summary for GreatVolume, Greatest % Increase, and Greatest % Decrease
11. Loop for summary
12. For greatest increase - must check if next value is larder - if yes, take over a new value and populate
13. For greatest decrease - must check if next value is smaler - if yes, take over a new value and populate
14. Write summary results
15. Repeat the whole code for the other worksheets
