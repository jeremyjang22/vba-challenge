# ucb-bootcamp-vba-homework
VBA homework to submit for UCB bootcamp

ASSIGNMENT INFO:

__Background__
You are well on your way to becoming a programmer and Excel master! In this homework assignment you will use VBA scripting to analyze real stock market data. Depending on your comfort level with VBA, you may choose to challenge yourself with a few of the challenge tasks.

__Before You Begin__ 
  -Create a new repository for this project called VBA-challenge. Do not add this homework to an existing repository.

  -Inside the new repository that you just created, add any VBA files you use for this assignment. These will be the main scripts to run for each analysis.

__Instructions__

  -Create a script that will loop through all the stocks for one year and output the following information.
      -The ticker symbol.
      -Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
      -The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
      -The total stock volume of the stock.
  -You should also have conditional formatting that will highlight positive change in green and negative change in red.


__CHALLENGES__

Your solution will also be able to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".

Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.


__Other Considerations__


Use the sheet alphabetical_testing.xlsx while developing your code. This data set is smaller and will allow you to test faster. Your code should run on this file in less than 3-5 minutes.


Make sure that the script acts the same on each sheet. The joy of VBA is to take the tediousness out of repetitive task and run over and over again with a click of the button.


__APPROACH__
Given this problem, I began by iterating through all of the sheets in the worksheet and each row in each sheet containing data in Column A. This guaranteed that I would be able to read in all the values in each sheet.

Assuming every sheet had some data in it, I grabbed the first ticker open value, which was in cell C2, and under the assumption I made, all of the open values would be there. Since I also needed to calculate the total Volume, I would also need to grab the volume at that first cell.

As I iterate through each row in the sheet, I grab the current name and volume. Additionally, to check for when I finish checking a stock, I also get the next ticker name one row below the current one. If the ticker names are the same, I add the volume value to what I currently have for the volume (by using total_volume += current_volume) and continue going down each row. If they aren't the same, then I know that I have finished going through that stock. I grab the closing value at the current row, and subtract it with the starting value to calculate the total difference and the percent difference. In the off-case that I would divide by zero for the percent change, I was told by the professor that I could just set it to zero. After crunching those number out, I take the name, the difference, the percent change, and the total volume and put it into a table on the side. Then I update the starting value to be the value at whichever row, but in Column C since that is where the <open> value begins for the newest stock. I also reset the total_volume to 0 and increase the index of the seperate table by one so I don't constantly override old results with new results.
  
Additionally, there were formatting blocks to do the conditional formatting for when total difference is negative, positive or no difference, in which I used red, green, and blue, respectively. For the percent change, I used Cells.NumberFormat(tickerStart, 11) = "0.00%" for two decimal places and formatting for percentage.

As for the BONUS, which required us to find the stock with the greatest percent increase, percent decrease, and total volume, I decided to iterate through the condensed table that the first loop created for me. And as I iterate through that table, I would keep track of the ___indexes___ of where each target value was. I used if statements to compare the current value to the current largest percent increase/percent decrease/total volume, and if they were larger than the current record, then I would update the indexes. And once I went through that table, I would use the indexes at the end and write them another part of the sheet where the user could know what these values were, and at where in the file they occurred.
